#Requires -Version 5.1
<#
.SYNOPSIS
    Zet gesynchroniseerde SharePoint bibliotheken om naar OneDrive snelkoppelingen ("Snelkoppeling toevoegen aan Mijn bestanden").

.DESCRIPTION
    WPF tool die per geselecteerde bibliotheek:
      1. De bibliotheek opzoekt via Microsoft Graph (drive en item ID).
      2. Alle bestanden op "alleen online" zet (Ruimte vrijmaken) en wacht tot alles veilig in de cloud staat.
      3. OneDrive afsluit, de sync configuratie van die bibliotheek verwijdert (met back-up) en de lokale map opruimt.
      4. OneDrive herstart en een snelkoppeling in Mijn bestanden aanmaakt.

    Veiligheidscriterium (per bestand, via de Cloud Files API):
      - alleen online (gedehydrateerd)                      -> veilig
      - lokaal aanwezig maar placeholder IN SYNC met cloud  -> veilig (inhoud staat identiek online)
      - geen placeholder (nooit geupload / uitgesloten)     -> NIET veilig
      - placeholder die niet in sync is (upload loopt nog)  -> NIET veilig
      - leesfout (geen toegang, pad onleesbaar)             -> NIET veilig
    Een bibliotheek wordt alleen ontkoppeld als er nul onveilige bestanden zijn.

    Scannen en vrijmaken draait op een achtergrondthread (met lange pad ondersteuning), dus het venster blijft
    reageren en de actie kan altijd worden geannuleerd.

    Draaien als de eindgebruiker zelf (HKCU), niet als een andere admin.

.NOTES
    Microsoft heeft geen ondersteunde API om een enkele bibliotheek te ontkoppelen. Deze tool past de OneDrive
    configuratie aan met OneDrive gestopt. Back-ups staan in %LOCALAPPDATA%\EOO\SPSyncNaarSnelkoppeling\Backup.
#>
[CmdletBinding()]
param()

#region STA afdwingen (WPF vereist Single Threaded Apartment)
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`""
    return
}
#endregion

#region Basisinstellingen en logging
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$script:AppRoot   = Join-Path $env:LOCALAPPDATA 'EOO\SPSyncNaarSnelkoppeling'
$script:LogFile   = Join-Path $script:AppRoot ("Log_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))
$script:BackupDir = Join-Path $script:AppRoot ("Backup\{0:yyyyMMdd_HHmmss}" -f (Get-Date))
$script:txtLog    = $null
$script:Window    = $null
$script:Cts       = $null     # CancellationTokenSource van de lopende actie
$script:Busy      = $false
$script:Critical  = $false    # $true zolang OneDrive gestopt is en de config wordt aangepast
$script:WeStoppedOneDrive = $false
New-Item -Path $script:AppRoot -ItemType Directory -Force | Out-Null

# Bestanden die OneDrive zelf niet synchroniseert en dus nooit een placeholder worden (bewust smal gehouden)
$script:IgnorePattern = '^(desktop\.ini|thumbs\.db|ehthumbs\.db|\.ds_store)$|^~\$|^\.~|^~.*\.tmp$|^\.[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'

# Instellingen voor het wachten
$script:PassIntervalSec   = 15   # pauze tussen twee scans
$script:DehydrateGraceSec = 120  # extra tijd om OneDrive lokale (wel gesynchroniseerde) bestanden te laten vrijmaken
$script:NoProgressHintMin = 3    # na zoveel minuten zonder voortgang een hint loggen

function Update-UI {
    # Laat de WPF dispatcher de wachtrij verwerken zodat het venster niet bevriest
    if ($script:Window) {
        try { $script:Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background) } catch { }
    }
}

function Test-Cancelled {
    return ($script:Cts -and $script:Cts.IsCancellationRequested)
}

function Wait-UI {
    # Wacht zonder het venster te blokkeren. Geeft $false terug als er geannuleerd is.
    param([double]$Seconds, [switch]$IgnoreCancel)
    $end = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $end) {
        if (-not $IgnoreCancel -and (Test-Cancelled)) { return $false }
        Start-Sleep -Milliseconds 100
        Update-UI
    }
    return $true
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'OK', 'Warn', 'Error')][string]$Level = 'Info'
    )
    $line  = '{0:HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level.ToUpper(), $Message
    $color = @{ Info = 'Gray'; OK = 'Green'; Warn = 'Yellow'; Error = 'Red' }[$Level]
    Write-Host $line -ForegroundColor $color
    Add-Content -LiteralPath $script:LogFile -Value $line -ErrorAction SilentlyContinue
    if ($script:txtLog) {
        $script:txtLog.AppendText("$line`r`n")
        $script:txtLog.ScrollToEnd()
        Update-UI
    }
}
#endregion

#region Datamodel en Cloud Files API (C#)
if (-not ('EooSyncLib' -as [type])) {
    Add-Type -TypeDefinition @'
using System.ComponentModel;
public class EooSyncLib : INotifyPropertyChanged {
    private bool _selected;
    private string _status;
    public bool Selected { get { return _selected; } set { _selected = value; Notify("Selected"); } }
    public string Status { get { return _status; } set { _status = value; Notify("Status"); } }
    public string Name { get; set; }
    public string MountPoint { get; set; }
    public string Url { get; set; }
    public string LibraryType { get; set; }
    public string ProviderKey { get; set; }
    public string DriveId { get; set; }
    public string ItemId { get; set; }
    public bool ShortcutExists { get; set; }
    public bool Ready { get; set; }
    public bool Unlinked { get; set; }
    public string RenamedPath { get; set; }
    public event PropertyChangedEventHandler PropertyChanged;
    private void Notify(string n) { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(n)); }
}
'@
}

if (-not ('EooCloudFiles' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;

public class EooScanProgress {
    public volatile int Files;
    public volatile int Dirs;
}

public class EooScanResult {
    public string Root;
    public long Total, Dehydrated, InSyncLocal, Ignored, Unpinned, UnpinFailed;
    public long NotSyncedCount, PendingCount, ErrorCount;
    public List<string> NotSynced = new List<string>();
    public List<string> Pending = new List<string>();
    public List<string> Errors = new List<string>();
    public bool Cancelled;
    public string Fatal;
    public long UnsafeCount { get { return NotSyncedCount + PendingCount + ErrorCount; } }
    internal static void Add(List<string> list, string item) { if (list.Count < 50) list.Add(item); }
}

public static class EooCloudFiles {
    const uint ATTR_DIRECTORY     = 0x10;
    const uint ATTR_REPARSE       = 0x400;
    const uint ATTR_OFFLINE       = 0x1000;
    const uint ATTR_PINNED        = 0x80000;
    const uint ATTR_UNPINNED      = 0x100000;
    const uint ATTR_RECALL_ON_DATA = 0x400000;
    const uint TAG_MOUNT_POINT    = 0xA0000003;
    const uint TAG_SYMLINK        = 0xA000000C;
    const int  HR_NOT_A_CLOUD_FILE = unchecked((int)0x80070178);
    const int  HR_MORE_DATA        = unchecked((int)0x800700EA);
    const int  CF_PIN_UNPINNED     = 2;
    const int  CF_PIN_EXCLUDED     = 3;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct WIN32_FIND_DATA {
        public uint dwFileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string cAlternateFileName;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr FindFirstFileExW(string lpFileName, int fInfoLevelId, out WIN32_FIND_DATA lpFindFileData, int fSearchOp, IntPtr lpSearchFilter, int dwAdditionalFlags);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool FindNextFileW(IntPtr hFindFile, out WIN32_FIND_DATA lpFindFileData);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool FindClose(IntPtr hFindFile);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern SafeFileHandle CreateFileW(string lpFileName, uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes, uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);
    [DllImport("cldapi.dll")]
    static extern int CfGetPlaceholderInfo(SafeFileHandle FileHandle, int InfoClass, byte[] InfoBuffer, uint InfoBufferLength, out uint ReturnedLength);
    [DllImport("cldapi.dll")]
    static extern int CfSetPinState(SafeFileHandle FileHandle, int PinState, int PinFlags, IntPtr Overlapped);

    static readonly IntPtr INVALID = new IntPtr(-1);

    public static string ToLongPath(string path) {
        if (path.StartsWith(@"\\?\")) return path;
        if (path.StartsWith(@"\\")) return @"\\?\UNC\" + path.Substring(2);
        return @"\\?\" + path;
    }

    static string Display(string longPath) {
        if (longPath.StartsWith(@"\\?\UNC\")) return @"\\" + longPath.Substring(8);
        if (longPath.StartsWith(@"\\?\")) return longPath.Substring(4);
        return longPath;
    }

    static SafeFileHandle OpenAttr(string longPath) {
        // FILE_READ_ATTRIBUTES, alle share modes, OPEN_EXISTING, BACKUP_SEMANTICS (ook mappen). Hydrateert niets.
        return CreateFileW(longPath, 0x80, 0x7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero);
    }

    static bool TryUnpin(SafeFileHandle h) {
        try { return CfSetPinState(h, CF_PIN_UNPINNED, 0, IntPtr.Zero) >= 0; } catch { return false; }
    }

    public static Task<EooScanResult> StartScan(string root, bool unpin, string ignorePattern, EooScanProgress progress, CancellationToken token) {
        return Task.Factory.StartNew(() => Scan(root, unpin, ignorePattern, progress, token),
            token, TaskCreationOptions.LongRunning, TaskScheduler.Default);
    }

    public static EooScanResult Scan(string root, bool unpin, string ignorePattern, EooScanProgress progress, CancellationToken token) {
        var r = new EooScanResult();
        r.Root = root;
        try {
            Regex ignore = string.IsNullOrEmpty(ignorePattern) ? null : new Regex(ignorePattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            byte[] buf = new byte[1024];
            var stack = new Stack<string>();
            string top = ToLongPath(root.TrimEnd('\\'));
            stack.Push(top);

            while (stack.Count > 0) {
                if (token.IsCancellationRequested) { r.Cancelled = true; return r; }
                string dir = stack.Pop();
                if (progress != null) progress.Dirs++;

                // Map zelf ook op 'niet vastgezet' zetten, zodat nieuwe bestanden ook online-only worden
                if (unpin) {
                    using (var hd = OpenAttr(dir)) { if (!hd.IsInvalid) TryUnpin(hd); }
                }

                WIN32_FIND_DATA fd;
                // FindExInfoBasic = 1, FindExSearchNameMatch = 0, FIND_FIRST_EX_LARGE_FETCH = 2
                IntPtr hFind = FindFirstFileExW(dir + @"\*", 1, out fd, 0, IntPtr.Zero, 2);
                if (hFind == INVALID) {
                    int err = Marshal.GetLastWin32Error();
                    if (err != 2 && err != 18) { // FILE_NOT_FOUND / NO_MORE_FILES = lege map
                        r.ErrorCount++;
                        EooScanResult.Add(r.Errors, Display(dir) + "  (map niet leesbaar, fout " + err + ")");
                    }
                    continue;
                }
                try {
                    do {
                        if (token.IsCancellationRequested) { r.Cancelled = true; return r; }
                        string name = fd.cFileName;
                        if (name == "." || name == "..") continue;
                        string full = dir + "\\" + name;
                        uint a = fd.dwFileAttributes;

                        if ((a & ATTR_DIRECTORY) != 0) {
                            // Geen junctions/symlinks volgen, cloud placeholder mappen wel
                            if ((a & ATTR_REPARSE) != 0 && (fd.dwReserved0 == TAG_MOUNT_POINT || fd.dwReserved0 == TAG_SYMLINK)) continue;
                            stack.Push(full);
                            continue;
                        }

                        if (progress != null) progress.Files++;
                        if (ignore != null && ignore.IsMatch(name)) { r.Ignored++; continue; }
                        r.Total++;

                        bool dehydrated = (a & (ATTR_RECALL_ON_DATA | ATTR_OFFLINE)) != 0;
                        if (dehydrated) {
                            r.Dehydrated++;
                            // Een vastgezet online-only bestand wordt anders weer gedownload
                            if (unpin && ((a & ATTR_PINNED) != 0 || (a & ATTR_UNPINNED) == 0)) {
                                using (var h = OpenAttr(full)) {
                                    if (!h.IsInvalid && TryUnpin(h)) r.Unpinned++; else r.UnpinFailed++;
                                }
                            }
                            continue;
                        }

                        // Lokaal aanwezig: placeholder status opvragen (alleen READ_ATTRIBUTES nodig)
                        using (var h = OpenAttr(full)) {
                            if (h.IsInvalid) {
                                int err = Marshal.GetLastWin32Error();
                                r.ErrorCount++;
                                EooScanResult.Add(r.Errors, Display(full) + "  (niet te openen, fout " + err + ")");
                                continue;
                            }
                            uint ret;
                            int hr = CfGetPlaceholderInfo(h, 0, buf, (uint)buf.Length, out ret);
                            if (hr == HR_NOT_A_CLOUD_FILE) {
                                r.NotSyncedCount++;
                                EooScanResult.Add(r.NotSynced, Display(full));
                                continue;
                            }
                            if (hr < 0 && hr != HR_MORE_DATA) {
                                r.ErrorCount++;
                                EooScanResult.Add(r.Errors, Display(full) + "  (status onbekend, 0x" + hr.ToString("X8") + ")");
                                continue;
                            }
                            int pin    = BitConverter.ToInt32(buf, 0);
                            int inSync = BitConverter.ToInt32(buf, 4);
                            if (pin == CF_PIN_EXCLUDED) {
                                r.NotSyncedCount++;
                                EooScanResult.Add(r.NotSynced, Display(full) + "  (uitgesloten van sync)");
                                continue;
                            }
                            if (inSync != 1) {
                                r.PendingCount++;
                                EooScanResult.Add(r.Pending, Display(full));
                                continue;
                            }
                            r.InSyncLocal++;
                            if (unpin && pin != CF_PIN_UNPINNED) {
                                if (TryUnpin(h)) r.Unpinned++; else r.UnpinFailed++;
                            }
                        }
                    } while (FindNextFileW(hFind, out fd));
                }
                finally { FindClose(hFind); }
            }
        }
        catch (Exception ex) {
            r.Fatal = ex.GetType().Name + ": " + ex.Message;
        }
        return r;
    }
}
'@
}
$script:Libraries = New-Object 'System.Collections.ObjectModel.ObservableCollection[EooSyncLib]'
#endregion

#region OneDrive helpers
function Get-OneDriveBusinessAccounts {
    # Alle zakelijke OneDrive accounts (Business1, Business2, ...)
    $root = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (-not (Test-Path $root)) { return @() }
    Get-ChildItem $root -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like 'Business*' } | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{
            Name       = $_.PSChildName
            RegPath    = $_.PSPath
            UserEmail  = $p.UserEmail
            UserFolder = $p.UserFolder
        }
    }
}

function Get-NormalizedPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return $Path.Trim().TrimEnd('\')
}

function Test-SamePath {
    param([string]$A, [string]$B)
    return [string]::Equals((Get-NormalizedPath $A), (Get-NormalizedPath $B), [StringComparison]::OrdinalIgnoreCase)
}

function Get-SyncedLibraries {
    # Gesynchroniseerde SharePoint bibliotheken uit de SyncEngines registratie
    $root = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
    if (-not (Test-Path $root)) { return @() }

    # Mounts binnen de eigen OneDrive map zijn al snelkoppelingen, die slaan we over
    $ownFolders = @(Get-OneDriveBusinessAccounts | ForEach-Object { Get-NormalizedPath $_.UserFolder } | Where-Object { $_ })

    foreach ($key in Get-ChildItem $root -ErrorAction SilentlyContinue) {
        $p     = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
        $mount = Get-NormalizedPath $p.MountPoint
        if (-not $mount -or -not $p.UrlNamespace) { continue }
        if ($p.LibraryType -in 'mysite', 'personal') { continue }
        if ($ownFolders | Where-Object { $mount.StartsWith("$_\", [StringComparison]::OrdinalIgnoreCase) }) { continue }

        $lib             = New-Object EooSyncLib
        $lib.Name        = Split-Path $mount -Leaf
        $lib.MountPoint  = $mount
        $lib.Url         = $p.UrlNamespace
        $lib.LibraryType = $p.LibraryType
        $lib.ProviderKey = $key.PSPath
        $lib.Status      = if (Test-Path -LiteralPath $mount) { 'Gesynchroniseerd' } else { 'Lokale map ontbreekt' }
        $lib
    }
}

function Get-OneDriveExe {
    $candidates = @(
        (Get-Process OneDrive -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Path),
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
        "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe",
        "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe"
    )
    $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
}

function Get-OneDriveProcess {
    # Alleen OneDrive processen van de huidige sessie
    Get-Process OneDrive -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq (Get-Process -Id $PID).SessionId }
}

function Start-OneDrive {
    if (Get-OneDriveProcess) { return $true }
    $exe = Get-OneDriveExe
    if (-not $exe) { Write-Log 'OneDrive.exe niet gevonden.' 'Error'; return $false }
    for ($try = 1; $try -le 3; $try++) {
        Write-Log "OneDrive starten (poging $try)..."
        try { Start-Process -FilePath $exe -ArgumentList '/background' -ErrorAction Stop } catch { Write-Log "  Starten mislukt: $($_.Exception.Message)" 'Warn' }
        for ($i = 0; $i -lt 20; $i++) {
            [void](Wait-UI 1 -IgnoreCancel)
            if (Get-OneDriveProcess) { $script:WeStoppedOneDrive = $false; [void](Wait-UI 5 -IgnoreCancel); return $true }
        }
    }
    return [bool](Get-OneDriveProcess)
}

function Stop-OneDrive {
    if (-not (Get-OneDriveProcess)) { return $true }
    $exe = Get-OneDriveExe
    Write-Log 'OneDrive netjes afsluiten...'
    if ($exe) { try { Start-Process -FilePath $exe -ArgumentList '/shutdown' -WindowStyle Hidden -ErrorAction Stop } catch { } }

    # Max 45 seconden wachten, daarna hard stoppen
    for ($i = 0; $i -lt 45 -and (Get-OneDriveProcess); $i++) { [void](Wait-UI 1 -IgnoreCancel) }
    if (Get-OneDriveProcess) {
        Write-Log 'OneDrive reageert niet, proces wordt geforceerd gestopt.' 'Warn'
        Get-OneDriveProcess | Stop-Process -Force -ErrorAction SilentlyContinue
        for ($i = 0; $i -lt 10 -and (Get-OneDriveProcess); $i++) { [void](Wait-UI 1 -IgnoreCancel) }
    }
    $script:WeStoppedOneDrive = $true
    # Even extra tijd zodat OneDrive zijn bestanden echt heeft losgelaten
    [void](Wait-UI 3 -IgnoreCancel)
    return -not (Get-OneDriveProcess)
}

function Test-AutoMountPolicy {
    # Intune/GPO "Bibliotheken automatisch synchroniseren" zet de sync gewoon terug
    foreach ($hive in 'HKCU:', 'HKLM:') {
        $key = "$hive\SOFTWARE\Policies\Microsoft\OneDrive\TenantAutoMount"
        if (Test-Path $key) {
            $values = (Get-Item $key).Property
            if ($values) { return $true }
        }
    }
    return $false
}
#endregion

#region Online controle (Cloud Files API op achtergrondthread)
function Invoke-CloudScan {
    # Scant een bibliotheek zonder de UI te blokkeren. Optioneel alles op 'niet vastgezet' (Ruimte vrijmaken).
    param([EooSyncLib]$Lib, [bool]$Unpin, [string]$Activity = 'Scannen')
    $progress = New-Object EooScanProgress
    $token    = if ($script:Cts) { $script:Cts.Token } else { [System.Threading.CancellationToken]::None }
    $task     = [EooCloudFiles]::StartScan($Lib.MountPoint, $Unpin, $script:IgnorePattern, $progress, $token)
    $tick     = 0
    while (-not $task.IsCompleted) {
        Start-Sleep -Milliseconds 100
        Update-UI
        if ((++$tick % 5) -eq 0) { $Lib.Status = "$Activity... $($progress.Files) bestanden" }
    }
    if ($task.IsFaulted) {
        $r = New-Object EooScanResult
        $r.Root  = $Lib.MountPoint
        $r.Fatal = $task.Exception.InnerException.Message
        return $r
    }
    if ($task.IsCanceled) {
        $r = New-Object EooScanResult
        $r.Cancelled = $true
        return $r
    }
    return $task.Result
}

function Write-ScanDetails {
    param([EooSyncLib]$Lib, $Scan)
    $groups = @(
        @{ Label = 'niet gesynchroniseerd (geen cloud bestand / uitgesloten)'; Count = $Scan.NotSyncedCount; Items = $Scan.NotSynced },
        @{ Label = 'wacht op upload (niet in sync)';                         Count = $Scan.PendingCount;   Items = $Scan.Pending },
        @{ Label = 'leesfout';                                                Count = $Scan.ErrorCount;     Items = $Scan.Errors }
    )
    foreach ($g in $groups) {
        if ($g.Count -le 0) { continue }
        Write-Log "[$($Lib.Name)] $($g.Count) bestand(en) $($g.Label):" 'Warn'
        $g.Items | Select-Object -First 10 | ForEach-Object { Write-Log "    $_" 'Warn' }
        if ($g.Count -gt 10) { Write-Log "    ... en nog $($g.Count - 10) (eerste 50 staan in het logbestand)" 'Warn' }
        $g.Items | Select-Object -Skip 10 | ForEach-Object { Add-Content -LiteralPath $script:LogFile -Value "    $_" -ErrorAction SilentlyContinue }
    }
}

function Wait-AllSafe {
    # Wacht tot alle kandidaten nul onveilige bestanden hebben, of tot de deadline / annulering.
    param([EooSyncLib[]]$Libs, [int]$TimeoutMinutes, [bool]$FreeUpSpace)

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $state    = @{}
    foreach ($l in $Libs) { $state[$l.MountPoint] = @{ Done = $false; Last = $null; LastUnsafe = [int64]-1; LastChange = Get-Date; Hinted = $false; SafeSince = $null } }
    $pass = 0

    while ($true) {
        $pass++
        $open = @($Libs | Where-Object { -not $state[$_.MountPoint].Done })
        if (-not $open) { break }

        foreach ($lib in $open) {
            if (Test-Cancelled) { break }
            $st = $state[$lib.MountPoint]

            if (-not (Test-Path -LiteralPath $lib.MountPoint)) {
                $lib.Status = 'Fout: lokale map ontbreekt'
                Write-Log "[$($lib.Name)] lokale map is verdwenen tijdens controle." 'Error'
                $st.Done = $true; continue
            }

            # Eerste ronde: Ruimte vrijmaken. Volgende rondes opnieuw, zodat net geuploade bestanden ook vrijkomen.
            $activity = if ($FreeUpSpace) { 'Ruimte vrijmaken en controleren' } else { 'Controleren' }
            $scan = Invoke-CloudScan -Lib $lib -Unpin $FreeUpSpace -Activity $activity
            $st.Last = $scan
            if ($scan.Cancelled) { break }

            if ($scan.Fatal) {
                Write-Log "[$($lib.Name)] scan mislukt: $($scan.Fatal)" 'Error'
                $lib.Status = 'Fout: scan mislukt'
                $st.Done = $true; continue
            }
            if ($pass -eq 1) {
                Write-Log ("[{0}] {1} bestanden: {2} alleen online, {3} lokaal en in sync, {4} onveilig." -f $lib.Name, $scan.Total, $scan.Dehydrated, $scan.InSyncLocal, $scan.UnsafeCount)
                if ($scan.UnpinFailed -gt 0) { Write-Log "[$($lib.Name)] $($scan.UnpinFailed) bestand(en) konden niet op 'alleen online' gezet worden (in gebruik?). Niet erg zolang ze in sync zijn." 'Warn' }
            }

            if ($scan.UnsafeCount -ne $st.LastUnsafe) { $st.LastUnsafe = $scan.UnsafeCount; $st.LastChange = Get-Date; $st.Hinted = $false }

            if ($scan.UnsafeCount -eq 0) {
                # Alles veilig. Geef OneDrive nog even de tijd om lokale kopieen vrij te maken, maar niet verplicht.
                if ($FreeUpSpace -and $scan.InSyncLocal -gt 0) {
                    if (-not $st.SafeSince) { $st.SafeSince = Get-Date }
                    if (((Get-Date) - $st.SafeSince).TotalSeconds -lt $script:DehydrateGraceSec -and (Get-Date) -lt $deadline) {
                        $lib.Status = "Veilig, OneDrive maakt nog $($scan.InSyncLocal) lokale kopie(en) vrij..."
                        continue
                    }
                }
                $lib.Ready  = $true
                $lib.Status = if ($scan.InSyncLocal -gt 0) { "Klaar: alles in cloud ($($scan.InSyncLocal) nog lokaal, wel in sync)" } else { 'Klaar: alles online' }
                Write-Log "[$($lib.Name)] alles staat veilig in de cloud ($($scan.Dehydrated) online-only, $($scan.InSyncLocal) lokaal in sync)." 'OK'
                $st.Done = $true
                continue
            }

            $st.SafeSince = $null
            $lib.Status = "Wacht: $($scan.PendingCount) upload, $($scan.NotSyncedCount) niet gesync., $($scan.ErrorCount) fout"
            if (-not $st.Hinted -and ((Get-Date) - $st.LastChange).TotalMinutes -ge $script:NoProgressHintMin) {
                Write-Log "[$($lib.Name)] al $($script:NoProgressHintMin) min geen voortgang. Controleer of OneDrive niet gepauzeerd is en of er sync fouten zijn (OneDrive icoon)." 'Warn'
                Write-ScanDetails -Lib $lib -Scan $scan
                $st.Hinted = $true
            }
        }

        if (Test-Cancelled) { break }
        if (-not @($Libs | Where-Object { -not $state[$_.MountPoint].Done })) { break }
        if ((Get-Date) -ge $deadline) { break }

        # OneDrive moet blijven draaien om te kunnen uploaden en vrijmaken
        if (-not (Get-OneDriveProcess)) {
            Write-Log 'OneDrive draait niet meer, opnieuw starten...' 'Warn'
            [void](Start-OneDrive)
        }
        if (-not (Wait-UI $script:PassIntervalSec)) { break }
    }

    # Resultaat voor alles wat niet klaar is
    foreach ($lib in $Libs) {
        $st = $state[$lib.MountPoint]
        if ($st.Done) { continue }
        if (Test-Cancelled) {
            $lib.Status = 'Geannuleerd'
            continue
        }
        $scan = $st.Last
        if ($scan) {
            $lib.Status = "Niet klaar: $($scan.UnsafeCount) onveilig bestand(en)"
            Write-Log "[$($lib.Name)] na $TimeoutMinutes min nog $($scan.UnsafeCount) onveilig bestand(en). Overgeslagen." 'Warn'
            Write-ScanDetails -Lib $lib -Scan $scan
        }
        else { $lib.Status = 'Niet klaar' }
    }
}
#endregion

#region Microsoft Graph helpers
function Get-MgContextSafe {
    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { return $null }
    try { Get-MgContext } catch { $null }
}

function Connect-Graph {
    if (Get-MgContextSafe) { return $true }

    # Alleen de authenticatie module is nodig, de rest gaat via Invoke-MgGraphRequest
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Log 'Module Microsoft.Graph.Authentication installeren (CurrentUser)...'
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
            }
            Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch { Write-Log "Installatie module mislukt: $($_.Exception.Message)" 'Error'; return $false }
    }
    try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
    catch { Write-Log "Module laden mislukt: $($_.Exception.Message)" 'Error'; return $false }

    try {
        Write-Log 'Aanmelden bij Microsoft 365 (Files.ReadWrite.All, Sites.Read.All)...'
        Connect-MgGraph -Scopes 'Files.ReadWrite.All', 'Sites.Read.All' -NoWelcome -ErrorAction Stop
    }
    catch { Write-Log "Aanmelden mislukt: $($_.Exception.Message)" 'Error'; return $false }

    $ctx = Get-MgContextSafe
    if (-not $ctx) { return $false }
    Write-Log "Verbonden als $($ctx.Account)" 'OK'

    # Snelkoppeling komt in de OneDrive van het Graph account, dus dat moet de OneDrive gebruiker zijn
    $odMails = @(Get-OneDriveBusinessAccounts | ForEach-Object { $_.UserEmail } | Where-Object { $_ })
    if ($odMails.Count -and ($odMails -notcontains $ctx.Account)) {
        Write-Log "Let op: OneDrive is gekoppeld aan $($odMails -join ', '), maar je bent aangemeld als $($ctx.Account)." 'Warn'
    }
    return $true
}

function Invoke-Graph {
    # Graph aanroep met retry bij throttling (429) en tijdelijke fouten (5xx / netwerk)
    param([string]$Method = 'GET', [string]$Uri, $Body, [int]$MaxTries = 4)
    $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($Body) {
        $params.Body        = ($Body | ConvertTo-Json -Depth 10)
        $params.ContentType = 'application/json'
    }
    for ($try = 1; ; $try++) {
        try { return Invoke-MgGraphRequest @params }
        catch {
            $msg  = $_.Exception.Message
            $code = 0
            try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            $transient = ($code -eq 429 -or $code -ge 500 -or $msg -match 'timed out|timeout|connection|TooManyRequests|ServiceUnavailable')
            if (-not $transient -or $try -ge $MaxTries) { throw }
            $delay = [math]::Min(30, [math]::Pow(2, $try) * 2)
            Write-Log "  Graph tijdelijk niet beschikbaar ($code), opnieuw over $delay s..." 'Warn'
            [void](Wait-UI $delay -IgnoreCancel)
        }
    }
}

function Get-GraphAll {
    # Volgt @odata.nextLink voor gepagineerde resultaten
    param([string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    while ($Uri) {
        $resp = Invoke-Graph -Uri $Uri
        foreach ($v in @($resp.value)) { if ($null -ne $v) { $items.Add($v) } }
        $Uri  = $resp.'@odata.nextLink'
    }
    $items
}

function Resolve-LibraryTarget {
    # Zoekt drive ID en item ID van de gesynchroniseerde bibliotheek of map
    param([EooSyncLib]$Lib)
    $g      = 'https://graph.microsoft.com/v1.0'
    $libUrl = [Uri]::UnescapeDataString($Lib.Url.Trim().TrimEnd('/'))
    $uri    = [Uri]$libUrl
    $segs   = @($uri.AbsolutePath.Trim('/') -split '/' | Where-Object { $_ })

    # Methode 1: site opzoeken en drive matchen op webUrl
    try {
        if ($segs.Count -ge 2 -and $segs[0] -in 'sites', 'teams') {
            $sitePath = ($segs[0..1] | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
            $site = Invoke-Graph -Uri "$g/sites/$($uri.Host):/$sitePath"
        }
        else {
            $site = Invoke-Graph -Uri "$g/sites/$($uri.Host)"
        }
        $drives = Get-GraphAll -Uri "$g/sites/$($site.id)/drives?`$select=id,name,webUrl"
        $drive = $drives | Where-Object {
            $d = [Uri]::UnescapeDataString($_.webUrl.TrimEnd('/'))
            $libUrl -eq $d -or $libUrl.StartsWith("$d/", [StringComparison]::OrdinalIgnoreCase)
        } | Sort-Object { $_.webUrl.Length } -Descending | Select-Object -First 1

        if ($drive) {
            $driveUrl = [Uri]::UnescapeDataString($drive.webUrl.TrimEnd('/'))
            $subPath  = $libUrl.Substring($driveUrl.Length).Trim('/')
            if ($subPath) {
                $enc  = ($subPath -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
                $item = Invoke-Graph -Uri "$g/drives/$($drive.id)/root:/$enc`?`$select=id,name"
            }
            else {
                $item = Invoke-Graph -Uri "$g/drives/$($drive.id)/root?`$select=id,name"
            }
            return [pscustomobject]@{ DriveId = $drive.id; ItemId = $item.id }
        }
    }
    catch { Write-Log "Site/drive zoeken mislukt voor $($Lib.Name): $($_.Exception.Message)" 'Warn' }

    # Methode 2: shares API met de URL (vangt subsites en afwijkende paden af)
    try {
        $b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Lib.Url.Trim().TrimEnd('/'))).TrimEnd('=').Replace('/', '_').Replace('+', '-')
        $item = Invoke-Graph -Uri "$g/shares/u!$b64/driveItem?`$select=id,name,parentReference"
        return [pscustomobject]@{ DriveId = $item.parentReference.driveId; ItemId = $item.id }
    }
    catch { Write-Log "Shares API mislukt voor $($Lib.Name): $($_.Exception.Message)" 'Warn' }

    return $null
}

function Get-ExistingShortcutIds {
    $children = Get-GraphAll -Uri "https://graph.microsoft.com/v1.0/me/drive/root/children?`$select=id,name,remoteItem&`$top=200"
    @($children | Where-Object { $_.remoteItem } | ForEach-Object { $_.remoteItem.id })
}

function New-OneDriveShortcut {
    param([EooSyncLib]$Lib)
    $body = @{
        name                                = $Lib.Name
        remoteItem                          = @{ id = $Lib.ItemId; parentReference = @{ driveId = $Lib.DriveId } }
        '@microsoft.graph.conflictBehavior' = 'rename'
    }
    Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/me/drive/root/children' -Body $body
}
#endregion

#region Sync configuratie verwijderen
function Backup-OneDriveConfig {
    # Geeft $true terug als de back-up compleet is. Zonder back-up wordt er niets aangepast.
    try {
        New-Item -Path $script:BackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        $ok = $true
        $exports = @(
            @{ Key = 'HKCU\Software\Microsoft\OneDrive\Accounts';       File = 'OneDrive_Accounts.reg'; Ps = 'HKCU:\Software\Microsoft\OneDrive\Accounts' },
            @{ Key = 'HKCU\Software\SyncEngines\Providers\OneDrive';    File = 'SyncEngines.reg';       Ps = 'HKCU:\Software\SyncEngines\Providers\OneDrive' }
        )
        foreach ($e in $exports) {
            if (-not (Test-Path $e.Ps)) { continue }
            & reg.exe export $e.Key (Join-Path $script:BackupDir $e.File) /y 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) { Write-Log "  Registry export mislukt: $($e.Key)" 'Error'; $ok = $false }
        }
        foreach ($acc in Get-OneDriveBusinessAccounts) {
            $src = Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\settings\$($acc.Name)"
            if (Test-Path -LiteralPath $src) {
                $dst = Join-Path $script:BackupDir $acc.Name
                New-Item -Path $dst -ItemType Directory -Force | Out-Null
                Get-ChildItem -LiteralPath $src -Filter '*.ini' -File -ErrorAction Stop | Copy-Item -Destination $dst -Force -ErrorAction Stop
            }
        }
        if ($ok) { Write-Log "Back-up gemaakt in $script:BackupDir" 'OK' }
        return $ok
    }
    catch {
        Write-Log "Back-up mislukt: $($_.Exception.Message)" 'Error'
        return $false
    }
}

function Remove-LinesFromIni {
    # Verwijdert regels met het pad of een scope ID, behoudt de originele encoding (meestal UTF-16 LE)
    param([string]$File, [string[]]$Needles)
    $bytes = $null
    for ($i = 0; $i -lt 5; $i++) {
        try { $bytes = [IO.File]::ReadAllBytes($File); break } catch { Start-Sleep -Milliseconds 500 }
    }
    if ($null -eq $bytes) { throw "Kan $File niet lezen" }

    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [Text.Encoding]::Unicode }
    elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $enc = New-Object Text.UTF8Encoding($true) }
    else { $enc = New-Object Text.UTF8Encoding($false) }

    $text  = [IO.File]::ReadAllText($File, $enc)
    $lines = $text -split "`r?`n"
    $keep  = New-Object System.Collections.Generic.List[string]
    $removed = 0
    foreach ($line in $lines) {
        $hit = $false
        foreach ($n in $Needles) { if ($n -and $line.IndexOf($n, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = $true; break } }
        if ($hit) { $removed++ } else { $keep.Add($line) }
    }
    if ($removed) {
        # Eerst naar tijdelijk bestand, dan vervangen: nooit een half geschreven ini achterlaten
        $tmp = "$File.eootmp"
        [IO.File]::WriteAllText($tmp, ($keep -join "`r`n"), $enc)
        Copy-Item -LiteralPath $tmp -Destination $File -Force -ErrorAction Stop
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
    return $removed
}

function Remove-LibrarySyncConfig {
    # Alleen uitvoeren met OneDrive gestopt. Geeft $true terug als de bibliotheek echt uit de config is.
    param([EooSyncLib]$Lib)
    $mount    = Get-NormalizedPath $Lib.MountPoint
    $scopeIds = New-Object System.Collections.Generic.List[string]
    $hits     = 0

    foreach ($acc in Get-OneDriveBusinessAccounts) {
        # ScopeId naar pad cache: scope ID's verzamelen en waarden verwijderen
        $cache = Join-Path $acc.RegPath 'ScopeIdToMountPointPathCache'
        if (Test-Path $cache) {
            $props = Get-ItemProperty $cache
            foreach ($name in (Get-Item $cache).Property) {
                if (Test-SamePath $props.$name $mount) {
                    if ($name.Length -ge 10) { $scopeIds.Add($name) }
                    try { Remove-ItemProperty -Path $cache -Name $name -ErrorAction Stop; $hits++ }
                    catch { Write-Log "  Registry waarde $name niet verwijderd: $($_.Exception.Message)" 'Warn' }
                }
            }
        }

        # Tenants: waardenaam is het lokale pad
        $tenants = Join-Path $acc.RegPath 'Tenants'
        if (Test-Path $tenants) {
            foreach ($t in Get-ChildItem $tenants) {
                foreach ($name in $t.Property) {
                    if (Test-SamePath $name $mount) {
                        try { Remove-ItemProperty -Path $t.PSPath -Name $name -ErrorAction Stop; $hits++ }
                        catch { Write-Log "  Tenant waarde niet verwijderd: $($_.Exception.Message)" 'Warn' }
                    }
                }
            }
        }

        # INI bestanden van het account (geen ClientPolicy of global)
        $settings = Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\settings\$($acc.Name)"
        if (Test-Path -LiteralPath $settings) {
            $needles = @("`"$mount`"", "`"$mount\`"") + $scopeIds
            foreach ($ini in Get-ChildItem -LiteralPath $settings -Filter '*.ini' -File |
                    Where-Object { $_.Name -notlike 'ClientPolicy*' -and $_.Name -ne 'global.ini' }) {
                try {
                    $n = Remove-LinesFromIni -File $ini.FullName -Needles $needles
                    if ($n) { Write-Log "  $n regel(s) verwijderd uit $($ini.Name)"; $hits += $n }
                }
                catch { Write-Log "  $($ini.Name) aanpassen mislukt: $($_.Exception.Message)" 'Warn' }
            }
        }
    }

    # SyncEngines registratie
    if ($Lib.ProviderKey -and (Test-Path $Lib.ProviderKey)) {
        try { Remove-Item -Path $Lib.ProviderKey -Recurse -Force -ErrorAction Stop; $hits++ }
        catch { Write-Log "  SyncEngines sleutel niet verwijderd: $($_.Exception.Message)" 'Warn' }
    }

    return ($hits -gt 1)
}

function Invoke-ProcessWithTimeout {
    # Start een proces en wacht met timeout zonder de UI te blokkeren. Hangt het, dan wordt het gestopt.
    param([string]$FilePath, [string]$Arguments, [int]$TimeoutSec)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = $FilePath
    $psi.Arguments       = $Arguments
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $end  = (Get-Date).AddSeconds($TimeoutSec)
    while (-not $proc.HasExited -and (Get-Date) -lt $end) { Start-Sleep -Milliseconds 200; Update-UI }
    if (-not $proc.HasExited) {
        try { $proc.Kill() } catch { }
        return -1
    }
    return $proc.ExitCode
}

function Remove-FolderRobust {
    # Verwijdert een map inclusief lange paden (rd met \\?\). Geeft $true als de map weg is.
    param([string]$Path, [int]$TimeoutSec = 300)
    if (-not (Test-Path -LiteralPath $Path)) { return $true }
    for ($try = 1; $try -le 3; $try++) {
        # Alleen-lezen/systeem attributen eraf, anders weigert rd sommige bestanden
        [void](Invoke-ProcessWithTimeout -FilePath "$env:SystemRoot\System32\attrib.exe" -Arguments ("-R -S -H `"{0}\*`" /S /D" -f $Path) -TimeoutSec 120)
        $code = Invoke-ProcessWithTimeout -FilePath "$env:SystemRoot\System32\cmd.exe" -Arguments ("/d /c rd /s /q `"{0}`"" -f [EooCloudFiles]::ToLongPath($Path)) -TimeoutSec $TimeoutSec
        if (-not (Test-Path -LiteralPath $Path)) { return $true }
        if ($code -eq -1) { Write-Log "  Verwijderen duurde langer dan $TimeoutSec s en is gestopt." 'Warn' }
        [void](Wait-UI 3 -IgnoreCancel)
    }
    return -not (Test-Path -LiteralPath $Path)
}

function Remove-LocalLibraryFolder {
    # Alle bestanden staan veilig in de cloud en de sync is al ontkoppeld, er gaat lokaal dus niets verloren
    param([EooSyncLib]$Lib)
    if (Remove-FolderRobust -Path $Lib.MountPoint) {
        Write-Log "  Lokale map verwijderd: $($Lib.MountPoint)" 'OK'
        Remove-EmptyParent -Path $Lib.MountPoint
        return $true
    }

    # Verwijderen lukt niet (bijv. bestand in gebruik). Hernoemen zodat de naam vrij is; later opnieuw proberen.
    $newPath = '{0}_ontkoppeld_{1:yyyyMMdd_HHmmss}' -f $Lib.MountPoint, (Get-Date)
    try {
        Rename-Item -LiteralPath $Lib.MountPoint -NewName (Split-Path $newPath -Leaf) -ErrorAction Stop
        $Lib.RenamedPath = $newPath
        Write-Log "  Map kon niet direct verwijderd worden, hernoemd naar $newPath (wordt na herstart opnieuw geprobeerd)." 'Warn'
    }
    catch {
        Write-Log "  Lokale map verwijderen en hernoemen mislukt: $($_.Exception.Message). Verwijder $($Lib.MountPoint) later handmatig." 'Warn'
    }
    return $false
}

function Remove-EmptyParent {
    param([string]$Path)
    $parent = Split-Path $Path -Parent
    if ($parent -and -not (Test-SamePath $parent $env:USERPROFILE) -and (Test-Path -LiteralPath $parent) -and
        -not (Get-ChildItem -LiteralPath $parent -Force -ErrorAction SilentlyContinue)) {
        Remove-Item -LiteralPath $parent -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path -LiteralPath $parent)) { Write-Log "  Lege map verwijderd: $parent" }
    }
}

function Test-PathStillSynced {
    # $true als OneDrive dit pad (nog/weer) als sync map kent. Dan NOOIT lokaal verwijderen.
    param([string]$Path)
    foreach ($l in @(Get-SyncedLibraries)) { if (Test-SamePath $l.MountPoint $Path) { return $true } }
    foreach ($acc in Get-OneDriveBusinessAccounts) {
        $cache = Join-Path $acc.RegPath 'ScopeIdToMountPointPathCache'
        if (Test-Path $cache) {
            $props = Get-ItemProperty $cache
            foreach ($name in (Get-Item $cache).Property) { if (Test-SamePath $props.$name $Path) { return $true } }
        }
    }
    return $false
}
#endregion

#region Hoofdproces
function Invoke-Preflight {
    # Stap 1 en 2: bibliotheek opzoeken in Graph en alles veilig in de cloud krijgen. Wijzigt niets aan de sync.
    param([EooSyncLib[]]$Selection, [int]$TimeoutMinutes, [bool]$FreeUpSpace)

    foreach ($lib in $Selection) { $lib.Ready = $false }
    if (-not (Connect-Graph)) { return }
    if (-not (Start-OneDrive)) { Write-Log 'OneDrive draait niet, kan niet controleren.' 'Error'; return }

    try { $existing = @(Get-ExistingShortcutIds) }
    catch { Write-Log "Bestaande snelkoppelingen ophalen mislukt: $($_.Exception.Message)" 'Warn'; $existing = @() }

    $candidates = New-Object System.Collections.Generic.List[EooSyncLib]
    foreach ($lib in $Selection) {
        if (Test-Cancelled) { return }
        Write-Log "[$($lib.Name)] opzoeken in SharePoint"
        $lib.Status = 'Zoeken in SharePoint...'
        $target = Resolve-LibraryTarget -Lib $lib
        if (-not $target -or -not $target.DriveId -or -not $target.ItemId) {
            $lib.Status = 'Fout: niet gevonden in SharePoint'
            Write-Log "[$($lib.Name)] niet gevonden via Graph, overgeslagen." 'Error'
            continue
        }
        $lib.DriveId        = $target.DriveId
        $lib.ItemId         = $target.ItemId
        $lib.ShortcutExists = $existing -contains $target.ItemId
        if ($lib.ShortcutExists) { Write-Log "[$($lib.Name)] snelkoppeling bestaat al, wordt niet opnieuw aangemaakt." }

        if (-not (Test-Path -LiteralPath $lib.MountPoint)) {
            $lib.Status = 'Fout: lokale map ontbreekt'
            Write-Log "[$($lib.Name)] lokale map ontbreekt, overgeslagen." 'Error'
            continue
        }
        $candidates.Add($lib)
    }
    if ($candidates.Count -eq 0) { return }

    Write-Log "Controle gestart voor $($candidates.Count) bibliotheek/bibliotheken (max $TimeoutMinutes min)."
    Wait-AllSafe -Libs $candidates.ToArray() -TimeoutMinutes $TimeoutMinutes -FreeUpSpace $FreeUpSpace
    if (Test-Cancelled) { Write-Log 'Controle geannuleerd.' 'Warn' }
}

function Invoke-Migration {
    # Stap 3 en 4: ontkoppelen, opruimen, OneDrive herstarten en snelkoppelingen maken
    param([EooSyncLib[]]$Selection, [bool]$DeleteLocal)

    if (Test-Cancelled) { return }
    $ready = @($Selection | Where-Object { $_.Ready })
    if (-not $ready) { Write-Log 'Geen bibliotheken klaar voor migratie.' 'Warn'; return }

    $list = ($ready | ForEach-Object { "  $($_.Name)" }) -join "`r`n"
    $msg  = "De volgende bibliotheken worden ontkoppeld en omgezet naar een snelkoppeling:`r`n`r`n$list`r`n`r`nOneDrive wordt hiervoor tijdelijk afgesloten. Doorgaan?"
    if ([System.Windows.MessageBox]::Show($script:Window, $msg, 'Bevestigen', 'YesNo', 'Warning') -ne 'Yes') {
        Write-Log 'Migratie geannuleerd door gebruiker.' 'Warn'; return
    }

    # Laatste controle vlak voor het stoppen van OneDrive (er kan net iets gewijzigd zijn)
    Write-Log 'Laatste controle voor ontkoppelen...'
    $final = New-Object System.Collections.Generic.List[EooSyncLib]
    foreach ($lib in $ready) {
        $scan = Invoke-CloudScan -Lib $lib -Unpin $false -Activity 'Laatste controle'
        if ($scan.Cancelled) { Write-Log 'Migratie geannuleerd.' 'Warn'; return }
        if ($scan.Fatal -or $scan.UnsafeCount -gt 0) {
            $lib.Ready  = $false
            $lib.Status = 'Niet klaar: gewijzigd na controle'
            Write-Log "[$($lib.Name)] is na de controle gewijzigd ($($scan.UnsafeCount) onveilig). Overgeslagen." 'Warn'
            Write-ScanDetails -Lib $lib -Scan $scan
            continue
        }
        $lib.Status = 'Klaar voor ontkoppelen'
        $final.Add($lib)
    }
    if ($final.Count -eq 0) { Write-Log 'Niets meer over om te migreren.' 'Warn'; return }

    # Vanaf hier niet meer annuleren of afsluiten: OneDrive staat uit
    $script:Critical = $true
    $ui.btnCancel.IsEnabled = $false
    try {
        if (-not (Stop-OneDrive)) { Write-Log 'OneDrive kon niet worden gestopt. Afgebroken, er is niets gewijzigd.' 'Error'; return }
        if (-not (Backup-OneDriveConfig)) { Write-Log 'Zonder volledige back-up wordt er niets aangepast. Afgebroken.' 'Error'; return }

        foreach ($lib in $final) {
            Write-Log "[$($lib.Name)] sync configuratie verwijderen"
            $lib.Status = 'Ontkoppelen...'
            try { $unlinked = Remove-LibrarySyncConfig -Lib $lib }
            catch { $unlinked = $false; Write-Log "[$($lib.Name)] fout bij ontkoppelen: $($_.Exception.Message)" 'Error' }

            if ($unlinked) {
                $lib.Unlinked = $true
                $lib.Status   = 'Ontkoppeld'
                Write-Log "[$($lib.Name)] ontkoppeld." 'OK'
                if ($DeleteLocal) {
                    $lib.Status = 'Lokale map opruimen...'
                    [void](Remove-LocalLibraryFolder -Lib $lib)
                    $lib.Status = 'Ontkoppeld'
                }
            }
            else {
                $lib.Status = 'Fout: config niet gevonden'
                Write-Log "[$($lib.Name)] sync configuratie niet (volledig) gevonden. Lokale map blijft staan." 'Error'
            }
        }
    }
    finally {
        # Nooit achterlaten met OneDrive uit
        if (-not (Start-OneDrive)) { Write-Log 'OneDrive kon niet worden gestart. Start OneDrive handmatig.' 'Error' }
        $script:Critical = $false
    }

    $unlinkedLibs = @($final | Where-Object { $_.Unlinked })
    if (-not $unlinkedLibs) { return }

    # Snelkoppelingen aanmaken (token kan verlopen zijn tijdens het wachten)
    [void](Connect-Graph)
    foreach ($lib in $unlinkedLibs) {
        if ($lib.ShortcutExists) { $lib.Status = 'Klaar (snelkoppeling bestond al)'; continue }
        $lib.Status = 'Snelkoppeling maken...'
        $done = $false
        for ($try = 1; $try -le 3 -and -not $done; $try++) {
            try {
                $sc = New-OneDriveShortcut -Lib $lib
                $lib.Status = 'Klaar: snelkoppeling toegevoegd'
                Write-Log "[$($lib.Name)] snelkoppeling '$($sc.name)' toegevoegd aan Mijn bestanden." 'OK'
                $done = $true
            }
            catch {
                Write-Log "[$($lib.Name)] snelkoppeling aanmaken mislukt (poging $try): $($_.Exception.Message)" 'Warn'
                if ($try -lt 3) { [void](Wait-UI 5 -IgnoreCancel) }
            }
        }
        if (-not $done) {
            $lib.Status = 'Fout: snelkoppeling mislukt'
            Write-Log "[$($lib.Name)] voeg de snelkoppeling handmatig toe via SharePoint > 'Snelkoppeling toevoegen aan Mijn bestanden'." 'Error'
        }
    }

    # Controle: komt OneDrive niet terug met de oude sync?
    Write-Log 'Controle na herstart (20 seconden wachten)...'
    [void](Wait-UI 20 -IgnoreCancel)
    foreach ($lib in $unlinkedLibs) {
        if (Test-PathStillSynced $lib.MountPoint) {
            Write-Log "[$($lib.Name)] staat na herstart weer in sync. Controleer Intune/GPO automatische sync." 'Warn'
            continue
        }
        # Uitgestelde opruiming van een hernoemde map, alleen als OneDrive die map zeker niet kent
        if ($lib.RenamedPath -and (Test-Path -LiteralPath $lib.RenamedPath)) {
            if (Test-PathStillSynced $lib.RenamedPath) {
                Write-Log "[$($lib.Name)] hernoemde map wordt door OneDrive gebruikt, niet verwijderd: $($lib.RenamedPath)" 'Warn'
            }
            elseif (Remove-FolderRobust -Path $lib.RenamedPath) {
                Write-Log "[$($lib.Name)] hernoemde map alsnog verwijderd." 'OK'
                Remove-EmptyParent -Path $lib.RenamedPath
            }
            else {
                Write-Log "[$($lib.Name)] map $($lib.RenamedPath) kon niet verwijderd worden. Na herstart van de pc handmatig verwijderen." 'Warn'
            }
        }
    }
    Write-Log 'Migratie afgerond.' 'OK'
}
#endregion

#region XAML
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SharePoint sync naar snelkoppeling" Height="720" Width="1100"
        WindowStartupLocation="CenterScreen" Background="#F4F6F8" FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Background" Value="#0F6CBD"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="180"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,12">
            <TextBlock Text="SharePoint sync naar snelkoppeling" FontSize="20" FontWeight="SemiBold" Foreground="#1B1B1B"/>
            <TextBlock x:Name="lblInfo" Foreground="#555" Margin="0,4,0,0" TextWrapping="Wrap"
                       Text="Selecteer de bibliotheken. De tool zet alles op alleen online, ontkoppelt de sync pas als elk bestand veilig in de cloud staat en voegt daarna een snelkoppeling toe aan Mijn bestanden."/>
        </StackPanel>

        <Border Grid.Row="1" Background="White" CornerRadius="4" Padding="12" Margin="0,0,0,12" BorderBrush="#DDD" BorderThickness="1">
            <WrapPanel>
                <CheckBox x:Name="chkFreeUp" Content="Eerst ruimte vrijmaken (alles alleen online)" IsChecked="True" Margin="0,0,24,0" VerticalAlignment="Center"/>
                <CheckBox x:Name="chkDelete" Content="Lokale map verwijderen na ontkoppelen" IsChecked="True" Margin="0,0,24,0" VerticalAlignment="Center"/>
                <TextBlock Text="Max. wachttijd totaal (min):" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBox x:Name="txtTimeout" Text="30" Width="50" VerticalAlignment="Center"/>
                <TextBlock x:Name="lblAccount" Margin="24,0,0,0" VerticalAlignment="Center" Foreground="#555"/>
            </WrapPanel>
        </Border>

        <DataGrid x:Name="grid" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False" HeadersVisibility="Column"
                  Background="White" BorderBrush="#DDD" RowHeight="28" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#EEE"
                  SelectionMode="Extended">
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="" Width="36">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Bibliotheek" Binding="{Binding Name}" Width="200" IsReadOnly="True"/>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="320" IsReadOnly="True"/>
                <DataGridTextColumn Header="Lokaal pad" Binding="{Binding MountPoint}" Width="*" IsReadOnly="True"/>
                <DataGridTextColumn Header="SharePoint URL" Binding="{Binding Url}" Width="*" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>

        <DockPanel Grid.Row="3" Margin="0,12,0,12" LastChildFill="False">
            <Button x:Name="btnRefresh" Content="Vernieuwen" Background="#6B6B6B"/>
            <Button x:Name="btnAll" Content="Alles selecteren" Background="#6B6B6B"/>
            <Button x:Name="btnConnect" Content="Aanmelden Microsoft 365" Background="#6B6B6B"/>
            <Button x:Name="btnMigrate" Content="Start migratie" DockPanel.Dock="Right" Margin="0"/>
            <Button x:Name="btnCheck" Content="Alleen controleren" DockPanel.Dock="Right" Background="#2B88D8"/>
            <Button x:Name="btnCancel" Content="Annuleren" DockPanel.Dock="Right" Background="#C50F1F" IsEnabled="False"/>
        </DockPanel>

        <TextBox x:Name="txtLog" Grid.Row="4" IsReadOnly="True" FontFamily="Consolas" FontSize="12" Background="#1E1E1E" Foreground="#DDD"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" TextWrapping="NoWrap" Padding="6"/>
    </Grid>
</Window>
'@
#endregion

#region Venster opbouwen
$reader        = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)
$ui = @{}
foreach ($n in 'lblInfo', 'chkFreeUp', 'chkDelete', 'txtTimeout', 'lblAccount', 'grid', 'btnRefresh', 'btnAll', 'btnConnect', 'btnMigrate', 'btnCheck', 'btnCancel', 'txtLog') {
    $ui[$n] = $script:Window.FindName($n)
}
$script:txtLog = $ui.txtLog
$ui.grid.ItemsSource = $script:Libraries

function Set-Busy {
    param([bool]$Busy)
    $script:Busy = $Busy
    foreach ($b in 'btnRefresh', 'btnAll', 'btnConnect', 'btnMigrate', 'btnCheck') { $ui[$b].IsEnabled = -not $Busy }
    foreach ($c in 'chkFreeUp', 'chkDelete', 'txtTimeout', 'grid') { $ui[$c].IsEnabled = -not $Busy }
    $ui.btnCancel.IsEnabled = $Busy
    if ($Busy) {
        if ($script:Cts) { $script:Cts.Dispose() }
        $script:Cts = New-Object System.Threading.CancellationTokenSource
    }
    $script:Window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::AppStarting } else { $null }
    Update-UI
}

function Import-LibraryList {
    $script:Libraries.Clear()
    foreach ($lib in Get-SyncedLibraries) { $script:Libraries.Add($lib) }
    Write-Log "$($script:Libraries.Count) gesynchroniseerde SharePoint bibliotheek/bibliotheken gevonden."

    $acc = @(Get-OneDriveBusinessAccounts)
    $ui.lblAccount.Text = if ($acc) { 'OneDrive: ' + (($acc | ForEach-Object { $_.UserEmail }) -join ', ') } else { 'Geen zakelijk OneDrive account gevonden' }
    if (Test-AutoMountPolicy) {
        Write-Log 'Policy TenantAutoMount actief: bibliotheken uit die policy komen na ontkoppelen terug. Pas eerst Intune/GPO aan.' 'Warn'
    }
}

function Get-Selection { @($script:Libraries | Where-Object { $_.Selected }) }

function Get-TimeoutValue {
    $t = 0
    if (-not [int]::TryParse($ui.txtTimeout.Text, [ref]$t) -or $t -lt 1) { $t = 30; $ui.txtTimeout.Text = '30' }
    $t
}
#endregion

#region Events
$ui.btnRefresh.Add_Click({ Import-LibraryList })

$ui.btnAll.Add_Click({
        # Alles al geselecteerd? Dan juist alles deselecteren
        $unselected = @($script:Libraries | Where-Object { -not $_.Selected }).Count
        $state = $unselected -gt 0
        foreach ($l in $script:Libraries) { $l.Selected = $state }
    })

$ui.btnConnect.Add_Click({
        Set-Busy $true
        try { [void](Connect-Graph) } finally { Set-Busy $false }
    })

$ui.btnCancel.Add_Click({
        if ($script:Critical) { return }
        if ($script:Cts -and -not $script:Cts.IsCancellationRequested) {
            Write-Log 'Annuleren aangevraagd, de huidige stap wordt netjes afgebroken...' 'Warn'
            $script:Cts.Cancel()
            $ui.btnCancel.IsEnabled = $false
        }
    })

$ui.btnCheck.Add_Click({
        $sel = Get-Selection
        if (-not $sel) { Write-Log 'Niets geselecteerd.' 'Warn'; return }
        Set-Busy $true
        try { Invoke-Preflight -Selection $sel -TimeoutMinutes (Get-TimeoutValue) -FreeUpSpace ([bool]$ui.chkFreeUp.IsChecked) }
        catch { Write-Log "Onverwachte fout: $($_.Exception.Message) (regel $($_.InvocationInfo.ScriptLineNumber))" 'Error' }
        finally { Set-Busy $false }
    })

$ui.btnMigrate.Add_Click({
        $sel = Get-Selection
        if (-not $sel) { Write-Log 'Niets geselecteerd.' 'Warn'; return }
        Set-Busy $true
        try {
            Invoke-Preflight -Selection $sel -TimeoutMinutes (Get-TimeoutValue) -FreeUpSpace ([bool]$ui.chkFreeUp.IsChecked)
            Invoke-Migration -Selection $sel -DeleteLocal ([bool]$ui.chkDelete.IsChecked)
        }
        catch {
            Write-Log "Onverwachte fout: $($_.Exception.Message) (regel $($_.InvocationInfo.ScriptLineNumber))" 'Error'
            # Nooit achterlaten met OneDrive uit
            [void](Start-OneDrive)
            $script:Critical = $false
        }
        finally { Set-Busy $false }
    })

$script:Window.Add_Closing({
        param($src, $e)
        if ($script:Critical) {
            [void][System.Windows.MessageBox]::Show($script:Window, 'OneDrive wordt op dit moment aangepast. Wacht tot deze stap klaar is.', 'Even geduld', 'OK', 'Information')
            $e.Cancel = $true
            return
        }
        if ($script:Busy) {
            if ([System.Windows.MessageBox]::Show($script:Window, 'Er loopt nog een actie. Afbreken en afsluiten?', 'Afsluiten', 'YesNo', 'Question') -ne 'Yes') {
                $e.Cancel = $true
                return
            }
            if ($script:Cts) { $script:Cts.Cancel() }
        }
    })

$script:Window.Add_ContentRendered({
        Write-Log "Logbestand: $script:LogFile"
        Import-LibraryList
    })
#endregion

#region Start
Write-Host 'SharePoint sync naar snelkoppeling gestart.' -ForegroundColor Cyan
[void]$script:Window.ShowDialog()
if ($script:Cts) { try { $script:Cts.Cancel() } catch { } }
if ($script:WeStoppedOneDrive -and -not (Get-OneDriveProcess)) { [void](Start-OneDrive) }
if (Get-MgContextSafe) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
Write-Host "Afgesloten. Log: $script:LogFile" -ForegroundColor Cyan
#endregion
