#Requires -Version 5.1
<#
.SYNOPSIS
    Zet gesynchroniseerde SharePoint bibliotheken om naar OneDrive snelkoppelingen ("Snelkoppeling toevoegen aan Mijn bestanden").

.DESCRIPTION
    WPF tool die per geselecteerde bibliotheek:
      1. De bibliotheek opzoekt via Microsoft Graph (drive en item ID).
      2. Alle bestanden op "alleen online" zet (Ruimte vrijmaken) en wacht tot echt alles in de cloud staat.
      3. OneDrive afsluit, de sync configuratie van die bibliotheek verwijdert (met back-up) en de lokale map opruimt.
      4. OneDrive herstart en een snelkoppeling in Mijn bestanden aanmaakt.

    Een bibliotheek wordt alleen ontkoppeld als ELK bestand online-only is. Eén lokaal bestand = overslaan.
    Draaien als de eindgebruiker zelf (HKCU), niet als een andere admin.

.NOTES
    Microsoft heeft geen ondersteunde API om één bibliotheek te ontkoppelen. Deze tool past de OneDrive
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
New-Item -Path $script:AppRoot -ItemType Directory -Force | Out-Null

# Bestanden die OneDrive zelf niet synchroniseert en dus nooit online-only worden
$script:IgnorePattern = '^(desktop\.ini|thumbs\.db|\.ds_store)$|^~\$|\.tmp$'

# Bestandsattributen van de Cloud Files API
$script:ATTR_RECALL_ON_DATA_ACCESS = 0x400000

function Update-UI {
    # Laat de WPF dispatcher de wachtrij verwerken zodat het venster niet bevriest
    if ($script:Window) {
        $script:Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
    }
}

function Wait-UI {
    param([int]$Seconds)
    $end = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $end) { Start-Sleep -Milliseconds 200; Update-UI }
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

#region Datamodel voor de DataGrid (INotifyPropertyChanged zodat status live bijwerkt)
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
    public event PropertyChangedEventHandler PropertyChanged;
    private void Notify(string n) { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(n)); }
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
    Get-ChildItem $root | Where-Object { $_.PSChildName -like 'Business*' } | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath
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

function Get-SyncedLibraries {
    # Gesynchroniseerde SharePoint bibliotheken uit de SyncEngines registratie
    $root = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
    if (-not (Test-Path $root)) { return @() }

    # Mounts binnen de eigen OneDrive map zijn al snelkoppelingen, die slaan we over
    $ownFolders = @(Get-OneDriveBusinessAccounts | ForEach-Object { Get-NormalizedPath $_.UserFolder } | Where-Object { $_ })

    foreach ($key in Get-ChildItem $root) {
        $p     = Get-ItemProperty $key.PSPath
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

function Start-OneDrive {
    if (Get-Process OneDrive -ErrorAction SilentlyContinue) { return $true }
    $exe = Get-OneDriveExe
    if (-not $exe) { Write-Log 'OneDrive.exe niet gevonden.' 'Error'; return $false }
    Write-Log 'OneDrive starten...'
    Start-Process -FilePath $exe -ArgumentList '/background'
    Wait-UI 10
    return [bool](Get-Process OneDrive -ErrorAction SilentlyContinue)
}

function Stop-OneDrive {
    $exe = Get-OneDriveExe
    if (-not (Get-Process OneDrive -ErrorAction SilentlyContinue)) { return $true }
    Write-Log 'OneDrive netjes afsluiten...'
    if ($exe) { Start-Process -FilePath $exe -ArgumentList '/shutdown' -WindowStyle Hidden }

    # Max 30 seconden wachten, daarna hard stoppen
    for ($i = 0; $i -lt 30 -and (Get-Process OneDrive -ErrorAction SilentlyContinue); $i++) { Wait-UI 1 }
    if (Get-Process OneDrive -ErrorAction SilentlyContinue) {
        Write-Log 'OneDrive reageert niet, proces wordt geforceerd gestopt.' 'Warn'
        Get-Process OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force
        Wait-UI 3
    }
    return -not (Get-Process OneDrive -ErrorAction SilentlyContinue)
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

#region Online-only controle
function Get-LocalOnlyFiles {
    # Geeft bestanden terug die NIET online-only zijn (dus mogelijk niet of niet volledig in de cloud)
    param([string]$Path)
    $errs = $null
    $files = Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue -ErrorVariable errs
    $result = @($files | Where-Object {
            $_.Name -notmatch $script:IgnorePattern -and
            -not ([int64]$_.Attributes -band $script:ATTR_RECALL_ON_DATA_ACCESS)
        })
    [pscustomobject]@{ Files = $result; Errors = @($errs).Count }
}

function Invoke-FreeUpSpace {
    # Zelfde als "Ruimte vrijmaken" in Verkenner: niet vastgezet, wel ontheven
    param([string]$Path)
    $attrib = Join-Path $env:SystemRoot 'System32\attrib.exe'
    $proc = Start-Process -FilePath $attrib -ArgumentList '-P', '+U', "`"$Path\*`"", '/S', '/D' -WindowStyle Hidden -PassThru
    while (-not $proc.HasExited) { Wait-UI 1 }
}

function Wait-AllOnline {
    param([EooSyncLib]$Lib, [int]$TimeoutMinutes)
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        $scan = Get-LocalOnlyFiles -Path $Lib.MountPoint
        if ($scan.Files.Count -eq 0 -and $scan.Errors -eq 0) { return $scan }
        $Lib.Status = "Wacht: $($scan.Files.Count) bestand(en) nog lokaal"
        Update-UI
        Wait-UI 10
    } while ((Get-Date) -lt $deadline)
    return $scan
}
#endregion

#region Microsoft Graph helpers
function Connect-Graph {
    if (Get-MgContextSafe) { return $true }

    # Alleen de authenticatie module is nodig, de rest gaat via Invoke-MgGraphRequest
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Log 'Module Microsoft.Graph.Authentication installeren (CurrentUser)...'
        try {
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
            }
            Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch { Write-Log "Installatie module mislukt: $($_.Exception.Message)" 'Error'; return $false }
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

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

function Get-MgContextSafe {
    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { return $null }
    Get-MgContext
}

function Invoke-Graph {
    param([string]$Method = 'GET', [string]$Uri, $Body)
    $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($Body) {
        $params.Body        = ($Body | ConvertTo-Json -Depth 10)
        $params.ContentType = 'application/json'
    }
    Invoke-MgGraphRequest @params
}

function Get-GraphAll {
    # Volgt @odata.nextLink voor gepagineerde resultaten
    param([string]$Uri)
    $items = @()
    while ($Uri) {
        $resp  = Invoke-Graph -Uri $Uri
        $items += @($resp.value)
        $Uri   = $resp.'@odata.nextLink'
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
    New-Item -Path $script:BackupDir -ItemType Directory -Force | Out-Null
    & reg.exe export 'HKCU\Software\Microsoft\OneDrive\Accounts' (Join-Path $script:BackupDir 'OneDrive_Accounts.reg') /y | Out-Null
    & reg.exe export 'HKCU\Software\SyncEngines\Providers\OneDrive' (Join-Path $script:BackupDir 'SyncEngines.reg') /y | Out-Null
    foreach ($acc in Get-OneDriveBusinessAccounts) {
        $src = Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\settings\$($acc.Name)"
        if (Test-Path $src) {
            $dst = Join-Path $script:BackupDir $acc.Name
            New-Item -Path $dst -ItemType Directory -Force | Out-Null
            Get-ChildItem -LiteralPath $src -Filter '*.ini' -File | Copy-Item -Destination $dst -Force
        }
    }
    Write-Log "Back-up gemaakt in $script:BackupDir" 'OK'
}

function Remove-LinesFromIni {
    # Verwijdert regels met het pad of een scope ID, behoudt de originele encoding (meestal UTF-16 LE)
    param([string]$File, [string[]]$Needles)
    $bytes = [IO.File]::ReadAllBytes($File)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [Text.Encoding]::Unicode }
    elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $enc = New-Object Text.UTF8Encoding($true) }
    else { $enc = New-Object Text.UTF8Encoding($false) }

    $text  = [IO.File]::ReadAllText($File, $enc)
    $lines = $text -split "`r?`n"
    $keep  = New-Object System.Collections.Generic.List[string]
    $removed = 0
    foreach ($line in $lines) {
        $hit = $false
        foreach ($n in $Needles) { if ($line.IndexOf($n, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = $true; break } }
        if ($hit) { $removed++ } else { $keep.Add($line) }
    }
    if ($removed) { [IO.File]::WriteAllText($File, ($keep -join "`r`n"), $enc) }
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
                if ((Get-NormalizedPath $props.$name) -eq $mount) {
                    if ($name.Length -ge 10) { $scopeIds.Add($name) }
                    Remove-ItemProperty -Path $cache -Name $name -ErrorAction SilentlyContinue
                    $hits++
                }
            }
        }

        # Tenants: waardenaam is het lokale pad
        $tenants = Join-Path $acc.RegPath 'Tenants'
        if (Test-Path $tenants) {
            foreach ($t in Get-ChildItem $tenants) {
                foreach ($name in $t.Property) {
                    if ((Get-NormalizedPath $name) -eq $mount) {
                        Remove-ItemProperty -Path $t.PSPath -Name $name -ErrorAction SilentlyContinue
                        $hits++
                    }
                }
            }
        }

        # INI bestanden van het account (geen ClientPolicy of global)
        $settings = Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\settings\$($acc.Name)"
        if (Test-Path $settings) {
            $needles = @("`"$mount`"", "`"$mount\`"") + $scopeIds
            Get-ChildItem -LiteralPath $settings -Filter '*.ini' -File |
                Where-Object { $_.Name -notlike 'ClientPolicy*' -and $_.Name -ne 'global.ini' } |
                ForEach-Object {
                    $n = Remove-LinesFromIni -File $_.FullName -Needles $needles
                    if ($n) { Write-Log "  $n regel(s) verwijderd uit $($_.Name)"; $hits += $n }
                }
        }
    }

    # SyncEngines registratie
    if ($Lib.ProviderKey -and (Test-Path $Lib.ProviderKey)) {
        Remove-Item -Path $Lib.ProviderKey -Recurse -Force -ErrorAction SilentlyContinue
        $hits++
    }

    return ($hits -gt 1)
}

function Remove-LocalLibraryFolder {
    # Alle bestanden zijn online-only placeholders, er gaat lokaal dus niets verloren
    param([EooSyncLib]$Lib)
    try {
        Remove-Item -LiteralPath $Lib.MountPoint -Recurse -Force -ErrorAction Stop
        Write-Log "  Lokale map verwijderd: $($Lib.MountPoint)" 'OK'

        # Lege bovenliggende tenantmap ook opruimen
        $parent = Split-Path $Lib.MountPoint -Parent
        if ($parent -and $parent -ne $env:USERPROFILE -and (Test-Path -LiteralPath $parent) -and
            -not (Get-ChildItem -LiteralPath $parent -Force -ErrorAction SilentlyContinue)) {
            Remove-Item -LiteralPath $parent -Force -ErrorAction SilentlyContinue
            Write-Log "  Lege map verwijderd: $parent"
        }
        return $true
    }
    catch {
        Write-Log "  Lokale map verwijderen mislukt: $($_.Exception.Message)" 'Warn'
        return $false
    }
}
#endregion

#region Hoofdproces
function Invoke-Preflight {
    # Stap 1 en 2: bibliotheek opzoeken in Graph en alles online-only maken. Wijzigt niets aan de sync.
    param([EooSyncLib[]]$Selection, [int]$TimeoutMinutes, [bool]$FreeUpSpace)

    if (-not (Connect-Graph)) { return }
    if (-not (Start-OneDrive)) { Write-Log 'OneDrive draait niet, kan niet controleren.' 'Error'; return }

    try { $existing = Get-ExistingShortcutIds }
    catch { Write-Log "Bestaande snelkoppelingen ophalen mislukt: $($_.Exception.Message)" 'Warn'; $existing = @() }

    foreach ($lib in $Selection) {
        $lib.Ready = $false
        Write-Log "[$($lib.Name)] controleren"

        # Graph doel bepalen
        $lib.Status = 'Zoeken in SharePoint...'
        $target = Resolve-LibraryTarget -Lib $lib
        if (-not $target) {
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

        # Ruimte vrijmaken en wachten tot alles online-only is
        if ($FreeUpSpace) {
            $lib.Status = 'Ruimte vrijmaken...'
            Write-Log "[$($lib.Name)] bestanden op alleen online zetten"
            Invoke-FreeUpSpace -Path $lib.MountPoint
            $scan = Wait-AllOnline -Lib $lib -TimeoutMinutes $TimeoutMinutes
        }
        else {
            $scan = Get-LocalOnlyFiles -Path $lib.MountPoint
        }

        if ($scan.Files.Count -eq 0 -and $scan.Errors -eq 0) {
            $lib.Ready  = $true
            $lib.Status = 'Klaar: alles online'
            Write-Log "[$($lib.Name)] alles staat online." 'OK'
        }
        else {
            $lib.Status = "Niet klaar: $($scan.Files.Count) lokaal, $($scan.Errors) leesfout(en)"
            Write-Log "[$($lib.Name)] $($scan.Files.Count) bestand(en) niet online-only, $($scan.Errors) leesfout(en). Overgeslagen." 'Warn'
            $scan.Files | Select-Object -First 10 | ForEach-Object { Write-Log "    $($_.FullName)" 'Warn' }
            if ($scan.Files.Count -gt 10) { Write-Log "    ... en nog $($scan.Files.Count - 10)" 'Warn' }
        }
    }
}

function Invoke-Migration {
    # Stap 3 en 4: ontkoppelen, opruimen, OneDrive herstarten en snelkoppelingen maken
    param([EooSyncLib[]]$Selection, [bool]$DeleteLocal)

    $ready = @($Selection | Where-Object { $_.Ready })
    if (-not $ready) { Write-Log 'Geen bibliotheken klaar voor migratie.' 'Warn'; return }

    $list = ($ready | ForEach-Object { "  $($_.Name)" }) -join "`r`n"
    $msg  = "De volgende bibliotheken worden ontkoppeld en omgezet naar een snelkoppeling:`r`n`r`n$list`r`n`r`nOneDrive wordt hiervoor tijdelijk afgesloten. Doorgaan?"
    if ([System.Windows.MessageBox]::Show($msg, 'Bevestigen', 'YesNo', 'Warning') -ne 'Yes') {
        Write-Log 'Migratie geannuleerd door gebruiker.' 'Warn'; return
    }

    # OneDrive stoppen en back-up maken
    if (-not (Stop-OneDrive)) { Write-Log 'OneDrive kon niet worden gestopt. Afgebroken.' 'Error'; return }
    Backup-OneDriveConfig

    # Per bibliotheek config verwijderen en eventueel de lokale map opruimen
    foreach ($lib in $ready) {
        Write-Log "[$($lib.Name)] sync configuratie verwijderen"
        if (Remove-LibrarySyncConfig -Lib $lib) {
            $lib.Unlinked = $true
            $lib.Status   = 'Ontkoppeld'
            Write-Log "[$($lib.Name)] ontkoppeld." 'OK'
            if ($DeleteLocal) { [void](Remove-LocalLibraryFolder -Lib $lib) }
        }
        else {
            $lib.Status = 'Fout: config niet gevonden'
            Write-Log "[$($lib.Name)] sync configuratie niet (volledig) gevonden. Lokale map blijft staan." 'Error'
        }
    }

    # OneDrive weer starten
    if (-not (Start-OneDrive)) { Write-Log 'OneDrive kon niet worden gestart. Start OneDrive handmatig.' 'Error' }

    # Snelkoppelingen aanmaken
    foreach ($lib in ($ready | Where-Object { $_.Unlinked })) {
        if ($lib.ShortcutExists) { $lib.Status = 'Klaar (snelkoppeling bestond al)'; continue }
        try {
            $sc = New-OneDriveShortcut -Lib $lib
            $lib.Status = 'Klaar: snelkoppeling toegevoegd'
            Write-Log "[$($lib.Name)] snelkoppeling '$($sc.name)' toegevoegd aan Mijn bestanden." 'OK'
        }
        catch {
            $lib.Status = 'Fout: snelkoppeling mislukt'
            Write-Log "[$($lib.Name)] snelkoppeling aanmaken mislukt: $($_.Exception.Message)" 'Error'
        }
    }

    # Controle: komt OneDrive niet terug met de oude sync?
    Write-Log 'Controle na herstart (20 seconden wachten)...'
    Wait-UI 20
    $stillSynced = @(Get-SyncedLibraries | ForEach-Object { $_.MountPoint })
    foreach ($lib in ($ready | Where-Object { $_.Unlinked })) {
        if ($stillSynced -contains $lib.MountPoint) {
            Write-Log "[$($lib.Name)] staat na herstart weer in sync. Controleer Intune/GPO automatische sync." 'Warn'
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
                       Text="Selecteer de bibliotheken. De tool zet alles op alleen online, ontkoppelt de sync pas als alles in de cloud staat en voegt daarna een snelkoppeling toe aan Mijn bestanden."/>
        </StackPanel>

        <Border Grid.Row="1" Background="White" CornerRadius="4" Padding="12" Margin="0,0,0,12" BorderBrush="#DDD" BorderThickness="1">
            <WrapPanel>
                <CheckBox x:Name="chkFreeUp" Content="Eerst ruimte vrijmaken (alles alleen online)" IsChecked="True" Margin="0,0,24,0" VerticalAlignment="Center"/>
                <CheckBox x:Name="chkDelete" Content="Lokale map verwijderen na ontkoppelen" IsChecked="True" Margin="0,0,24,0" VerticalAlignment="Center"/>
                <TextBlock Text="Max. wachttijd per bibliotheek (min):" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBox x:Name="txtTimeout" Text="15" Width="50" VerticalAlignment="Center"/>
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
                <DataGridTextColumn Header="Bibliotheek" Binding="{Binding Name}" Width="220" IsReadOnly="True"/>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="250" IsReadOnly="True"/>
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
foreach ($n in 'lblInfo', 'chkFreeUp', 'chkDelete', 'txtTimeout', 'lblAccount', 'grid', 'btnRefresh', 'btnAll', 'btnConnect', 'btnMigrate', 'btnCheck', 'txtLog') {
    $ui[$n] = $script:Window.FindName($n)
}
$script:txtLog = $ui.txtLog
$ui.grid.ItemsSource = $script:Libraries

function Set-Busy {
    param([bool]$Busy)
    foreach ($b in 'btnRefresh', 'btnAll', 'btnConnect', 'btnMigrate', 'btnCheck') { $ui[$b].IsEnabled = -not $Busy }
    $script:Window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
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
    if (-not [int]::TryParse($ui.txtTimeout.Text, [ref]$t) -or $t -lt 1) { $t = 15; $ui.txtTimeout.Text = '15' }
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

$ui.btnCheck.Add_Click({
        $sel = Get-Selection
        if (-not $sel) { Write-Log 'Niets geselecteerd.' 'Warn'; return }
        Set-Busy $true
        try { Invoke-Preflight -Selection $sel -TimeoutMinutes (Get-TimeoutValue) -FreeUpSpace ([bool]$ui.chkFreeUp.IsChecked) }
        catch { Write-Log "Onverwachte fout: $($_.Exception.Message)" 'Error' }
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
            Write-Log "Onverwachte fout: $($_.Exception.Message)" 'Error'
            # Nooit achterlaten met OneDrive uit
            [void](Start-OneDrive)
        }
        finally { Set-Busy $false }
    })

$script:Window.Add_ContentRendered({
        Write-Log "Logbestand: $script:LogFile"
        Import-LibraryList
    })
#endregion

#region Start
Write-Host 'SharePoint sync naar snelkoppeling gestart.' -ForegroundColor Cyan
[void]$script:Window.ShowDialog()
if (Get-MgContextSafe) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
Write-Host "Afgesloten. Log: $script:LogFile" -ForegroundColor Cyan
#endregion
