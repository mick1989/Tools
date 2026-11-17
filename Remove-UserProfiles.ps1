# Remove-UserProfiles.ps1
# EOO Profiel Opschoner - bulk verwijderen van gebruikersprofielen op Intune managed (Entra ID joined) systemen.
#
# Let op: op een Entra ID joined device zijn gebruikers GEEN lokale accounts. Er valt dus niets te
# verwijderen met Remove-LocalUser; wat er lokaal staat is het PROFIEL (map + ProfileList + caches).
# Dit script verwijdert:
#   - Het profiel via Win32_UserProfile (map C:\Users\<naam> + ProfileList registersleutel)
#   - Optioneel: Entra ID identiteitscache (IdentityStore\Cache + LogonCache) en ProfileGuid
#   - Optioneel: lidmaatschap van de lokale Administrators groep (SID-gebaseerd)
# Het Entra ID account zelf blijft bestaan; de gebruiker kan zich gewoon opnieuw aanmelden.
#
# Tabblad "Lokale accounts": lokale gebruikers aanmaken/bewerken en per account lokale groepen
# instellen (Administrators, Users, Remote Desktop Users, ...). Groepslidmaatschap gaat via ADSI,
# omdat Get-LocalGroupMember faalt op groepen met Entra ID / verweesde SIDs.
#
# Vereist: administrator, Windows PowerShell 5.1, 64-bit proces, STA (wordt automatisch geregeld;
#          vanuit PowerShell 7 wordt herstart in 5.1 i.v.m. de LocalAccounts module)
# Log: C:\ProgramData\EOO\Logs\ProfielOpschoner_yyyyMMdd.log

#region --- Elevatie / 64-bit / STA ---------------------------------------------------------------
$isAdmin  = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$is32on64 = [Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess
$isSta    = [Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
$isCore   = $PSVersionTable.PSEdition -eq 'Core'

if (-not $isAdmin -or $is32on64 -or -not $isSta -or $isCore) {
    $psExe = if ($is32on64) { "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe" }
             else           { "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe" }
    $psArgs = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$PSCommandPath`""
    try {
        if (-not $isAdmin) { Start-Process -FilePath $psExe -ArgumentList $psArgs -Verb RunAs }
        else               { Start-Process -FilePath $psExe -ArgumentList $psArgs }
    } catch {
        Write-Warning "Kon script niet (verhoogd) herstarten: $($_.Exception.Message)"
    }
    exit
}
#endregion

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

#region --- C# helpers (data-item met INotifyPropertyChanged + snelle mapgrootte) -----------------
if (-not ('EooProfileItem' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;

public class EooProfileItem : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;
    private void Notify(string p) { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(p)); }

    private bool _selected;
    public bool Selected {
        get { return _selected; }
        set { _selected = CanDelete && value; Notify("Selected"); }
    }
    public bool     CanDelete   { get; set; }
    public bool     Loaded      { get; set; }
    public string   UserName    { get; set; }
    public string   DisplayName { get; set; }
    public string   Type        { get; set; }
    public string   Sid         { get; set; }
    public string   ProfilePath { get; set; }
    public DateTime? LastUsed   { get; set; }
    public string   Reason      { get; set; }

    private long _sizeBytes = -1;
    public long SizeBytes { get { return _sizeBytes; } set { _sizeBytes = value; Notify("SizeBytes"); } }

    private string _sizeText = "-";
    public string SizeText { get { return _sizeText; } set { _sizeText = value; Notify("SizeText"); } }

    private string _status = "";
    public string Status { get { return _status; } set { _status = value; Notify("Status"); } }
}

public static class EooDirSize {
    // Iteratief, slaat reparse points (junctions / OneDrive placeholders) over, negeert access denied
    public static long Get(string root) {
        long total = 0;
        if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return 0;
        var stack = new Stack<DirectoryInfo>();
        stack.Push(new DirectoryInfo(root));
        while (stack.Count > 0) {
            DirectoryInfo dir = stack.Pop();
            try {
                foreach (FileSystemInfo fsi in dir.EnumerateFileSystemInfos()) {
                    try {
                        if ((fsi.Attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint) continue;
                        if ((fsi.Attributes & FileAttributes.Directory) == FileAttributes.Directory) stack.Push((DirectoryInfo)fsi);
                        else total += ((FileInfo)fsi).Length;
                    } catch { }
                }
            } catch { }
        }
        return total;
    }
}
"@
}
if (-not ('EooGroupItem' -as [type])) {
    Add-Type -TypeDefinition @"
using System.ComponentModel;

public class EooGroupItem : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;
    public string Name        { get; set; }
    public string Sid         { get; set; }
    public string Description { get; set; }

    private bool _isChecked;
    public bool IsChecked {
        get { return _isChecked; }
        set { _isChecked = value; var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs("IsChecked")); }
    }
}
"@
}
#endregion

#region --- Logging -----------------------------------------------------------------------------
$logDir = Join-Path $env:ProgramData 'EOO\Logs'
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$script:sync = [hashtable]::Synchronized(@{
    Queue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    LogFile = Join-Path $logDir ("ProfielOpschoner_{0}.log" -f (Get-Date -Format 'yyyyMMdd'))
    Cancel  = $false
})

function Write-UiLog {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    $script:sync.Queue.Enqueue($line)
    try { Add-Content -LiteralPath $script:sync.LogFile -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd'), $line) -Encoding UTF8 } catch { }
}
#endregion

#region --- Profielen ophalen --------------------------------------------------------------------
function Get-MachineSid {
    try {
        $acc = Get-CimInstance -ClassName Win32_UserAccount -Filter "LocalAccount=True" -ErrorAction Stop | Select-Object -First 1
        if ($acc) { return ($acc.SID -replace '-\d+$', '') }
    } catch { }
    return $null
}

function Resolve-ProfileName {
    param([string]$Sid, [string]$Path)
    $upn = $null; $display = $null

    # Entra ID: UPN + weergavenaam staan in de IdentityStore cache
    $ic = "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$Sid\IdentityCache\$Sid"
    if (Test-Path -LiteralPath $ic) {
        $p = Get-ItemProperty -LiteralPath $ic -ErrorAction SilentlyContinue
        if ($p) { $upn = $p.UserName; $display = $p.DisplayName }
    }
    if (-not $upn) {
        try { $upn = (New-Object Security.Principal.SecurityIdentifier $Sid).Translate([Security.Principal.NTAccount]).Value } catch { }
    }
    if (-not $upn) { $upn = Split-Path -Path $Path -Leaf }
    [pscustomobject]@{ UserName = $upn; DisplayName = $display }
}

function Get-ProfileLastUsed {
    param([string]$Sid, [string]$Path, $CimLastUse)
    # Betrouwbaarste bron: LocalProfileUnloadTime (Win10 1903+). LastUseTime van WMI wordt o.a. door AV bijgewerkt.
    $key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Sid"
    $p = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
    if ($p -and $null -ne $p.LocalProfileUnloadTimeHigh -and $null -ne $p.LocalProfileUnloadTimeLow) {
        try {
            $hi = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int32]$p.LocalProfileUnloadTimeHigh), 0)
            $lo = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int32]$p.LocalProfileUnloadTimeLow), 0)
            $ft = ([int64]$hi -shl 32) -bor [int64]$lo
            if ($ft -gt 0) { return [DateTime]::FromFileTime($ft) }
        } catch { }
    }
    $ntuser = Join-Path $Path 'NTUSER.DAT'
    if (Test-Path -LiteralPath $ntuser) {
        try { return (Get-Item -LiteralPath $ntuser -Force).LastWriteTime } catch { }
    }
    if ($CimLastUse) { return [DateTime]$CimLastUse }
    return $null
}

function Get-ProfileItems {
    $machineSid = Get-MachineSid
    $currentSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $exclude    = @($txtExclude.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
                Where-Object { -not $_.Special -and $_.LocalPath -and $_.LocalPath -notlike "$env:windir*" }

    foreach ($prof in $profiles) {
        $sid  = $prof.SID
        $path = $prof.LocalPath
        $name = Resolve-ProfileName -Sid $sid -Path $path

        $type = if     ($sid -like 'S-1-12-1-*')                       { 'Entra ID' }
                elseif ($machineSid -and $sid -like "$machineSid-*")   { 'Lokaal' }
                elseif ($sid -like 'S-1-5-21-*')                        { 'AD domein' }
                else                                                    { 'Overig' }

        $item = New-Object EooProfileItem
        $item.Sid         = $sid
        $item.ProfilePath = $path
        $item.UserName    = $name.UserName
        $item.DisplayName = $name.DisplayName
        $item.Type        = $type
        $item.Loaded      = [bool]$prof.Loaded
        $item.LastUsed    = Get-ProfileLastUsed -Sid $sid -Path $path -CimLastUse $prof.LastUseTime

        $reason = $null
        if     ($sid -eq $currentSid) { $reason = 'Huidige gebruiker' }
        elseif ($prof.Loaded)         { $reason = 'In gebruik (aangemeld)' }
        else {
            $leaf = Split-Path -Path $path -Leaf
            foreach ($pat in $exclude) {
                if ($leaf -like $pat -or $item.UserName -like $pat -or $item.UserName -like "*\$pat" -or $item.UserName -like "$pat@*") {
                    $reason = "Uitgesloten ($pat)"; break
                }
            }
        }
        if (-not (Test-Path -LiteralPath $path)) { $item.Status = 'Map ontbreekt' }

        $item.CanDelete = -not $reason
        $item.Reason    = $reason
        $item
    }
}

function Update-ProfileList {
    $script:items.Clear()
    try {
        $list = @(Get-ProfileItems | Sort-Object Type, UserName)
        foreach ($i in $list) { $script:items.Add($i) }
        Write-UiLog ("{0} profielen gevonden ({1} Entra ID)." -f $list.Count, @($list | Where-Object Type -eq 'Entra ID').Count)
    } catch {
        Write-UiLog "Profielen ophalen mislukt: $($_.Exception.Message)" 'ERROR'
    }
    $script:view.Refresh()
}
#endregion

#region --- Achtergrond workers (runspaces) ------------------------------------------------------
$script:sizeWorker = {
    param($Items, $Sync)
    function Format-Size([long]$b) {
        if ($b -ge 1GB) { return '{0:N2} GB' -f ($b / 1GB) }
        if ($b -ge 1MB) { return '{0:N0} MB' -f ($b / 1MB) }
        return '{0:N0} KB' -f ($b / 1KB)
    }
    foreach ($it in $Items) {
        if ($Sync.Cancel) { break }
        $it.SizeText = 'bezig...'
        $b = [EooDirSize]::Get($it.ProfilePath)
        $it.SizeBytes = $b
        $it.SizeText  = Format-Size $b
    }
}

$script:deleteWorker = {
    param($Items, $Opt, $Sync)

    function WLog([string]$Message, [string]$Level = 'INFO') {
        $line = "{0} [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
        $Sync.Queue.Enqueue($line)
        try { Add-Content -LiteralPath $Sync.LogFile -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd'), $line) -Encoding UTF8 } catch { }
    }

    function Remove-RegKeyLogged([string]$Path, [string]$Label) {
        if (Test-Path -LiteralPath $Path) {
            if ($Opt.WhatIf) { WLog "   [TEST] zou registersleutel verwijderen: $Path"; return }
            try { Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop; WLog "   $Label opgeruimd" }
            catch { WLog "   $Label opruimen mislukt: $($_.Exception.Message)" 'WARN' }
        }
    }

    function Remove-FolderForced([string]$Path) {
        # Eigenaar naar Administrators (SID, taalonafhankelijk), volledige rechten, daarna rd met long-path prefix
        & icacls.exe "$Path" /setowner '*S-1-5-32-544' /T /C /Q 2>&1 | Out-Null
        & icacls.exe "$Path" /grant '*S-1-5-32-544:(OI)(CI)F' /T /C /Q 2>&1 | Out-Null
        & cmd.exe /c "rd /s /q `"\\?\$Path`"" 2>&1 | Out-Null
        return -not (Test-Path -LiteralPath $Path)
    }

    # Administrators groep (naam is taalafhankelijk, dus via SID vertalen)
    $adminGroup = $null
    if ($Opt.RemoveAdmin) {
        try {
            $grpName = (New-Object Security.Principal.SecurityIdentifier 'S-1-5-32-544').Translate([Security.Principal.NTAccount]).Value.Split('\')[-1]
            $adminGroup = [ADSI]"WinNT://$env:COMPUTERNAME/$grpName,group"
        } catch { WLog "Kan lokale Administrators groep niet openen: $($_.Exception.Message)" 'WARN' }
    }

    $ok = 0; $fail = 0; $skip = 0
    foreach ($it in $Items) {
        if ($Sync.Cancel) { WLog 'Afgebroken door gebruiker.' 'WARN'; break }

        $sid   = $it.Sid
        $label = "$($it.UserName) ($sid)"
        $it.Status = 'Bezig...'
        WLog "Verwerken: $label"

        try {
            $cim = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$sid'" -ErrorAction Stop
            if (-not $cim) {
                WLog "   Profiel niet (meer) gevonden in WMI" 'WARN'
                $it.Status = 'Niet gevonden'; $skip++; continue
            }
            if ($cim.Loaded) {
                WLog "   Profiel is geladen (gebruiker aangemeld of proces actief) - overgeslagen" 'WARN'
                $it.Status = 'In gebruik'; $skip++; continue
            }
            $path = $cim.LocalPath

            # 1. Profiel zelf
            $removed = $false
            if ($Opt.WhatIf) {
                WLog "   [TEST] zou profiel verwijderen: $path"
                $removed = $true
            } else {
                try {
                    Remove-CimInstance -InputObject $cim -ErrorAction Stop
                    $removed = $true
                    WLog "   Profiel verwijderd via WMI ($path)"
                } catch {
                    WLog "   WMI-verwijdering mislukt: $($_.Exception.Message)" 'WARN'
                }

                if ((Test-Path -LiteralPath $path)) {
                    if ($Opt.Force) {
                        WLog "   Map bestaat nog, geforceerd verwijderen..."
                        if (Remove-FolderForced $path) { WLog "   Map geforceerd verwijderd" }
                        else { WLog "   Map kon niet volledig verwijderd worden: $path" 'WARN' }
                    } else {
                        WLog "   Map bestaat nog: $path (zet 'Forceer' aan om restanten op te ruimen)" 'WARN'
                    }
                }
                if (-not $removed -and $Opt.Force) {
                    Remove-RegKeyLogged "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" 'ProfileList-sleutel'
                    $removed = -not (Test-Path -LiteralPath $path) -and
                               -not (Test-Path -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid")
                }
            }

            # 2. ProfileGuid verwijzingen
            if ($removed) {
                Get-ChildItem -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileGuid' -ErrorAction SilentlyContinue |
                    Where-Object { (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue).SidString -eq $sid } |
                    ForEach-Object { Remove-RegKeyLogged $_.PSPath 'ProfileGuid-verwijzing' }
            }

            # 3. Entra ID identiteitscache (gecachte aanmelding)
            if ($removed -and $Opt.CleanIdentity -and $sid -like 'S-1-12-1-*') {
                Remove-RegKeyLogged "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$sid" 'IdentityStore cache'
                Get-ChildItem -LiteralPath 'HKLM:\SOFTWARE\Microsoft\IdentityStore\LogonCache' -ErrorAction SilentlyContinue | ForEach-Object {
                    Remove-RegKeyLogged (Join-Path $_.PSPath "Sid2Name\$sid") 'LogonCache Sid2Name'
                    Get-ChildItem -LiteralPath (Join-Path $_.PSPath 'Name2Sid') -ErrorAction SilentlyContinue |
                        Where-Object { (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue).Sid -eq $sid } |
                        ForEach-Object { Remove-RegKeyLogged $_.PSPath 'LogonCache Name2Sid' }
                }
            }

            # 4. Lokale Administrators groep
            if ($removed -and $adminGroup) {
                try {
                    $memberSids = @($adminGroup.Invoke('Members')) | ForEach-Object {
                        $bytes = $_.GetType().InvokeMember('objectSid', 'GetProperty', $null, $_, $null)
                        (New-Object Security.Principal.SecurityIdentifier($bytes, 0)).Value
                    }
                    if ($memberSids -contains $sid) {
                        if ($Opt.WhatIf) { WLog "   [TEST] zou verwijderen uit lokale Administrators" }
                        else { $adminGroup.Remove("WinNT://$sid"); WLog "   Verwijderd uit lokale Administrators" }
                    }
                } catch { WLog "   Administrators-lidmaatschap aanpassen mislukt: $($_.Exception.Message)" 'WARN' }
            }

            if ($removed) { $it.Status = if ($Opt.WhatIf) { 'Test OK' } else { 'Verwijderd' }; $ok++ }
            else          { $it.Status = 'Fout'; $fail++ }
        } catch {
            WLog "   Fout: $($_.Exception.Message)" 'ERROR'
            $it.Status = 'Fout'; $fail++
        }
    }
    $mode = if ($Opt.WhatIf) { ' (TESTMODUS - niets gewijzigd)' } else { '' }
    WLog ("Klaar{0}: {1} gelukt, {2} mislukt, {3} overgeslagen." -f $mode, $ok, $fail, $skip)
}

$script:job = $null

function Start-Worker {
    param([scriptblock]$Script, [object[]]$Arguments, [string]$Kind)
    $script:sync.Cancel = $false
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Script.ToString())
    foreach ($a in $Arguments) { [void]$ps.AddArgument($a) }
    $script:job = @{ PS = $ps; RS = $rs; Handle = $ps.BeginInvoke(); Kind = $Kind }
    Set-Busy $true
}

function Complete-Worker {
    $j = $script:job
    $script:job = $null
    try { [void]$j.PS.EndInvoke($j.Handle) } catch { Write-UiLog "Worker fout: $($_.Exception.Message)" 'ERROR' }
    foreach ($e in $j.PS.Streams.Error) { Write-UiLog "Worker: $e" 'ERROR' }
    $j.PS.Dispose(); $j.RS.Close(); $j.RS.Dispose()
    Set-Busy $false
    if ($j.Kind -eq 'Delete' -and -not $chkWhatIf.IsChecked) { Update-ProfileList }
}
#endregion

#region --- Lokale accounts ----------------------------------------------------------------------
$script:editAccount = $null   # $null = nieuw account, anders de rij van het account dat bewerkt wordt

function Get-LocalUserGroupNames {
    param([string]$Name)
    # Via ADSI: Get-LocalGroupMember faalt op groepen met Entra ID / verweesde SIDs
    try {
        $u = [ADSI]"WinNT://$env:COMPUTERNAME/$Name,user"
        @($u.Invoke('Groups')) | ForEach-Object { $_.GetType().InvokeMember('Name', 'GetProperty', $null, $_, $null) }
    } catch { @() }
}

function Update-GroupList {
    $script:groupItems.Clear()
    # Meest gebruikte groepen bovenaan: Administrators, Users, Remote Desktop Users
    $order = @{ 'S-1-5-32-544' = 0; 'S-1-5-32-545' = 1; 'S-1-5-32-555' = 2 }
    try {
        Get-LocalGroup -ErrorAction Stop |
            Sort-Object @{ e = { if ($order.ContainsKey($_.SID.Value)) { $order[$_.SID.Value] } else { 9 } } }, Name |
            ForEach-Object {
                $g = New-Object EooGroupItem
                $g.Name = $_.Name
                $g.Sid  = $_.SID.Value
                if ($_.Description) { $g.Description = $_.Description }
                $script:groupItems.Add($g)
            }
    } catch {
        Write-UiLog "Lokale groepen ophalen mislukt: $($_.Exception.Message)" 'ERROR'
    }
}

function Update-AccountList {
    param([string]$SelectName)
    $script:accounts.Clear()
    try {
        foreach ($u in @(Get-LocalUser -ErrorAction Stop | Sort-Object Name)) {
            $script:accounts.Add([pscustomobject]@{
                Name           = $u.Name
                FullName       = $u.FullName
                Description    = $u.Description
                IsEnabled      = [bool]$u.Enabled
                EnabledText    = if ($u.Enabled) { 'Ja' } else { 'Nee' }
                PwNeverExpires = $null -eq $u.PasswordExpires
                UserMayChange  = [bool]$u.UserMayChangePassword
                LastLogon      = $u.LastLogon
                Groups         = (@(Get-LocalUserGroupNames $u.Name) | Sort-Object) -join ', '
                Sid            = $u.SID.Value
            })
        }
    } catch {
        Write-UiLog "Lokale accounts ophalen mislukt: $($_.Exception.Message)" 'ERROR'
    }
    if ($SelectName) {
        $dgAccounts.SelectedItem = $script:accounts | Where-Object Name -eq $SelectName | Select-Object -First 1
    }
}

function Reset-AccountForm {
    $script:editAccount = $null
    $dgAccounts.SelectedItem = $null
    $txtAccMode.Text        = 'Nieuw lokaal account'
    $txtAccName.IsReadOnly  = $false
    $txtAccName.Text        = ''
    $txtAccFullName.Text    = ''
    $txtAccDesc.Text        = ''
    $pwdAcc.Clear(); $pwdAcc2.Clear()
    $txtPwHint.Text         = ''
    $chkAccNeverExpires.IsChecked = $true
    $chkAccMustChange.IsChecked   = $false
    $chkAccCannotChange.IsChecked = $false
    $chkAccDisabled.IsChecked     = $false
    foreach ($g in $script:groupItems) { $g.IsChecked = $g.Sid -eq 'S-1-5-32-545' }
    $btnAccSave.Content = 'Account aanmaken'
}

function Show-AccountInForm {
    param($Row)
    $script:editAccount = $Row
    $txtAccMode.Text        = "Account bewerken: $($Row.Name)"
    $txtAccName.IsReadOnly  = $true
    $txtAccName.Text        = $Row.Name
    $txtAccFullName.Text    = $Row.FullName
    $txtAccDesc.Text        = $Row.Description
    $pwdAcc.Clear(); $pwdAcc2.Clear()
    $txtPwHint.Text         = 'Leeg laten = wachtwoord niet wijzigen'
    $mustChange = $false
    try { $mustChange = ([ADSI]"WinNT://$env:COMPUTERNAME/$($Row.Name),user").PasswordExpired.Value -eq 1 } catch { }
    $chkAccNeverExpires.IsChecked = $Row.PwNeverExpires
    $chkAccMustChange.IsChecked   = $mustChange
    $chkAccCannotChange.IsChecked = -not $Row.UserMayChange
    $chkAccDisabled.IsChecked     = -not $Row.IsEnabled
    $member = @(Get-LocalUserGroupNames $Row.Name)
    foreach ($g in $script:groupItems) { $g.IsChecked = $member -contains $g.Name }
    $btnAccSave.Content = 'Wijzigingen opslaan'
}

function Test-AccountForm {
    # Geeft een foutmelding terug, of $null als alles klopt
    $name = $txtAccName.Text.Trim()
    if (-not $script:editAccount) {
        if (-not $name)                       { return 'Vul een gebruikersnaam in.' }
        if ($name.Length -gt 20)              { return 'Gebruikersnaam mag maximaal 20 tekens zijn.' }
        if ($name -match '["/\\\[\]:;|=,+*?<>]' -or $name -match '^[\s\.]+$') {
            return 'Gebruikersnaam bevat ongeldige tekens ( " / \ [ ] : ; | = , + * ? < > ).'
        }
        if (Get-LocalUser -Name $name -ErrorAction SilentlyContinue) { return "Er bestaat al een lokaal account '$name'." }
        if (Get-LocalGroup -Name $name -ErrorAction SilentlyContinue) { return "'$name' is al de naam van een lokale groep." }
        if ($pwdAcc.SecurePassword.Length -eq 0) { return 'Vul een wachtwoord in.' }
    }
    if ($pwdAcc.Password -cne $pwdAcc2.Password) { return 'De wachtwoorden komen niet overeen.' }
    if ($chkAccNeverExpires.IsChecked -and $chkAccMustChange.IsChecked) {
        return "'Wachtwoord verloopt nooit' en 'Wachtwoord wijzigen bij volgende aanmelding' gaan niet samen."
    }
    if ($chkAccCannotChange.IsChecked -and $chkAccMustChange.IsChecked) {
        return "'Gebruiker kan wachtwoord niet wijzigen' en 'Wachtwoord wijzigen bij volgende aanmelding' gaan niet samen."
    }
    return $null
}

function Set-AccountGroups {
    param([string]$Name, [string[]]$Add, [string[]]$Remove)
    foreach ($g in $Add) {
        try { ([ADSI]"WinNT://$env:COMPUTERNAME/$g,group").Add("WinNT://$env:COMPUTERNAME/$Name,user"); Write-UiLog "   Toegevoegd aan groep '$g'" }
        catch { Write-UiLog "   Toevoegen aan groep '$g' mislukt: $($_.Exception.Message)" 'ERROR' }
    }
    foreach ($g in $Remove) {
        try { ([ADSI]"WinNT://$env:COMPUTERNAME/$g,group").Remove("WinNT://$env:COMPUTERNAME/$Name,user"); Write-UiLog "   Verwijderd uit groep '$g'" }
        catch { Write-UiLog "   Verwijderen uit groep '$g' mislukt: $($_.Exception.Message)" 'ERROR' }
    }
}

function Set-MustChangePassword {
    param([string]$Name, [bool]$Value)
    try {
        $u = [ADSI]"WinNT://$env:COMPUTERNAME/$Name,user"
        $u.PasswordExpired = [int]$Value
        $u.SetInfo()
    } catch { Write-UiLog "   'Wachtwoord wijzigen bij aanmelding' instellen mislukt: $($_.Exception.Message)" 'WARN' }
}

function Save-Account {
    $err = Test-AccountForm
    if ($err) { [System.Windows.MessageBox]::Show($err, 'Lokale accounts', 'OK', 'Warning') | Out-Null; return }

    $name     = $txtAccName.Text.Trim()
    $wanted   = @($script:groupItems | Where-Object IsChecked | ForEach-Object Name)
    $current  = if ($script:editAccount) { @(Get-LocalUserGroupNames $name) } else { @() }
    $toAdd    = @($wanted  | Where-Object { $current -notcontains $_ })
    $toRemove = @($current | Where-Object { $wanted  -notcontains $_ -and ($script:groupItems.Name -contains $_) })
    $adminGrp = ($script:groupItems | Where-Object Sid -eq 'S-1-5-32-544').Name
    $newPw    = $pwdAcc.SecurePassword.Length -gt 0

    # Bevestiging met overzicht van de wijzigingen
    $lines = @()
    if (-not $script:editAccount) { $lines += "Nieuw lokaal account: $name" } else { $lines += "Account wijzigen: $name" }
    if ($script:editAccount -and $newPw) { $lines += ' - Wachtwoord wordt opnieuw ingesteld' }
    if ($toAdd)    { $lines += " - Toevoegen aan: $($toAdd -join ', ')" }
    if ($toRemove) { $lines += " - Verwijderen uit: $($toRemove -join ', ')" }
    if ($toAdd -contains $adminGrp) { $lines += "`r`nLet op: dit account krijgt lokale administratorrechten." }
    if (-not $wanted) { $lines += "`r`nLet op: het account zit straks in geen enkele groep en kan zich mogelijk niet aanmelden." }
    $res = [System.Windows.MessageBox]::Show(($lines -join "`r`n") + "`r`n`r`nDoorgaan?", 'Lokale accounts', 'YesNo', 'Question')
    if ($res -ne 'Yes') { return }

    $mustChange = [bool]$chkAccMustChange.IsChecked
    try {
        if (-not $script:editAccount) {
            $p = @{
                Name                     = $name
                Password                 = $pwdAcc.SecurePassword
                PasswordNeverExpires     = [bool]$chkAccNeverExpires.IsChecked
                UserMayNotChangePassword = [bool]$chkAccCannotChange.IsChecked
                Disabled                 = [bool]$chkAccDisabled.IsChecked
                ErrorAction              = 'Stop'
            }
            if ($txtAccFullName.Text.Trim()) { $p.FullName    = $txtAccFullName.Text.Trim() }
            if ($txtAccDesc.Text.Trim())     { $p.Description = $txtAccDesc.Text.Trim() }
            New-LocalUser @p | Out-Null
            Write-UiLog "Lokaal account '$name' aangemaakt door $([Security.Principal.WindowsIdentity]::GetCurrent().Name)."
        } else {
            $p = @{
                Name                  = $name
                FullName              = $txtAccFullName.Text.Trim()
                Description           = $txtAccDesc.Text.Trim()
                PasswordNeverExpires  = [bool]$chkAccNeverExpires.IsChecked
                UserMayChangePassword = -not $chkAccCannotChange.IsChecked
                ErrorAction           = 'Stop'
            }
            if ($newPw) { $p.Password = $pwdAcc.SecurePassword }
            Set-LocalUser @p
            if ($chkAccDisabled.IsChecked -and $script:editAccount.IsEnabled)         { Disable-LocalUser -Name $name -ErrorAction Stop; Write-UiLog "   Account uitgeschakeld" }
            elseif (-not $chkAccDisabled.IsChecked -and -not $script:editAccount.IsEnabled) { Enable-LocalUser -Name $name -ErrorAction Stop; Write-UiLog "   Account ingeschakeld" }
            Write-UiLog ("Lokaal account '{0}' bijgewerkt{1}." -f $name, $(if ($newPw) { ' (wachtwoord opnieuw ingesteld)' } else { '' }))
        }
    } catch {
        Write-UiLog "Account '$name' opslaan mislukt: $($_.Exception.Message)" 'ERROR'
        [System.Windows.MessageBox]::Show("Opslaan mislukt:`r`n$($_.Exception.Message)", 'Lokale accounts', 'OK', 'Error') | Out-Null
        return
    }

    # PasswordExpired alleen aanraken bij nieuw account met vinkje, of als de waarde wijzigt
    if (-not $script:editAccount) { if ($mustChange) { Set-MustChangePassword $name $true } }
    else {
        $was = $false
        try { $was = ([ADSI]"WinNT://$env:COMPUTERNAME/$name,user").PasswordExpired.Value -eq 1 } catch { }
        if ($was -ne $mustChange) { Set-MustChangePassword $name $mustChange }
    }
    Set-AccountGroups -Name $name -Add $toAdd -Remove $toRemove

    Update-AccountList -SelectName $name
}
#endregion

#region --- GUI ----------------------------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="EOO - Profiel Opschoner" Height="800" Width="1300" MinHeight="620" MinWidth="1000"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <SolidColorBrush x:Key="BrushBg"         Color="#FAF9F6"/>
    <SolidColorBrush x:Key="BrushCard"       Color="#FFFFFF"/>
    <SolidColorBrush x:Key="BrushBorder"     Color="#E4E1DB"/>
    <SolidColorBrush x:Key="BrushAccent"     Color="#208AD0"/>
    <SolidColorBrush x:Key="BrushAccentDark" Color="#03173D"/>
    <SolidColorBrush x:Key="BrushAccentSoft" Color="#E6F2FB"/>
    <SolidColorBrush x:Key="BrushSubText"    Color="#646C7A"/>
    <SolidColorBrush x:Key="BrushDanger"     Color="#DC2626"/>
    <SolidColorBrush x:Key="BrushConsoleDefault" Color="#CBD5E1"/>

    <!-- Wit, afgeronde knop; hover = EOO blauw -->
    <Style x:Key="OutlineButton" TargetType="Button">
      <Setter Property="Height" Value="32"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{StaticResource BrushCard}" BorderBrush="{StaticResource BrushBorder}"
                    BorderThickness="1" CornerRadius="8" Padding="14,0">
              <Border.Effect>
                <DropShadowEffect BlurRadius="6" ShadowDepth="1" Opacity="0.08" Color="#000000"/>
              </Border.Effect>
              <TextBlock x:Name="PART_Text" Text="{TemplateBinding Content}" Foreground="{StaticResource BrushAccentDark}"
                         VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="{StaticResource BrushAccent}"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BrushAccent}"/>
                <Setter TargetName="PART_Text" Property="Foreground" Value="White"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- EOO gradient knop -->
    <Style x:Key="PrimaryButton" TargetType="Button">
      <Setter Property="Height" Value="32"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" CornerRadius="8" Padding="16,0">
              <Border.Background>
                <LinearGradientBrush StartPoint="1,0" EndPoint="0,0">
                  <GradientStop Color="#208AD0" Offset="0"/>
                  <GradientStop Color="#03173D" Offset="1"/>
                </LinearGradientBrush>
              </Border.Background>
              <Border.Effect>
                <DropShadowEffect BlurRadius="8" ShadowDepth="1" Opacity="0.18" Color="#03173D"/>
              </Border.Effect>
              <TextBlock Text="{TemplateBinding Content}" Foreground="White" FontWeight="SemiBold"
                         VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Opacity" Value="0.9"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Rode actieknop -->
    <Style x:Key="DangerButton" TargetType="Button">
      <Setter Property="Height" Value="36"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{StaticResource BrushDanger}" CornerRadius="8" Padding="18,0">
              <Border.Effect>
                <DropShadowEffect BlurRadius="8" ShadowDepth="1" Opacity="0.18" Color="#7F1D1D"/>
              </Border.Effect>
              <TextBlock Text="{TemplateBinding Content}" Foreground="White" FontWeight="SemiBold"
                         VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#B91C1C"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Input" TargetType="TextBox">
      <Setter Property="Height" Value="30"/>
      <Setter Property="Padding" Value="6,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="{StaticResource BrushAccentDark}"/>
      <Setter Property="Background" Value="{StaticResource BrushCard}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BrushBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BrushAccent}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="PwInput" TargetType="PasswordBox">
      <Setter Property="Height" Value="30"/>
      <Setter Property="Padding" Value="6,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="{StaticResource BrushAccentDark}"/>
      <Setter Property="Background" Value="{StaticResource BrushCard}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BrushBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="PasswordBox">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BrushAccent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Tabbladen: platte tekst, actief = EOO blauwe onderlijn -->
    <Style TargetType="TabItem">
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="Bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,3"
                    CornerRadius="6,6,0,0" Padding="16,8" Margin="0,0,4,0">
              <ContentPresenter x:Name="Cp" ContentSource="Header" TextElement.FontWeight="SemiBold"
                                TextElement.Foreground="{StaticResource BrushSubText}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="{StaticResource BrushAccentSoft}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BrushAccent}"/>
                <Setter TargetName="Cp" Property="TextElement.Foreground" Value="{StaticResource BrushAccentDark}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,18,0"/>
      <Setter Property="Foreground" Value="{StaticResource BrushAccentDark}"/>
    </Style>

    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
      <Setter Property="Foreground" Value="{StaticResource BrushSubText}"/>
    </Style>

    <Style x:Key="SectionTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Foreground" Value="{StaticResource BrushAccentDark}"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>

    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="{StaticResource BrushAccentDark}"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="8,7"/>
      <Setter Property="BorderBrush" Value="{StaticResource BrushAccent}"/>
      <Setter Property="BorderThickness" Value="0,0,1,3"/>
    </Style>

    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="6,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{StaticResource BrushAccentSoft}"/>
          <Setter Property="Foreground" Value="{StaticResource BrushAccentDark}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <Grid Background="{StaticResource BrushBg}">
    <Grid.RowDefinitions>
      <RowDefinition Height="96"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <!-- Header (EOO gradient) -->
    <Grid Grid.Row="0">
      <Grid.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
          <GradientStop Color="#208AD0" Offset="0"/>
          <GradientStop Color="#03173D" Offset="1"/>
        </LinearGradientBrush>
      </Grid.Background>
      <Image x:Name="imgLogo" Width="200" Height="84" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="14,0,0,0" Stretch="Uniform"/>
      <TextBlock x:Name="txtLogoFallback" Text="EOO" Foreground="White" FontSize="24" FontWeight="Bold"
                 HorizontalAlignment="Left" VerticalAlignment="Center" Margin="22,0,0,0" Visibility="Collapsed"/>
      <StackPanel Margin="232,0,280,0" VerticalAlignment="Center">
        <TextBlock Text="Profiel Opschoner" Foreground="White" FontSize="18" FontWeight="SemiBold"/>
        <TextBlock Foreground="#CBD5E1" FontSize="12" Margin="0,3,0,0" TextWrapping="Wrap"
                   Text="Verwijdert lokale gebruikersprofielen (Entra ID / AD / lokaal) en beheert lokale accounts en hun groepslidmaatschap."/>
      </StackPanel>
      <StackPanel HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,18,0">
        <TextBlock x:Name="txtHost" Foreground="White" FontSize="15" FontWeight="SemiBold" HorizontalAlignment="Right"/>
        <TextBlock x:Name="txtRunAs" Foreground="#CBD5E1" FontSize="11" HorizontalAlignment="Right" Margin="0,2,0,0"/>
      </StackPanel>
      <Border Height="3" VerticalAlignment="Bottom" Background="{StaticResource BrushAccent}"/>
    </Grid>

    <Grid Grid.Row="1" Margin="16,8,16,16">
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="150"/>
      </Grid.RowDefinitions>

      <TabControl x:Name="tabMain" Grid.Row="0" Background="Transparent" BorderThickness="0" Padding="0,12,0,0">
      <TabItem Header="Profielen opschonen">
      <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Filter / toolbar -->
      <DockPanel Grid.Row="0" Margin="0,0,0,10" LastChildFill="False">
        <TextBlock Text="Zoeken" Style="{StaticResource Lbl}"/>
        <TextBox x:Name="txtSearch" Width="220" Style="{StaticResource Input}" Margin="0,0,16,0"/>
        <CheckBox x:Name="chkEntraOnly" Content="Alleen Entra ID" IsChecked="True"/>
        <CheckBox x:Name="chkHideProtected" Content="Verberg beschermde profielen"/>
        <TextBlock Text="Uitsluiten" Style="{StaticResource Lbl}"/>
        <TextBox x:Name="txtExclude" Width="200" Style="{StaticResource Input}" Margin="0,0,8,0"
                 Text="defaultuser*, *admin*" ToolTip="Komma-gescheiden. Wildcards toegestaan. Matcht op mapnaam en UPN/gebruikersnaam. Klik daarna op Vernieuwen."/>
        <Button x:Name="btnRefresh" Content="Vernieuwen" Style="{StaticResource PrimaryButton}" DockPanel.Dock="Right" Margin="0"/>
        <Button x:Name="btnSize" Content="Groottes berekenen" Style="{StaticResource OutlineButton}" DockPanel.Dock="Right"/>
      </DockPanel>

      <!-- Selectie -->
      <DockPanel Grid.Row="1" Margin="0,0,0,10" LastChildFill="False">
        <Button x:Name="btnSelectAll" Content="Alles aanvinken" Style="{StaticResource OutlineButton}"/>
        <Button x:Name="btnSelectNone" Content="Niets aanvinken" Style="{StaticResource OutlineButton}"/>
        <Button x:Name="btnSelectRows" Content="Gemarkeerde rijen aanvinken" Style="{StaticResource OutlineButton}"
                ToolTip="Markeer rijen met Shift/Ctrl en vink ze in een keer aan"/>
        <TextBlock Text="Niet gebruikt sinds" Style="{StaticResource Lbl}" Margin="10,0,6,0"/>
        <TextBox x:Name="txtDays" Width="56" Text="90" Style="{StaticResource Input}" HorizontalContentAlignment="Center"/>
        <TextBlock Text="dagen" Style="{StaticResource Lbl}" Margin="6,0,8,0"/>
        <Button x:Name="btnSelectOld" Content="Aanvinken" Style="{StaticResource OutlineButton}"/>
        <Border DockPanel.Dock="Right" Background="{StaticResource BrushAccentSoft}" CornerRadius="12" Padding="12,4" VerticalAlignment="Center">
          <TextBlock x:Name="txtCount" Foreground="{StaticResource BrushAccentDark}" FontWeight="SemiBold"/>
        </Border>
      </DockPanel>

      <!-- Profielen -->
      <Border Grid.Row="2" Background="{StaticResource BrushCard}" BorderBrush="{StaticResource BrushBorder}" BorderThickness="1" CornerRadius="8">
        <Border.Effect>
          <DropShadowEffect BlurRadius="10" ShadowDepth="1" Opacity="0.06" Color="#000000"/>
        </Border.Effect>
        <DataGrid x:Name="dgProfiles" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                  IsReadOnly="True" SelectionMode="Extended" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                  HorizontalGridLinesBrush="#EFEDE8" Background="{StaticResource BrushCard}" BorderThickness="0" RowHeight="28"
                  AlternatingRowBackground="#FCFBF9" Foreground="{StaticResource BrushAccentDark}" Margin="1">
          <DataGrid.RowStyle>
            <Style TargetType="DataGridRow">
              <Style.Triggers>
                <DataTrigger Binding="{Binding CanDelete}" Value="False">
                  <Setter Property="Foreground" Value="#A3A9B3"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Selected}" Value="True">
                  <Setter Property="Background" Value="#FDECEC"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="Verwijderd">
                  <Setter Property="Foreground" Value="#15803D"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="Fout">
                  <Setter Property="Foreground" Value="{StaticResource BrushDanger}"/>
                </DataTrigger>
              </Style.Triggers>
            </Style>
          </DataGrid.RowStyle>
          <DataGrid.Columns>
            <DataGridTemplateColumn Width="36" SortMemberPath="Selected">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                            IsEnabled="{Binding CanDelete}" HorizontalAlignment="Center" Margin="0"/>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTextColumn Header="Gebruiker" Binding="{Binding UserName}" Width="2*"/>
            <DataGridTextColumn Header="Weergavenaam" Binding="{Binding DisplayName}" Width="1.3*"/>
            <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="85"/>
            <DataGridTextColumn Header="Laatst gebruikt" Binding="{Binding LastUsed, StringFormat=dd-MM-yyyy HH:mm}" Width="130"/>
            <DataGridTextColumn Header="Grootte" Binding="{Binding SizeText}" SortMemberPath="SizeBytes" Width="90"/>
            <DataGridTextColumn Header="Profielmap" Binding="{Binding ProfilePath}" Width="1.5*"/>
            <DataGridTextColumn Header="Opmerking" Binding="{Binding Reason}" Width="1.2*"/>
            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
            <DataGridTextColumn Header="SID" Binding="{Binding Sid}" Width="1.5*"/>
          </DataGrid.Columns>
        </DataGrid>
      </Border>

      <!-- Opties -->
      <Border Grid.Row="3" Background="{StaticResource BrushCard}" BorderBrush="{StaticResource BrushBorder}" BorderThickness="1"
              CornerRadius="8" Padding="14,10" Margin="0,12,0,0">
        <StackPanel>
          <TextBlock Text="Opties" Style="{StaticResource SectionTitle}"/>
          <WrapPanel>
            <CheckBox x:Name="chkWhatIf" Content="Testmodus (niets verwijderen, alleen loggen)" IsChecked="True" FontWeight="SemiBold" Foreground="#B45309"/>
            <CheckBox x:Name="chkCleanIdentity" Content="Entra ID identiteitscache opruimen" IsChecked="True"
                      ToolTip="Verwijdert IdentityStore\Cache en LogonCache voor deze SID (gecachte offline aanmelding)."/>
            <CheckBox x:Name="chkRemoveAdmin" Content="Uit lokale Administrators groep halen"
                      ToolTip="Let op: als Intune (Account protection / Local users and groups) het lidmaatschap uitrolt, komt het terug."/>
            <CheckBox x:Name="chkForce" Content="Forceer bij fout (eigenaarschap + rd)"
                      ToolTip="Als WMI faalt of de map blijft staan: eigenaar naar Administrators, rechten geven en map + ProfileList-sleutel verwijderen."/>
          </WrapPanel>
        </StackPanel>
      </Border>

      <!-- Acties -->
      <DockPanel Grid.Row="4" Margin="0,12,0,12" LastChildFill="False">
        <Button x:Name="btnDelete" Content="Verwijder aangevinkte profielen" Style="{StaticResource DangerButton}"/>
        <Button x:Name="btnCancel" Content="Stoppen" Style="{StaticResource OutlineButton}" Height="36" IsEnabled="False"/>
        <Button x:Name="btnOpenLog" Content="Logmap openen" Style="{StaticResource OutlineButton}" Height="36" DockPanel.Dock="Right" Margin="0"/>
        <TextBlock x:Name="txtBusy" VerticalAlignment="Center" Foreground="{StaticResource BrushAccent}" FontWeight="SemiBold" Margin="8,0,0,0"/>
      </DockPanel>
      </Grid>
      </TabItem>

      <!-- Lokale accounts -->
      <TabItem Header="Lokale accounts">
        <Grid Margin="0,0,0,12">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="14"/>
            <ColumnDefinition Width="430"/>
          </Grid.ColumnDefinitions>

          <!-- Overzicht bestaande accounts -->
          <DockPanel Grid.Column="0">
            <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10" LastChildFill="False">
              <Button x:Name="btnAccNew" Content="Nieuw account" Style="{StaticResource PrimaryButton}"/>
              <Button x:Name="btnAccRefresh" Content="Vernieuwen" Style="{StaticResource OutlineButton}"/>
              <TextBlock Text="Selecteer een account om het te bewerken of groepen aan te passen." Style="{StaticResource Lbl}" Margin="8,0,0,0"/>
            </DockPanel>
            <Border Background="{StaticResource BrushCard}" BorderBrush="{StaticResource BrushBorder}" BorderThickness="1" CornerRadius="8">
              <Border.Effect>
                <DropShadowEffect BlurRadius="10" ShadowDepth="1" Opacity="0.06" Color="#000000"/>
              </Border.Effect>
              <DataGrid x:Name="dgAccounts" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                        IsReadOnly="True" SelectionMode="Single" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                        HorizontalGridLinesBrush="#EFEDE8" Background="{StaticResource BrushCard}" BorderThickness="0" RowHeight="28"
                        AlternatingRowBackground="#FCFBF9" Foreground="{StaticResource BrushAccentDark}" Margin="1">
                <DataGrid.RowStyle>
                  <Style TargetType="DataGridRow">
                    <Style.Triggers>
                      <DataTrigger Binding="{Binding IsEnabled}" Value="False">
                        <Setter Property="Foreground" Value="#A3A9B3"/>
                      </DataTrigger>
                    </Style.Triggers>
                  </Style>
                </DataGrid.RowStyle>
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Gebruiker" Binding="{Binding Name}" Width="1.2*"/>
                  <DataGridTextColumn Header="Volledige naam" Binding="{Binding FullName}" Width="1.2*"/>
                  <DataGridTextColumn Header="Actief" Binding="{Binding EnabledText}" Width="60"/>
                  <DataGridTextColumn Header="Groepen" Binding="{Binding Groups}" Width="2*"/>
                  <DataGridTextColumn Header="Laatste aanmelding" Binding="{Binding LastLogon, StringFormat=dd-MM-yyyy HH:mm}" Width="135"/>
                  <DataGridTextColumn Header="Beschrijving" Binding="{Binding Description}" Width="1.5*"/>
                </DataGrid.Columns>
              </DataGrid>
            </Border>
          </DockPanel>

          <!-- Formulier aanmaken / bewerken -->
          <Border Grid.Column="2" Background="{StaticResource BrushCard}" BorderBrush="{StaticResource BrushBorder}" BorderThickness="1"
                  CornerRadius="8" Padding="14,10">
            <Border.Effect>
              <DropShadowEffect BlurRadius="10" ShadowDepth="1" Opacity="0.06" Color="#000000"/>
            </Border.Effect>
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <TextBlock x:Name="txtAccMode" Grid.Row="0" Text="Nieuw lokaal account" Style="{StaticResource SectionTitle}" Margin="0,0,0,8"/>

              <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="130"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Text="Gebruikersnaam" Style="{StaticResource Lbl}"/>
                <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtAccName" Style="{StaticResource Input}" MaxLength="20"/>
                <TextBlock Grid.Row="1" Text="Volledige naam" Style="{StaticResource Lbl}"/>
                <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtAccFullName" Style="{StaticResource Input}" MaxLength="256"/>
                <TextBlock Grid.Row="2" Text="Beschrijving" Style="{StaticResource Lbl}"/>
                <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtAccDesc" Style="{StaticResource Input}" MaxLength="48"/>
                <TextBlock Grid.Row="3" Text="Wachtwoord" Style="{StaticResource Lbl}"/>
                <PasswordBox Grid.Row="3" Grid.Column="1" x:Name="pwdAcc" Style="{StaticResource PwInput}"/>
                <TextBlock Grid.Row="4" Text="Bevestig wachtwoord" Style="{StaticResource Lbl}"/>
                <PasswordBox Grid.Row="4" Grid.Column="1" x:Name="pwdAcc2" Style="{StaticResource PwInput}"/>
                <TextBlock Grid.Row="5" Grid.Column="1" x:Name="txtPwHint" FontSize="11" Foreground="{StaticResource BrushSubText}" Margin="2,0,0,4"/>
              </Grid>

              <StackPanel Grid.Row="2" Margin="0,4,0,8">
                <CheckBox x:Name="chkAccNeverExpires" Content="Wachtwoord verloopt nooit" IsChecked="True" Margin="0,3"/>
                <CheckBox x:Name="chkAccMustChange" Content="Wachtwoord wijzigen bij volgende aanmelding" Margin="0,3"/>
                <CheckBox x:Name="chkAccCannotChange" Content="Gebruiker kan wachtwoord niet wijzigen" Margin="0,3"/>
                <CheckBox x:Name="chkAccDisabled" Content="Account uitgeschakeld" Margin="0,3"/>
              </StackPanel>

              <TextBlock Grid.Row="3" Text="Lokale groepen" Style="{StaticResource SectionTitle}"/>
              <Border Grid.Row="4" BorderBrush="{StaticResource BrushBorder}" BorderThickness="1" CornerRadius="6" MinHeight="90">
                <ListBox x:Name="lstGroups" BorderThickness="0" Background="Transparent" Margin="2"
                         ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                  <ListBox.ItemTemplate>
                    <DataTemplate>
                      <CheckBox IsChecked="{Binding IsChecked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                Content="{Binding Name}" ToolTip="{Binding Description}" Margin="2,2"/>
                    </DataTemplate>
                  </ListBox.ItemTemplate>
                </ListBox>
              </Border>

              <DockPanel Grid.Row="5" Margin="0,12,0,0" LastChildFill="False">
                <Button x:Name="btnAccSave" Content="Account aanmaken" Style="{StaticResource PrimaryButton}"/>
                <Button x:Name="btnAccClear" Content="Leegmaken" Style="{StaticResource OutlineButton}" DockPanel.Dock="Right" Margin="0"/>
              </DockPanel>
            </Grid>
          </Border>
        </Grid>
      </TabItem>
      </TabControl>

      <!-- Log -->
      <Border Grid.Row="1" Background="{StaticResource BrushAccentDark}" CornerRadius="8" Padding="2">
        <TextBox x:Name="txtLog" IsReadOnly="True" FontFamily="Consolas" FontSize="12" Background="Transparent" BorderThickness="0"
                 Foreground="{StaticResource BrushConsoleDefault}" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                 TextWrapping="NoWrap" Padding="8,6"/>
      </Border>
    </Grid>
  </Grid>
</Window>
'@

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
foreach ($n in $xaml.SelectNodes('//*[@*[local-name()="Name"]]')) {
    Set-Variable -Name $n.Name -Value $window.FindName($n.Name) -Scope Script
}

$txtHost.Text  = "$env:COMPUTERNAME"
$txtRunAs.Text = "Uitgevoerd als $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"

# EOO logo (zelfde bron als de Install Tool); zonder internet tekst-fallback
try {
    $logoBytes = (Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Easy-Office-Online/software/refs/heads/main/EOO_Logo_rgb2.png' -UseBasicParsing -TimeoutSec 5).Content
    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit(); $bmp.StreamSource = New-Object System.IO.MemoryStream(, [byte[]]$logoBytes); $bmp.CacheOption = 'OnLoad'; $bmp.EndInit()
    $imgLogo.Source = $bmp
} catch {
    $imgLogo.Visibility = 'Collapsed'
    $txtLogoFallback.Visibility = 'Visible'
}

$script:items = New-Object 'System.Collections.ObjectModel.ObservableCollection[EooProfileItem]'
$dgProfiles.ItemsSource = $script:items
$script:view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:items)
$script:view.Filter = [Predicate[object]] {
    param($o)
    if ($chkEntraOnly.IsChecked -and $o.Type -ne 'Entra ID') { return $false }
    if ($chkHideProtected.IsChecked -and -not $o.CanDelete)   { return $false }
    $q = $txtSearch.Text.Trim()
    if ($q) {
        $hay = "$($o.UserName) $($o.DisplayName) $($o.ProfilePath) $($o.Sid)"
        if ($hay -notlike "*$q*") { return $false }
    }
    return $true
}

$script:accounts = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$dgAccounts.ItemsSource = $script:accounts
$script:groupItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[EooGroupItem]'
$lstGroups.ItemsSource = $script:groupItems

function Set-Busy([bool]$Busy) {
    foreach ($c in @($btnRefresh, $btnSize, $btnDelete, $btnSelectAll, $btnSelectNone, $btnSelectRows, $btnSelectOld, $txtExclude)) { $c.IsEnabled = -not $Busy }
    $btnCancel.IsEnabled = $Busy
    $txtBusy.Text = if ($Busy) { 'Bezig...' } else { '' }
}

function Get-VisibleItems { foreach ($o in $script:view) { $o } }
#endregion

#region --- Events -------------------------------------------------------------------------------
$txtSearch.Add_TextChanged({ $script:view.Refresh() })
$chkEntraOnly.Add_Click({ $script:view.Refresh() })
$chkHideProtected.Add_Click({ $script:view.Refresh() })

$btnRefresh.Add_Click({ Update-ProfileList })

$btnSelectAll.Add_Click({  foreach ($o in Get-VisibleItems) { $o.Selected = $true } })
$btnSelectNone.Add_Click({ foreach ($o in $script:items)    { $o.Selected = $false } })
$btnSelectRows.Add_Click({ foreach ($o in @($dgProfiles.SelectedItems)) { $o.Selected = $true } })

$btnSelectOld.Add_Click({
    $days = 0
    if (-not [int]::TryParse($txtDays.Text, [ref]$days) -or $days -lt 0) {
        [System.Windows.MessageBox]::Show('Vul een geldig aantal dagen in.', 'Profiel Opschoner', 'OK', 'Warning') | Out-Null
        return
    }
    $cutoff = (Get-Date).AddDays(-$days)
    $n = 0
    foreach ($o in Get-VisibleItems) {
        if ($o.CanDelete -and $o.LastUsed -and $o.LastUsed -lt $cutoff) { $o.Selected = $true; $n++ }
    }
    Write-UiLog "$n profielen aangevinkt die langer dan $days dagen niet gebruikt zijn."
})

$btnSize.Add_Click({
    $list = @(Get-VisibleItems)
    if (-not $list) { return }
    Write-UiLog "Groottes berekenen voor $($list.Count) profielen..."
    Start-Worker -Script $script:sizeWorker -Arguments @($list, $script:sync) -Kind 'Size'
})

$btnCancel.Add_Click({
    $script:sync.Cancel = $true
    Write-UiLog 'Stoppen aangevraagd, huidig profiel wordt nog afgerond...' 'WARN'
})

$btnOpenLog.Add_Click({ Start-Process explorer.exe -ArgumentList "`"$logDir`"" })

$btnDelete.Add_Click({
    $sel = @($script:items | Where-Object { $_.Selected -and $_.CanDelete })
    if (-not $sel) {
        [System.Windows.MessageBox]::Show('Geen profielen aangevinkt.', 'Profiel Opschoner', 'OK', 'Information') | Out-Null
        return
    }
    $whatIf = [bool]$chkWhatIf.IsChecked
    $names  = ($sel | Select-Object -First 15 | ForEach-Object { " - $($_.UserName)" }) -join "`r`n"
    if ($sel.Count -gt 15) { $names += "`r`n ... en nog $($sel.Count - 15)" }

    $msg = if ($whatIf) { "TESTMODUS: er wordt niets verwijderd.`r`n`r`n$($sel.Count) profielen worden gesimuleerd:`r`n$names" }
           else         { "Weet je zeker dat je $($sel.Count) profielen DEFINITIEF wilt verwijderen?`r`nAlle lokale data (Bureaublad, Documenten buiten OneDrive, AppData) gaat verloren.`r`n`r`n$names" }
    $icon = if ($whatIf) { 'Information' } else { 'Warning' }
    $res  = [System.Windows.MessageBox]::Show($msg, 'Profiel Opschoner', 'YesNo', $icon)
    if ($res -ne 'Yes') { return }

    $opt = @{
        WhatIf        = $whatIf
        CleanIdentity = [bool]$chkCleanIdentity.IsChecked
        RemoveAdmin   = [bool]$chkRemoveAdmin.IsChecked
        Force         = [bool]$chkForce.IsChecked
    }
    $mode = if ($whatIf) { 'TEST' } else { 'ECHT' }
    Write-UiLog ("Start verwijderen ({0}) van {1} profielen door {2}. Opties: identiteitscache={3}, admins={4}, forceer={5}" -f `
        $mode, $sel.Count, [Security.Principal.WindowsIdentity]::GetCurrent().Name, $opt.CleanIdentity, $opt.RemoveAdmin, $opt.Force)
    foreach ($s in $sel) { $s.Status = 'Wachtrij' }
    Start-Worker -Script $script:deleteWorker -Arguments @($sel, $opt, $script:sync) -Kind 'Delete'
})

# Lokale accounts
$btnAccNew.Add_Click({ Reset-AccountForm; $txtAccName.Focus() | Out-Null })
$btnAccClear.Add_Click({ Reset-AccountForm })
$btnAccRefresh.Add_Click({
    $keep = if ($script:editAccount) { $script:editAccount.Name } else { $null }
    Update-GroupList
    Update-AccountList -SelectName $keep
    if (-not $keep) { Reset-AccountForm }
})
$btnAccSave.Add_Click({ Save-Account })
$dgAccounts.Add_SelectionChanged({
    $row = $dgAccounts.SelectedItem
    if ($row) { Show-AccountInForm $row }
})

# Timer: log leegtrekken, teller bijwerken, worker afronden
$script:timer = New-Object System.Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromMilliseconds(250)
$script:timer.Add_Tick({
    $line = $null
    $added = $false
    while ($script:sync.Queue.TryDequeue([ref]$line)) { $txtLog.AppendText("$line`r`n"); $added = $true }
    if ($added) { $txtLog.ScrollToEnd() }

    $selCount = @($script:items | Where-Object Selected).Count
    $visCount = @(Get-VisibleItems).Count
    $txtCount.Text = "$selCount aangevinkt  |  $visCount van $($script:items.Count) zichtbaar"
    $btnDelete.Content = if ($chkWhatIf.IsChecked) { "Test verwijderen ($selCount)" } else { "Verwijder aangevinkte profielen ($selCount)" }

    if ($script:job -and $script:job.Handle.IsCompleted) { Complete-Worker }
})

$window.Add_Closing({
    param($s, $e)
    if ($script:job) {
        $r = [System.Windows.MessageBox]::Show('Er loopt nog een taak. Toch afsluiten?', 'Profiel Opschoner', 'YesNo', 'Warning')
        if ($r -ne 'Yes') { $e.Cancel = $true; return }
        $script:sync.Cancel = $true
    }
    $script:timer.Stop()
})

$window.Add_ContentRendered({
    Write-UiLog "Gestart op $env:COMPUTERNAME door $([Security.Principal.WindowsIdentity]::GetCurrent().Name). Log: $($script:sync.LogFile)"
    Update-ProfileList
    Update-GroupList
    Update-AccountList
    Reset-AccountForm
})
#endregion

$script:timer.Start()
[void]$window.ShowDialog()
