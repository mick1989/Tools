#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Audit (WPF) - zoek in het Unified Audit Log naar bestands- en
    mapactiviteiten (verplaatsen, hernoemen, verwijderen, kopieren, ...) in
    SharePoint en OneDrive, met een apart admin-account.
.DESCRIPTION
    - Inloggen via Connect-ExchangeOnline met je admin-account (UPN als hint,
      je kiest zelf het account). Optioneel een klant-tenant via GDAP
      (DelegatedOrganization).
    - Zoekperiode (van t/m, inclusief), activiteiten via vinkjes en snelkeuzes,
      recordtype, gebruikers (server-side), sites en bestand/map-namen
      (client-side, komma-gescheiden of regex).
    - Grote periodes worden in blokken opgehaald (per dag / 6 uur / uur).
      Raakt een blok de limiet van 50.000, dan wordt het automatisch gesplitst.
    - Resultaten: filteren, sorteren, details (volledige AuditData),
      bron- of doelmap openen in de browser, kopieren, exporteren naar CSV (;).
    Het zoeken draait op de achtergrond; het venster blijft bruikbaar en je
    kunt tussentijds stoppen (wat al binnen is blijft staan).
.NOTES
    Vereist: Windows PowerShell 5.1, module ExchangeOnlineManagement (3.x).
    Rol:     'View-Only Audit Logs' of 'Audit Logs' (bijv. via Compliance
             Management / Global Reader). Bij GDAP: rol in de relatie nodig.
    Bewaartermijn audit log: standaard 180 dagen.
#>

# Windows PowerShell 5.1 in STA nodig (WPF)
if ($PSVersionTable.PSEdition -eq 'Core' -or [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $self = $MyInvocation.MyCommand.Path
    if ($self) {
        Start-Process powershell.exe -ArgumentList "-STA -NoProfile -ExecutionPolicy Bypass -File `"$self`""
    }
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

#region --- Status en instellingen ---

$script:AppName      = 'SharePoint Audit'
$script:SettingsDir  = Join-Path $env:LOCALAPPDATA 'SharePointAuditTool'
$script:SettingsFile = Join-Path $script:SettingsDir 'settings.json'
$script:Window       = $null
$script:Connected    = $false
$script:ConnLabel    = ''
$script:Job          = $null
$script:Table        = $null
$script:OpChecks     = [ordered]@{}

$script:Sync = [hashtable]::Synchronized(@{
    Log     = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    Cancel  = $false
    Status  = ''
    Pct     = -1
    Table   = $null
    Summary = ''
})

# Activiteiten voor SharePointFileOperation (namen zoals in Purview)
$script:AllOps = @(
    'FileMoved','FolderMoved','FileRenamed','FolderRenamed','FileCopied','FolderCopied',
    'FileRecycled','FolderRecycled','FileDeleted','FolderDeleted',
    'FileDeletedFirstStageRecycleBin','FileDeletedSecondStageRecycleBin',
    'FolderDeletedFirstStageRecycleBin','FolderDeletedSecondStageRecycleBin',
    'FileRestored','FolderRestored',
    'FileUploaded','FileModified','FileModifiedExtended','FolderCreated','FolderModified',
    'FileCheckedIn','FileCheckedOut','FileCheckOutDiscarded',
    'FileAccessed','FileAccessedExtended','FileDownloaded','FilePreviewed','FileSyncDownloadedFull','FileSyncUploadedFull'
)

$script:Presets = [ordered]@{
    'Verplaatsen/verwijderen' = @('FileMoved','FolderMoved','FolderRenamed','FileRecycled','FolderRecycled','FolderDeleted','FileCopied')
    'Verplaatsen/hernoemen'   = @('FileMoved','FolderMoved','FileRenamed','FolderRenamed','FileCopied','FolderCopied')
    'Verwijderen'             = @('FileRecycled','FolderRecycled','FileDeleted','FolderDeleted','FileDeletedFirstStageRecycleBin','FileDeletedSecondStageRecycleBin','FolderDeletedFirstStageRecycleBin','FolderDeletedSecondStageRecycleBin')
    'Herstellen'              = @('FileRestored','FolderRestored')
    'Wijzigen/uploaden'       = @('FileUploaded','FileModified','FolderCreated','FolderModified','FileRenamed','FolderRenamed')
    'Openen/downloaden'       = @('FileAccessed','FileDownloaded','FilePreviewed','FileSyncDownloadedFull')
}

function Get-Settings {
    try {
        if (Test-Path $script:SettingsFile) { return (Get-Content $script:SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json) }
    } catch { }
    return $null
}

function Save-Settings {
    try {
        if (-not (Test-Path $script:SettingsDir)) { New-Item -ItemType Directory -Path $script:SettingsDir -Force | Out-Null }
        $ops = @($script:OpChecks.Keys | Where-Object { $script:OpChecks[$_].IsChecked })
        $o = [ordered]@{
            AdminUpn   = $script:TxtAdmin.Text.Trim()
            Tenant     = $script:TxtTenant.Text.Trim()
            NoWam      = [bool]$script:ChkNoWam.IsChecked
            Sites      = $script:TxtSites.Text
            SiteRegex  = [bool]$script:ChkRegex.IsChecked
            Operations = $ops
            Chunk      = $script:CmbChunk.SelectedIndex
            RecordType = $script:CmbRecord.SelectedIndex
        }
        [System.IO.File]::WriteAllText($script:SettingsFile, ($o | ConvertTo-Json -Depth 4), [System.Text.Encoding]::UTF8)
    } catch { }
}

#endregion

#region --- UI basis-helpers ---

function New-Brush {
    param([string]$Hex)
    $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
    $b.Freeze()
    return $b
}

function Show-Msg {
    param([string]$Text, [string]$Title = 'Info', [string]$Button = 'OK', [string]$Icon = 'Information')
    if ($script:Window -and $script:Window.IsLoaded) { return [System.Windows.MessageBox]::Show($script:Window, $Text, $Title, $Button, $Icon) }
    return [System.Windows.MessageBox]::Show($Text, $Title, $Button, $Icon)
}

function Write-Log {
    param([string]$Message, [string]$Type = 'info')
    if ($null -eq $script:LogPara) { return }
    $ts = Get-Date -Format 'HH:mm:ss'
    $color = switch ($Type) {
        'success' { '#50C850' }
        'error'   { '#F05050' }
        'warning' { '#FFB432' }
        'api'     { '#82B4FF' }
        default   { '#C8C8C8' }
    }
    $prefix = switch ($Type) {
        'success' { '[OK]  ' }
        'error'   { '[ERR] ' }
        'warning' { '[WRN] ' }
        'api'     { '[API] ' }
        default   { '[INF] ' }
    }
    $r1 = New-Object System.Windows.Documents.Run("$ts ")
    $r1.Foreground = New-Brush '#828282'
    $r2 = New-Object System.Windows.Documents.Run("$prefix$Message")
    $r2.Foreground = New-Brush $color
    $script:LogPara.Inlines.Add($r1)
    $script:LogPara.Inlines.Add($r2)
    $script:LogPara.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
    $script:LogBox.ScrollToEnd()
}

function Set-Status {
    param([string]$Text, [string]$Type = 'info')
    $script:LblStatus.Text = $Text
    $script:StatusBar.Background = switch ($Type) {
        'success' { New-Brush '#107C10' }
        'error'   { New-Brush '#C42B1C' }
        'warning' { New-Brush '#B26200' }
        default   { New-Brush '#0078D4' }
    }
}

function Update-Enabled {
    $busy = ($null -ne $script:Job)
    $script:HeaderPanel.IsEnabled = -not $busy
    $script:SearchArea.IsEnabled  = $script:Connected -and -not $busy
    $script:BtnSearch.IsEnabled   = $script:Connected -and -not $busy
    $script:BtnStop.IsEnabled     = $busy -and ($script:Job.Kind -eq 'search')
    $script:BtnConnect.IsEnabled  = -not $busy
    $script:BtnDisconnect.IsEnabled = $script:Connected -and -not $busy
    $script:Progress.Visibility   = if ($busy) { 'Visible' } else { 'Collapsed' }
    $script:Window.Cursor         = if ($busy -and $script:Job.Kind -ne 'search') { [System.Windows.Input.Cursors]::AppStarting } else { $null }
}

function Get-XamlControls {
    param($Root, [string]$Xaml)
    $h = @{}
    foreach ($m in [regex]::Matches($Xaml, 'x:Name="([^"]+)"')) {
        $n = $m.Groups[1].Value
        $c = $Root.FindName($n)
        if ($c) { $h[$n] = $c }
    }
    return $h
}

#endregion

#region --- Achtergrond-runspace (Exchange Online sessie leeft hier) ---

function Initialize-Worker {
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('Sync', $script:Sync)
    $script:Rs = $rs
}

function Start-Work {
    param([string]$Kind, [string]$Script, $Params)
    if ($script:Job) { return }
    $script:Sync.Cancel = $false
    $script:Sync.Pct    = -1
    $script:Sync.Status = ''
    $ps = [powershell]::Create()
    $ps.Runspace = $script:Rs
    [void]$ps.AddScript($Script)
    if ($null -ne $Params) { [void]$ps.AddArgument($Params) }
    $script:Job = @{ Kind = $Kind; PS = $ps; Handle = $ps.BeginInvoke(); Started = Get-Date }
    Update-Enabled
}

function Invoke-Tick {
    $item = $null
    while ($script:Sync.Log.TryDequeue([ref]$item)) { Write-Log ([string]$item[0]) ([string]$item[1]) }

    if (-not $script:Job) { return }
    $j = $script:Job
    if ($script:Sync.Status) {
        $el = [int]((Get-Date) - $j.Started).TotalSeconds
        $script:LblStatus.Text = "$($script:Sync.Status)   ($el s)"
    }
    $pct = [double]$script:Sync.Pct
    if ($pct -ge 0) { $script:Progress.IsIndeterminate = $false; $script:Progress.Value = $pct }
    else { $script:Progress.IsIndeterminate = $true }

    if (-not $j.Handle.IsCompleted) { return }

    $script:Job = $null
    $err = $null
    try { [void]$j.PS.EndInvoke($j.Handle) }
    catch {
        $ex = $_.Exception
        while ($ex.InnerException -and $ex -is [System.Management.Automation.MethodInvocationException]) { $ex = $ex.InnerException }
        $err = $ex
    }
    foreach ($e in $j.PS.Streams.Error) { if (-not $err) { $err = $e.Exception } }
    $j.PS.Dispose()
    while ($script:Sync.Log.TryDequeue([ref]$item)) { Write-Log ([string]$item[0]) ([string]$item[1]) }

    switch ($j.Kind) {
        'connect'    { Complete-Connect $err }
        'disconnect' { Complete-Disconnect $err }
        'search'     { Complete-Search $err }
    }
    Update-Enabled
}

#endregion

#region --- Scripts voor de achtergrond ---

$script:ConnectScript = @'
param($P)
function L { param($m, $t = 'info') $Sync.Log.Enqueue([object[]]@($m, $t)) }
$Sync.ConnUser = $null; $Sync.ConnOrg = $null; $Sync.AuditOn = $null; $Sync.HasSearch = $false
Import-Module ExchangeOnlineManagement -ErrorAction Stop
L "Module ExchangeOnlineManagement $((Get-Module ExchangeOnlineManagement).Version) geladen"
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }

$cmd = Get-Command Connect-ExchangeOnline
$prm = @{ UserPrincipalName = $P.Upn; ShowBanner = $false; ErrorAction = 'Stop' }
if ($P.Tenant) { $prm.DelegatedOrganization = $P.Tenant }
if ($P.NoWam) {
    if ($cmd.Parameters.ContainsKey('DisableWAM')) { $prm.DisableWAM = $true }
    else { L 'Deze moduleversie kent -DisableWAM niet; standaard login gebruikt' 'warning' }
}
$Sync.Status = 'Inloggen... (rond de login af in het venster dat opent)'
L ("Connect-ExchangeOnline -UserPrincipalName {0}{1}" -f $P.Upn, $(if ($P.Tenant) { " -DelegatedOrganization $($P.Tenant)" } else { '' })) 'api'
Connect-ExchangeOnline @prm

$ci = @(Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' }) | Select-Object -Last 1
if ($ci) {
    $Sync.ConnUser = $ci.UserPrincipalName
    $Sync.ConnOrg  = if ($ci.DelegatedOrganization) { $ci.DelegatedOrganization } else { $ci.TenantID }
}
$Sync.HasSearch = [bool](Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)
if ($Sync.HasSearch) {
    try { $Sync.AuditOn = [bool](Get-AdminAuditLogConfig -ErrorAction Stop).UnifiedAuditLogIngestionEnabled } catch { $Sync.AuditOn = $null }
}
'@

$script:DisconnectScript = @'
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop | Out-Null } catch { }
'@

$script:SearchScript = @'
param($P)
function L { param($m, $t = 'info') $Sync.Log.Enqueue([object[]]@($m, $t)) }

$inv = [System.Globalization.CultureInfo]::InvariantCulture
$utcStyle = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal

$t = New-Object System.Data.DataTable 'Audit'
foreach ($c in 'Tijd','TijdUTC','Gebruiker','Activiteit','Site','BronMap','BronBestand','DoelMap','DoelBestand','Item','ClientIP','Programma','Workload','Raw') { [void]$t.Columns.Add($c, [string]) }
$Sync.Table = $t
$Sync.Summary = ''
$seen = New-Object 'System.Collections.Generic.HashSet[string]'
$opCount = @{}

# Blokken maken
$queue = New-Object 'System.Collections.Generic.Queue[object]'
$s = $P.Start
while ($s -lt $P.End) {
    $e = if ($P.ChunkHours -gt 0) { $s.AddHours($P.ChunkHours) } else { $P.End }
    if ($e -gt $P.End) { $e = $P.End }
    $queue.Enqueue([object[]]@($s, $e))
    $s = $e
}
$totalBlocks = $queue.Count
$doneBlocks  = 0
$rawTotal    = 0
$failed      = 0
L ("Zoeken {0:dd-MM-yyyy HH:mm} t/m {1:dd-MM-yyyy HH:mm} (lokale tijd) in {2} blok(ken)" -f $P.Start, $P.End.AddSeconds(-1), $totalBlocks) 'api'

while ($queue.Count -gt 0) {
    if ($Sync.Cancel) { L 'Gestopt door gebruiker; resultaten tot nu toe blijven staan' 'warning'; break }
    $blk = $queue.Dequeue(); $bs = $blk[0]; $be = $blk[1]
    $label = '{0:dd-MM HH:mm} - {1:dd-MM HH:mm}' -f $bs, $be

    $prm = @{
        StartDate      = $bs.ToUniversalTime()
        EndDate        = $be.ToUniversalTime()
        SessionCommand = 'ReturnLargeSet'
        ResultSize     = 5000
        ErrorAction    = 'Stop'
    }
    if ($P.RecordType)    { $prm.RecordType = $P.RecordType }
    if ($P.Ops.Count)     { $prm.Operations = [string[]]$P.Ops }
    if ($P.Users.Count)   { $prm.UserIds    = [string[]]$P.Users }

    $attempt = 0; $ok = $false; $recs = $null; $total = 0
    while (-not $ok -and $attempt -lt 3 -and -not $Sync.Cancel) {
        $attempt++
        $prm.SessionId = [guid]::NewGuid().ToString()
        $recs = New-Object 'System.Collections.Generic.List[object]'
        $total = 0; $bad = $false
        try {
            while (-not $Sync.Cancel) {
                $Sync.Status = "Blok $($doneBlocks + 1)/$totalBlocks ($label): $($recs.Count) opgehaald..."
                $batch = @(Search-UnifiedAuditLog @prm)
                if ($batch.Count -eq 0) { break }
                $total = [int]$batch[0].ResultCount
                if ($total -le 0) { $bad = $true; break }
                foreach ($b in $batch) { $recs.Add($b) }
                $maxIdx = [int](($batch | Measure-Object -Property ResultIndex -Maximum).Maximum)
                if ($maxIdx -ge $total -or $recs.Count -ge $total) { break }
            }
        } catch {
            $bad = $true
            L "Blok $label poging $attempt/3: $($_.Exception.Message)" 'warning'
        }
        if ($bad) { if ($attempt -lt 3) { Start-Sleep -Seconds (3 * $attempt) }; continue }
        $ok = $true
    }
    if ($Sync.Cancel -and -not $ok) { continue }
    if (-not $ok) { L "Blok $label mislukt na 3 pogingen - overgeslagen" 'error'; $failed++; $doneBlocks++; continue }

    # Limiet van 50.000 per sessie: blok splitsen
    if ($total -ge 50000 -and ($be - $bs).TotalMinutes -gt 20) {
        $mid = $bs.AddTicks([long](($be - $bs).Ticks / 2))
        $queue.Enqueue([object[]]@($bs, $mid)); $queue.Enqueue([object[]]@($mid, $be))
        $totalBlocks++
        L "Blok $label heeft $total+ resultaten (limiet 50.000) - wordt gesplitst" 'warning'
        continue
    }
    if ($total -ge 50000) { L "Blok $label zit op de limiet van 50.000 - mogelijk onvolledig" 'warning' }

    $rawTotal += $recs.Count
    $added = 0
    foreach ($r in $recs) {
        $id = [string]$r.Identity
        if ($id -and -not $seen.Add($id)) { continue }
        try { $a = $r.AuditData | ConvertFrom-Json } catch { continue }
        $site = [string]$a.SiteUrl
        if ($P.SiteRegex -and -not ("$site $($a.ObjectId)" -match $P.SiteRegex)) { continue }
        if ($P.NameRegex -and -not ("$($a.SourceFileName) $($a.DestinationFileName) $($a.SourceRelativeUrl) $($a.DestinationRelativeUrl) $($a.ObjectId)" -match $P.NameRegex)) { continue }

        $tl = ''; $tu = [string]$a.CreationTime
        $dt = [datetime]::MinValue
        if ([datetime]::TryParse($tu, $inv, $utcStyle, [ref]$dt)) {
            $tu = $dt.ToString('yyyy-MM-dd HH:mm:ss')
            $tl = $dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
        }
        $op = [string]$a.Operation
        $row = $t.NewRow()
        $row.Tijd        = $tl
        $row.TijdUTC     = $tu
        $row.Gebruiker   = [string]$a.UserId
        $row.Activiteit  = $op
        $row.Site        = $site
        $row.BronMap     = [string]$a.SourceRelativeUrl
        $row.BronBestand = [string]$a.SourceFileName
        $row.DoelMap     = [string]$a.DestinationRelativeUrl
        $row.DoelBestand = [string]$a.DestinationFileName
        $row.Item        = [string]$a.ObjectId
        $row.ClientIP    = [string]$a.ClientIP
        $row.Programma   = [string]$a.UserAgent
        $row.Workload    = [string]$a.Workload
        $row.Raw         = [string]$r.AuditData
        $t.Rows.Add($row)
        $opCount[$op] = 1 + [int]$opCount[$op]
        $added++
    }
    $doneBlocks++
    $Sync.Pct = [math]::Round(100 * $doneBlocks / [math]::Max(1, $totalBlocks), 0)
    L "Blok $label : $($recs.Count) records, $added na filter (totaal $($t.Rows.Count))"
}

$Sync.Summary = (($opCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key): $($_.Value)" }) -join '   ')
$Sync.RawTotal = $rawTotal
$Sync.Failed = $failed
'@

#endregion

#region --- Acties ---

function Test-Module {
    if (Get-Module -ListAvailable ExchangeOnlineManagement) { return $true }
    $a = Show-Msg "De module ExchangeOnlineManagement is niet gevonden.`n`nNu installeren voor de huidige gebruiker (PowerShell Gallery)?" 'Module ontbreekt' 'YesNo' 'Question'
    if ($a -ne 'Yes') { return $false }
    try {
        Set-Status 'Module installeren...'
        $script:Window.Cursor = [System.Windows.Input.Cursors]::Wait
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log 'Module ExchangeOnlineManagement geinstalleerd' 'success'
        return $true
    } catch {
        Write-Log "Installatie mislukt: $($_.Exception.Message)" 'error'
        Show-Msg "Installatie mislukt:`n$($_.Exception.Message)" 'Fout' 'OK' 'Error' | Out-Null
        return $false
    } finally { $script:Window.Cursor = $null; Set-Status 'Niet verbonden' }
}

function Invoke-Connect {
    $upn = $script:TxtAdmin.Text.Trim()
    if ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        Show-Msg 'Vul je admin-account in (bijv. admin@bedrijf.onmicrosoft.com).' 'Admin-account' 'OK' 'Warning' | Out-Null
        $script:TxtAdmin.Focus() | Out-Null
        return
    }
    if (-not (Test-Module)) { return }
    $script:Connected = $false
    $script:LblConn.Text = 'Bezig met inloggen...'
    $script:LblConn.Foreground = New-Brush '#FFC864'
    Set-Status 'Inloggen... (rond de login af in het venster dat opent)'
    Save-Settings
    Start-Work 'connect' $script:ConnectScript @{ Upn = $upn; Tenant = $script:TxtTenant.Text.Trim(); NoWam = [bool]$script:ChkNoWam.IsChecked }
}

function Complete-Connect {
    param($Err)
    if ($Err) {
        $script:Connected = $false
        $script:LblConn.Text = 'Niet verbonden'
        $script:LblConn.Foreground = New-Brush '#FFC864'
        Write-Log "Inloggen mislukt: $($Err.Message)" 'error'
        if ($Err.Message -match 'window|handle|WAM|broker') {
            Write-Log "Tip: vink 'Browser-login (geen WAM)' aan en probeer opnieuw." 'warning'
        }
        Set-Status 'Inloggen mislukt' 'error'
        return
    }
    $user = [string]$script:Sync.ConnUser
    $org  = [string]$script:Sync.ConnOrg
    if (-not $script:Sync.HasSearch) {
        Write-Log "Verbonden als $user, maar Search-UnifiedAuditLog is niet beschikbaar. Het account mist de rol 'View-Only Audit Logs' / 'Audit Logs' (of die zit niet in de GDAP-relatie)." 'error'
        Start-Work 'disconnect' $script:DisconnectScript $null
        Set-Status 'Geen rechten op het audit log' 'error'
        return
    }
    $script:Connected = $true
    $script:ConnLabel = if ($org) { "$user  |  $org" } else { $user }
    $script:LblConn.Text = $script:ConnLabel
    $script:LblConn.ToolTip = $script:ConnLabel
    $script:LblConn.Foreground = New-Brush '#B4F0B4'
    Write-Log "Verbonden als $user$(if ($org) { " ($org)" })" 'success'
    $typed = $script:TxtAdmin.Text.Trim()
    if ($user -and $typed -and $user -ne $typed) {
        Write-Log "Let op: ingelogd account ($user) is niet het ingevulde account ($typed)." 'warning'
    }
    if ($script:Sync.AuditOn -eq $false) {
        Write-Log 'Unified audit log lijkt UIT te staan in deze tenant (UnifiedAuditLogIngestionEnabled = False). Er komen dan geen resultaten.' 'warning'
    }
    Set-Status "Verbonden - kies periode en activiteiten en klik Zoeken" 'success'
}

function Invoke-Disconnect {
    Start-Work 'disconnect' $script:DisconnectScript $null
    Set-Status 'Uitloggen...'
}

function Complete-Disconnect {
    param($Err)
    if (-not $script:Connected) { return }
    $script:Connected = $false
    $script:LblConn.Text = 'Niet verbonden'
    $script:LblConn.ToolTip = $null
    $script:LblConn.Foreground = New-Brush '#FFC864'
    Write-Log 'Uitgelogd' 'success'
    Set-Status 'Niet verbonden'
}

function ConvertTo-TermRegex {
    param([string]$Text, [bool]$IsRegex)
    $Text = $Text.Trim()
    if (-not $Text) { return $null }
    if ($IsRegex) { return $Text }
    $terms = @($Text -split '[,;|]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_) })
    if (-not $terms.Count) { return $null }
    return ($terms -join '|')
}

function Invoke-Search {
    if (-not $script:DpFrom.SelectedDate -or -not $script:DpTo.SelectedDate) {
        Show-Msg 'Kies een begin- en einddatum.' 'Periode' 'OK' 'Warning' | Out-Null; return
    }
    $start = $script:DpFrom.SelectedDate.Date
    $end   = $script:DpTo.SelectedDate.Date.AddDays(1)
    if ($end -le $start) { Show-Msg 'De einddatum ligt voor de begindatum.' 'Periode' 'OK' 'Warning' | Out-Null; return }
    if ($end -gt (Get-Date).Date.AddDays(1)) { $end = (Get-Date).Date.AddDays(1) }
    if ($start -lt (Get-Date).Date.AddDays(-180)) {
        Write-Log 'Begindatum ligt meer dan 180 dagen terug; zonder Audit Premium is ouder dan 180 dagen niet meer beschikbaar.' 'warning'
    }

    $ops = @($script:OpChecks.Keys | Where-Object { $script:OpChecks[$_].IsChecked })
    $extra = @($script:TxtExtraOps.Text -split '[,;\s]+' | Where-Object { $_ })
    $ops = @($ops + $extra | Select-Object -Unique)
    if (-not $ops.Count) {
        $a = Show-Msg "Geen activiteiten gekozen: alle activiteiten van dit recordtype worden opgehaald. Dat kan veel resultaten en tijd kosten.`n`nDoorgaan?" 'Alle activiteiten' 'YesNo' 'Question'
        if ($a -ne 'Yes') { return }
    }

    $siteRx = $null; $nameRx = $null
    try {
        $siteRx = ConvertTo-TermRegex $script:TxtSites.Text ([bool]$script:ChkRegex.IsChecked)
        if ($siteRx) { [void][regex]::new($siteRx) }
        $nameRx = ConvertTo-TermRegex $script:TxtNames.Text $false
    } catch {
        Show-Msg "Ongeldige regex bij Sites:`n$($_.Exception.Message)" 'Regex' 'OK' 'Warning' | Out-Null; return
    }

    $users = @($script:TxtUsers.Text -split '[,;\s]+' | Where-Object { $_ })
    $chunkHours = switch ($script:CmbChunk.SelectedIndex) { 0 { 24 } 1 { 6 } 2 { 1 } default { 0 } }
    $rt = [string]$script:CmbRecord.SelectedItem.Tag

    $p = @{
        Start = $start; End = $end; Ops = $ops; Users = $users; RecordType = $rt
        ChunkHours = $chunkHours; SiteRegex = $siteRx; NameRegex = $nameRx
    }
    Write-Log ("Activiteiten: {0}" -f $(if ($ops.Count) { $ops -join ', ' } else { '(alle)' }))
    if ($siteRx) { Write-Log "Sitefilter: $siteRx" }
    if ($nameRx) { Write-Log "Bestand/map-filter: $nameRx" }
    if ($users.Count) { Write-Log "Gebruikers: $($users -join ', ')" }

    $script:Grid.ItemsSource = $null
    $script:Table = $null
    $script:TxtDetail.Text = ''
    $script:LblCount.Text = ''
    $script:LblSummary.Text = ''
    Save-Settings
    Set-Status 'Zoeken...'
    Start-Work 'search' $script:SearchScript $p
}

function Stop-Search {
    if (-not $script:Job -or $script:Job.Kind -ne 'search') { return }
    if (-not $script:Sync.Cancel) {
        $script:Sync.Cancel = $true
        Write-Log 'Stoppen na het lopende verzoek... (nogmaals klikken = direct afbreken)' 'warning'
    } else {
        try { [void]$script:Job.PS.BeginStop($null, $null) } catch { }
    }
}

function Complete-Search {
    param($Err)
    $stopped = $Err -is [System.Management.Automation.PipelineStoppedException]
    if ($Err -and -not $stopped) {
        Write-Log "Zoeken mislukt: $($Err.Message)" 'error'
        if ($Err.Message -match 'session|token|expired|verlopen|Connect-ExchangeOnline') { Write-Log 'Sessie mogelijk verlopen; log opnieuw in.' 'warning' }
    }
    $script:Table = $script:Sync.Table
    $n = 0
    if ($script:Table) {
        $script:Table.DefaultView.Sort = 'Tijd ASC'
        $script:Grid.ItemsSource = $script:Table.DefaultView
        $n = $script:Table.Rows.Count
        # Activiteit-filter vullen
        $script:CmbOpFilter.Items.Clear()
        [void]$script:CmbOpFilter.Items.Add('(alle activiteiten)')
        foreach ($o in @($script:Table.Rows | ForEach-Object { $_.Activiteit } | Sort-Object -Unique)) { [void]$script:CmbOpFilter.Items.Add($o) }
        $script:CmbOpFilter.SelectedIndex = 0
    }
    $script:LblSummary.Text = [string]$script:Sync.Summary
    Update-Filter
    if ($stopped -or $script:Sync.Cancel) {
        Set-Status "Gestopt - $n regels" 'warning'
    } elseif ($Err) {
        Set-Status "Zoeken mislukt - $n regels" 'error'
    } elseif ($n -eq 0) {
        Write-Log "Geen resultaten ($($script:Sync.RawTotal) records opgehaald voor filtering)." 'warning'
        Set-Status 'Geen resultaten gevonden voor deze periode/filters' 'warning'
    } else {
        $f = [int]$script:Sync.Failed
        Write-Log "Klaar: $n regels ($($script:Sync.RawTotal) records opgehaald)$(if ($f) { ", $f blok(ken) mislukt" })" 'success'
        Set-Status "Klaar - $n regels gevonden" $(if ($f) { 'warning' } else { 'success' })
    }
}

function Get-LikeEscaped {
    param([string]$s)
    $s = $s.Replace("'", "''")
    return ($s -replace '([\[\]\*%])', '[$1]')
}

function Update-Filter {
    if (-not $script:Table) { $script:LblCount.Text = ''; return }
    $parts = @()
    $q = $script:TxtFilter.Text.Trim()
    if ($q) {
        $e = Get-LikeEscaped $q
        $cols = 'Gebruiker','Activiteit','Site','BronMap','BronBestand','DoelMap','DoelBestand','Item','ClientIP'
        $parts += '(' + (($cols | ForEach-Object { "$_ LIKE '%$e%'" }) -join ' OR ') + ')'
    }
    if ($script:CmbOpFilter.SelectedIndex -gt 0) {
        $parts += "Activiteit = '$(([string]$script:CmbOpFilter.SelectedItem).Replace("'", "''"))'"
    }
    try { $script:Table.DefaultView.RowFilter = ($parts -join ' AND ') } catch { $script:Table.DefaultView.RowFilter = '' }
    $script:LblCount.Text = "$($script:Table.DefaultView.Count) van $($script:Table.Rows.Count) regels"
}

function Get-SelectedRow {
    $rv = $script:Grid.SelectedItem
    if ($rv -is [System.Data.DataRowView]) { return $rv.Row }
    return $null
}

function Show-Detail {
    $r = Get-SelectedRow
    if (-not $r) { $script:TxtDetail.Text = ''; return }
    $sb = New-Object System.Text.StringBuilder
    $src = Get-ItemUrl $r 'Bron'
    $dst = Get-ItemUrl $r 'Doel'
    [void]$sb.AppendLine("Tijd (lokaal): $($r.Tijd)    UTC: $($r.TijdUTC)")
    [void]$sb.AppendLine("Gebruiker:     $($r.Gebruiker)")
    [void]$sb.AppendLine("Activiteit:    $($r.Activiteit)")
    if ($src) { [void]$sb.AppendLine("Bron:          $src/$($r.BronBestand)") }
    if ($dst) { [void]$sb.AppendLine("Doel:          $dst/$($r.DoelBestand)") }
    [void]$sb.AppendLine('')
    try { [void]$sb.Append(($r.Raw | ConvertFrom-Json | ConvertTo-Json -Depth 10)) } catch { [void]$sb.Append($r.Raw) }
    $script:TxtDetail.Text = $sb.ToString()
}

function Get-ItemUrl {
    param($Row, [string]$Which)
    $site = ([string]$Row.Site).TrimEnd('/')
    $rel  = if ($Which -eq 'Doel') { [string]$Row.DoelMap } else { [string]$Row.BronMap }
    if (-not $site -or -not $rel) { return $null }
    return "$site/$($rel.Trim('/'))"
}

function Open-Folder {
    param([string]$Which)
    $r = Get-SelectedRow
    if (-not $r) { Show-Msg 'Selecteer eerst een regel.' 'Openen' | Out-Null; return }
    $u = Get-ItemUrl $r $Which
    if (-not $u -and $Which -eq 'Bron' -and $r.Item) { $u = [string]$r.Item; $u = $u.Substring(0, [math]::Max(0, $u.LastIndexOf('/'))) }
    if (-not $u) { Show-Msg "Deze regel heeft geen $($Which.ToLower())map." 'Openen' | Out-Null; return }
    try { Start-Process $u } catch { Write-Log "Openen mislukt: $u" 'error' }
}

function Get-ExportColumns {
    $cols = @()
    foreach ($c in $script:Grid.Columns) {
        $path = $null
        try { $path = $c.Binding.Path.Path } catch { }
        if ($path) { $cols += @{ Header = [string]$c.Header; Path = $path } }
    }
    return $cols
}

function Export-Results {
    if (-not $script:Table -or $script:Table.DefaultView.Count -eq 0) { Show-Msg 'Geen resultaten om te exporteren.' 'Export' | Out-Null; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (puntkomma)|*.csv'
    $org = ([string]$script:Sync.ConnOrg) -replace '\.onmicrosoft\.com$', '' -replace '[^\w\.-]', '_'
    $dlg.FileName = "Audit_$(if ($org) { "$($org)_" })$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if (-not $dlg.ShowDialog($script:Window)) { return }
    try {
        $cols = Get-ExportColumns
        $rows = foreach ($rv in $script:Table.DefaultView) {
            $o = [ordered]@{}
            foreach ($c in $cols) { $o[$c.Header] = [string]$rv.Row[$c.Path] }
            [pscustomobject]$o
        }
        $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Write-Log "Export: $($script:Table.DefaultView.Count) regels naar $($dlg.FileName)" 'success'
        Set-Status 'Export klaar' 'success'
    } catch {
        Write-Log "Export mislukt: $($_.Exception.Message)" 'error'
    }
}

function Copy-Selection {
    $sel = @($script:Grid.SelectedItems)
    if (-not $sel.Count) { return }
    $cols = Get-ExportColumns
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine((($cols | ForEach-Object { $_.Header }) -join "`t"))
    foreach ($rv in $sel) { [void]$sb.AppendLine((($cols | ForEach-Object { [string]$rv.Row[$_.Path] }) -join "`t")) }
    [System.Windows.Clipboard]::SetText($sb.ToString())
    Write-Log "$($sel.Count) regel(s) gekopieerd" 'success'
}

function Set-Preset {
    param([string]$Name)
    $set = if ($Name -and $script:Presets.Contains($Name)) { $script:Presets[$Name] } else { @() }
    foreach ($k in $script:OpChecks.Keys) { $script:OpChecks[$k].IsChecked = ($set -contains $k) }
}

function Set-Range {
    param([int]$Days)
    $script:DpTo.SelectedDate = (Get-Date).Date
    $script:DpFrom.SelectedDate = (Get-Date).Date.AddDays(-($Days - 1))
}

#endregion

#region --- XAML ---

$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SharePoint Audit" Width="1400" Height="920" MinWidth="1000" MinHeight="680"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>
<Style TargetType="Button">
  <Setter Property="Foreground" Value="White"/>
  <Setter Property="Background" Value="#0078D4"/>
  <Setter Property="BorderThickness" Value="0"/>
  <Setter Property="Padding" Value="12,5"/>
  <Setter Property="Margin" Value="0,0,8,6"/>
  <Setter Property="Cursor" Value="Hand"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="Button">
        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.85"/></Trigger>
          <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.7"/></Trigger>
          <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.4"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>
<Style x:Key="BtnGreen" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#107C10"/></Style>
<Style x:Key="BtnRed"   TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#C42B1C"/></Style>
<Style x:Key="BtnGray"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#7A7A7A"/></Style>
<Style x:Key="BtnPreset" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="#E1ECF7"/><Setter Property="Foreground" Value="#004E8C"/><Setter Property="Padding" Value="8,3"/><Setter Property="Margin" Value="0,0,6,6"/>
</Style>
<Style x:Key="BtnDark"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="#005A9E"/><Setter Property="Padding" Value="14,7"/><Setter Property="Margin" Value="8,0,0,0"/>
</Style>
<Style TargetType="DataGrid">
  <Setter Property="AutoGenerateColumns" Value="False"/>
  <Setter Property="IsReadOnly" Value="True"/>
  <Setter Property="CanUserAddRows" Value="False"/>
  <Setter Property="CanUserDeleteRows" Value="False"/>
  <Setter Property="CanUserResizeRows" Value="False"/>
  <Setter Property="HeadersVisibility" Value="Column"/>
  <Setter Property="GridLinesVisibility" Value="Horizontal"/>
  <Setter Property="HorizontalGridLinesBrush" Value="#E6E6E6"/>
  <Setter Property="AlternatingRowBackground" Value="#F3F8FE"/>
  <Setter Property="Background" Value="White"/>
  <Setter Property="BorderBrush" Value="#DDDDDD"/>
  <Setter Property="BorderThickness" Value="1"/>
  <Setter Property="SelectionUnit" Value="FullRow"/>
  <Setter Property="SelectionMode" Value="Extended"/>
  <Setter Property="MinRowHeight" Value="24"/>
  <Setter Property="EnableRowVirtualization" Value="True"/>
  <Setter Property="EnableColumnVirtualization" Value="True"/>
</Style>
<Style TargetType="DataGridColumnHeader">
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="Padding" Value="6,5"/>
  <Setter Property="Background" Value="#F0F0F0"/>
  <Setter Property="BorderBrush" Value="#DDDDDD"/>
  <Setter Property="BorderThickness" Value="0,0,1,1"/>
</Style>
<Style x:Key="CellText" TargetType="TextBlock">
  <Setter Property="Margin" Value="6,0"/>
  <Setter Property="VerticalAlignment" Value="Center"/>
  <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
  <Setter Property="ToolTip" Value="{Binding Text, RelativeSource={RelativeSource Self}}"/>
</Style>
<Style x:Key="CellOp" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="Foreground" Value="#333333"/>
  <Style.Triggers>
    <Trigger Property="Text" Value="FileMoved"><Setter Property="Foreground" Value="#0063B1"/></Trigger>
    <Trigger Property="Text" Value="FolderMoved"><Setter Property="Foreground" Value="#0063B1"/></Trigger>
    <Trigger Property="Text" Value="FileRenamed"><Setter Property="Foreground" Value="#8A2BE2"/></Trigger>
    <Trigger Property="Text" Value="FolderRenamed"><Setter Property="Foreground" Value="#8A2BE2"/></Trigger>
    <Trigger Property="Text" Value="FileCopied"><Setter Property="Foreground" Value="#107C10"/></Trigger>
    <Trigger Property="Text" Value="FolderCopied"><Setter Property="Foreground" Value="#107C10"/></Trigger>
    <Trigger Property="Text" Value="FileRecycled"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="FolderRecycled"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="FileDeleted"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="FolderDeleted"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="FileRestored"><Setter Property="Foreground" Value="#B26200"/></Trigger>
    <Trigger Property="Text" Value="FolderRestored"><Setter Property="Foreground" Value="#B26200"/></Trigger>
  </Style.Triggers>
</Style>
<Style TargetType="GroupBox">
  <Setter Property="Margin" Value="0,0,0,6"/>
  <Setter Property="Padding" Value="8,6"/>
  <Setter Property="BorderBrush" Value="#CCCCCC"/>
  <Setter Property="Background" Value="White"/>
</Style>
<Style TargetType="CheckBox"><Setter Property="Margin" Value="0,0,14,6"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
<Style x:Key="Lbl" TargetType="TextBlock"><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="Margin" Value="0,0,6,6"/><Setter Property="Foreground" Value="#333333"/></Style>
<Style x:Key="FieldBox" TargetType="TextBox">
  <Setter Property="Height" Value="26"/>
  <Setter Property="VerticalContentAlignment" Value="Center"/>
  <Setter Property="Margin" Value="0,0,12,6"/>
  <Setter Property="Padding" Value="4,0"/>
</Style>
<Style x:Key="Sep" TargetType="Border"><Setter Property="Width" Value="1"/><Setter Property="Height" Value="22"/><Setter Property="Background" Value="#CCCCCC"/><Setter Property="Margin" Value="4,0,12,6"/></Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*" MinHeight="300"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="130" MinHeight="50"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0078D4">
      <DockPanel Margin="16,0,12,0">
        <StackPanel x:Name="HeaderPanel" DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Admin-account:" Foreground="White" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="TxtAdmin" Width="220" Height="26" VerticalContentAlignment="Center"
                   ToolTip="Je aparte admin-account. Wordt als hint in het inlogscherm gezet."/>
          <TextBlock Text="Tenant (optioneel, GDAP):" Foreground="White" VerticalAlignment="Center" Margin="12,0,6,0"/>
          <TextBox x:Name="TxtTenant" Width="190" Height="26" VerticalContentAlignment="Center"
                   ToolTip="Leeg = thuis-tenant van het admin-account. Klant via GDAP: klant.onmicrosoft.com of tenant-ID."/>
          <CheckBox x:Name="ChkNoWam" Content="Browser-login (geen WAM)" Foreground="White" Margin="12,0,0,0"
                    ToolTip="Gebruik dit als het Windows-inlogvenster (WAM) je eigen account pakt of niet opent."/>
          <TextBlock x:Name="LblConn" Text="Niet verbonden" Foreground="#FFC864" VerticalAlignment="Center" Margin="14,0,0,0" MaxWidth="300" TextTrimming="CharacterEllipsis"/>
          <Button x:Name="BtnConnect" Content="Inloggen" Style="{StaticResource BtnDark}"/>
          <Button x:Name="BtnDisconnect" Content="Uitloggen" Style="{StaticResource BtnDark}" IsEnabled="False"/>
        </StackPanel>
        <TextBlock Text="SharePoint Audit" Foreground="White" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </DockPanel>
    </Border>

    <GroupBox Grid.Row="1" Header="Zoeken in het audit log" Margin="8,8,8,0">
      <StackPanel>
        <StackPanel x:Name="SearchArea" IsEnabled="False">
          <WrapPanel>
            <TextBlock Text="Van:" Style="{StaticResource Lbl}"/>
            <DatePicker x:Name="DpFrom" Width="125" Margin="0,0,10,6"/>
            <TextBlock Text="t/m:" Style="{StaticResource Lbl}"/>
            <DatePicker x:Name="DpTo" Width="125" Margin="0,0,10,6"/>
            <Button x:Name="BtnR1"  Content="Vandaag"  Style="{StaticResource BtnPreset}"/>
            <Button x:Name="BtnR7"  Content="7 dagen"  Style="{StaticResource BtnPreset}"/>
            <Button x:Name="BtnR30" Content="30 dagen" Style="{StaticResource BtnPreset}"/>
            <Button x:Name="BtnR90" Content="90 dagen" Style="{StaticResource BtnPreset}"/>
            <Border Style="{StaticResource Sep}"/>
            <TextBlock Text="Recordtype:" Style="{StaticResource Lbl}"/>
            <ComboBox x:Name="CmbRecord" Width="210" Height="26" Margin="0,0,12,6">
              <ComboBoxItem Content="SharePointFileOperation" Tag="SharePointFileOperation"/>
              <ComboBoxItem Content="SharePoint" Tag="SharePoint"/>
              <ComboBoxItem Content="SharePointSharingOperation" Tag="SharePointSharingOperation"/>
              <ComboBoxItem Content="OneDrive" Tag="OneDrive"/>
              <ComboBoxItem Content="(alle recordtypes)" Tag=""/>
            </ComboBox>
            <TextBlock Text="Ophalen:" Style="{StaticResource Lbl}"/>
            <ComboBox x:Name="CmbChunk" Width="120" Height="26" Margin="0,0,12,6"
                      ToolTip="Grote periodes in blokken ophalen. Per blok max. 50.000 resultaten; bij overschrijding wordt automatisch gesplitst.">
              <ComboBoxItem Content="Per dag"/>
              <ComboBoxItem Content="Per 6 uur"/>
              <ComboBoxItem Content="Per uur"/>
              <ComboBoxItem Content="Hele periode"/>
            </ComboBox>
          </WrapPanel>
          <WrapPanel>
            <TextBlock Text="Snelkeuze:" Style="{StaticResource Lbl}" FontWeight="SemiBold"/>
            <WrapPanel x:Name="PnlPresets"/>
          </WrapPanel>
          <Border Background="#F7F9FC" BorderBrush="#E3E8EF" BorderThickness="1" CornerRadius="4" Padding="8,6,8,0" Margin="0,0,0,6">
            <WrapPanel x:Name="PnlOps"/>
          </Border>
          <WrapPanel>
            <TextBlock Text="Extra activiteiten:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtExtraOps" Width="200" Style="{StaticResource FieldBox}" ToolTip="Andere activiteitnamen, komma-gescheiden (bijv. SharingSet,AnonymousLinkCreated)"/>
            <TextBlock Text="Sites bevatten:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtSites" Width="230" Style="{StaticResource FieldBox}" ToolTip="Komma-gescheiden delen van de site-URL, bijv. Drachten, Heerenveen, Finance. Leeg = alle sites."/>
            <CheckBox x:Name="ChkRegex" Content="regex" ToolTip="Sites-veld als reguliere expressie gebruiken"/>
            <TextBlock Text="Bestand/map bevat:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtNames" Width="170" Style="{StaticResource FieldBox}" ToolTip="Komma-gescheiden; zoekt in bron/doel bestandsnaam en pad"/>
            <TextBlock Text="Gebruikers:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtUsers" Width="200" Style="{StaticResource FieldBox}" ToolTip="UPN's, komma-gescheiden. Wordt server-side gefilterd (sneller). Leeg = iedereen."/>
          </WrapPanel>
        </StackPanel>
        <WrapPanel>
          <Button x:Name="BtnSearch" Content="  Zoeken  " Style="{StaticResource BtnGreen}" IsEnabled="False" FontWeight="SemiBold"/>
          <Button x:Name="BtnStop" Content="Stoppen" Style="{StaticResource BtnRed}" IsEnabled="False"/>
          <ProgressBar x:Name="Progress" Width="260" Height="14" Margin="4,0,0,6" Visibility="Collapsed" IsIndeterminate="True"/>
        </WrapPanel>
      </StackPanel>
    </GroupBox>

    <Grid Grid.Row="2" Margin="8,4,8,0">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*" MinHeight="120"/>
        <RowDefinition Height="5"/>
        <RowDefinition Height="170" MinHeight="40"/>
      </Grid.RowDefinitions>
      <WrapPanel Grid.Row="0">
        <TextBlock Text="Filter resultaten:" Style="{StaticResource Lbl}"/>
        <TextBox x:Name="TxtFilter" Width="240" Style="{StaticResource FieldBox}" ToolTip="Zoekt in gebruiker, activiteit, site, paden, bestandsnamen en IP"/>
        <ComboBox x:Name="CmbOpFilter" Width="190" Height="26" Margin="0,0,12,6"/>
        <Button x:Name="BtnOpenDst" Content="Open doelmap" ToolTip="Map waar het item naartoe is verplaatst/gekopieerd (dubbelklik op een regel doet hetzelfde)"/>
        <Button x:Name="BtnOpenSrc" Content="Open bronmap" Style="{StaticResource BtnGray}"/>
        <Button x:Name="BtnCopy" Content="Kopieer selectie" Style="{StaticResource BtnGray}"/>
        <Button x:Name="BtnExport" Content="Export CSV" Style="{StaticResource BtnGray}"/>
        <TextBlock x:Name="LblCount" Style="{StaticResource Lbl}" Foreground="Gray" Margin="6,0,0,6"/>
      </WrapPanel>
      <TextBlock Grid.Row="1" x:Name="LblSummary" Foreground="#005A9E" TextWrapping="Wrap" Margin="0,0,0,6"/>
      <DataGrid Grid.Row="2" x:Name="Grid">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Tijd (lokaal)" Binding="{Binding Tijd}" Width="135" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Tijd UTC" Binding="{Binding TijdUTC}" Width="135" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Gebruiker" Binding="{Binding Gebruiker}" Width="190" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Activiteit" Binding="{Binding Activiteit}" Width="120" ElementStyle="{StaticResource CellOp}"/>
          <DataGridTextColumn Header="Site" Binding="{Binding Site}" Width="220" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Bron map" Binding="{Binding BronMap}" Width="200" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Bron bestand" Binding="{Binding BronBestand}" Width="160" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Doel map" Binding="{Binding DoelMap}" Width="200" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Doel bestand" Binding="{Binding DoelBestand}" Width="160" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Client-IP" Binding="{Binding ClientIP}" Width="110" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Workload" Binding="{Binding Workload}" Width="85" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Programma" Binding="{Binding Programma}" Width="160" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Item" Binding="{Binding Item}" Width="300" ElementStyle="{StaticResource CellText}"/>
        </DataGrid.Columns>
      </DataGrid>
      <GridSplitter Grid.Row="3" HorizontalAlignment="Stretch" Background="#DDDDDD"/>
      <TextBox Grid.Row="4" x:Name="TxtDetail" IsReadOnly="True" FontFamily="Consolas" FontSize="11.5" TextWrapping="NoWrap"
               VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Background="White" BorderBrush="#DDDDDD"/>
    </Grid>

    <GridSplitter Grid.Row="3" HorizontalAlignment="Stretch" Background="#DDDDDD" Margin="0,4,0,0"/>
    <RichTextBox Grid.Row="4" x:Name="LogBox" IsReadOnly="True" Background="#141414" Foreground="#C8C8C8" BorderThickness="0"
                 FontFamily="Consolas" FontSize="11.5" VerticalScrollBarVisibility="Auto" Padding="6,4">
      <FlowDocument PagePadding="0"><Paragraph x:Name="LogPara" Margin="0"/></FlowDocument>
    </RichTextBox>
    <Border x:Name="StatusBar" Grid.Row="5" Background="#0078D4">
      <TextBlock x:Name="LblStatus" Text="Niet verbonden" Foreground="White" VerticalAlignment="Center" Margin="12,0"/>
    </Border>
  </Grid>
</Window>
'@

#endregion

#region --- Opbouw venster ---

$script:Window = [Windows.Markup.XamlReader]::Parse($script:MainXaml)
$ctl = Get-XamlControls $script:Window $script:MainXaml
foreach ($k in $ctl.Keys) { Set-Variable -Scope Script -Name $k -Value $ctl[$k] }

# Activiteit-vinkjes en snelkeuzes
foreach ($op in $script:AllOps) {
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content = $op
    $cb.Width = 225
    [void]$script:PnlOps.Children.Add($cb)
    $script:OpChecks[$op] = $cb
}
$presetStyle = $script:Window.FindResource('BtnPreset')
foreach ($name in @($script:Presets.Keys) + @('Geen')) {
    $b = New-Object System.Windows.Controls.Button
    $b.Content = $name
    $b.Tag = $name
    $b.Style = $presetStyle
    $b.Add_Click({ param($s, $e) Set-Preset ([string]$s.Tag) })
    [void]$script:PnlPresets.Children.Add($b)
}

# Standaardwaarden + opgeslagen instellingen
$script:CmbRecord.SelectedIndex = 0
$script:CmbChunk.SelectedIndex  = 0
$script:CmbOpFilter.Items.Add('(alle activiteiten)') | Out-Null
$script:CmbOpFilter.SelectedIndex = 0
Set-Range 7
Set-Preset 'Verplaatsen/verwijderen'
$cfg = Get-Settings
if ($cfg) {
    if ($cfg.AdminUpn) { $script:TxtAdmin.Text  = $cfg.AdminUpn }
    if ($cfg.Tenant)   { $script:TxtTenant.Text = $cfg.Tenant }
    $script:ChkNoWam.IsChecked = [bool]$cfg.NoWam
    if ($cfg.Sites)    { $script:TxtSites.Text  = $cfg.Sites }
    $script:ChkRegex.IsChecked = [bool]$cfg.SiteRegex
    if ($null -ne $cfg.Operations) {
        $saved = @($cfg.Operations)
        foreach ($k in $script:OpChecks.Keys) { $script:OpChecks[$k].IsChecked = ($saved -contains $k) }
    }
    if ($null -ne $cfg.Chunk -and [int]$cfg.Chunk -ge 0 -and [int]$cfg.Chunk -lt 4) { $script:CmbChunk.SelectedIndex = [int]$cfg.Chunk }
    if ($null -ne $cfg.RecordType -and [int]$cfg.RecordType -ge 0 -and [int]$cfg.RecordType -lt 5) { $script:CmbRecord.SelectedIndex = [int]$cfg.RecordType }
}

# Events
$script:BtnConnect.Add_Click({ Invoke-Connect })
$script:BtnDisconnect.Add_Click({ Invoke-Disconnect })
$script:TxtAdmin.Add_KeyDown({ param($s, $e) if ($e.Key -eq 'Return') { Invoke-Connect } })
$script:BtnSearch.Add_Click({ Invoke-Search })
$script:BtnStop.Add_Click({ Stop-Search })
$script:BtnR1.Add_Click({ Set-Range 1 })
$script:BtnR7.Add_Click({ Set-Range 7 })
$script:BtnR30.Add_Click({ Set-Range 30 })
$script:BtnR90.Add_Click({ Set-Range 90 })
$script:TxtFilter.Add_TextChanged({ Update-Filter })
$script:CmbOpFilter.Add_SelectionChanged({ Update-Filter })
$script:Grid.Add_SelectionChanged({ Show-Detail })
$script:Grid.Add_MouseDoubleClick({ Open-Folder 'Doel' })
$script:BtnOpenDst.Add_Click({ Open-Folder 'Doel' })
$script:BtnOpenSrc.Add_Click({ Open-Folder 'Bron' })
$script:BtnCopy.Add_Click({ Copy-Selection })
$script:BtnExport.Add_Click({ Export-Results })

$script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$script:Timer.Interval = [TimeSpan]::FromMilliseconds(250)
$script:Timer.Add_Tick({ Invoke-Tick })

$script:Window.Add_Loaded({
    Initialize-Worker
    $script:Timer.Start()
    Update-Enabled
    Write-Log 'SharePoint Audit gestart. Vul je admin-account in en klik Inloggen.'
    Write-Log "Benodigde rol: 'View-Only Audit Logs' of 'Audit Logs' (bij GDAP: in de relatie)."
    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { Write-Log 'Module ExchangeOnlineManagement niet gevonden - wordt aangeboden bij inloggen.' 'warning' }
})

$script:Window.Add_Closing({
    Save-Settings
    $script:Timer.Stop()
    try {
        if ($script:Job) { $script:Sync.Cancel = $true; $script:Job.PS.Stop() }
        if ($script:Connected -and $script:Rs) {
            $ps = [powershell]::Create(); $ps.Runspace = $script:Rs
            [void]$ps.AddScript($script:DisconnectScript)
            $h = $ps.BeginInvoke()
            [void]$h.AsyncWaitHandle.WaitOne(5000)
        }
    } catch { }
    try { $script:Rs.Close() } catch { }
})

[void]$script:Window.ShowDialog()

#endregion
