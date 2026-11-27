#Requires -Version 5.1
<#
.SYNOPSIS
    Teams Audit (WPF) - zoek in het Unified Audit Log naar Microsoft Teams-
    activiteit: wie voegt wie toe aan / verwijdert wie uit een chat, team of
    kanaal, wie verlaat een chat, berichten verwijderd/bewerkt, teams en
    kanalen aangemaakt/verwijderd, instellingen, apps/bots/tabs, vergaderingen,
    beheer- en beveiligingsgebeurtenissen.
.DESCRIPTION
    - Inloggen via Connect-ExchangeOnline met je aparte admin-account (UPN als
      hint). Optioneel een klant-tenant via GDAP (DelegatedOrganization).
    - Zoekperiode met datum en tijd, activiteiten per categorie (vinkjes en
      snelkeuzes), uitvoerders (server-side), betrokkene en team/chat-naam
      (client-side).
    - Optioneel ook Entra/M365-groepslidmaatschap (Add/Remove member to group),
      omdat lidmaatschapswijzigingen via Entra/M365-admin in het Teams-log een
      teameigenaar als uitvoerder tonen in plaats van de echte uitvoerder.
    - Grote periodes in blokken; bij 50.000 resultaten wordt automatisch
      gesplitst. Dedupe op record-ID.
    - Tabbladen: Alle gebeurtenissen (met Nederlandse omschrijving),
      Lidmaatschap (een regel per betrokken persoon: wie deed wat bij wie),
      Overzicht (per uitvoerder, activiteit, team/chat en betrokkene).
    - Filteren, rechtsklikmenu (filter op uitvoerder/chat/betrokkene),
      details (volledige AuditData), kopieren, CSV-export, HTML-rapport.
.NOTES
    Vereist: Windows PowerShell 5.1, module ExchangeOnlineManagement (3.x).
    Rol:     'View-Only Audit Logs' of 'Audit Logs'. Bij GDAP: rol in de relatie.
    Bewaartermijn audit log: standaard 180 dagen.
    Let op: inhoud van berichten staat NIET in het audit log (alleen metadata).
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

$script:AppName      = 'Teams Audit'
$script:SettingsDir  = Join-Path $env:LOCALAPPDATA 'TeamsAuditTool'
$script:SettingsFile = Join-Path $script:SettingsDir 'settings.json'
$script:Window       = $null
$script:Connected    = $false
$script:ConnLabel    = ''
$script:Job          = $null
$script:Table        = $null
$script:Members      = $null
$script:OpChecks     = [ordered]@{}
$script:CatChecks    = @{}
$script:LastSearch   = $null
$script:Domains      = @()

$script:Sync = [hashtable]::Synchronized(@{
    Log     = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    Cancel  = $false
    Status  = ''
    Pct     = -1
    Table   = $null
    Members = $null
    Summary = ''
})

# Activiteiten per categorie (operation-namen zoals in Purview)
$script:OpCats = [ordered]@{
    'Lidmaatschap'        = @('MemberAdded','MemberRemoved','MemberRoleChanged')
    'Teams en kanalen'    = @('TeamCreated','TeamDeleted','TeamSettingChanged','ChannelAdded','ChannelDeleted','ChannelSettingChanged')
    'Gedeelde kanalen'    = @('InviteSent','InviteeResponded','ChannelOwnerResponded','TerminatedSharing','SharingRestored')
    'Chats en berichten'  = @('ChatCreated','ChatUpdated','MessageSent','MessageUpdated','MessageDeleted','GraphMessageUpdated','GraphMessageUndeleted','MessageCreatedHasLink','MessageEditedHasLink','ReactedToMessage')
    'Lezen en ophalen'    = @('ChatRetrieved','MessageRead','MessagesListed','MessageHostedContentRead','MessageHostedContentsListed','MessagesExported','SubscribedToMessages')
    'Vergaderingen'       = @('MeetingDetail','MeetingParticipantDetail','CallParticipantDetail','SensitivityLabelApplied','SensitivityLabelChanged','SensitivityLabelRemoved','SensitiveContentShared')
    'Apps, bots en tabs'  = @('AppInstalled','AppUpgraded','AppUninstalled','BotAddedToTeam','BotRemovedFromTeam','ConnectorAdded','ConnectorRemoved','ConnectorUpdated','TabAdded','TabRemoved','TabUpdated','AppPublishedToCatalog','AppUpdatedInCatalog','AppDeletedFromCatalog')
    'Beheer en beveiliging' = @('TeamsTenantSettingChanged','TeamsAdminAction','UserBlocked','UserUnblocked','UserAccepted','TeamsImpersonationDetected','SecurityRiskInMeetingDetected','SecurityRiskInCallDetected','CreateThreadProbe','TeamsSessionStarted')
}

$script:Friendly = @{
    MemberAdded = 'Leden toegevoegd'; MemberRemoved = 'Leden verwijderd'; MemberRoleChanged = 'Rol van lid gewijzigd'
    TeamCreated = 'Team aangemaakt'; TeamDeleted = 'Team verwijderd'; TeamSettingChanged = 'Teaminstelling gewijzigd'
    ChannelAdded = 'Kanaal toegevoegd'; ChannelDeleted = 'Kanaal verwijderd'; ChannelSettingChanged = 'Kanaalinstelling gewijzigd'
    InviteSent = 'Uitnodiging gedeeld kanaal verstuurd'; InviteeResponded = 'Gereageerd op uitnodiging gedeeld kanaal'
    ChannelOwnerResponded = 'Kanaaleigenaar reageerde op uitnodiging'; TerminatedSharing = 'Delen van kanaal beeindigd'; SharingRestored = 'Delen van kanaal hersteld'
    ChatCreated = 'Chat aangemaakt'; ChatUpdated = 'Chat bijgewerkt'; ChatRetrieved = 'Chat opgehaald'
    MessageSent = 'Bericht verstuurd'; MessageUpdated = 'Bericht bewerkt'; MessageDeleted = 'Bericht verwijderd'
    GraphMessageUpdated = 'Bericht bewerkt (Graph)'; GraphMessageUndeleted = 'Verwijderd bericht hersteld'
    MessageCreatedHasLink = 'Bericht met link verstuurd'; MessageEditedHasLink = 'Bericht met link bewerkt'; ReactedToMessage = 'Reactie op bericht'
    MessageRead = 'Bericht gelezen'; MessagesListed = 'Berichten opgehaald'; MessageHostedContentRead = 'Inhoud in bericht gelezen'
    MessageHostedContentsListed = 'Inhoud in berichten opgehaald'; MessagesExported = 'Berichten geexporteerd'; SubscribedToMessages = 'Geabonneerd op berichten'
    MeetingDetail = 'Vergaderdetails'; MeetingParticipantDetail = 'Deelnemers vergadering'; CallParticipantDetail = 'Deelnemers gesprek'
    SensitivityLabelApplied = 'Gevoeligheidslabel toegepast'; SensitivityLabelChanged = 'Gevoeligheidslabel gewijzigd'; SensitivityLabelRemoved = 'Gevoeligheidslabel verwijderd'
    SensitiveContentShared = 'Gevoelige inhoud gedeeld in scherm'
    AppInstalled = 'App geinstalleerd'; AppUpgraded = 'App bijgewerkt'; AppUninstalled = 'App verwijderd'
    BotAddedToTeam = 'Bot toegevoegd'; BotRemovedFromTeam = 'Bot verwijderd'
    ConnectorAdded = 'Connector toegevoegd'; ConnectorRemoved = 'Connector verwijderd'; ConnectorUpdated = 'Connector bijgewerkt'
    TabAdded = 'Tab toegevoegd'; TabRemoved = 'Tab verwijderd'; TabUpdated = 'Tab bijgewerkt'
    AppPublishedToCatalog = 'App gepubliceerd in catalogus'; AppUpdatedInCatalog = 'App bijgewerkt in catalogus'; AppDeletedFromCatalog = 'App verwijderd uit catalogus'
    TeamsTenantSettingChanged = 'Tenantinstelling Teams gewijzigd'; TeamsAdminAction = 'Teams-beheeractie'
    UserBlocked = 'Gebruiker geblokkeerd'; UserUnblocked = 'Gebruiker gedeblokkeerd'; UserAccepted = 'Bericht van onbekende geaccepteerd'
    TeamsImpersonationDetected = 'Imitatiepoging gedetecteerd'; SecurityRiskInMeetingDetected = 'Beveiligingsrisico in vergadering'
    SecurityRiskInCallDetected = 'Beveiligingsrisico in gesprek'; CreateThreadProbe = 'Chat-aanmaak geprobeerd (probe)'; TeamsSessionStarted = 'Teams-sessie gestart'
    'Add member to group.' = 'Lid toegevoegd aan groep (Entra)'; 'Remove member from group.' = 'Lid verwijderd uit groep (Entra)'
    'Add owner to group.' = 'Eigenaar toegevoegd aan groep (Entra)'; 'Remove owner from group.' = 'Eigenaar verwijderd uit groep (Entra)'
}

$script:Presets = [ordered]@{
    'Wie verwijdert wie'   = @('MemberRemoved','MemberAdded','MemberRoleChanged')
    'Alles verwijderd'     = @('MemberRemoved','MessageDeleted','TeamDeleted','ChannelDeleted','AppUninstalled','BotRemovedFromTeam','ConnectorRemoved','TabRemoved','TerminatedSharing')
    'Teams/kanalen beheer' = @('MemberAdded','MemberRemoved','MemberRoleChanged','TeamCreated','TeamDeleted','TeamSettingChanged','ChannelAdded','ChannelDeleted','ChannelSettingChanged')
    'Berichten'            = @('ChatCreated','ChatUpdated','MessageSent','MessageUpdated','MessageDeleted','GraphMessageUpdated','GraphMessageUndeleted','MessageCreatedHasLink','MessageEditedHasLink')
    'Vergaderingen'        = @('MeetingDetail','MeetingParticipantDetail','CallParticipantDetail')
    'Apps/bots/tabs'       = $script:OpCats['Apps, bots en tabs']
    'Beheer/beveiliging'   = $script:OpCats['Beheer en beveiliging']
    'Alles behalve lezen'  = @($script:OpCats.Keys | Where-Object { $_ -notin 'Lezen en ophalen','Vergaderingen' } | ForEach-Object { $script:OpCats[$_] })
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
            Operations = $ops
            ExtraOps   = $script:TxtExtraOps.Text
            Chunk      = $script:CmbChunk.SelectedIndex
            RecordType = $script:CmbRecord.SelectedIndex
            Entra      = [bool]$script:ChkEntra.IsChecked
            OnlyTeams  = [bool]$script:ChkOnlyTeams.IsChecked
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
        default   { New-Brush '#464EB8' }
    }
}

function Update-Enabled {
    $busy = ($null -ne $script:Job)
    $script:HeaderPanel.IsEnabled   = -not $busy
    $script:SearchArea.IsEnabled    = $script:Connected -and -not $busy
    $script:BtnSearch.IsEnabled     = $script:Connected -and -not $busy
    $script:BtnStop.IsEnabled       = $busy -and ($script:Job.Kind -eq 'search')
    $script:BtnConnect.IsEnabled    = -not $busy
    $script:BtnDisconnect.IsEnabled = $script:Connected -and -not $busy
    $script:Progress.Visibility     = if ($busy) { 'Visible' } else { 'Collapsed' }
    $script:Window.Cursor           = if ($busy -and $script:Job.Kind -ne 'search') { [System.Windows.Input.Cursors]::AppStarting } else { $null }
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
$Sync.ConnUser = $null; $Sync.ConnOrg = $null; $Sync.AuditOn = $null; $Sync.HasSearch = $false; $Sync.Domains = @()
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
try {
    $Sync.Domains = @(Get-AcceptedDomain -ErrorAction Stop | ForEach-Object { ([string]$_.DomainName).ToLower() })
    L "Domeinen van de tenant: $($Sync.Domains -join ', ')"
} catch { L 'Domeinen niet op te halen; extern/gast wordt alleen herkend op rol en #EXT#' 'warning' }
'@

$script:DisconnectScript = @'
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop | Out-Null } catch { }
'@

$script:SearchScript = @'
param($P)
function L { param($m, $t = 'info') $Sync.Log.Enqueue([object[]]@($m, $t)) }

$inv = [System.Globalization.CultureInfo]::InvariantCulture
$utcStyle = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
$F = $P.Friendly
$domains = @($P.Domains)

function New-T {
    param([string]$Name, [string[]]$Cols, [string[]]$IntCols = @())
    $t = New-Object System.Data.DataTable $Name
    foreach ($c in $Cols) { if ($IntCols -contains $c) { [void]$t.Columns.Add($c, [int]) } else { [void]$t.Columns.Add($c, [string]) } }
    return ,$t
}

$ev = New-T 'Events' @('Tijd','TijdUTC','Gebruiker','Actie','Activiteit','Omschrijving','Type','TeamChat','Kanaal','Betrokkenen','Extern','Instelling','OudeWaarde','NieuweWaarde','App','Extra','BerichtId','ChatId','TeamId','ClientIP','UserType','Workload','RecordType','Raw')
$mb = New-T 'Leden' @('Tijd','Uitvoerder','Actie','Betrokkene','BetrokkeneUPN','Rol','Extern','Type','TeamChat','Kanaal','Bron','Activiteit','ChatId','TeamId','Raw')
$Sync.Table = $ev
$Sync.Members = $mb
$Sync.Summary = ''
$Sync.Stats = $null
$seen = New-Object 'System.Collections.Generic.HashSet[string]'
$opCount = @{}

function Test-Ext {
    param([string]$Upn, $Role)
    if ("$Role" -eq '2' -or "$Role" -eq 'Guest') { return $true }
    if ($Upn -match '#EXT#') { return $true }
    if ($domains.Count -and $Upn -match '@([^@]+)$') { return -not ($domains -contains $Matches[1].ToLower()) }
    return $false
}

function Get-RoleNl {
    param($r)
    switch ("$r") { '0' { 'Lid' } 'Member' { 'Lid' } '1' { 'Eigenaar' } 'Owner' { 'Eigenaar' } '2' { 'Gast' } 'Guest' { 'Gast' } default { "$r" } }
}

function Get-UserTypeNl {
    param($u)
    switch ("$u") { '0' { 'Gebruiker' } 'Regular' { 'Gebruiker' } '2' { 'Admin' } 'Admin' { 'Admin' } '3' { 'DC-admin' } '4' { 'Systeem' } 'System' { 'Systeem' } '5' { 'Applicatie' } 'Application' { 'Applicatie' } '6' { 'Service principal' } 'ServicePrincipal' { 'Service principal' } '9' { 'Partner (GDAP)' } 'PartnerTechnician' { 'Partner (GDAP)' } '10' { 'Gast' } 'Guest' { 'Gast' } default { "$u" } }
}

function Get-Val {
    param($v)
    if ($null -eq $v) { return '' }
    if ($v -is [string] -or $v -is [ValueType]) { return [string]$v }
    if ($v -is [array]) {
        $parts = foreach ($x in $v) { if ($x -is [string] -or $x -is [ValueType]) { [string]$x } else { ($x | ConvertTo-Json -Compress -Depth 4) } }
        return ($parts -join ', ')
    }
    return ($v | ConvertTo-Json -Compress -Depth 4)
}

function Cut { param([string]$s, [int]$n = 300) if ($s.Length -gt $n) { return $s.Substring(0, $n) + '...' }; return $s }

# Blokken maken: per query x tijdsblok
$queue = New-Object 'System.Collections.Generic.Queue[object]'
foreach ($qi in 0..($P.Queries.Count - 1)) {
    $s = $P.Start
    while ($s -lt $P.End) {
        $e = if ($P.ChunkHours -gt 0) { $s.AddHours($P.ChunkHours) } else { $P.End }
        if ($e -gt $P.End) { $e = $P.End }
        $queue.Enqueue([object[]]@($qi, $s, $e))
        $s = $e
    }
}
$totalBlocks = $queue.Count
$doneBlocks  = 0
$rawTotal    = 0
$failed      = 0
$skippedWl   = 0
L ("Zoeken {0:dd-MM-yyyy HH:mm} t/m {1:dd-MM-yyyy HH:mm} (lokale tijd) in {2} blok(ken), {3} zoekopdracht(en)" -f $P.Start, $P.End.AddSeconds(-1), $totalBlocks, $P.Queries.Count) 'api'

while ($queue.Count -gt 0) {
    if ($Sync.Cancel) { L 'Gestopt door gebruiker; resultaten tot nu toe blijven staan' 'warning'; break }
    $blk = $queue.Dequeue(); $q = $P.Queries[[int]$blk[0]]; $bs = $blk[1]; $be = $blk[2]
    $label = '{0} {1:dd-MM HH:mm} - {2:dd-MM HH:mm}' -f $q.Name, $bs, $be

    $prm = @{
        StartDate      = $bs.ToUniversalTime()
        EndDate        = $be.ToUniversalTime()
        SessionCommand = 'ReturnLargeSet'
        ResultSize     = 5000
        ErrorAction    = 'Stop'
    }
    if ($q.RecordType)  { $prm.RecordType = $q.RecordType }
    if ($q.Ops.Count)   { $prm.Operations = [string[]]$q.Ops }
    if ($P.Users.Count) { $prm.UserIds    = [string[]]$P.Users }

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

    if ($total -ge 50000 -and ($be - $bs).TotalMinutes -gt 20) {
        $mid = $bs.AddTicks([long](($be - $bs).Ticks / 2))
        $queue.Enqueue([object[]]@($blk[0], $bs, $mid)); $queue.Enqueue([object[]]@($blk[0], $mid, $be))
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
        $op = [string]$a.Operation
        $isEntra = ($q.Kind -eq 'Entra')
        if (-not $isEntra -and $P.OnlyTeams -and [string]$a.Workload -ne 'MicrosoftTeams') { $skippedWl++; continue }

        $tl = ''; $tu = [string]$a.CreationTime
        $dt = [datetime]::MinValue
        if ([datetime]::TryParse($tu, $inv, $utcStyle, [ref]$dt)) {
            $tu = $dt.ToString('yyyy-MM-dd HH:mm:ss')
            $tl = $dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
        }
        $actor = [string]$a.UserId
        $friendly = if ($F.ContainsKey($op)) { [string]$F[$op] } else { $op }

        # Doelpersonen verzamelen: @{Name; Upn; Role; Ext; Self}
        $targets = New-Object 'System.Collections.Generic.List[object]'
        $type = ''; $ctxName = ''; $channel = ''; $chatId = ''; $teamId = ''
        $setting = ''; $oldV = ''; $newV = ''; $app = ''; $msgId = ''; $extra = ''

        if ($isEntra) {
            $tgt = ''
            foreach ($t in @($a.Target)) { $tid = [string]$t.ID; if ($tid -match '@' -and -not $tgt) { $tgt = $tid } }
            if (-not $tgt) { $tgt = [string]$a.ObjectId }
            foreach ($mp in @($a.ModifiedProperties)) {
                if ([string]$mp.Name -eq 'Group.DisplayName') { $ctxName = ([string]$mp.NewValue).Trim('"', ' ') }
                if ([string]$mp.Name -eq 'Group.ObjectID')    { $teamId  = ([string]$mp.NewValue).Trim('"', ' ') }
            }
            if (-not $ctxName -and -not $teamId) { continue }
            $type = 'Groep (Entra/M365)'
            $role = if ($op -match 'owner') { '1' } else { '0' }
            $targets.Add(@{ Name = $tgt; Upn = $tgt; Role = (Get-RoleNl $role); Ext = (Test-Ext $tgt $role); Self = ($tgt -and $tgt -eq $actor) })
        } else {
            $msg = $null
            if ($a.Messages) { $msg = @($a.Messages)[0] }
            $pick = { param($n) $v = [string]$a.$n; if (-not $v -and $msg) { $v = [string]$msg.$n }; $v }
            $teamName = & $pick 'TeamName'
            $chatName = & $pick 'ChatName'
            $channel  = & $pick 'ChannelName'
            $chatId   = & $pick 'ChatThreadId'
            $teamId   = & $pick 'TeamGuid'
            if (-not $teamId) { $teamId = & $pick 'AADGroupId' }
            $ct = [string]$a.CommunicationType
            $type = switch ($ct) {
                'GroupChat'   { 'Groepschat' }
                'OneOnOne'    { '1-op-1 chat' }
                'OneOnOneChat'{ '1-op-1 chat' }
                'MeetingChat' { 'Vergaderchat' }
                'Channel'     { 'Kanaal' }
                'Team'        { if ($channel) { 'Kanaal' } else { 'Team' } }
                ''            {
                    if ($op -match 'Meeting|Call') { 'Vergadering/gesprek' }
                    elseif ($teamName) { if ($channel) { 'Kanaal' } else { 'Team' } }
                    elseif ($chatId -or $chatName) { 'Chat' }
                    elseif ($op -match 'Tenant|Admin|Catalog|Organization') { 'Tenant/beheer' }
                    else { '' }
                }
                default       { $ct }
            }
            $ctxName = if ($teamName) { $teamName } elseif ($chatName) { $chatName } elseif ([string]$a.ItemName) { [string]$a.ItemName } else { '' }
            if (-not $ctxName -and $chatId) { $ctxName = '(chat zonder naam) ' + $(if ($chatId.Length -gt 24) { $chatId.Substring(0, 24) + '...' } else { $chatId }) }

            foreach ($m in @($a.Members)) {
                if ($null -eq $m) { continue }
                $u = [string]$m.UPN
                $n = if ([string]$m.DisplayName) { [string]$m.DisplayName } else { $u }
                $targets.Add(@{ Name = $n; Upn = $u; Role = (Get-RoleNl $m.Role); Ext = (Test-Ext $u $m.Role); Self = ($u -and $u -eq $actor) })
            }
            if ($targets.Count -eq 0 -and $a.Attendees) {
                foreach ($m in @($a.Attendees)) {
                    if ($null -eq $m) { continue }
                    $u = [string]$m.UPN; if (-not $u) { $u = [string]$m.UserObjectId }
                    $n = if ([string]$m.DisplayName) { [string]$m.DisplayName } else { $u }
                    $targets.Add(@{ Name = $n; Upn = $u; Role = [string]$m.RecipientType; Ext = (Test-Ext $u $null); Self = $false })
                }
            }

            $setting = [string]$a.Name
            $oldV = Cut (Get-Val $a.OldValue)
            $newV = Cut (Get-Val $a.NewValue)
            $app = [string]$a.AddOnName
            if (-not $app -and $a.TabType) { $app = "Tab: $([string]$a.TabType)" }
            if (-not $app -and $a.AzureADAppId) { $app = "App-ID: $([string]$a.AzureADAppId)" }
            $msgId = [string]$a.MessageId
            if (-not $msgId -and $msg) { $msgId = [string]$msg.Id }

            $ex = New-Object 'System.Collections.Generic.List[string]'
            foreach ($k in 'ChannelType','JoinTime','LeaveTime','StartTime','EndTime','MeetingDetailId','MessageSizeInBytes','MessageURLs','DeviceInformation','ArtifactsShared','ParticipantInfo','ExtraProperties','ResourceTenantId') {
                $pv = $a.PSObject.Properties[$k]
                if ($pv -and $null -ne $pv.Value -and "$($pv.Value)" -ne '') { $ex.Add("$k=$(Cut (Get-Val $pv.Value) 200)") }
            }
            $extra = $ex -join '  |  '
        }

        # Client-side filters
        $tgtText = (($targets | ForEach-Object { "$($_.Name) $($_.Upn)" }) -join ' ')
        if ($P.TargetRx -and -not ($tgtText -match $P.TargetRx)) { continue }
        if ($P.CtxRx -and -not ("$ctxName $channel $chatId $teamId" -match $P.CtxRx)) { continue }

        # Omschrijving
        $who  = (($targets | ForEach-Object { if ($_.Upn -and $_.Name -ne $_.Upn) { "$($_.Name) <$($_.Upn)>" } else { $_.Name } }) -join ', ')
        $whoR = (($targets | ForEach-Object { "$($_.Name) -> $($_.Role)" }) -join ', ')
        $typeLbl = switch ($type) { 'Kanaal' { 'team' } 'Groep (Entra/M365)' { 'groep' } default { $type.ToLower() } }
        $where = if ($ctxName) { "$typeLbl '$ctxName'" } elseif ($type) { $type.ToLower() } else { '' }
        if ($channel -and $where -notmatch [regex]::Escape($channel)) { $where += " / kanaal '$channel'" }
        $actorShow = if ($actor) { $actor } else { '(onbekend)' }
        $allSelf = ($targets.Count -gt 0 -and @($targets | Where-Object { -not $_.Self }).Count -eq 0)
        $desc = switch -Regex ($op) {
            '^MemberAdded$|^Add member to group\.$|^Add owner to group\.$' { "$actorShow voegde $who toe aan $where" }
            '^MemberRemoved$|^Remove member from group\.$|^Remove owner from group\.$' {
                if ($allSelf) { "$actorShow heeft $where verlaten" } else { "$actorShow verwijderde $who uit $where" }
            }
            '^MemberRoleChanged$' { "$actorShow wijzigde de rol van $whoR in $where" }
            default {
                $d = "$actorShow - $friendly"
                if ($where) { $d += " ($where)" }
                if ($setting -or $oldV -or $newV) { $d += ": $setting '$oldV' -> '$newV'" }
                if ($app) { $d += ": $app" }
                if ($who -and $op -notmatch 'Participant') { $d += " [betrokken: $(Cut $who 200)]" }
                $d
            }
        }
        if ($isEntra) { $desc += ' (via Entra/M365-beheer)' }

        $anyExt = (@($targets | Where-Object { $_.Ext }).Count -gt 0)
        $row = $ev.NewRow()
        $row.Tijd = $tl; $row.TijdUTC = $tu; $row.Gebruiker = $actor
        $row.Actie = $friendly; $row.Activiteit = $op; $row.Omschrijving = $desc
        $row.Type = $type; $row.TeamChat = $ctxName; $row.Kanaal = $channel
        $row.Betrokkenen = Cut $who 1000; $row.Extern = if ($anyExt) { 'Ja' } else { '' }
        $row.Instelling = $setting; $row.OudeWaarde = $oldV; $row.NieuweWaarde = $newV
        $row.App = $app; $row.Extra = $extra; $row.BerichtId = $msgId
        $row.ChatId = $chatId; $row.TeamId = $teamId
        $row.ClientIP = [string]$a.ClientIP; $row.UserType = Get-UserTypeNl $a.UserType
        $row.Workload = [string]$a.Workload; $row.RecordType = [string]$r.RecordType
        $row.Raw = [string]$r.AuditData
        $ev.Rows.Add($row)

        if ($op -match '^Member(Added|Removed|RoleChanged)$' -or $isEntra) {
            foreach ($t in $targets) {
                $act = switch -Regex ($op) {
                    'Added|^Add '        { if ($op -match 'owner') { 'Eigenaar toegevoegd' } else { 'Toegevoegd' } }
                    'Removed|^Remove '   { if ($t.Self) { 'Zelf verlaten' } elseif ($op -match 'owner') { 'Eigenaar verwijderd' } else { 'Verwijderd' } }
                    'RoleChanged'        { 'Rol gewijzigd' }
                    default              { $op }
                }
                $m = $mb.NewRow()
                $m.Tijd = $tl; $m.Uitvoerder = $actor; $m.Actie = $act
                $m.Betrokkene = [string]$t.Name; $m.BetrokkeneUPN = [string]$t.Upn; $m.Rol = [string]$t.Role
                $m.Extern = if ($t.Ext) { 'Ja' } else { '' }
                $m.Type = $type; $m.TeamChat = $ctxName; $m.Kanaal = $channel
                $m.Bron = if ($isEntra) { 'Entra/M365' } else { 'Teams' }
                $m.Activiteit = $op; $m.ChatId = $chatId; $m.TeamId = $teamId
                $m.Raw = [string]$r.AuditData
                $mb.Rows.Add($m)
            }
        }
        $opCount[$op] = 1 + [int]$opCount[$op]
        $added++
    }
    $doneBlocks++
    $Sync.Pct = [math]::Round(100 * $doneBlocks / [math]::Max(1, $totalBlocks), 0)
    L "Blok $label : $($recs.Count) records, $added na filter (totaal $($ev.Rows.Count) gebeurtenissen, $($mb.Rows.Count) lidmaatschapsregels)"
}
if ($skippedWl) { L "$skippedWl record(s) van andere workloads (bijv. Viva Engage) overgeslagen" }

# --- Overzichten ---
$Sync.Status = 'Overzichten maken...'
$stActor = New-T 'PerUitvoerder' @('Uitvoerder','Totaal','LedenToegevoegd','LedenVerwijderd','ZelfVerlaten','BerichtenVerwijderd','Instellingen','Overig') @('Totaal','LedenToegevoegd','LedenVerwijderd','ZelfVerlaten','BerichtenVerwijderd','Instellingen','Overig')
$stOp    = New-T 'PerActiviteit' @('Activiteit','Omschrijving','Aantal','Uitvoerders') @('Aantal','Uitvoerders')
$stCtx   = New-T 'PerTeamChat' @('Type','Naam','Gebeurtenissen','LedenToegevoegd','LedenVerwijderd','BerichtenVerwijderd','Uitvoerders') @('Gebeurtenissen','LedenToegevoegd','LedenVerwijderd','BerichtenVerwijderd','Uitvoerders')
$stTgt   = New-T 'PerBetrokkene' @('Betrokkene','UPN','Extern','KeerToegevoegd','KeerVerwijderd','ZelfVerlaten','VerwijderdDoor') @('KeerToegevoegd','KeerVerwijderd','ZelfVerlaten')

$ha = @{}; $ho = @{}; $hc = @{}; $ht = @{}
foreach ($row in $ev.Rows) {
    $k = if ($row.Gebruiker) { $row.Gebruiker } else { '(onbekend)' }
    if (-not $ha.ContainsKey($k)) { $ha[$k] = @{ T = 0; A = 0; R = 0; S = 0; M = 0; I = 0; O = 0 } }
    $x = $ha[$k]; $x.T++
    switch -Regex ($row.Activiteit) {
        '^MemberAdded$|^Add (member|owner)' { $x.A++ }
        '^MemberRemoved$|^Remove (member|owner)' { if ($row.Omschrijving -match 'verlaten') { $x.S++ } else { $x.R++ } }
        '^MessageDeleted$' { $x.M++ }
        'SettingChanged|RoleChanged' { $x.I++ }
        default { $x.O++ }
    }
    $o = $row.Activiteit
    if (-not $ho.ContainsKey($o)) { $ho[$o] = @{ N = 0; D = $row.Actie; U = @{} } }
    $ho[$o].N++; $ho[$o].U[$k] = 1
    if ($row.TeamChat -or $row.ChatId -or $row.TeamId) {
        $ck = "$($row.Type)|$($row.TeamChat)"
        if (-not $hc.ContainsKey($ck)) { $hc[$ck] = @{ Type = $row.Type; Name = $row.TeamChat; N = 0; A = 0; R = 0; M = 0; U = @{} } }
        $c = $hc[$ck]; $c.N++; $c.U[$k] = 1
        if ($row.Activiteit -match '^MessageDeleted$') { $c.M++ }
    }
}
foreach ($m in $mb.Rows) {
    $ck = "$($m.Type)|$($m.TeamChat)"
    if ($hc.ContainsKey($ck)) {
        if ($m.Actie -match 'toegevoegd') { $hc[$ck].A++ }
        elseif ($m.Actie -match 'verwijderd|verlaten') { $hc[$ck].R++ }
    }
    $tk = if ($m.BetrokkeneUPN) { $m.BetrokkeneUPN.ToLower() } else { $m.Betrokkene }
    if (-not $ht.ContainsKey($tk)) { $ht[$tk] = @{ Name = $m.Betrokkene; Upn = $m.BetrokkeneUPN; Ext = $m.Extern; A = 0; R = 0; S = 0; By = @{} } }
    $t = $ht[$tk]
    if ($m.Actie -match 'toegevoegd') { $t.A++ }
    elseif ($m.Actie -eq 'Zelf verlaten') { $t.S++ }
    elseif ($m.Actie -match 'verwijderd') { $t.R++; $t.By[$m.Uitvoerder] = 1 + [int]$t.By[$m.Uitvoerder] }
}
foreach ($k in $ha.Keys) { $x = $ha[$k]; [void]$stActor.Rows.Add($k, $x.T, $x.A, $x.R, $x.S, $x.M, $x.I, $x.O) }
foreach ($k in $ho.Keys) { $x = $ho[$k]; [void]$stOp.Rows.Add($k, $x.D, $x.N, $x.U.Count) }
foreach ($k in $hc.Keys) { $x = $hc[$k]; [void]$stCtx.Rows.Add($x.Type, $x.Name, $x.N, $x.A, $x.R, $x.M, $x.U.Count) }
foreach ($k in $ht.Keys) {
    $x = $ht[$k]
    $by = (($x.By.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key) ($($_.Value)x)" }) -join ', ')
    [void]$stTgt.Rows.Add($x.Name, $x.Upn, $x.Ext, $x.A, $x.R, $x.S, $by)
}
$stActor.DefaultView.Sort = 'Totaal DESC'
$stOp.DefaultView.Sort    = 'Aantal DESC'
$stCtx.DefaultView.Sort   = 'Gebeurtenissen DESC'
$stTgt.DefaultView.Sort   = 'KeerVerwijderd DESC'
$Sync.Stats = @{ Actor = $stActor; Op = $stOp; Ctx = $stCtx; Tgt = $stTgt }

$Sync.Summary = (($opCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$(if ($F.ContainsKey($_.Key)) { $F[$_.Key] } else { $_.Key }): $($_.Value)" }) -join '   ')
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
    $script:Domains = @($script:Sync.Domains)
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
    Set-Status 'Verbonden - kies periode en activiteiten en klik Zoeken' 'success'
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
    param([string]$Text)
    $Text = $Text.Trim()
    if (-not $Text) { return $null }
    $terms = @($Text -split '[,;|]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_) })
    if (-not $terms.Count) { return $null }
    return ($terms -join '|')
}

function Get-TimeOfDay {
    param([string]$Text, [string]$Default)
    $t = $Text.Trim(); if (-not $t) { $t = $Default }
    $ts = [TimeSpan]::Zero
    if ($t -match '^\d{1,2}$') { $t = "$($t):00" }
    if ([TimeSpan]::TryParse($t, [ref]$ts) -and $ts.TotalHours -lt 24 -and $ts -ge [TimeSpan]::Zero) { return $ts }
    return $null
}

function Get-SelectedOps {
    $ops = @($script:OpChecks.Keys | Where-Object { $script:OpChecks[$_].IsChecked })
    $extra = @($script:TxtExtraOps.Text -split '[,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    return @($ops + $extra | Select-Object -Unique)
}

function Invoke-Search {
    if (-not $script:DpFrom.SelectedDate -or -not $script:DpTo.SelectedDate) {
        Show-Msg 'Kies een begin- en einddatum.' 'Periode' 'OK' 'Warning' | Out-Null; return
    }
    $tFrom = Get-TimeOfDay $script:TxtFromTime.Text '00:00'
    $tTo   = Get-TimeOfDay $script:TxtToTime.Text '23:59'
    if ($null -eq $tFrom -or $null -eq $tTo) { Show-Msg 'Ongeldige tijd. Gebruik UU:MM, bijv. 08:30.' 'Tijd' 'OK' 'Warning' | Out-Null; return }
    $start = $script:DpFrom.SelectedDate.Date.Add($tFrom)
    $end   = $script:DpTo.SelectedDate.Date.Add($tTo).AddMinutes(1)
    if ($end -le $start) { Show-Msg 'Het einde ligt voor het begin.' 'Periode' 'OK' 'Warning' | Out-Null; return }
    if ($end -gt (Get-Date).AddMinutes(1)) { $end = (Get-Date).AddMinutes(1) }
    if ($start -lt (Get-Date).Date.AddDays(-180)) {
        Write-Log 'Begindatum ligt meer dan 180 dagen terug; zonder Audit Premium is ouder dan 180 dagen niet meer beschikbaar.' 'warning'
    }

    $ops = Get-SelectedOps
    $rt = [string]$script:CmbRecord.SelectedItem.Tag
    if (-not $ops.Count) {
        $a = Show-Msg "Geen activiteiten gekozen: alle Teams-activiteiten worden opgehaald (ook gelezen berichten e.d.). Dat kan veel resultaten en tijd kosten.`n`nDoorgaan?" 'Alle activiteiten' 'YesNo' 'Question'
        if ($a -ne 'Yes') { return }
        if (-not $rt) { $rt = 'MicrosoftTeams' }
    }
    $queries = @(@{ Name = 'Teams'; Kind = 'Teams'; RecordType = $rt; Ops = $ops })
    if ($script:ChkEntra.IsChecked) {
        $queries += @{ Name = 'Entra'; Kind = 'Entra'; RecordType = 'AzureActiveDirectory'; Ops = @('Add member to group.','Remove member from group.','Add owner to group.','Remove owner from group.') }
    }

    $users = @($script:TxtUsers.Text -split '[,;\s]+' | Where-Object { $_ })
    $tgtRx = ConvertTo-TermRegex $script:TxtTarget.Text
    $ctxRx = ConvertTo-TermRegex $script:TxtCtx.Text
    $chunkHours = switch ($script:CmbChunk.SelectedIndex) { 0 { 24 } 1 { 6 } 2 { 1 } default { 0 } }

    $p = @{
        Start = $start; End = $end; Queries = $queries; Users = $users
        ChunkHours = $chunkHours; TargetRx = $tgtRx; CtxRx = $ctxRx
        OnlyTeams = [bool]$script:ChkOnlyTeams.IsChecked
        Friendly = $script:Friendly; Domains = $script:Domains
    }
    $script:LastSearch = @{ Start = $start; End = $end; Ops = $ops; Users = $users; Target = $script:TxtTarget.Text; Ctx = $script:TxtCtx.Text; Entra = [bool]$script:ChkEntra.IsChecked }
    Write-Log ("Activiteiten: {0}" -f $(if ($ops.Count) { $ops -join ', ' } else { '(alle)' }))
    if ($users.Count) { Write-Log "Uitvoerders: $($users -join ', ')" }
    if ($tgtRx) { Write-Log "Betrokkene-filter: $tgtRx" }
    if ($ctxRx) { Write-Log "Team/chat-filter: $ctxRx" }
    if ($script:ChkEntra.IsChecked) { Write-Log 'Ook Entra/M365-groepslidmaatschap wordt opgehaald' }

    Clear-Results
    Save-Settings
    Set-Status 'Zoeken...'
    Start-Work 'search' $script:SearchScript $p
}

function Clear-Results {
    $script:Grid.ItemsSource = $null
    $script:GridMembers.ItemsSource = $null
    foreach ($g in $script:GridStActor, $script:GridStOp, $script:GridStCtx, $script:GridStTgt) { $g.ItemsSource = $null }
    $script:Table = $null
    $script:Members = $null
    $script:TxtDetail.Text = ''
    $script:LblCount.Text = ''
    $script:LblSummary.Text = ''
    $script:TabEvents.Header = 'Alle gebeurtenissen'
    $script:TabMembers.Header = 'Lidmaatschap (wie deed wat bij wie)'
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

function Set-ComboItems {
    param($Combo, [string]$AllText, [string[]]$Items)
    $Combo.Items.Clear()
    [void]$Combo.Items.Add($AllText)
    foreach ($i in $Items) { if ($i) { [void]$Combo.Items.Add($i) } }
    $Combo.SelectedIndex = 0
}

function Complete-Search {
    param($Err)
    $stopped = $Err -is [System.Management.Automation.PipelineStoppedException]
    if ($Err -and -not $stopped) {
        Write-Log "Zoeken mislukt: $($Err.Message)" 'error'
        if ($Err.Message -match 'session|token|expired|verlopen|Connect-ExchangeOnline') { Write-Log 'Sessie mogelijk verlopen; log opnieuw in.' 'warning' }
    }
    $script:Table   = $script:Sync.Table
    $script:Members = $script:Sync.Members
    $n = 0; $nm = 0
    if ($script:Table) {
        $script:Table.DefaultView.Sort = 'Tijd ASC'
        $script:Grid.ItemsSource = $script:Table.DefaultView
        $n = $script:Table.Rows.Count
        Set-ComboItems $script:CmbOpFilter '(alle activiteiten)' @($script:Table.Rows | ForEach-Object { $_.Activiteit } | Sort-Object -Unique)
        Set-ComboItems $script:CmbTypeFilter '(alle typen)' @($script:Table.Rows | ForEach-Object { $_.Type } | Sort-Object -Unique)
    }
    if ($script:Members) {
        $script:Members.DefaultView.Sort = 'Tijd ASC'
        $script:GridMembers.ItemsSource = $script:Members.DefaultView
        $nm = $script:Members.Rows.Count
    }
    $st = $script:Sync.Stats
    if ($st) {
        $script:GridStActor.ItemsSource = $st.Actor.DefaultView
        $script:GridStOp.ItemsSource    = $st.Op.DefaultView
        $script:GridStCtx.ItemsSource   = $st.Ctx.DefaultView
        $script:GridStTgt.ItemsSource   = $st.Tgt.DefaultView
    }
    $script:LblSummary.Text = [string]$script:Sync.Summary
    Update-Filter
    if ($nm -gt 0 -and $n -gt 0 -and -not ($script:Table.Rows | Where-Object { $_.Activiteit -notmatch '^Member|group\.$' } | Select-Object -First 1)) {
        $script:Tabs.SelectedIndex = 1
    }
    if ($stopped -or $script:Sync.Cancel) {
        Set-Status "Gestopt - $n gebeurtenissen, $nm lidmaatschapsregels" 'warning'
    } elseif ($Err) {
        Set-Status "Zoeken mislukt - $n gebeurtenissen" 'error'
    } elseif ($n -eq 0) {
        Write-Log "Geen resultaten ($($script:Sync.RawTotal) records opgehaald voor filtering)." 'warning'
        Set-Status 'Geen resultaten gevonden voor deze periode/filters' 'warning'
    } else {
        $f = [int]$script:Sync.Failed
        Write-Log "Klaar: $n gebeurtenissen, $nm lidmaatschapsregels ($($script:Sync.RawTotal) records opgehaald)$(if ($f) { ", $f blok(ken) mislukt" })" 'success'
        Set-Status "Klaar - $n gebeurtenissen, $nm lidmaatschapsregels" $(if ($f) { 'warning' } else { 'success' })
    }
}

function Get-LikeEscaped {
    param([string]$s)
    $s = $s.Replace("'", "''")
    return ($s -replace '([\[\]\*%])', '[$1]')
}

function Get-RowFilter {
    param([string[]]$TextCols)
    $parts = @()
    $q = $script:TxtFilter.Text.Trim()
    if ($q) {
        foreach ($word in @($q -split '\s+' | Where-Object { $_ })) {
            $e = Get-LikeEscaped $word
            $parts += '(' + (($TextCols | ForEach-Object { "[$_] LIKE '%$e%'" }) -join ' OR ') + ')'
        }
    }
    if ($script:CmbOpFilter.SelectedIndex -gt 0) {
        $parts += "[Activiteit] = '$(([string]$script:CmbOpFilter.SelectedItem).Replace("'", "''"))'"
    }
    if ($script:CmbTypeFilter.SelectedIndex -gt 0) {
        $parts += "[Type] = '$(([string]$script:CmbTypeFilter.SelectedItem).Replace("'", "''"))'"
    }
    if ($script:ChkExtOnly.IsChecked) { $parts += "[Extern] = 'Ja'" }
    return ($parts -join ' AND ')
}

function Update-Filter {
    if (-not $script:Table) { $script:LblCount.Text = ''; return }
    $evCols = 'Gebruiker','Actie','Activiteit','Omschrijving','Type','TeamChat','Kanaal','Betrokkenen','Instelling','OudeWaarde','NieuweWaarde','App','ClientIP','ChatId','TeamId','UserType','BerichtId'
    $mbCols = 'Uitvoerder','Actie','Betrokkene','BetrokkeneUPN','Rol','Type','TeamChat','Kanaal','Bron','Activiteit','ChatId','TeamId'
    try { $script:Table.DefaultView.RowFilter = (Get-RowFilter $evCols) } catch { $script:Table.DefaultView.RowFilter = '' }
    if ($script:Members) {
        try { $script:Members.DefaultView.RowFilter = (Get-RowFilter $mbCols) } catch { $script:Members.DefaultView.RowFilter = '' }
    }
    $ne = $script:Table.DefaultView.Count
    $nm = if ($script:Members) { $script:Members.DefaultView.Count } else { 0 }
    $script:TabEvents.Header  = "Alle gebeurtenissen ($ne)"
    $script:TabMembers.Header = "Lidmaatschap - wie deed wat bij wie ($nm)"
    $script:LblCount.Text = "$ne van $($script:Table.Rows.Count) gebeurtenissen  |  $nm van $(if ($script:Members) { $script:Members.Rows.Count } else { 0 }) lidmaatschapsregels"
}

function Get-ActiveGrid {
    switch ($script:Tabs.SelectedIndex) {
        0 { return $script:Grid }
        1 { return $script:GridMembers }
        default {
            foreach ($g in $script:GridStActor, $script:GridStOp, $script:GridStCtx, $script:GridStTgt) { if ($g.IsKeyboardFocusWithin) { return $g } }
            return $script:GridStActor
        }
    }
}

function Get-SelectedRow {
    param($G)
    if (-not $G) { $G = Get-ActiveGrid }
    $rv = $G.SelectedItem
    if ($rv -is [System.Data.DataRowView]) { return $rv.Row }
    return $null
}

function Show-Detail {
    param($G)
    $r = Get-SelectedRow $G
    if (-not $r) { return }
    $sb = New-Object System.Text.StringBuilder
    $cols = $r.Table.Columns | Where-Object { $_.ColumnName -ne 'Raw' }
    foreach ($c in $cols) {
        $v = [string]$r[$c.ColumnName]
        if ($v) { [void]$sb.AppendLine(('{0,-15} {1}' -f ($c.ColumnName + ':'), $v)) }
    }
    if ($r.Table.Columns.Contains('Raw')) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('--- Volledige AuditData ---')
        try { [void]$sb.Append(($r.Raw | ConvertFrom-Json | ConvertTo-Json -Depth 10)) } catch { [void]$sb.Append($r.Raw) }
    }
    $script:TxtDetail.Text = $sb.ToString()
    $script:TxtDetail.ScrollToHome()
}

function Get-ExportColumns {
    param($G)
    $cols = @()
    if ($G.AutoGenerateColumns) {
        $tbl = $G.ItemsSource.Table
        foreach ($c in $tbl.Columns) { $cols += @{ Header = $c.ColumnName; Path = $c.ColumnName } }
        return $cols
    }
    foreach ($c in $G.Columns) {
        $path = $null
        try { $path = $c.Binding.Path.Path } catch { }
        if ($path) { $cols += @{ Header = [string]$c.Header; Path = $path } }
    }
    return $cols
}

function Get-OrgTag {
    $org = ([string]$script:Sync.ConnOrg) -replace '\.onmicrosoft\.com$', '' -replace '[^\w\.-]', '_'
    if ($org) { return "$($org)_" } else { return '' }
}

function Export-Results {
    $g = Get-ActiveGrid
    $view = $g.ItemsSource
    if (-not $view -or $view.Count -eq 0) { Show-Msg 'Geen resultaten om te exporteren op dit tabblad.' 'Export' | Out-Null; return }
    $name = switch ($script:Tabs.SelectedIndex) { 0 { 'Gebeurtenissen' } 1 { 'Lidmaatschap' } default { $view.Table.TableName } }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (puntkomma)|*.csv'
    $dlg.FileName = "TeamsAudit_$(Get-OrgTag)$($name)_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if (-not $dlg.ShowDialog($script:Window)) { return }
    try {
        $cols = Get-ExportColumns $g
        $rows = foreach ($rv in $view) {
            $o = [ordered]@{}
            foreach ($c in $cols) { $o[$c.Header] = [string]$rv.Row[$c.Path] }
            [pscustomobject]$o
        }
        $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Write-Log "Export: $($view.Count) regels naar $($dlg.FileName)" 'success'
        Set-Status 'Export klaar' 'success'
    } catch {
        Write-Log "Export mislukt: $($_.Exception.Message)" 'error'
    }
}

function ConvertTo-HtmlTable {
    param($View, [string[]]$Cols, [int]$Max = 5000)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table><thead><tr>')
    foreach ($c in $Cols) { [void]$sb.Append("<th>$([System.Net.WebUtility]::HtmlEncode($c))</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    $i = 0
    foreach ($rv in $View) {
        if ($i -ge $Max) { break }
        $cls = ''
        if ($rv.Row.Table.Columns.Contains('Actie')) {
            $act = [string]$rv.Row['Actie']
            if ($act -match 'erwijderd') { $cls = ' class="rem"' } elseif ($act -match 'oegevoegd') { $cls = ' class="add"' } elseif ($act -match 'verlaten') { $cls = ' class="left"' }
        }
        [void]$sb.Append("<tr$cls>")
        foreach ($c in $Cols) { [void]$sb.Append("<td>$([System.Net.WebUtility]::HtmlEncode([string]$rv.Row[$c]))</td>") }
        [void]$sb.Append('</tr>')
        $i++
    }
    [void]$sb.Append('</tbody></table>')
    if ($View.Count -gt $Max) { [void]$sb.Append("<p class='note'>Alleen de eerste $Max van $($View.Count) regels getoond; gebruik CSV-export voor alles.</p>") }
    return $sb.ToString()
}

function Export-Report {
    if (-not $script:Table -or $script:Table.Rows.Count -eq 0) { Show-Msg 'Geen resultaten voor een rapport.' 'Rapport' | Out-Null; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'HTML-rapport|*.html'
    $dlg.FileName = "TeamsAudit_$(Get-OrgTag)Rapport_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if (-not $dlg.ShowDialog($script:Window)) { return }
    try {
        $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }
        $ls = $script:LastSearch
        $period = if ($ls) { '{0:dd-MM-yyyy HH:mm} t/m {1:dd-MM-yyyy HH:mm}' -f $ls.Start, $ls.End.AddMinutes(-1) } else { '' }
        $filt = $script:TxtFilter.Text.Trim()
        $st = $script:Sync.Stats
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append(@'
<!doctype html><html lang="nl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Teams Audit rapport</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222;background:#fafafa}
h1{color:#464EB8;margin-bottom:4px}h2{color:#464EB8;margin-top:28px;border-bottom:2px solid #464EB8;padding-bottom:4px}
.meta{color:#555;margin-bottom:16px}.meta b{color:#222}
table{border-collapse:collapse;width:100%;font-size:12px;background:#fff}
th{background:#464EB8;color:#fff;text-align:left;padding:6px;position:sticky;top:0}
td{border-bottom:1px solid #e3e3e3;padding:5px 6px;vertical-align:top}
tr:nth-child(even) td{background:#f5f6fc}
tr.rem td{background:#fde7e9}tr.add td{background:#e6f4e6}tr.left td{background:#fff4ce}
.note{color:#888;font-size:12px}.cards{display:flex;gap:12px;flex-wrap:wrap}
.card{background:#fff;border:1px solid #ddd;border-radius:6px;padding:10px 16px;min-width:150px}
.card .n{font-size:24px;font-weight:600;color:#464EB8}
</style></head><body>
'@)
        $removed = @($script:Members.Rows | Where-Object { $_.Actie -match 'erwijderd' }).Count
        $addedN  = @($script:Members.Rows | Where-Object { $_.Actie -match 'oegevoegd' }).Count
        $leftN   = @($script:Members.Rows | Where-Object { $_.Actie -eq 'Zelf verlaten' }).Count
        $msgDel  = @($script:Table.Rows | Where-Object { $_.Activiteit -eq 'MessageDeleted' }).Count
        [void]$sb.Append("<h1>Teams Audit rapport</h1><div class='meta'>Tenant: <b>$(& $enc $script:Sync.ConnOrg)</b> &nbsp;|&nbsp; Periode: <b>$(& $enc $period)</b> &nbsp;|&nbsp; Gemaakt: <b>$(Get-Date -Format 'dd-MM-yyyy HH:mm')</b> door $(& $enc $script:Sync.ConnUser)")
        if ($filt) { [void]$sb.Append(" &nbsp;|&nbsp; Filter: <b>$(& $enc $filt)</b>") }
        if ($ls -and $ls.Users.Count) { [void]$sb.Append(" &nbsp;|&nbsp; Uitvoerders: <b>$(& $enc ($ls.Users -join ', '))</b>") }
        if ($ls -and $ls.Target) { [void]$sb.Append(" &nbsp;|&nbsp; Betrokkene: <b>$(& $enc $ls.Target)</b>") }
        if ($ls -and $ls.Ctx) { [void]$sb.Append(" &nbsp;|&nbsp; Team/chat: <b>$(& $enc $ls.Ctx)</b>") }
        [void]$sb.Append('</div>')
        [void]$sb.Append("<div class='cards'><div class='card'><div class='n'>$($script:Table.Rows.Count)</div>gebeurtenissen</div><div class='card'><div class='n'>$addedN</div>leden toegevoegd</div><div class='card'><div class='n'>$removed</div>leden verwijderd</div><div class='card'><div class='n'>$leftN</div>zelf verlaten</div><div class='card'><div class='n'>$msgDel</div>berichten verwijderd</div></div>")

        [void]$sb.Append('<h2>Lidmaatschap - wie deed wat bij wie</h2>')
        if ($script:Members -and $script:Members.DefaultView.Count) {
            [void]$sb.Append((ConvertTo-HtmlTable $script:Members.DefaultView @('Tijd','Uitvoerder','Actie','Betrokkene','BetrokkeneUPN','Rol','Extern','Type','TeamChat','Kanaal','Bron')))
        } else { [void]$sb.Append("<p class='note'>Geen lidmaatschapswijzigingen in de selectie.</p>") }
        if ($st) {
            [void]$sb.Append('<h2>Per uitvoerder</h2>');  [void]$sb.Append((ConvertTo-HtmlTable $st.Actor.DefaultView @($st.Actor.Columns | ForEach-Object { $_.ColumnName })))
            [void]$sb.Append('<h2>Per betrokkene</h2>');  [void]$sb.Append((ConvertTo-HtmlTable $st.Tgt.DefaultView @($st.Tgt.Columns | ForEach-Object { $_.ColumnName })))
            [void]$sb.Append('<h2>Per team / chat</h2>'); [void]$sb.Append((ConvertTo-HtmlTable $st.Ctx.DefaultView @($st.Ctx.Columns | ForEach-Object { $_.ColumnName })))
            [void]$sb.Append('<h2>Per activiteit</h2>');  [void]$sb.Append((ConvertTo-HtmlTable $st.Op.DefaultView @($st.Op.Columns | ForEach-Object { $_.ColumnName })))
        }
        [void]$sb.Append('<h2>Alle gebeurtenissen (huidige filter)</h2>')
        [void]$sb.Append((ConvertTo-HtmlTable $script:Table.DefaultView @('Tijd','Gebruiker','Actie','Omschrijving','Type','TeamChat','Kanaal','Betrokkenen','Extern','ClientIP')))
        [void]$sb.Append("<p class='note'>Bron: Microsoft 365 Unified Audit Log. Tijden in lokale tijd. Inhoud van berichten wordt niet gelogd. Lidmaatschapswijzigingen via Entra/M365-beheer tonen in het Teams-log een teameigenaar als uitvoerder.</p></body></html>")
        [System.IO.File]::WriteAllText($dlg.FileName, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
        Write-Log "Rapport opgeslagen: $($dlg.FileName)" 'success'
        Start-Process $dlg.FileName
    } catch {
        Write-Log "Rapport mislukt: $($_.Exception.Message)" 'error'
    }
}

function Copy-Selection {
    $g = Get-ActiveGrid
    $sel = @($g.SelectedItems)
    if (-not $sel.Count) { return }
    $cols = Get-ExportColumns $g
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine((($cols | ForEach-Object { $_.Header }) -join "`t"))
    foreach ($rv in $sel) { [void]$sb.AppendLine((($cols | ForEach-Object { [string]$rv.Row[$_.Path] }) -join "`t")) }
    [System.Windows.Clipboard]::SetText($sb.ToString())
    Write-Log "$($sel.Count) regel(s) gekopieerd" 'success'
}

function Set-FilterFrom {
    param([string]$What)
    $r = Get-SelectedRow
    if (-not $r) { return }
    $cols = $r.Table.Columns
    $v = switch ($What) {
        'actor'  { if ($cols.Contains('Gebruiker')) { $r.Gebruiker } elseif ($cols.Contains('Uitvoerder')) { $r.Uitvoerder } }
        'ctx'    { if ($cols.Contains('ChatId') -and $r.ChatId) { $r.ChatId } elseif ($cols.Contains('TeamId') -and $r.TeamId) { $r.TeamId } elseif ($cols.Contains('TeamChat')) { $r.TeamChat } elseif ($cols.Contains('Naam')) { $r.Naam } }
        'target' { if ($cols.Contains('BetrokkeneUPN') -and $r.BetrokkeneUPN) { $r.BetrokkeneUPN } elseif ($cols.Contains('UPN')) { $r.UPN } elseif ($cols.Contains('Betrokkene')) { $r.Betrokkene } }
    }
    $v = [string]$v
    if (-not $v) { Write-Log 'Deze regel heeft daar geen waarde voor' 'warning'; return }
    $script:TxtFilter.Text = $v
    if ($script:Tabs.SelectedIndex -gt 1) { $script:Tabs.SelectedIndex = 0 }
}

function New-GridMenu {
    $cm = New-Object System.Windows.Controls.ContextMenu
    $items = @(
        @('Filter op uitvoerder', { Set-FilterFrom 'actor' }),
        @('Filter op deze chat / dit team', { Set-FilterFrom 'ctx' }),
        @('Filter op betrokkene', { Set-FilterFrom 'target' }),
        @('-', $null),
        @('Kopieer selectie', { Copy-Selection }),
        @('Wis filters', { Clear-Filters })
    )
    foreach ($it in $items) {
        if ($it[0] -eq '-') { [void]$cm.Items.Add((New-Object System.Windows.Controls.Separator)); continue }
        $mi = New-Object System.Windows.Controls.MenuItem
        $mi.Header = $it[0]
        $mi.Add_Click($it[1])
        [void]$cm.Items.Add($mi)
    }
    return $cm
}

function Clear-Filters {
    $script:TxtFilter.Text = ''
    if ($script:CmbOpFilter.Items.Count) { $script:CmbOpFilter.SelectedIndex = 0 }
    if ($script:CmbTypeFilter.Items.Count) { $script:CmbTypeFilter.SelectedIndex = 0 }
    $script:ChkExtOnly.IsChecked = $false
    Update-Filter
}

function Set-Preset {
    param([string]$Name)
    $set = if ($Name -and $script:Presets.Contains($Name)) { $script:Presets[$Name] } else { @() }
    foreach ($k in $script:OpChecks.Keys) { $script:OpChecks[$k].IsChecked = ($set -contains $k) }
    Update-OpCount
}

function Update-OpCount {
    $n = @($script:OpChecks.Keys | Where-Object { $script:OpChecks[$_].IsChecked }).Count
    $script:ExpOps.Header = "Activiteiten ($n van $($script:OpChecks.Count) gekozen) - klik om uit/in te klappen"
    foreach ($cat in $script:OpCats.Keys) {
        $ops = $script:OpCats[$cat]
        $c = @($ops | Where-Object { $script:OpChecks[$_].IsChecked }).Count
        $script:CatChecks[$cat].IsChecked = if ($c -eq 0) { $false } elseif ($c -eq $ops.Count) { $true } else { $null }
    }
}

function Set-Range {
    param([int]$Days)
    $script:DpTo.SelectedDate = (Get-Date).Date
    $script:DpFrom.SelectedDate = (Get-Date).Date.AddDays(-($Days - 1))
    $script:TxtFromTime.Text = '00:00'
    $script:TxtToTime.Text = '23:59'
}

#endregion

#region --- XAML ---

$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Teams Audit" Width="1500" Height="960" MinWidth="1050" MinHeight="700"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>
<Style TargetType="Button">
  <Setter Property="Foreground" Value="White"/>
  <Setter Property="Background" Value="#464EB8"/>
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
  <Setter Property="Background" Value="#E8E9F8"/><Setter Property="Foreground" Value="#33378A"/><Setter Property="Padding" Value="8,3"/><Setter Property="Margin" Value="0,0,6,6"/>
</Style>
<Style x:Key="BtnDark"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="#33378A"/><Setter Property="Padding" Value="14,7"/><Setter Property="Margin" Value="8,0,0,0"/>
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
  <Setter Property="AlternatingRowBackground" Value="#F5F6FC"/>
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
<Style x:Key="CellExt" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="Foreground" Value="#B26200"/>
</Style>
<Style x:Key="CellAct" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="Foreground" Value="#333333"/>
  <Style.Triggers>
    <Trigger Property="Text" Value="Verwijderd"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="Eigenaar verwijderd"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="Toegevoegd"><Setter Property="Foreground" Value="#107C10"/></Trigger>
    <Trigger Property="Text" Value="Eigenaar toegevoegd"><Setter Property="Foreground" Value="#107C10"/></Trigger>
    <Trigger Property="Text" Value="Zelf verlaten"><Setter Property="Foreground" Value="#B26200"/></Trigger>
    <Trigger Property="Text" Value="Rol gewijzigd"><Setter Property="Foreground" Value="#8A2BE2"/></Trigger>
    <Trigger Property="Text" Value="Leden verwijderd"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="Leden toegevoegd"><Setter Property="Foreground" Value="#107C10"/></Trigger>
    <Trigger Property="Text" Value="Bericht verwijderd"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="Team verwijderd"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="Kanaal verwijderd"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="Lid verwijderd uit groep (Entra)"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
    <Trigger Property="Text" Value="Lid toegevoegd aan groep (Entra)"><Setter Property="Foreground" Value="#107C10"/></Trigger>
    <Trigger Property="Text" Value="Rol van lid gewijzigd"><Setter Property="Foreground" Value="#8A2BE2"/></Trigger>
    <Trigger Property="Text" Value="Bericht bewerkt"><Setter Property="Foreground" Value="#0063B1"/></Trigger>
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
<Style x:Key="StHead" TargetType="TextBlock"><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Foreground" Value="#33378A"/><Setter Property="Margin" Value="0,4,0,4"/></Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*" MinHeight="300"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="120" MinHeight="50"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#464EB8">
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
        <StackPanel Orientation="Vertical" VerticalAlignment="Center">
          <TextBlock Text="Teams Audit" Foreground="White" FontSize="20" FontWeight="SemiBold"/>
          <TextBlock Text="Wie voegt toe, verwijdert, verlaat, wijzigt - chats, teams, kanalen, berichten, apps en vergaderingen" Foreground="#DADCF5" FontSize="11"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <GroupBox Grid.Row="1" Header="Zoeken in het audit log" Margin="8,8,8,0">
      <StackPanel>
        <StackPanel x:Name="SearchArea" IsEnabled="False">
          <WrapPanel>
            <TextBlock Text="Van:" Style="{StaticResource Lbl}"/>
            <DatePicker x:Name="DpFrom" Width="120" Margin="0,0,4,6"/>
            <TextBox x:Name="TxtFromTime" Width="50" Style="{StaticResource FieldBox}" Text="00:00" ToolTip="Begintijd (UU:MM)"/>
            <TextBlock Text="t/m:" Style="{StaticResource Lbl}"/>
            <DatePicker x:Name="DpTo" Width="120" Margin="0,0,4,6"/>
            <TextBox x:Name="TxtToTime" Width="50" Style="{StaticResource FieldBox}" Text="23:59" ToolTip="Eindtijd (UU:MM, inclusief die minuut)"/>
            <Button x:Name="BtnR1"   Content="Vandaag"  Style="{StaticResource BtnPreset}"/>
            <Button x:Name="BtnR7"   Content="7 dagen"  Style="{StaticResource BtnPreset}"/>
            <Button x:Name="BtnR30"  Content="30 dagen" Style="{StaticResource BtnPreset}"/>
            <Button x:Name="BtnR90"  Content="90 dagen" Style="{StaticResource BtnPreset}"/>
            <Button x:Name="BtnR180" Content="180 dagen" Style="{StaticResource BtnPreset}"/>
            <Border Style="{StaticResource Sep}"/>
            <TextBlock Text="Recordtype:" Style="{StaticResource Lbl}"/>
            <ComboBox x:Name="CmbRecord" Width="230" Height="26" Margin="0,0,12,6"
                      ToolTip="Automatisch = geen recordtype-filter, alleen op activiteit (vangt ook Teams-admin/add-on recordtypes).">
              <ComboBoxItem Content="Automatisch (op activiteit)" Tag=""/>
              <ComboBoxItem Content="MicrosoftTeams" Tag="MicrosoftTeams"/>
              <ComboBoxItem Content="MicrosoftTeamsAdmin" Tag="MicrosoftTeamsAdmin"/>
              <ComboBoxItem Content="MicrosoftTeamsDevice" Tag="MicrosoftTeamsDevice"/>
              <ComboBoxItem Content="MicrosoftTeamsAnalytics" Tag="MicrosoftTeamsAnalytics"/>
            </ComboBox>
            <TextBlock Text="Ophalen:" Style="{StaticResource Lbl}"/>
            <ComboBox x:Name="CmbChunk" Width="110" Height="26" Margin="0,0,12,6"
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
          <Expander x:Name="ExpOps" Header="Activiteiten" IsExpanded="False" Margin="0,0,0,6">
            <Border Background="#F7F8FD" BorderBrush="#E1E3F2" BorderThickness="1" CornerRadius="4" Padding="8,6,8,0" Margin="0,4,0,0">
              <ScrollViewer MaxHeight="230" VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="PnlOps"/>
              </ScrollViewer>
            </Border>
          </Expander>
          <WrapPanel>
            <TextBlock Text="Uitvoerders:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtUsers" Width="210" Style="{StaticResource FieldBox}" ToolTip="UPN's van wie de actie uitvoerde, komma-gescheiden. Server-side (sneller). Leeg = iedereen."/>
            <TextBlock Text="Betrokkene bevat:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtTarget" Width="190" Style="{StaticResource FieldBox}" ToolTip="Naam of UPN van de persoon die is toegevoegd/verwijderd (deel van tekst, komma-gescheiden). Toont alleen gebeurtenissen met die persoon als betrokkene."/>
            <TextBlock Text="Team/chat bevat:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtCtx" Width="190" Style="{StaticResource FieldBox}" ToolTip="Deel van teamnaam, chatnaam, kanaalnaam, chat-ID of team-ID (komma-gescheiden)."/>
            <TextBlock Text="Extra activiteiten:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtExtraOps" Width="170" Style="{StaticResource FieldBox}" ToolTip="Andere operation-namen, komma-gescheiden"/>
          </WrapPanel>
          <WrapPanel>
            <CheckBox x:Name="ChkEntra" Content="Ook Entra/M365-groepslidmaatschap ophalen"
                      ToolTip="Wijzigingen via Entra ID / M365-admin / Graph tonen in het Teams-log een teameigenaar als uitvoerder. Met dit vinkje worden ook 'Add/Remove member/owner to/from group' uit het Entra-log opgehaald met de echte uitvoerder."/>
            <CheckBox x:Name="ChkOnlyTeams" Content="Alleen Teams-workload" IsChecked="True"
                      ToolTip="Sla records van andere diensten met dezelfde activiteitnaam over (bijv. MessageDeleted in Viva Engage)."/>
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
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <WrapPanel Grid.Row="0">
        <TextBlock Text="Filter resultaten:" Style="{StaticResource Lbl}"/>
        <TextBox x:Name="TxtFilter" Width="230" Style="{StaticResource FieldBox}" ToolTip="Woorden (alle moeten voorkomen) in uitvoerder, betrokkene, team/chat, omschrijving, IP, ID's..."/>
        <ComboBox x:Name="CmbOpFilter" Width="180" Height="26" Margin="0,0,8,6"/>
        <ComboBox x:Name="CmbTypeFilter" Width="140" Height="26" Margin="0,0,8,6"/>
        <CheckBox x:Name="ChkExtOnly" Content="Alleen extern/gast"/>
        <Button x:Name="BtnClearFilter" Content="Wis filters" Style="{StaticResource BtnGray}"/>
        <Border Style="{StaticResource Sep}"/>
        <Button x:Name="BtnCopy" Content="Kopieer selectie" Style="{StaticResource BtnGray}"/>
        <Button x:Name="BtnExport" Content="Export CSV (tabblad)" Style="{StaticResource BtnGray}"/>
        <Button x:Name="BtnReport" Content="HTML-rapport"/>
        <TextBlock x:Name="LblCount" Style="{StaticResource Lbl}" Foreground="Gray" Margin="6,0,0,6"/>
      </WrapPanel>
      <TextBlock Grid.Row="1" x:Name="LblSummary" Foreground="#33378A" TextWrapping="Wrap" Margin="0,0,0,6"/>
      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*" MinWidth="400"/>
          <ColumnDefinition Width="5"/>
          <ColumnDefinition Width="400" MinWidth="120"/>
        </Grid.ColumnDefinitions>
        <TabControl x:Name="Tabs" Grid.Column="0">
          <TabItem x:Name="TabEvents" Header="Alle gebeurtenissen">
            <DataGrid x:Name="Grid">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Tijd (lokaal)" Binding="{Binding Tijd}" Width="130" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Uitvoerder" Binding="{Binding Gebruiker}" Width="190" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Actie" Binding="{Binding Actie}" Width="160" ElementStyle="{StaticResource CellAct}"/>
                <DataGridTextColumn Header="Omschrijving" Binding="{Binding Omschrijving}" Width="420" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="95" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Team / chat" Binding="{Binding TeamChat}" Width="180" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Kanaal" Binding="{Binding Kanaal}" Width="110" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Betrokkenen" Binding="{Binding Betrokkenen}" Width="220" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Extern" Binding="{Binding Extern}" Width="50" ElementStyle="{StaticResource CellExt}"/>
                <DataGridTextColumn Header="Instelling" Binding="{Binding Instelling}" Width="120" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Oud" Binding="{Binding OudeWaarde}" Width="100" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Nieuw" Binding="{Binding NieuweWaarde}" Width="100" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="App" Binding="{Binding App}" Width="120" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Soort uitvoerder" Binding="{Binding UserType}" Width="100" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Client-IP" Binding="{Binding ClientIP}" Width="110" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Activiteit" Binding="{Binding Activiteit}" Width="140" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Extra" Binding="{Binding Extra}" Width="250" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Tijd UTC" Binding="{Binding TijdUTC}" Width="130" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Chat-ID" Binding="{Binding ChatId}" Width="160" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Team-ID" Binding="{Binding TeamId}" Width="160" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Bericht-ID" Binding="{Binding BerichtId}" Width="120" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Workload" Binding="{Binding Workload}" Width="100" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Recordtype" Binding="{Binding RecordType}" Width="120" ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
          </TabItem>
          <TabItem x:Name="TabMembers" Header="Lidmaatschap (wie deed wat bij wie)">
            <DataGrid x:Name="GridMembers">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Tijd (lokaal)" Binding="{Binding Tijd}" Width="130" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Uitvoerder (wie)" Binding="{Binding Uitvoerder}" Width="200" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Actie" Binding="{Binding Actie}" Width="125" ElementStyle="{StaticResource CellAct}"/>
                <DataGridTextColumn Header="Betrokkene (bij wie)" Binding="{Binding Betrokkene}" Width="180" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="UPN betrokkene" Binding="{Binding BetrokkeneUPN}" Width="210" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Rol" Binding="{Binding Rol}" Width="70" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Extern" Binding="{Binding Extern}" Width="50" ElementStyle="{StaticResource CellExt}"/>
                <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="110" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Team / chat (waar)" Binding="{Binding TeamChat}" Width="220" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Kanaal" Binding="{Binding Kanaal}" Width="120" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Bron" Binding="{Binding Bron}" Width="80" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Chat-ID" Binding="{Binding ChatId}" Width="160" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Team-ID" Binding="{Binding TeamId}" Width="160" ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
          </TabItem>
          <TabItem x:Name="TabStats" Header="Overzicht">
            <Grid Margin="4">
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="8"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <TextBlock Grid.Row="0" Grid.Column="0" Text="Per uitvoerder" Style="{StaticResource StHead}"/>
              <DataGrid Grid.Row="1" Grid.Column="0" x:Name="GridStActor" AutoGenerateColumns="True"/>
              <TextBlock Grid.Row="0" Grid.Column="2" Text="Per betrokkene (wie werd toegevoegd/verwijderd, en door wie)" Style="{StaticResource StHead}"/>
              <DataGrid Grid.Row="1" Grid.Column="2" x:Name="GridStTgt" AutoGenerateColumns="True"/>
              <TextBlock Grid.Row="2" Grid.Column="0" Text="Per team / chat" Style="{StaticResource StHead}"/>
              <DataGrid Grid.Row="3" Grid.Column="0" x:Name="GridStCtx" AutoGenerateColumns="True"/>
              <TextBlock Grid.Row="2" Grid.Column="2" Text="Per activiteit" Style="{StaticResource StHead}"/>
              <DataGrid Grid.Row="3" Grid.Column="2" x:Name="GridStOp" AutoGenerateColumns="True"/>
            </Grid>
          </TabItem>
        </TabControl>
        <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#DDDDDD"/>
        <DockPanel Grid.Column="2">
          <TextBlock DockPanel.Dock="Top" Text="Details geselecteerde regel" Style="{StaticResource StHead}" Margin="6,4,0,4"/>
          <TextBox x:Name="TxtDetail" IsReadOnly="True" FontFamily="Consolas" FontSize="11.5" TextWrapping="NoWrap"
                   VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Background="White" BorderBrush="#DDDDDD"/>
        </DockPanel>
      </Grid>
    </Grid>

    <GridSplitter Grid.Row="3" HorizontalAlignment="Stretch" Background="#DDDDDD" Margin="0,4,0,0"/>
    <RichTextBox Grid.Row="4" x:Name="LogBox" IsReadOnly="True" Background="#141414" Foreground="#C8C8C8" BorderThickness="0"
                 FontFamily="Consolas" FontSize="11.5" VerticalScrollBarVisibility="Auto" Padding="6,4">
      <FlowDocument PagePadding="0"><Paragraph x:Name="LogPara" Margin="0"/></FlowDocument>
    </RichTextBox>
    <Border x:Name="StatusBar" Grid.Row="5" Background="#464EB8">
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

# Activiteit-vinkjes per categorie
foreach ($cat in $script:OpCats.Keys) {
    $dp = New-Object System.Windows.Controls.DockPanel
    $dp.Margin = New-Object System.Windows.Thickness(0, 0, 0, 4)
    $cc = New-Object System.Windows.Controls.CheckBox
    $cc.Content = $cat
    $cc.Tag = $cat
    $cc.Width = 175
    $cc.FontWeight = [System.Windows.FontWeights]::SemiBold
    $cc.Foreground = New-Brush '#33378A'
    $cc.VerticalAlignment = 'Top'
    $cc.ToolTip = 'Hele categorie aan/uit'
    $cc.Add_Click({
        param($s, $e)
        $v = ($s.IsChecked -eq $true)
        foreach ($o in $script:OpCats[[string]$s.Tag]) { $script:OpChecks[$o].IsChecked = $v }
        Update-OpCount
    })
    [System.Windows.Controls.DockPanel]::SetDock($cc, 'Left')
    [void]$dp.Children.Add($cc)
    $script:CatChecks[$cat] = $cc
    $wp = New-Object System.Windows.Controls.WrapPanel
    foreach ($op in $script:OpCats[$cat]) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $op
        $cb.Width = 215
        $cb.ToolTip = [string]$script:Friendly[$op]
        $cb.Add_Click({ Update-OpCount })
        [void]$wp.Children.Add($cb)
        $script:OpChecks[$op] = $cb
    }
    [void]$dp.Children.Add($wp)
    [void]$script:PnlOps.Children.Add($dp)
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

# Rechtsklikmenu's
foreach ($g in $script:Grid, $script:GridMembers, $script:GridStActor, $script:GridStOp, $script:GridStCtx, $script:GridStTgt) {
    $g.ContextMenu = New-GridMenu
    $g.Add_SelectionChanged({ param($s, $e) Show-Detail $s })
}

# Standaardwaarden + opgeslagen instellingen
$script:CmbRecord.SelectedIndex = 0
$script:CmbChunk.SelectedIndex  = 0
Set-ComboItems $script:CmbOpFilter '(alle activiteiten)' @()
Set-ComboItems $script:CmbTypeFilter '(alle typen)' @()
Set-Range 7
Set-Preset 'Wie verwijdert wie'
$cfg = Get-Settings
if ($cfg) {
    if ($cfg.AdminUpn) { $script:TxtAdmin.Text  = $cfg.AdminUpn }
    if ($cfg.Tenant)   { $script:TxtTenant.Text = $cfg.Tenant }
    $script:ChkNoWam.IsChecked = [bool]$cfg.NoWam
    if ($null -ne $cfg.Operations) {
        $saved = @($cfg.Operations)
        foreach ($k in $script:OpChecks.Keys) { $script:OpChecks[$k].IsChecked = ($saved -contains $k) }
    }
    if ($cfg.ExtraOps) { $script:TxtExtraOps.Text = $cfg.ExtraOps }
    if ($null -ne $cfg.Chunk -and [int]$cfg.Chunk -ge 0 -and [int]$cfg.Chunk -lt 4) { $script:CmbChunk.SelectedIndex = [int]$cfg.Chunk }
    if ($null -ne $cfg.RecordType -and [int]$cfg.RecordType -ge 0 -and [int]$cfg.RecordType -lt $script:CmbRecord.Items.Count) { $script:CmbRecord.SelectedIndex = [int]$cfg.RecordType }
    $script:ChkEntra.IsChecked = [bool]$cfg.Entra
    if ($null -ne $cfg.OnlyTeams) { $script:ChkOnlyTeams.IsChecked = [bool]$cfg.OnlyTeams }
}
Update-OpCount

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
$script:BtnR180.Add_Click({ Set-Range 180 })
$script:TxtFilter.Add_TextChanged({ Update-Filter })
$script:CmbOpFilter.Add_SelectionChanged({ Update-Filter })
$script:CmbTypeFilter.Add_SelectionChanged({ Update-Filter })
$script:ChkExtOnly.Add_Click({ Update-Filter })
$script:BtnClearFilter.Add_Click({ Clear-Filters })
$script:BtnCopy.Add_Click({ Copy-Selection })
$script:BtnExport.Add_Click({ Export-Results })
$script:BtnReport.Add_Click({ Export-Report })
$script:Grid.Add_MouseDoubleClick({ Set-FilterFrom 'ctx' })
$script:GridMembers.Add_MouseDoubleClick({ Set-FilterFrom 'target' })

$script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$script:Timer.Interval = [TimeSpan]::FromMilliseconds(250)
$script:Timer.Add_Tick({ Invoke-Tick })

$script:Window.Add_Loaded({
    Initialize-Worker
    $script:Timer.Start()
    Update-Enabled
    Write-Log 'Teams Audit gestart. Vul je admin-account in en klik Inloggen.'
    Write-Log "Benodigde rol: 'View-Only Audit Logs' of 'Audit Logs' (bij GDAP: in de relatie)."
    Write-Log 'Tip: dubbelklik op een gebeurtenis = filter op die chat/team; dubbelklik in Lidmaatschap = filter op die persoon. Rechtsklik voor meer.'
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
