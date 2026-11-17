#Requires -Version 5.1
<#
.SYNOPSIS
    Conditional Access Beheer (WPF) - CA-policies en alles wat erbij hoort
    interactief beheren, met een apart admin-account.
.DESCRIPTION
    Tabbladen:
      Policies           overzicht, details in leesbare vorm, nieuw (leeg of uit
                         sjabloon), bewerken (formulier + JSON), dupliceren,
                         status aan / alleen rapport / uit, verwijderen.
                         Bulk: gebruiker/groep uitsluiten of de uitsluiting
                         weer weghalen in alle geselecteerde policies, en de
                         noodaccounts in 1 keer uitsluiten.
      Locaties           named locations (IP-bereiken en landen) maken,
                         wijzigen, verwijderen; toont in welke policies ze
                         gebruikt worden.
      Auth.sterktes      ingebouwde en eigen authenticatiesterktes.
      Auth.contexten     authentication contexts (c1..c99).
      Gebruiker-impact   welke policies gelden voor een gebruiker (via groepen
                         en rollen) + recente aanmeldingen met CA-resultaat.
      Back-up / import   alles exporteren naar JSON en importeren in een andere
                         tenant (via GDAP), noodaccounts per tenant instellen.

    Inloggen: altijd met account-kiezer (Prompt=SelectAccount), zodat je met
    je aparte admin-account inlogt en niet automatisch met je eigen account.
    Het token blijft alleen in dit proces (geen cache op schijf). Het laatst
    gebruikte admin-account wordt als hint onthouden.

    Voor elke wijziging of verwijdering wordt automatisch een JSON-back-up van
    het object gemaakt in %LOCALAPPDATA%\CABeheerTool\Backups\<tenant>.
    Bij opslaan van een policy die aan staat en jezelf zou blokkeren, volgt een
    waarschuwing.
.NOTES
    Vereist: Microsoft.Graph.Authentication (2.33.0 aanbevolen)
      Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser
    Rollen: Conditional Access Administrator of Security Administrator
            (lezen aanmeldlogboeken: Reports Reader / Security Reader).
#>

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $script = $MyInvocation.MyCommand.Path
    if ($script) {
        Start-Process powershell.exe -ArgumentList "-STA -NoProfile -ExecutionPolicy Bypass -File `"$script`"" -Wait
    }
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Windows.Forms

#region --- Globale state ---

$script:GraphRoot   = "https://graph.microsoft.com/v1.0"
$script:CaRoot      = "https://graph.microsoft.com/v1.0/identity/conditionalAccess"
$script:GraphScopes = @(
    'Policy.Read.All', 'Policy.ReadWrite.ConditionalAccess', 'Application.Read.All',
    'Directory.Read.All', 'AuditLog.Read.All'
)
$script:GraphClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools
$script:Window        = $null
$script:Connected     = $false
$script:Msal          = $null
$script:MsalAuthority = $null
$script:Auth          = $null    # account + verloopmoment van het token
$script:Me            = $null    # ingelogde admin (id, upn, groepen, rollen)
$script:TenantId      = $null
$script:LongOp        = $false
$script:StopRequested = $false

$script:Policies      = @()
$script:Locations     = @()
$script:Strengths     = @()
$script:Contexts      = @()
$script:Names         = @{}      # id -> weergavenaam (gebruikers, groepen, rollen, apps)
$script:PolicyTable   = $null
$script:ImpactUser    = $null
$script:ImportItems   = $null

$script:AppDir    = Join-Path $env:LOCALAPPDATA 'CABeheerTool'
$script:Settings  = $null

$script:StateLabels = @{
    enabled = 'Aan'; disabled = 'Uit'; enabledForReportingButNotEnforced = 'Alleen rapport'
}

# Bekende waarden die geen object-ID zijn
$script:SpecialNames = @{
    All = 'Alle'; None = 'Geen'; GuestsOrExternalUsers = 'Gasten/externen'
    Office365 = 'Office 365'; MicrosoftAdminPortals = 'Microsoft-beheerportals'
    AllTrusted = 'Alle vertrouwde locaties'
    '00000000-0000-0000-0000-000000000000' = 'MFA-vertrouwde IP-adressen'
}

# Beheerrollen (template-ID's, in elke tenant gelijk) voor de beheerder-sjablonen
$script:AdminRoles = [ordered]@{
    '62e90394-69f5-4237-9190-012177145e10' = 'Global Administrator'
    'e8611ab8-c189-46e8-94e1-60213ab1f814' = 'Privileged Role Administrator'
    '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' = 'Privileged Authentication Administrator'
    '194ae4cb-b126-40b2-bd5b-6091b380977d' = 'Security Administrator'
    'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' = 'Conditional Access Administrator'
    '29232cdf-9323-42fd-ade2-1d097af3e4de' = 'Exchange Administrator'
    'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' = 'SharePoint Administrator'
    'fe930be7-5e62-47db-91af-98c3a49a38b1' = 'User Administrator'
    '729827e3-9c14-49f7-bb1b-9608f156bbb8' = 'Helpdesk Administrator'
    'c4e39bd9-1100-46d3-8c65-fb160da0071f' = 'Authentication Administrator'
    '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3' = 'Application Administrator'
    '158c047a-c907-4556-b7ef-446551a6b5f7' = 'Cloud Application Administrator'
    '966707d0-3269-4727-9be2-8c3a10f19b9d' = 'Password Administrator'
    'b0f54661-2d74-4c50-afa3-1ec803f12efe' = 'Billing Administrator'
    '3a2c62db-5318-420d-8d74-23affee5d9d5' = 'Intune Administrator'
    '17315797-102d-40b4-93e0-432062caca18' = 'Compliance Administrator'
}

$script:GrantLabels = [ordered]@{
    mfa                  = 'MFA vereisen'
    compliantDevice      = 'Apparaat compliant'
    domainJoinedDevice   = 'Hybrid joined apparaat'
    approvedApplication  = 'Goedgekeurde client-app'
    compliantApplication = 'App-beveiligingsbeleid'
    passwordChange       = 'Wachtwoord wijzigen'
}

$script:ClientAppTypes = [ordered]@{
    browser                     = 'Browser'
    mobileAppsAndDesktopClients = 'Mobiele apps en desktop'
    exchangeActiveSync          = 'Exchange ActiveSync (legacy)'
    other                       = 'Andere clients (legacy)'
}

$script:Platforms = @('android', 'iOS', 'windows', 'macOS', 'linux', 'windowsPhone')

# Terugval als de lijst niet via Graph opgehaald kan worden
$script:DefaultCombinations = @(
    'windowsHelloForBusiness', 'fido2', 'x509CertificateMultiFactor', 'deviceBasedPush',
    'temporaryAccessPassOneTime', 'temporaryAccessPassMultiUse', 'password,microsoftAuthenticatorPush',
    'password,softwareOath', 'password,hardwareOath', 'password,sms', 'password,voice',
    'federatedMultiFactor', 'microsoftAuthenticatorPush,federatedSingleFactor',
    'softwareOath,federatedSingleFactor', 'hardwareOath,federatedSingleFactor',
    'sms,federatedSingleFactor', 'voice,federatedSingleFactor',
    'x509CertificateSingleFactor', 'sms', 'federatedSingleFactor', 'password'
)
$script:Combinations = $null

#endregion

#region --- UI basis-helpers ---

function New-Brush {
    param([string]$Hex)
    $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
    $b.Freeze()
    return $b
}

function Update-UI {
    if ($script:Window) {
        $script:Window.Dispatcher.Invoke([action]{ }, [System.Windows.Threading.DispatcherPriority]::Background)
    }
}

function Show-Msg {
    param([string]$Text, [string]$Title = "Info", [string]$Button = "OK", [string]$Icon = "Information")
    if ($script:Window -and $script:Window.IsLoaded) {
        return [System.Windows.MessageBox]::Show($script:Window, $Text, $Title, $Button, $Icon)
    }
    return [System.Windows.MessageBox]::Show($Text, $Title, $Button, $Icon)
}

function Write-Log {
    param([string]$Message, [string]$Type = "info")
    if ($null -eq $script:LogPara) { return }
    $ts = Get-Date -Format "HH:mm:ss"
    $color = switch ($Type) {
        "success" { "#50C850" }
        "error"   { "#F05050" }
        "warning" { "#FFB432" }
        "api"     { "#82B4FF" }
        default   { "#C8C8C8" }
    }
    $prefix = switch ($Type) {
        "success" { "[OK]  " }
        "error"   { "[ERR] " }
        "warning" { "[WRN] " }
        "api"     { "[API] " }
        default   { "[INF] " }
    }
    $r1 = New-Object System.Windows.Documents.Run("$ts ")
    $r1.Foreground = New-Brush "#828282"
    $r2 = New-Object System.Windows.Documents.Run("$prefix$Message")
    $r2.Foreground = New-Brush $color
    $script:LogPara.Inlines.Add($r1)
    $script:LogPara.Inlines.Add($r2)
    $script:LogPara.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
    $script:LogBox.ScrollToEnd()
}

function Set-Status {
    param([string]$Text, [string]$Type = "info")
    if ($null -eq $script:LblStatus) { return }
    $script:LblStatus.Text = $Text
    $script:StatusBar.Background = switch ($Type) {
        "success" { New-Brush "#107C10" }
        "error"   { New-Brush "#C42B1C" }
        default   { New-Brush "#0078D4" }
    }
    Update-UI
}

function Enter-Busy {
    param([string]$Text)
    $script:MainArea.IsEnabled = $false
    $script:HeaderPanel.IsEnabled = $false
    $script:Window.Cursor = [System.Windows.Input.Cursors]::Wait
    if ($Text) { Set-Status $Text }
}

function Exit-Busy {
    if ($script:LongOp) { return }
    $script:MainArea.IsEnabled = $true
    $script:HeaderPanel.IsEnabled = $true
    $script:Window.Cursor = $null
}

function Start-LongOp {
    param([string]$Text)
    Enter-Busy $Text
    $script:LongOp = $true
    $script:StopRequested = $false
    $script:BtnStop.IsEnabled = $true
    $script:BtnStop.Visibility = 'Visible'
}

function Stop-LongOp {
    $script:LongOp = $false
    $script:BtnStop.Visibility = 'Collapsed'
    Exit-Busy
}

function New-Table {
    param([string[]]$Columns)
    $t = New-Object System.Data.DataTable
    foreach ($c in $Columns) { [void]$t.Columns.Add($c) }
    return ,$t
}

function Set-GridRows {
    param($Grid, [string[]]$Columns, $Rows, [switch]$PassThru)
    $t = New-Table $Columns
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $dr = $t.NewRow()
        foreach ($c in $Columns) { $dr[$c] = [string]$r[$c] }
        [void]$t.Rows.Add($dr)
    }
    $Grid.ItemsSource = $t.DefaultView
    if ($PassThru) { return ,$t }
}

function Get-SelectedRows {
    param($Grid)
    return @($Grid.SelectedItems | ForEach-Object { $_.Row })
}

function ConvertTo-LikeLiteral {
    param([string]$Text)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        $s = [string]$ch
        if ('[', ']', '*', '%' -contains $s) { [void]$sb.Append("[$s]") }
        elseif ($s -eq "'") { [void]$sb.Append("''") }
        else { [void]$sb.Append($s) }
    }
    return $sb.ToString()
}

function ConvertTo-Hashtable {
    # PSCustomObject (uit ConvertFrom-Json) -> hashtable, recursief.
    # Arrays blijven arrays, ook met 0 of 1 element.
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $h = [ordered]@{}
        foreach ($k in @($InputObject.Keys)) { $h[[string]$k] = ConvertTo-Hashtable $InputObject[$k] }
        return $h
    }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) { $h[$p.Name] = ConvertTo-Hashtable $p.Value }
        return $h
    }
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $list = @(foreach ($i in $InputObject) { ,(ConvertTo-Hashtable $i) })
        return ,$list
    }
    return $InputObject
}

function Copy-Deep {
    param($Object)
    if ($null -eq $Object) { return $null }
    return (ConvertTo-Hashtable ($Object | ConvertTo-Json -Depth 30 | ConvertFrom-Json))
}

function ConvertTo-JsonText {
    param($Object)
    $json = $Object | ConvertTo-Json -Depth 30
    # PS 5.1 escapet tekens als ' < > &; voor de leesbaarheid terugzetten
    $json = $json -replace '\\u0027', "'" -replace '\\u003c', '<' -replace '\\u003e', '>' -replace '\\u0026', '&'
    return (Format-Json $json)
}

function Format-Json {
    # PS 5.1 ConvertTo-Json springt erg ruim in; netjes met 2 spaties
    param([string]$Json)
    $indent = 0
    $out = New-Object System.Text.StringBuilder
    foreach ($line in ($Json -split "`r?`n")) {
        $t = $line.Trim()
        if (-not $t) { continue }
        if ($t -match '^[\}\]]') { $indent = [Math]::Max(0, $indent - 1) }
        $t = $t -replace '":\s+', '": '
        [void]$out.AppendLine((' ' * (2 * $indent)) + $t)
        if ($t -match '[\{\[]$') { $indent++ }
    }
    return $out.ToString().TrimEnd()
}

function Get-Val {
    # Veilig een geneste waarde lezen: Get-Val $policy 'conditions.users.includeUsers'
    param($Object, [string]$Path)
    $cur = $Object
    foreach ($p in $Path.Split('.')) {
        if ($null -eq $cur) { return $null }
        if ($cur -is [System.Collections.IDictionary]) { $cur = $cur[$p] } else { $cur = $cur.$p }
    }
    return $cur
}

function Get-Arr {
    param($Object, [string]$Path)
    $v = Get-Val $Object $Path
    if ($null -eq $v) { return @() }
    return @($v | Where-Object { $null -ne $_ -and "$_" -ne '' })
}

function Test-Guid {
    param([string]$Text)
    $g = [guid]::Empty
    return [guid]::TryParse($Text, [ref]$g)
}

#endregion

#region --- Instellingen en back-ups ---

function Get-Settings {
    if ($script:Settings) { return $script:Settings }
    $file = Join-Path $script:AppDir 'settings.json'
    $s = $null
    if (Test-Path $file) {
        try { $s = ConvertTo-Hashtable (Get-Content $file -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { }
    }
    if (-not $s) { $s = [ordered]@{} }
    if (-not $s.Contains('LastAdmin'))  { $s.LastAdmin = '' }
    if (-not $s.Contains('LastTenant')) { $s.LastTenant = '' }
    if (-not $s.Contains('BreakGlass') -or $null -eq $s.BreakGlass) { $s.BreakGlass = [ordered]@{} }
    $script:Settings = $s
    return $s
}

function Save-Settings {
    try {
        if (-not (Test-Path $script:AppDir)) { [void](New-Item -ItemType Directory -Path $script:AppDir -Force) }
        $file = Join-Path $script:AppDir 'settings.json'
        $script:Settings | ConvertTo-Json -Depth 10 | Set-Content -Path $file -Encoding UTF8
    }
    catch { Write-Log "Instellingen opslaan mislukt: $_" "warning" }
}

function Get-BreakGlass {
    # Noodaccounts van de huidige tenant: lijst van @{ id; upn }
    $s = Get-Settings
    if (-not $script:TenantId -or -not $s.BreakGlass.Contains($script:TenantId)) { return @() }
    return @($s.BreakGlass[$script:TenantId] | Where-Object { $_ })
}

function Set-BreakGlass {
    param($List)
    $s = Get-Settings
    $s.BreakGlass[$script:TenantId] = @($List)
    Save-Settings
}

function Get-BackupDir {
    $d = Join-Path $script:AppDir "Backups\$($script:TenantId)"
    if (-not (Test-Path $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    return $d
}

function Save-Backup {
    # Slaat een object op voordat het gewijzigd of verwijderd wordt
    param($Object, [string]$Kind, [string]$Action)
    if ($null -eq $Object) { return }
    try {
        $name = ([string]$Object.displayName) -replace '[\\/:*?"<>|]', '_'
        if ($name.Length -gt 60) { $name = $name.Substring(0, 60) }
        $file = Join-Path (Get-BackupDir) ("{0}_{1}_{2}_{3}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $Kind, $Action, $name)
        $wrap = [ordered]@{ kind = $Kind; action = $Action; date = (Get-Date).ToString('s'); tenant = $script:TenantId; object = $Object }
        ConvertTo-JsonText $wrap | Set-Content -Path $file -Encoding UTF8
    }
    catch { Write-Log "Back-up maken mislukt: $_" "warning" }
}

#endregion


#region --- Modules en verbinding ---

function Initialize-Modules {
    # Vanaf module-versie 2.34.0 dwingt Microsoft.Graph.Authentication WAM
    # af (geen account-kiezer). 2.33.0 gebruikt nog de normale browser-login.
    try {
        Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -ErrorAction Stop
    }
    catch {
        try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
        catch { return $false }
    }
    # MSAL van de Graph-module laden voor de eigen login met account-kiezer
    $script:MsalAvailable = [bool]('Microsoft.Identity.Client.PublicClientApplicationBuilder' -as [type])
    if (-not $script:MsalAvailable) {
        $base = (Get-Module Microsoft.Graph.Authentication).ModuleBase
        $sub  = if ($PSVersionTable.PSEdition -eq 'Core') { 'Core' } else { 'Desktop' }
        $dll  = Join-Path $base "Dependencies\$sub\Microsoft.Identity.Client.dll"
        if (-not (Test-Path $dll)) {
            $dll = Get-ChildItem $base -Recurse -Filter 'Microsoft.Identity.Client.dll' -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -match "\\$sub\\" } | Select-Object -First 1 -ExpandProperty FullName
        }
        if ($dll) {
            try { Add-Type -Path $dll -ErrorAction Stop; $script:MsalAvailable = $true } catch { }
        }
    }
    return $true
}

function Wait-MsalTask {
    # Wacht op een MSAL-task terwijl de UI blijft reageren
    param($Task)
    while (-not $Task.IsCompleted) { Update-UI; Start-Sleep -Milliseconds 100 }
    if ($Task.IsCanceled) { throw "Inloggen geannuleerd of verlopen" }
    if ($Task.IsFaulted) { throw $Task.Exception.GetBaseException().Message }
    return $Task.Result
}

function Get-MsalApp {
    param([string]$Tenant)
    $authority = "https://login.microsoftonline.com/$(if ($Tenant) { $Tenant } else { 'organizations' })"
    if (-not $script:Msal -or $script:MsalAuthority -ne $authority) {
        # Geen token-cache op schijf: alleen in dit proces
        $script:Msal = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($script:GraphClientId).
            WithAuthority($authority).WithRedirectUri('http://localhost').Build()
        $script:MsalAuthority = $authority
    }
    return $script:Msal
}

function Connect-GraphWithToken {
    param($Result)
    $sec = ConvertTo-SecureString $Result.AccessToken -AsPlainText -Force
    Connect-MgGraph -AccessToken $sec -NoWelcome -ErrorAction Stop | Out-Null
    $script:Auth = @{ Account = $Result.Account; Expires = $Result.ExpiresOn.LocalDateTime; Scopes = @($Result.Scopes) }
}

function Connect-ToGraph {
    # Altijd de account-kiezer tonen, zodat je bewust met je admin-account
    # inlogt (ook als de browser al met je gewone account is aangemeld).
    param([string]$Tenant, [string]$LoginHint)
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    $script:Auth = $null
    if ($script:MsalAvailable) {
        $app = Get-MsalApp $Tenant
        # Eerder account in dit proces vergeten, anders kiest MSAL het stil opnieuw
        foreach ($a in @((Wait-MsalTask $app.GetAccountsAsync()))) { [void](Wait-MsalTask $app.RemoveAsync($a)) }
        $b = $app.AcquireTokenInteractive([string[]]$script:GraphScopes).
            WithUseEmbeddedWebView($false).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)
        if ($LoginHint) { $b = $b.WithLoginHint($LoginHint) }
        $cts = New-Object System.Threading.CancellationTokenSource ([TimeSpan]::FromMinutes(5))
        $r = Wait-MsalTask $b.ExecuteAsync($cts.Token)
        Connect-GraphWithToken $r
        return $r.Account.Username
    }
    # Terugval: standaard login van de module, token alleen in dit proces
    Write-Log "MSAL niet beschikbaar - standaard Graph-login (browser kiest mogelijk automatisch een account)" "warning"
    $p = @{ Scopes = $script:GraphScopes; NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
    if ($Tenant) { $p.TenantId = $Tenant }
    Connect-MgGraph @p | Out-Null
    return (Get-MgContext).Account
}

function Update-GraphToken {
    # Een sessie met -AccessToken ververst zichzelf niet: kort voor het
    # verlopen stil een nieuw token halen en opnieuw verbinden.
    if (-not $script:Auth -or (Get-Date) -lt $script:Auth.Expires.AddMinutes(-5)) { return }
    try {
        $r = Wait-MsalTask $script:Msal.AcquireTokenSilent([string[]]$script:GraphScopes, $script:Auth.Account).ExecuteAsync()
        Connect-GraphWithToken $r
        Write-Log "Graph-token vernieuwd (geldig tot $($script:Auth.Expires.ToString('HH:mm')))" "info"
    }
    catch {
        Write-Log "Token vernieuwen mislukt, klik zo nodig op 'Opnieuw inloggen': $_" "warning"
        $script:Auth.Expires = (Get-Date).AddHours(1)   # niet bij elke call opnieuw proberen
    }
}

function Disconnect-All {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    if ($script:Msal) {
        try { foreach ($a in @((Wait-MsalTask $script:Msal.GetAccountsAsync()))) { [void](Wait-MsalTask $script:Msal.RemoveAsync($a)) } } catch { }
    }
    $script:Auth = $null
    $script:Connected = $false
    $script:Me = $null
}

#endregion

#region --- Graph helpers ---

function Invoke-Graph {
    param([string]$Uri, [string]$Method = "GET", $Body = $null, [hashtable]$Headers = $null, [switch]$AllowNotFound, [switch]$Quiet)
    Update-GraphToken
    $short = $Uri -replace 'https://graph\.microsoft\.com/(v1\.0|beta)/', ''
    if (-not $Quiet) { Write-Log "$Method $short" "api" }
    $p = @{ Uri = $Uri; Method = $Method }
    if ($null -ne $Body) {
        $p.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 30 }
        $p.ContentType = "application/json; charset=utf-8"
    }
    if ($Headers) { $p.Headers = $Headers }
    try {
        return Invoke-MgGraphRequest @p
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails.Message) {
            try {
                $detail = ($_.ErrorDetails.Message | ConvertFrom-Json)
                if ($detail.error.message) { $errMsg = $detail.error.message }
                else { $errMsg = "$errMsg - $($_.ErrorDetails.Message)" }
            }
            catch { $errMsg = "$errMsg - $($_.ErrorDetails.Message)" }
        }
        if ($AllowNotFound -and $errMsg -match 'does not exist|NotFound|404') { return $null }
        Write-Log "FOUT bij $Method $short`: $errMsg" "error"
        throw $errMsg
    }
}

function Get-GraphPaged {
    param([string]$Uri, [hashtable]$Headers = $null, [int]$Max = 0)
    $all = New-Object System.Collections.ArrayList
    do {
        $r = Invoke-Graph -Uri $Uri -Headers $Headers
        foreach ($v in @($r.value)) { [void]$all.Add($v) }
        $Uri = $r['@odata.nextLink']
        if ($Max -and $all.Count -ge $Max) { break }
        if ($Uri) { Set-Status "$($all.Count) geladen..." }
    } while ($Uri -and -not $script:StopRequested)
    # Losse elementen teruggeven; aanroepers pakken zelf in met @()
    return $all.ToArray()
}

function Invoke-GraphBatch {
    # Losse GET-requests via /$batch (max 20 per aanroep).
    # Requests: @{ id; method; url (relatief) }. Geeft id -> response terug.
    param([System.Collections.IList]$Requests)
    $results = @{}
    $total = $Requests.Count
    for ($i = 0; $i -lt $total; $i += 20) {
        $todo = @($Requests[$i..([Math]::Min($i + 19, $total - 1))])
        for ($attempt = 1; $attempt -le 4 -and $todo.Count -gt 0; $attempt++) {
            $payload = @{ requests = @(foreach ($r in $todo) { [ordered]@{ id = [string]$r.id; method = $r.method; url = $r.url } }) }
            $resp = Invoke-Graph "$($script:GraphRoot)/`$batch" -Method POST -Body $payload -Quiet
            $retry = @()
            foreach ($x in @($resp.responses)) {
                if ([int]$x.status -eq 429 -and $attempt -lt 4) { $retry += @($todo | Where-Object { [string]$_.id -eq [string]$x.id }) }
                else { $results[[string]$x.id] = $x }
            }
            $todo = $retry
            if ($todo.Count -gt 0) { Start-Sleep -Seconds 3 }
        }
    }
    return $results
}

function Get-CleanObject {
    # Alleen-lezen velden en @odata-annotaties weghalen voor POST/PATCH
    param($Object, [string[]]$Drop = @('id', 'createdDateTime', 'modifiedDateTime', 'templateId', 'partialEnablementStrategy', 'deletedDateTime'))
    $c = Copy-Deep $Object
    foreach ($k in @($c.Keys)) {
        if ($Drop -contains $k -or ($k -like '*@odata*' -and $k -ne '@odata.type')) { $c.Remove($k) }
    }
    Remove-ODataContext $c
    return $c
}

function Remove-ODataContext {
    param($Node)
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($k in @($Node.Keys)) {
            if ($k -like '*@odata.context' -or $k -like '*@odata.id') { $Node.Remove($k); continue }
            Remove-ODataContext $Node[$k]
        }
    }
    elseif ($Node -is [System.Collections.IEnumerable] -and $Node -isnot [string]) {
        foreach ($i in $Node) { Remove-ODataContext $i }
    }
}

function Get-PolicyBody {
    # Schone body voor een CA-policy (POST of PATCH)
    param($Policy)
    $b = Get-CleanObject $Policy
    $gc = $b['grantControls']
    if ($gc -and $gc['authenticationStrength']) {
        $gc['authenticationStrength'] = [ordered]@{ id = [string]$gc['authenticationStrength']['id'] }
    }
    return $b
}

#endregion

#region --- Data laden en namen ---

function Get-AllCaData {
    Set-Status "Policies laden..."
    $script:Policies  = @(Get-GraphPaged "$($script:CaRoot)/policies")
    Set-Status "Named locations laden..."
    $script:Locations = @(Get-GraphPaged "$($script:CaRoot)/namedLocations")
    Set-Status "Authenticatiesterktes laden..."
    try { $script:Strengths = @(Get-GraphPaged "$($script:CaRoot)/authenticationStrength/policies") }
    catch { $script:Strengths = @() }
    try { $script:Contexts = @(Get-GraphPaged "$($script:CaRoot)/authenticationContextClassReferences") }
    catch { $script:Contexts = @() }
    foreach ($l in $script:Locations) { $script:Names[[string]$l.id] = [string]$l.displayName }
    foreach ($s in $script:Strengths) { $script:Names[[string]$s.id] = [string]$s.displayName }
    foreach ($c in $script:Contexts) { $script:Names["ctx:$($c.id)"] = [string]$c.displayName }
    Resolve-PolicyNames $script:Policies
    Write-Log "Geladen: $($script:Policies.Count) policies, $($script:Locations.Count) locaties, $($script:Strengths.Count) sterktes, $($script:Contexts.Count) contexten" "success"
}

function Resolve-PolicyNames {
    # Zoekt de namen op van alle gebruikers, groepen, rollen en apps in de policies
    param($PolicyList)
    $objIds = New-Object System.Collections.Generic.HashSet[string]
    $appIds = New-Object System.Collections.Generic.HashSet[string]
    foreach ($p in @($PolicyList)) {
        foreach ($path in 'conditions.users.includeUsers', 'conditions.users.excludeUsers', 'conditions.users.includeGroups', 'conditions.users.excludeGroups') {
            foreach ($id in (Get-Arr $p $path)) { if ((Test-Guid $id) -and -not $script:Names.ContainsKey($id)) { [void]$objIds.Add($id) } }
        }
        foreach ($path in 'conditions.applications.includeApplications', 'conditions.applications.excludeApplications') {
            foreach ($id in (Get-Arr $p $path)) { if ((Test-Guid $id) -and -not $script:Names.ContainsKey($id)) { [void]$appIds.Add($id) } }
        }
    }
    # Rollen: alle roldefinities 1x ophalen
    if (-not $script:RolesLoaded) {
        try {
            $roles = Get-GraphPaged "$($script:GraphRoot)/roleManagement/directory/roleDefinitions?`$select=id,displayName,templateId"
            $script:RoleList = @($roles | Sort-Object { $_.displayName })
            foreach ($r in $roles) { $script:Names[[string]$r.templateId] = [string]$r.displayName; $script:Names[[string]$r.id] = [string]$r.displayName }
            $script:RolesLoaded = $true
        }
        catch { foreach ($k in $script:AdminRoles.Keys) { $script:Names[$k] = $script:AdminRoles[$k] } }
    }
    $ids = @($objIds | Where-Object { -not $script:Names.ContainsKey($_) })
    for ($i = 0; $i -lt $ids.Count; $i += 1000) {
        $chunk = @($ids[$i..([Math]::Min($i + 999, $ids.Count - 1))])
        try {
            $r = Invoke-Graph "$($script:GraphRoot)/directoryObjects/getByIds" -Method POST -Body @{ ids = $chunk; types = @('user', 'group') } -Quiet
            foreach ($o in @($r.value)) {
                $n = if ($o.userPrincipalName) { "$($o.displayName) ($($o.userPrincipalName))" } else { [string]$o.displayName }
                $script:Names[[string]$o.id] = $n
            }
        }
        catch { }
    }
    $apps = @($appIds)
    if ($apps.Count) {
        $reqs = New-Object System.Collections.ArrayList
        $n = 0
        foreach ($a in $apps) { [void]$reqs.Add(@{ id = "$n"; method = 'GET'; url = "/servicePrincipals?`$filter=appId eq '$a'&`$select=appId,displayName" }); $n++ }
        try {
            $res = Invoke-GraphBatch $reqs
            foreach ($k in $res.Keys) {
                foreach ($sp in @($res[$k].body.value)) { $script:Names[[string]$sp.appId] = [string]$sp.displayName }
            }
        }
        catch { }
    }
}

function Get-Name {
    param([string]$Id)
    if (-not $Id) { return '' }
    if ($script:SpecialNames.ContainsKey($Id)) { return $script:SpecialNames[$Id] }
    if ($script:Names.ContainsKey($Id)) { return $script:Names[$Id] }
    return "$Id (onbekend/verwijderd)"
}

function Get-NameList {
    param($Ids, [int]$Max = 0)
    $list = @($Ids | ForEach-Object { Get-Name $_ })
    if ($Max -and $list.Count -gt $Max) { return (($list[0..($Max - 1)] -join ', ') + " +$($list.Count - $Max)") }
    return ($list -join ', ')
}

#endregion


#region --- Policies: samenvatting en evaluatie ---

function Get-UsersShort {
    param($P)
    $iu = Get-Arr $P 'conditions.users.includeUsers'
    $ig = Get-Arr $P 'conditions.users.includeGroups'
    $ir = Get-Arr $P 'conditions.users.includeRoles'
    $parts = @()
    if ($iu -contains 'All') { $parts += 'Alle gebruikers' }
    elseif ($iu -contains 'None') { $parts += 'Geen' }
    else {
        $u = @($iu | Where-Object { $_ -ne 'GuestsOrExternalUsers' })
        if ($u.Count) { $parts += "$($u.Count) gebr." }
        if ($iu -contains 'GuestsOrExternalUsers' -or (Get-Val $P 'conditions.users.includeGuestsOrExternalUsers')) { $parts += 'gasten' }
    }
    if ($ig.Count) { $parts += "$($ig.Count) groep(en)" }
    if ($ir.Count) { $parts += "$($ir.Count) rol(len)" }
    if (-not $parts.Count -and (Get-Val $P 'conditions.clientApplications')) { $parts += 'Workload identities' }
    $ex = (Get-Arr $P 'conditions.users.excludeUsers').Count + (Get-Arr $P 'conditions.users.excludeGroups').Count + (Get-Arr $P 'conditions.users.excludeRoles').Count
    $txt = $parts -join ', '
    if ($ex) { $txt += " (excl. $ex)" }
    return $txt
}

function Get-AppsShort {
    param($P)
    $ia = Get-Arr $P 'conditions.applications.includeApplications'
    $ua = Get-Arr $P 'conditions.applications.includeUserActions'
    $ac = Get-Arr $P 'conditions.applications.includeAuthenticationContextClassReferences'
    if ($ua.Count) { return "Actie: $($ua -join ', ')" }
    if ($ac.Count) { return "Auth.context: $(($ac | ForEach-Object { Get-Name "ctx:$_" }) -join ', ')" }
    if ($ia -contains 'All') { $t = 'Alle resources' }
    elseif ($ia.Count -eq 1) { $t = Get-Name $ia[0] }
    else { $t = "$($ia.Count) apps" }
    $ex = (Get-Arr $P 'conditions.applications.excludeApplications').Count
    if ($ex) { $t += " (excl. $ex)" }
    return $t
}

function Get-GrantShort {
    param($P)
    $gc = Get-Val $P 'grantControls'
    $parts = @()
    if ($gc) {
        $bc = Get-Arr $gc 'builtInControls'
        if ($bc -contains 'block') { return 'BLOKKEREN' }
        foreach ($c in $bc) { $parts += $(if ($script:GrantLabels.Contains($c)) { $script:GrantLabels[$c] } else { $c }) }
        $as = Get-Val $gc 'authenticationStrength'
        if ($as -and $as.id) { $parts += "Sterkte: $(Get-Name ([string]$as.id))" }
        if ((Get-Arr $gc 'termsOfUse').Count) { $parts += 'Gebruiksvoorwaarden' }
        $op = if ($gc.operator -eq 'AND') { ' EN ' } else { ' OF ' }
        $txt = $parts -join $op
    }
    else { $txt = '' }
    $sc = Get-Val $P 'sessionControls'
    $s = @()
    if ($sc) {
        if ((Get-Val $sc 'signInFrequency.isEnabled')) {
            $f = $sc.signInFrequency
            $s += $(if ($f.frequencyInterval -eq 'everyTime') { 'aanmelden elke keer' } else { "aanmelden elke $($f.value) $($f.type)" })
        }
        if ((Get-Val $sc 'persistentBrowser.isEnabled')) { $s += "browser: $($sc.persistentBrowser.mode)" }
        if ((Get-Val $sc 'applicationEnforcedRestrictions.isEnabled')) { $s += 'app-beperkingen' }
        if ((Get-Val $sc 'cloudAppSecurity.isEnabled')) { $s += 'Defender for Cloud Apps' }
    }
    if ($s.Count) { $txt = (@($txt, "Sessie: $($s -join ', ')") | Where-Object { $_ }) -join ' | ' }
    return $txt
}

function Get-PolicySummary {
    # Leesbare uitleg van een policy voor het detailpaneel
    param($P)
    $sb = New-Object System.Text.StringBuilder
    $add = { param($t) [void]$sb.AppendLine($t) }
    $list = {
        param([string]$Label, $Ids, [string]$Prefix = '')
        $ids = @($Ids)
        if ($ids.Count) {
            [void]$sb.AppendLine("  ${Label}:")
            foreach ($i in $ids) { [void]$sb.AppendLine("    - " + (Get-Name "$Prefix$i")) }
        }
    }
    & $add "$($P.displayName)"
    & $add ("Status: {0}    Gewijzigd: {1}" -f $script:StateLabels[[string]$P.state], $(if ($P.modifiedDateTime) { ([datetime]$P.modifiedDateTime).ToLocalTime().ToString('dd-MM-yyyy HH:mm') } else { '-' }))
    & $add "ID: $($P.id)"
    & $add ""
    & $add "GEBRUIKERS"
    & $list 'Opnemen gebruikers' (Get-Arr $P 'conditions.users.includeUsers')
    & $list 'Opnemen groepen' (Get-Arr $P 'conditions.users.includeGroups')
    & $list 'Opnemen rollen' (Get-Arr $P 'conditions.users.includeRoles')
    if (Get-Val $P 'conditions.users.includeGuestsOrExternalUsers') { & $add "  Opnemen gasten/externen: $((Get-Val $P 'conditions.users.includeGuestsOrExternalUsers.guestOrExternalUserTypes'))" }
    & $list 'Uitsluiten gebruikers' (Get-Arr $P 'conditions.users.excludeUsers')
    & $list 'Uitsluiten groepen' (Get-Arr $P 'conditions.users.excludeGroups')
    & $list 'Uitsluiten rollen' (Get-Arr $P 'conditions.users.excludeRoles')
    if (Get-Val $P 'conditions.users.excludeGuestsOrExternalUsers') { & $add "  Uitsluiten gasten/externen: $((Get-Val $P 'conditions.users.excludeGuestsOrExternalUsers.guestOrExternalUserTypes'))" }
    $bg = @(Get-BreakGlass)
    if ($bg.Count) {
        $missing = @($bg | Where-Object { (Get-Arr $P 'conditions.users.excludeUsers') -notcontains $_.id })
        if ($missing.Count) { & $add "  LET OP: noodaccount(s) NIET uitgesloten: $(($missing | ForEach-Object { $_.upn }) -join ', ')" }
        else { & $add "  Noodaccounts zijn uitgesloten." }
    }
    & $add ""
    & $add "DOELRESOURCES"
    & $list 'Opnemen' (Get-Arr $P 'conditions.applications.includeApplications')
    & $list 'Uitsluiten' (Get-Arr $P 'conditions.applications.excludeApplications')
    & $list 'Gebruikersacties' (Get-Arr $P 'conditions.applications.includeUserActions')
    & $list 'Authenticatiecontext' (Get-Arr $P 'conditions.applications.includeAuthenticationContextClassReferences') 'ctx:'
    $af = Get-Val $P 'conditions.applications.applicationFilter'
    if ($af) { & $add "  App-filter ($($af.mode)): $($af.rule)" }
    & $add ""
    & $add "VOORWAARDEN"
    $cat = Get-Arr $P 'conditions.clientAppTypes'
    if ($cat.Count -and $cat -notcontains 'all') { & $add "  Client-apps: $(($cat | ForEach-Object { if ($script:ClientAppTypes.Contains($_)) { $script:ClientAppTypes[$_] } else { $_ } }) -join ', ')" }
    $pi = Get-Arr $P 'conditions.platforms.includePlatforms'
    if ($pi.Count) { & $add "  Platforms: $($pi -join ', ')$(if ((Get-Arr $P 'conditions.platforms.excludePlatforms').Count) { ' - excl. ' + ((Get-Arr $P 'conditions.platforms.excludePlatforms') -join ', ') })" }
    & $list 'Locaties opnemen' (Get-Arr $P 'conditions.locations.includeLocations')
    & $list 'Locaties uitsluiten' (Get-Arr $P 'conditions.locations.excludeLocations')
    $ur = Get-Arr $P 'conditions.userRiskLevels'
    $sr = Get-Arr $P 'conditions.signInRiskLevels'
    if ($ur.Count) { & $add "  Gebruikersrisico: $($ur -join ', ')" }
    if ($sr.Count) { & $add "  Aanmeldrisico: $($sr -join ', ')" }
    $df = Get-Val $P 'conditions.devices.deviceFilter'
    if ($df) { & $add "  Apparaatfilter ($($df.mode)): $($df.rule)" }
    $tm = Get-Val $P 'conditions.authenticationFlows.transferMethods'
    if ($tm) { & $add "  Authenticatiestromen: $tm" }
    & $add ""
    & $add "TOEGANG"
    & $add "  $(Get-GrantShort $P)"
    return $sb.ToString()
}

function Test-InScope {
    # Valt een gebruiker (id, groepen, rollen, gast) binnen include/exclude
    # van de gebruikersvoorwaarde? Geeft @{ Result; Reason } terug.
    param($P, $U)
    $c = Get-Val $P 'conditions.users'
    if (-not $c) { return @{ Result = 'Niet van toepassing'; Reason = 'Geen gebruikersvoorwaarde (workload identities)' } }
    $why = { param($list, $kind)
        foreach ($i in @($list)) {
            if ($kind -eq 'user' -and ($i -eq 'All' -or $i -eq $U.id)) { return $(if ($i -eq 'All') { 'Alle gebruikers' } else { 'Direct' }) }
            if ($kind -eq 'user' -and $i -eq 'GuestsOrExternalUsers' -and $U.isGuest) { return 'Gast' }
            if ($kind -eq 'group' -and $U.groups -contains $i) { return "Groep $(Get-Name $i)" }
            if ($kind -eq 'role' -and $U.roles -contains $i) { return "Rol $(Get-Name $i)" }
        }
        return $null
    }
    $ex = (& $why (Get-Arr $c 'excludeUsers') 'user'), (& $why (Get-Arr $c 'excludeGroups') 'group'), (& $why (Get-Arr $c 'excludeRoles') 'role') | Where-Object { $_ }
    if (-not $ex -and $U.isGuest -and (Get-Val $c 'excludeGuestsOrExternalUsers')) { $ex = @('Gast/extern') }
    $in = (& $why (Get-Arr $c 'includeUsers') 'user'), (& $why (Get-Arr $c 'includeGroups') 'group'), (& $why (Get-Arr $c 'includeRoles') 'role') | Where-Object { $_ }
    if (-not $in -and $U.isGuest -and (Get-Val $c 'includeGuestsOrExternalUsers')) { $in = @('Gast/extern') }
    if ($ex) { return @{ Result = 'Uitgesloten'; Reason = (@($ex) -join '; ') } }
    if ($in) { return @{ Result = 'Van toepassing'; Reason = (@($in) -join '; ') } }
    return @{ Result = 'Niet van toepassing'; Reason = '' }
}

function Get-UserContext {
    # Gebruiker + alle (geneste) groepen en actieve rollen voor de evaluatie
    param([string]$UserId)
    $u = Invoke-Graph "$($script:GraphRoot)/users/$($UserId)?`$select=id,displayName,userPrincipalName,userType,accountEnabled"
    $groups = New-Object System.Collections.ArrayList
    $roles = New-Object System.Collections.ArrayList
    foreach ($m in (Get-GraphPaged "$($script:GraphRoot)/users/$UserId/transitiveMemberOf?`$select=id,roleTemplateId")) {
        if ($m['@odata.type'] -eq '#microsoft.graph.group') { [void]$groups.Add([string]$m.id) }
        elseif ($m['@odata.type'] -eq '#microsoft.graph.directoryRole') { [void]$roles.Add([string]$m.roleTemplateId) }
    }
    return @{
        id = [string]$u.id; upn = [string]$u.userPrincipalName; displayName = [string]$u.displayName
        isGuest = ($u.userType -eq 'Guest'); groups = @($groups); roles = @($roles)
    }
}

function Test-SelfLockout {
    # Waarschuwt als een policy die aan staat de ingelogde admin blokkeert
    # voor alle resources (bijv. vergeten jezelf/noodaccount uit te sluiten).
    param($P)
    $warn = @(Get-LockoutWarnings $P)
    if (-not $warn.Count) { return $true }
    $r = Show-Msg "$($warn -join "`n`n")`n`nToch doorgaan en de policy AAN zetten?" "Pas op: buitensluiten" "YesNo" "Warning"
    return ($r -eq 'Yes')
}

function Get-LockoutWarnings {
    param($P)
    if ($P.state -ne 'enabled' -or -not $script:Me) { return @() }
    $gc = Get-Val $P 'grantControls'
    $block = (Get-Arr $gc 'builtInControls') -contains 'block'
    $allApps = (Get-Arr $P 'conditions.applications.includeApplications') -contains 'All' -or
               (Get-Arr $P 'conditions.applications.includeApplications') -contains 'MicrosoftAdminPortals'
    $scope = Test-InScope $P $script:Me
    $warn = @()
    if ($block -and $allApps -and $scope.Result -eq 'Van toepassing') {
        $warn += "Deze policy BLOKKEERT jouw eigen account ($($script:Me.upn)) voor $(if ((Get-Arr $P 'conditions.applications.includeApplications') -contains 'All') { 'alle resources' } else { 'de beheerportals' })."
    }
    $bg = @(Get-BreakGlass)
    if ((Get-Arr $P 'conditions.users.includeUsers') -contains 'All' -and $bg.Count) {
        $missing = @($bg | Where-Object { (Get-Arr $P 'conditions.users.excludeUsers') -notcontains $_.id })
        if ($missing.Count) { $warn += "Noodaccount(s) niet uitgesloten: $(($missing | ForEach-Object { $_.upn }) -join ', ')" }
    }
    elseif ((Get-Arr $P 'conditions.users.includeUsers') -contains 'All' -and -not $bg.Count) {
        $warn += "Er zijn voor deze tenant geen noodaccounts ingesteld (tab Back-up / import)."
    }
    return $warn
}

#endregion


#region --- Kiezers en kleine dialogen ---

$script:KindLabels = [ordered]@{ user = 'Gebruikers'; group = 'Groepen'; role = 'Rollen'; app = 'Apps'; location = 'Locaties' }

function Search-Objects {
    param([string]$Kind, [string]$Text)
    $t = $Text.Trim()
    $hdr = @{ ConsistencyLevel = 'eventual' }
    $q = $t.Replace('"', '')
    $out = @()
    switch ($Kind) {
        'user' {
            $uri = if ($q) { "$($script:GraphRoot)/users?`$search=`"displayName:$q`" OR `"userPrincipalName:$q`"&`$select=id,displayName,userPrincipalName&`$top=50&`$orderby=displayName" }
                   else { "$($script:GraphRoot)/users?`$select=id,displayName,userPrincipalName&`$top=50&`$orderby=displayName&`$count=true" }
            $r = Invoke-Graph $uri -Headers $hdr -Quiet
            $out = @($r.value | ForEach-Object { @{ id = [string]$_.id; kind = 'user'; name = "$($_.displayName) ($($_.userPrincipalName))" } })
        }
        'group' {
            $uri = if ($q) { "$($script:GraphRoot)/groups?`$search=`"displayName:$q`"&`$select=id,displayName,mail&`$top=50&`$orderby=displayName" }
                   else { "$($script:GraphRoot)/groups?`$select=id,displayName,mail&`$top=50&`$orderby=displayName&`$count=true" }
            $r = Invoke-Graph $uri -Headers $hdr -Quiet
            $out = @($r.value | ForEach-Object { @{ id = [string]$_.id; kind = 'group'; name = [string]$_.displayName } })
        }
        'role' {
            $roles = if ($script:RoleList) { $script:RoleList } else { @($script:AdminRoles.Keys | ForEach-Object { @{ templateId = $_; displayName = $script:AdminRoles[$_] } }) }
            $out = @($roles | Where-Object { -not $q -or $_.displayName -like "*$q*" } | ForEach-Object { @{ id = [string]$_.templateId; kind = 'role'; name = [string]$_.displayName } })
        }
        'app' {
            $out += @(@{ id = 'Office365'; kind = 'app'; name = 'Office 365' }, @{ id = 'MicrosoftAdminPortals'; kind = 'app'; name = 'Microsoft-beheerportals' } |
                Where-Object { -not $q -or $_.name -like "*$q*" })
            $uri = if ($q) { "$($script:GraphRoot)/servicePrincipals?`$search=`"displayName:$q`"&`$select=appId,displayName&`$top=50" }
                   else { "$($script:GraphRoot)/servicePrincipals?`$select=appId,displayName&`$top=50&`$orderby=displayName&`$count=true" }
            $r = Invoke-Graph $uri -Headers $hdr -Quiet
            $out += @($r.value | ForEach-Object { @{ id = [string]$_.appId; kind = 'app'; name = [string]$_.displayName } })
        }
        'location' {
            $out += @(@{ id = 'AllTrusted'; kind = 'location'; name = 'Alle vertrouwde locaties' })
            $out += @($script:Locations | Sort-Object { $_.displayName } | ForEach-Object { @{ id = [string]$_.id; kind = 'location'; name = [string]$_.displayName } })
            if ($q) { $out = @($out | Where-Object { $_.name -like "*$q*" }) }
        }
    }
    foreach ($o in $out) { $script:Names[$o.id] = $o.name }
    return $out
}

function Show-ObjectPicker {
    # Modale zoekdialoog. Alle handlers vuren tijdens ShowDialog(), dus lokale
    # variabelen zijn bereikbaar. Geeft een lijst @{ id; kind; name } terug.
    param([string[]]$Kinds = @('user'), [string]$Title = 'Kiezen', [switch]$Single, $Owner = $script:Window)
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="620" Height="560" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <DockPanel Margin="12">
    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
      <ComboBox x:Name="CmbKind" DockPanel.Dock="Left" Width="120" Margin="0,0,8,0" DisplayMemberPath="Label" SelectedValuePath="Kind"/>
      <Button x:Name="BtnFind" DockPanel.Dock="Right" Content="Zoeken" Margin="8,0,0,0"/>
      <TextBox x:Name="TxtFind" Height="26" VerticalContentAlignment="Center"/>
    </DockPanel>
    <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}" Text="Typ (een deel van) de naam en druk op Enter. Ctrl/Shift voor meerdere."/>
    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
      <Button x:Name="BtnOk" Content="Toevoegen" Style="{StaticResource BtnGreen}"/>
      <Button x:Name="BtnCancel" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True"/>
    </StackPanel>
    <ListBox x:Name="LstRes" SelectionMode="Extended" DisplayMemberPath="name"/>
  </DockPanel>
</Window>
'@
    $w = ConvertFrom-Xaml $xaml.Replace('{{STYLES}}', $script:StylesXaml)
    $w.Title = $Title
    if ($Owner) { $w.Owner = $Owner }
    $cmb = $w.FindName('CmbKind'); $txt = $w.FindName('TxtFind'); $lst = $w.FindName('LstRes')
    if ($Single) { $lst.SelectionMode = 'Single' }
    $cmb.ItemsSource = @($Kinds | ForEach-Object { [pscustomobject]@{ Kind = $_; Label = $script:KindLabels[$_] } })
    $cmb.SelectedIndex = 0
    if ($Kinds.Count -eq 1) { $cmb.Visibility = 'Collapsed' }
    $find = {
        try {
            $w.Cursor = [System.Windows.Input.Cursors]::Wait
            $res = @(Search-Objects $cmb.SelectedValue $txt.Text)
            $lst.ItemsSource = @($res | ForEach-Object { [pscustomobject]$_ })
        }
        catch { [System.Windows.MessageBox]::Show($w, "Zoeken mislukt:`n$_", 'Fout', 'OK', 'Error') | Out-Null }
        finally { $w.Cursor = $null }
    }
    $w.FindName('BtnFind').Add_Click($find)
    $txt.Add_KeyDown({ if ($_.Key -eq 'Return') { & $find } })
    $cmb.Add_SelectionChanged({ & $find })
    $lst.Add_MouseDoubleClick({ if ($lst.SelectedItem) { $w.DialogResult = $true } })
    $w.FindName('BtnOk').Add_Click({ if ($lst.SelectedItems.Count) { $w.DialogResult = $true } })
    $w.Add_ContentRendered({ $txt.Focus(); & $find })
    if ($w.ShowDialog()) {
        return @($lst.SelectedItems | ForEach-Object { @{ id = $_.id; kind = $_.kind; name = $_.name } })
    }
    return @()
}

function Show-InputDialog {
    # Eenvoudig formulier met velden: @( @{ Key; Label; Value; Multi; Options } )
    # Options = lijst -> ComboBox. Geeft hashtable Key -> waarde of $null.
    param([string]$Title, $Fields, [string]$Hint = '', [int]$Width = 520, $Owner = $script:Window)
    $w = New-Object System.Windows.Window
    $w.Title = $Title; $w.Width = $Width; $w.SizeToContent = 'Height'; $w.WindowStartupLocation = 'CenterOwner'
    $w.FontFamily = 'Segoe UI'; $w.FontSize = 12; $w.Background = New-Brush '#F5F5F5'; $w.ResizeMode = 'NoResize'
    if ($Owner) { $w.Owner = $Owner }
    $w.Resources = (ConvertFrom-Xaml "<ResourceDictionary xmlns=`"http://schemas.microsoft.com/winfx/2006/xaml/presentation`" xmlns:x=`"http://schemas.microsoft.com/winfx/2006/xaml`">$($script:StylesXaml)</ResourceDictionary>")
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = '14'
    if ($Hint) {
        $h = New-Object System.Windows.Controls.TextBlock
        $h.Text = $Hint; $h.TextWrapping = 'Wrap'; $h.Foreground = New-Brush '#555555'; $h.Margin = '0,0,0,10'
        [void]$sp.Children.Add($h)
    }
    $ctl = @{}
    foreach ($f in $Fields) {
        $l = New-Object System.Windows.Controls.TextBlock
        $l.Text = $f.Label; $l.Margin = '0,0,0,3'
        [void]$sp.Children.Add($l)
        if ($f.Options) {
            $c = New-Object System.Windows.Controls.ComboBox
            $c.ItemsSource = @($f.Options); $c.Margin = '0,0,0,8'
            $c.SelectedItem = $(if ($f.Value) { $f.Value } else { @($f.Options)[0] })
        }
        elseif ($f.Check) {
            $c = New-Object System.Windows.Controls.CheckBox
            $c.Content = $f.Check; $c.IsChecked = [bool]$f.Value; $c.Margin = '0,0,0,8'
        }
        else {
            $c = New-Object System.Windows.Controls.TextBox
            $c.Text = [string]$f.Value; $c.Margin = '0,0,0,8'; $c.Padding = '4,3'
            if ($f.Multi) {
                $c.AcceptsReturn = $true; $c.Height = 150; $c.TextWrapping = 'NoWrap'
                $c.VerticalScrollBarVisibility = 'Auto'; $c.FontFamily = 'Consolas'
            }
        }
        $ctl[$f.Key] = $c
        [void]$sp.Children.Add($c)
    }
    $bp = New-Object System.Windows.Controls.StackPanel
    $bp.Orientation = 'Horizontal'; $bp.HorizontalAlignment = 'Right'; $bp.Margin = '0,6,0,0'
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = 'OK'; $ok.Style = $w.Resources['BtnGreen']; $ok.IsDefault = -not ($Fields | Where-Object { $_.Multi })
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = 'Annuleren'; $cancel.Style = $w.Resources['BtnGray']; $cancel.IsCancel = $true
    [void]$bp.Children.Add($ok); [void]$bp.Children.Add($cancel)
    [void]$sp.Children.Add($bp)
    $w.Content = $sp
    $ok.Add_Click({ $w.DialogResult = $true })
    if (-not $w.ShowDialog()) { return $null }
    $res = @{}
    foreach ($k in $ctl.Keys) {
        $c = $ctl[$k]
        $res[$k] = if ($c -is [System.Windows.Controls.ComboBox]) { [string]$c.SelectedItem }
                   elseif ($c -is [System.Windows.Controls.CheckBox]) { [bool]$c.IsChecked }
                   else { $c.Text }
    }
    return $res
}

function Show-TextWindow {
    param([string]$Title, [string]$Text)
    $w = New-Object System.Windows.Window
    $w.Title = $Title; $w.Width = 820; $w.Height = 640; $w.WindowStartupLocation = 'CenterOwner'; $w.Owner = $script:Window
    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Text = $Text; $tb.IsReadOnly = $true; $tb.FontFamily = 'Consolas'; $tb.FontSize = 12
    $tb.VerticalScrollBarVisibility = 'Auto'; $tb.HorizontalScrollBarVisibility = 'Auto'; $tb.Padding = '6'
    $w.Content = $tb
    [void]$w.ShowDialog()
}

#endregion


#region --- Policy-editor ---

$script:EditorXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="980" Height="780" MinWidth="820" MinHeight="600" WindowStartupLocation="CenterOwner"
        FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>
{{STYLES}}
<Style x:Key="IdList" TargetType="ListBox">
  <Setter Property="Height" Value="110"/>
  <Setter Property="SelectionMode" Value="Extended"/>
  <Setter Property="DisplayMemberPath" Value="name"/>
  <Setter Property="Margin" Value="0,0,0,4"/>
</Style>
<Style x:Key="Small" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Padding" Value="10,3"/><Setter Property="Margin" Value="0,0,6,8"/>
</Style>
<Style x:Key="SmallGray" TargetType="Button" BasedOn="{StaticResource Small}"><Setter Property="Background" Value="#7A7A7A"/></Style>
  </Window.Resources>
  <DockPanel Margin="12">
    <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
      <Grid.ColumnDefinitions><ColumnDefinition Width="60"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="170"/></Grid.ColumnDefinitions>
      <TextBlock Text="Naam" Style="{StaticResource FieldLabel}"/>
      <TextBox x:Name="TxtName" Grid.Column="1" Style="{StaticResource FieldBox}"/>
      <TextBlock Grid.Column="2" Text="Status" Style="{StaticResource FieldLabel}" Margin="14,0,8,6"/>
      <ComboBox x:Name="CmbState" Grid.Column="3" Height="26" Margin="0,0,0,6" DisplayMemberPath="Label" SelectedValuePath="Value"/>
    </Grid>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="BtnSave" Content="Opslaan in tenant" Style="{StaticResource BtnGreen}"/>
        <Button x:Name="BtnCancel" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True"/>
      </StackPanel>
      <TextBlock x:Name="LblHint" Foreground="Gray" VerticalAlignment="Center" TextWrapping="Wrap"
                 Text="Velden die het formulier niet kent, blijven behouden. Op tab JSON kun je alles direct aanpassen."/>
    </DockPanel>
    <TabControl x:Name="Tabs">
      <TabItem Header="Gebruikers">
        <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel Margin="10">
          <GroupBox Header="Opnemen"><StackPanel>
            <StackPanel Orientation="Horizontal">
              <RadioButton x:Name="RbUNone" Content="Geen" Margin="0,0,18,6"/>
              <RadioButton x:Name="RbUAll" Content="Alle gebruikers" Margin="0,0,18,6"/>
              <RadioButton x:Name="RbUSel" Content="Gebruikers, groepen en rollen selecteren" Margin="0,0,18,6"/>
            </StackPanel>
            <StackPanel x:Name="PnlUSel">
              <CheckBox x:Name="ChkUGuests" Content="Gasten en externe gebruikers (alle typen)" Margin="0,0,0,6"/>
              <ListBox x:Name="LstUInc" Style="{StaticResource IdList}"/>
              <WrapPanel>
                <Button x:Name="BtnUIncAdd" Content="Toevoegen..." Style="{StaticResource Small}"/>
                <Button x:Name="BtnUIncAdmins" Content="Beheerrollen toevoegen" Style="{StaticResource Small}"/>
                <Button x:Name="BtnUIncDel" Content="Verwijderen" Style="{StaticResource SmallGray}"/>
              </WrapPanel>
            </StackPanel>
            <TextBlock x:Name="LblGuestObj" Style="{StaticResource Hint}" Visibility="Collapsed"
                       Text="Deze policy gebruikt een uitgebreide gasten-instelling (includeGuestsOrExternalUsers); die blijft behouden en is via JSON aan te passen."/>
          </StackPanel></GroupBox>
          <GroupBox Header="Uitsluiten"><StackPanel>
            <CheckBox x:Name="ChkXGuests" Content="Gasten en externe gebruikers (alle typen)" Margin="0,0,0,6"/>
            <ListBox x:Name="LstUExc" Style="{StaticResource IdList}"/>
            <WrapPanel>
              <Button x:Name="BtnUExcAdd" Content="Toevoegen..." Style="{StaticResource Small}"/>
              <Button x:Name="BtnUExcBg" Content="Noodaccounts toevoegen" Style="{StaticResource Small}"/>
              <Button x:Name="BtnUExcMe" Content="Mijzelf toevoegen" Style="{StaticResource Small}"/>
              <Button x:Name="BtnUExcDel" Content="Verwijderen" Style="{StaticResource SmallGray}"/>
            </WrapPanel>
          </StackPanel></GroupBox>
        </StackPanel></ScrollViewer>
      </TabItem>
      <TabItem Header="Doelresources">
        <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel Margin="10">
          <GroupBox Header="Waarop is de policy van toepassing"><StackPanel>
            <RadioButton x:Name="RbAAll" Content="Alle resources (voorheen 'alle cloud-apps')"/>
            <RadioButton x:Name="RbASel" Content="Geselecteerde apps"/>
            <RadioButton x:Name="RbAAction" Content="Gebruikersactie"/>
            <RadioButton x:Name="RbACtx" Content="Authenticatiecontext"/>
            <RadioButton x:Name="RbANone" Content="Geen"/>
          </StackPanel></GroupBox>
          <GroupBox x:Name="GbASel" Header="Geselecteerde apps"><StackPanel>
            <ListBox x:Name="LstAInc" Style="{StaticResource IdList}"/>
            <WrapPanel>
              <Button x:Name="BtnAIncAdd" Content="Toevoegen..." Style="{StaticResource Small}"/>
              <Button x:Name="BtnAIncDel" Content="Verwijderen" Style="{StaticResource SmallGray}"/>
            </WrapPanel>
          </StackPanel></GroupBox>
          <GroupBox x:Name="GbAExc" Header="Apps uitsluiten"><StackPanel>
            <ListBox x:Name="LstAExc" Style="{StaticResource IdList}"/>
            <WrapPanel>
              <Button x:Name="BtnAExcAdd" Content="Toevoegen..." Style="{StaticResource Small}"/>
              <Button x:Name="BtnAExcDel" Content="Verwijderen" Style="{StaticResource SmallGray}"/>
            </WrapPanel>
          </StackPanel></GroupBox>
          <GroupBox x:Name="GbAAction" Header="Gebruikersactie"><StackPanel>
            <RadioButton x:Name="RbActReg" Content="Beveiligingsgegevens registreren (registerSecurityInformation)"/>
            <RadioButton x:Name="RbActDev" Content="Apparaten registreren of joinen (registerOrJoinDevices)"/>
          </StackPanel></GroupBox>
          <GroupBox x:Name="GbACtx" Header="Authenticatiecontext"><StackPanel x:Name="PnlCtx"/></GroupBox>
        </StackPanel></ScrollViewer>
      </TabItem>
      <TabItem Header="Voorwaarden">
        <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel Margin="10">
          <GroupBox Header="Client-apps (niets aangevinkt = alle)"><WrapPanel x:Name="PnlClientApps"/></GroupBox>
          <GroupBox Header="Apparaatplatforms"><StackPanel>
            <CheckBox x:Name="ChkPlat" Content="Platformvoorwaarde gebruiken" Margin="0,0,0,6"/>
            <StackPanel x:Name="PnlPlat">
              <TextBlock Text="Opnemen (niets aangevinkt = alle platforms):" Margin="0,0,0,4"/>
              <WrapPanel x:Name="PnlPlatInc" Margin="0,0,0,6"/>
              <TextBlock Text="Uitsluiten:" Margin="0,0,0,4"/>
              <WrapPanel x:Name="PnlPlatExc"/>
            </StackPanel>
          </StackPanel></GroupBox>
          <GroupBox Header="Locaties"><StackPanel>
            <CheckBox x:Name="ChkLoc" Content="Locatievoorwaarde gebruiken" Margin="0,0,0,6"/>
            <StackPanel x:Name="PnlLoc">
              <StackPanel Orientation="Horizontal">
                <RadioButton x:Name="RbLAll" Content="Alle locaties" Margin="0,0,18,6"/>
                <RadioButton x:Name="RbLSel" Content="Geselecteerde locaties" Margin="0,0,18,6"/>
              </StackPanel>
              <ListBox x:Name="LstLInc" Style="{StaticResource IdList}" Height="80"/>
              <WrapPanel>
                <Button x:Name="BtnLIncAdd" Content="Toevoegen..." Style="{StaticResource Small}"/>
                <Button x:Name="BtnLIncDel" Content="Verwijderen" Style="{StaticResource SmallGray}"/>
              </WrapPanel>
              <TextBlock Text="Uitsluiten:" Margin="0,0,0,4"/>
              <ListBox x:Name="LstLExc" Style="{StaticResource IdList}" Height="80"/>
              <WrapPanel>
                <Button x:Name="BtnLExcAdd" Content="Toevoegen..." Style="{StaticResource Small}"/>
                <Button x:Name="BtnLExcDel" Content="Verwijderen" Style="{StaticResource SmallGray}"/>
              </WrapPanel>
            </StackPanel>
          </StackPanel></GroupBox>
          <GroupBox Header="Risico (vereist Entra ID P2)"><Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="140"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
            <TextBlock Text="Gebruikersrisico" Style="{StaticResource FieldLabel}"/>
            <WrapPanel x:Name="PnlURisk" Grid.Column="1"/>
            <TextBlock Grid.Row="1" Text="Aanmeldrisico" Style="{StaticResource FieldLabel}"/>
            <WrapPanel x:Name="PnlSRisk" Grid.Row="1" Grid.Column="1"/>
          </Grid></GroupBox>
          <GroupBox Header="Apparaatfilter"><Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="140"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
            <TextBlock Text="Modus" Style="{StaticResource FieldLabel}"/>
            <ComboBox x:Name="CmbDevMode" Grid.Column="1" Width="200" HorizontalAlignment="Left" Height="26" Margin="0,0,0,6"/>
            <TextBlock Grid.Row="1" Text="Regel (leeg = geen)" Style="{StaticResource FieldLabel}"/>
            <TextBox x:Name="TxtDevRule" Grid.Row="1" Grid.Column="1" Style="{StaticResource FieldBox}" FontFamily="Consolas"
                     ToolTip="Bijv.: device.isCompliant -ne True -or device.trustType -ne &quot;ServerAD&quot;"/>
          </Grid></GroupBox>
          <GroupBox Header="Authenticatiestromen"><WrapPanel>
            <CheckBox x:Name="ChkFlowDc" Content="Device code flow" Margin="0,0,18,0"/>
            <CheckBox x:Name="ChkFlowAt" Content="Authenticatieoverdracht"/>
          </WrapPanel></GroupBox>
        </StackPanel></ScrollViewer>
      </TabItem>
      <TabItem Header="Toegang en sessie">
        <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel Margin="10">
          <GroupBox Header="Toekennen"><StackPanel>
            <StackPanel Orientation="Horizontal">
              <RadioButton x:Name="RbBlock" Content="Toegang blokkeren" Margin="0,0,18,6"/>
              <RadioButton x:Name="RbGrant" Content="Toegang verlenen met" Margin="0,0,18,6"/>
              <RadioButton x:Name="RbNoGrant" Content="Geen toekenning (alleen sessie)" Margin="0,0,18,6"/>
            </StackPanel>
            <StackPanel x:Name="PnlGrant" Margin="18,0,0,0">
              <WrapPanel x:Name="PnlGrantCtl"/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,6">
                <TextBlock Text="Authenticatiesterkte:" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <ComboBox x:Name="CmbStrength" Width="300" Height="26" DisplayMemberPath="Label" SelectedValuePath="Id"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal">
                <RadioButton x:Name="RbOpAnd" Content="Alle geselecteerde vereisen (EN)" Margin="0,0,18,0"/>
                <RadioButton x:Name="RbOpOr" Content="Een van de geselecteerde (OF)"/>
              </StackPanel>
            </StackPanel>
          </StackPanel></GroupBox>
          <GroupBox Header="Sessie"><StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
              <CheckBox x:Name="ChkSif" Content="Aanmeldfrequentie" VerticalAlignment="Center" Width="150"/>
              <RadioButton x:Name="RbSifTime" Content="Elke" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <TextBox x:Name="TxtSifVal" Width="50" Height="24" VerticalContentAlignment="Center" Margin="0,0,6,0"/>
              <ComboBox x:Name="CmbSifType" Width="90" Height="24" Margin="0,0,18,0"/>
              <RadioButton x:Name="RbSifEvery" Content="Elke keer" VerticalAlignment="Center"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
              <CheckBox x:Name="ChkPb" Content="Permanente browsersessie" VerticalAlignment="Center" Width="150"/>
              <ComboBox x:Name="CmbPb" Width="160" Height="24"/>
            </StackPanel>
            <CheckBox x:Name="ChkAer" Content="Door app afgedwongen beperkingen (Exchange/SharePoint)" Margin="0,0,0,6"/>
            <CheckBox x:Name="ChkRes" Content="Standaard tolerantie (resilience defaults) uitschakelen" Margin="0,0,0,6"/>
            <TextBlock x:Name="LblSessOther" Style="{StaticResource Hint}" Visibility="Collapsed"/>
          </StackPanel></GroupBox>
        </StackPanel></ScrollViewer>
      </TabItem>
      <TabItem x:Name="TabJson" Header="JSON">
        <DockPanel Margin="6">
          <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}"
                     Text="De volledige policy. Wijzigingen hier worden overgenomen zodra je een ander tabblad kiest of opslaat."/>
          <TextBox x:Name="TxtJson" AcceptsReturn="True" AcceptsTab="True" FontFamily="Consolas" FontSize="12"
                   VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
        </DockPanel>
      </TabItem>
    </TabControl>
  </DockPanel>
</Window>
'@

function New-IdItem {
    param([string]$Id, [string]$Kind)
    $n = Get-Name $Id
    $label = switch ($Kind) { 'group' { "[groep] $n" } 'role' { "[rol] $n" } 'user' { if ($Id -eq 'GuestsOrExternalUsers') { $n } else { "[gebruiker] $n" } } default { $n } }
    return [pscustomobject]@{ id = $Id; kind = $Kind; name = $label }
}

function Add-IdItems {
    param($ListBox, $Items)
    $cur = @($ListBox.Items | ForEach-Object { $_.id })
    foreach ($i in @($Items)) {
        if ($cur -contains $i.id) { continue }
        [void]$ListBox.Items.Add((New-IdItem $i.id $i.kind))
        $cur += $i.id
    }
}

function Remove-SelectedIdItems {
    param($ListBox)
    foreach ($i in @($ListBox.SelectedItems)) { $ListBox.Items.Remove($i) }
}

function Get-IdItems {
    param($ListBox, [string]$Kind)
    return @($ListBox.Items | Where-Object { -not $Kind -or $_.kind -eq $Kind } | ForEach-Object { [string]$_.id })
}

#endregion


#region --- Policy-editor: logica ---

function Show-PolicyEditor {
    # Geeft de aangepaste policy (hashtable) terug, of $null bij annuleren.
    # Alle handlers vuren tijdens ShowDialog(), dus lokale variabelen werken.
    param($Policy, [string]$Title = 'Policy bewerken')
    $base = if ($Policy) { Copy-Deep $Policy } else { [ordered]@{} }
    $w = ConvertFrom-Xaml $script:EditorXaml.Replace('{{STYLES}}', $script:StylesXaml)
    $w.Title = $Title
    $script:EdBase = $base
    $w.Owner = $script:Window
    $c = @{}
    foreach ($n in 'TxtName', 'CmbState', 'BtnSave', 'Tabs', 'TabJson', 'TxtJson',
                   'RbUNone', 'RbUAll', 'RbUSel', 'PnlUSel', 'ChkUGuests', 'LstUInc', 'BtnUIncAdd', 'BtnUIncAdmins', 'BtnUIncDel', 'LblGuestObj',
                   'ChkXGuests', 'LstUExc', 'BtnUExcAdd', 'BtnUExcBg', 'BtnUExcMe', 'BtnUExcDel',
                   'RbAAll', 'RbASel', 'RbAAction', 'RbACtx', 'RbANone', 'GbASel', 'GbAExc', 'GbAAction', 'GbACtx', 'PnlCtx',
                   'LstAInc', 'BtnAIncAdd', 'BtnAIncDel', 'LstAExc', 'BtnAExcAdd', 'BtnAExcDel', 'RbActReg', 'RbActDev',
                   'PnlClientApps', 'ChkPlat', 'PnlPlat', 'PnlPlatInc', 'PnlPlatExc', 'ChkLoc', 'PnlLoc', 'RbLAll', 'RbLSel',
                   'LstLInc', 'BtnLIncAdd', 'BtnLIncDel', 'LstLExc', 'BtnLExcAdd', 'BtnLExcDel', 'PnlURisk', 'PnlSRisk',
                   'CmbDevMode', 'TxtDevRule', 'ChkFlowDc', 'ChkFlowAt',
                   'RbBlock', 'RbGrant', 'RbNoGrant', 'PnlGrant', 'PnlGrantCtl', 'CmbStrength', 'RbOpAnd', 'RbOpOr',
                   'ChkSif', 'RbSifTime', 'TxtSifVal', 'CmbSifType', 'RbSifEvery', 'ChkPb', 'CmbPb', 'ChkAer', 'ChkRes', 'LblSessOther') {
        $c[$n] = $w.FindName($n)
        if ($null -eq $c[$n]) { throw "Editor-element '$n' niet gevonden" }
    }

    # --- Vaste keuzelijsten en dynamische checkboxen ---
    $c.CmbState.ItemsSource = @(
        [pscustomobject]@{ Value = 'enabledForReportingButNotEnforced'; Label = 'Alleen rapport' }
        [pscustomobject]@{ Value = 'enabled'; Label = 'Aan' }
        [pscustomobject]@{ Value = 'disabled'; Label = 'Uit' }
    )
    $c.CmbDevMode.ItemsSource = @('include', 'exclude')
    $c.CmbSifType.ItemsSource = @('hours', 'days')
    $c.CmbPb.ItemsSource = @('always', 'never')
    $chk = @{}
    $mk = {
        param($Panel, [string]$Key, [string]$Label)
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $Label; $cb.Tag = $Key; $cb.Margin = '0,0,18,6'
        [void]$Panel.Children.Add($cb)
        return $cb
    }
    foreach ($k in $script:ClientAppTypes.Keys) { $chk["ca_$k"] = & $mk $c.PnlClientApps $k $script:ClientAppTypes[$k] }
    foreach ($k in $script:Platforms) { $chk["pi_$k"] = & $mk $c.PnlPlatInc $k $k; $chk["px_$k"] = & $mk $c.PnlPlatExc $k $k }
    foreach ($k in 'high', 'medium', 'low') { $chk["ur_$k"] = & $mk $c.PnlURisk $k $k; $chk["sr_$k"] = & $mk $c.PnlSRisk $k $k }
    foreach ($k in $script:GrantLabels.Keys) { $chk["g_$k"] = & $mk $c.PnlGrantCtl $k $script:GrantLabels[$k] }
    foreach ($ctx in @($script:Contexts | Sort-Object { $_.id })) {
        $chk["ctx_$($ctx.id)"] = & $mk $c.PnlCtx ([string]$ctx.id) "$($ctx.id) - $($ctx.displayName)$(if (-not $ctx.isAvailable) { ' (niet gepubliceerd)' })"
    }
    if (-not $script:Contexts.Count) { [void]$c.PnlCtx.Children.Add((New-Object System.Windows.Controls.TextBlock -Property @{ Text = 'Geen authenticatiecontexten in deze tenant (tab Auth.contexten).' })) }
    $strengths = @([pscustomobject]@{ Id = ''; Label = '(geen)' }) + @($script:Strengths | Sort-Object { $_.policyType -ne 'builtIn' }, { $_.displayName } |
        ForEach-Object { [pscustomobject]@{ Id = [string]$_.id; Label = "$($_.displayName)$(if ($_.policyType -eq 'builtIn') { ' (ingebouwd)' })" } })
    $c.CmbStrength.ItemsSource = $strengths

    # --- Policy -> formulier ---
    $load = {
        param($P)
        $c.TxtName.Text = [string]$P.displayName
        $c.CmbState.SelectedValue = $(if ($P.state) { [string]$P.state } else { 'enabledForReportingButNotEnforced' })
        # Gebruikers
        $iu = Get-Arr $P 'conditions.users.includeUsers'
        $c.LstUInc.Items.Clear(); $c.LstUExc.Items.Clear()
        if ($iu -contains 'All') { $c.RbUAll.IsChecked = $true }
        elseif (-not $iu.Count -and -not (Get-Arr $P 'conditions.users.includeGroups').Count -and -not (Get-Arr $P 'conditions.users.includeRoles').Count -and -not (Get-Val $P 'conditions.users.includeGuestsOrExternalUsers')) { $c.RbUNone.IsChecked = $true }
        elseif ($iu -contains 'None') { $c.RbUNone.IsChecked = $true }
        else { $c.RbUSel.IsChecked = $true }
        $c.ChkUGuests.IsChecked = ($iu -contains 'GuestsOrExternalUsers')
        Add-IdItems $c.LstUInc @($iu | Where-Object { 'All', 'None', 'GuestsOrExternalUsers' -notcontains $_ } | ForEach-Object { @{ id = $_; kind = 'user' } })
        Add-IdItems $c.LstUInc @(Get-Arr $P 'conditions.users.includeGroups' | ForEach-Object { @{ id = $_; kind = 'group' } })
        Add-IdItems $c.LstUInc @(Get-Arr $P 'conditions.users.includeRoles' | ForEach-Object { @{ id = $_; kind = 'role' } })
        $c.LblGuestObj.Visibility = $(if ((Get-Val $P 'conditions.users.includeGuestsOrExternalUsers') -or (Get-Val $P 'conditions.users.excludeGuestsOrExternalUsers')) { 'Visible' } else { 'Collapsed' })
        $xu = Get-Arr $P 'conditions.users.excludeUsers'
        $c.ChkXGuests.IsChecked = ($xu -contains 'GuestsOrExternalUsers')
        Add-IdItems $c.LstUExc @($xu | Where-Object { $_ -ne 'GuestsOrExternalUsers' } | ForEach-Object { @{ id = $_; kind = 'user' } })
        Add-IdItems $c.LstUExc @(Get-Arr $P 'conditions.users.excludeGroups' | ForEach-Object { @{ id = $_; kind = 'group' } })
        Add-IdItems $c.LstUExc @(Get-Arr $P 'conditions.users.excludeRoles' | ForEach-Object { @{ id = $_; kind = 'role' } })
        # Doelresources
        $ia = Get-Arr $P 'conditions.applications.includeApplications'
        $ua = Get-Arr $P 'conditions.applications.includeUserActions'
        $ac = Get-Arr $P 'conditions.applications.includeAuthenticationContextClassReferences'
        $c.LstAInc.Items.Clear(); $c.LstAExc.Items.Clear()
        if ($ua.Count) { $c.RbAAction.IsChecked = $true }
        elseif ($ac.Count) { $c.RbACtx.IsChecked = $true }
        elseif ($ia -contains 'All') { $c.RbAAll.IsChecked = $true }
        elseif ($ia.Count -and $ia -notcontains 'None') { $c.RbASel.IsChecked = $true }
        else { $c.RbANone.IsChecked = $true }
        Add-IdItems $c.LstAInc @($ia | Where-Object { 'All', 'None' -notcontains $_ } | ForEach-Object { @{ id = $_; kind = 'app' } })
        Add-IdItems $c.LstAExc @(Get-Arr $P 'conditions.applications.excludeApplications' | ForEach-Object { @{ id = $_; kind = 'app' } })
        $c.RbActDev.IsChecked = ($ua -contains 'urn:user:registerdevice' -or $ua -contains 'registerOrJoinDevices')
        $c.RbActReg.IsChecked = -not $c.RbActDev.IsChecked
        foreach ($k in @($chk.Keys | Where-Object { $_ -like 'ctx_*' })) { $chk[$k].IsChecked = ($ac -contains $chk[$k].Tag) }
        # Voorwaarden
        $cat = Get-Arr $P 'conditions.clientAppTypes'
        foreach ($k in $script:ClientAppTypes.Keys) { $chk["ca_$k"].IsChecked = ($cat -contains $k) }
        $pl = Get-Val $P 'conditions.platforms'
        $c.ChkPlat.IsChecked = [bool]$pl
        foreach ($k in $script:Platforms) {
            $chk["pi_$k"].IsChecked = ((Get-Arr $pl 'includePlatforms') -contains $k)
            $chk["px_$k"].IsChecked = ((Get-Arr $pl 'excludePlatforms') -contains $k)
        }
        $lo = Get-Val $P 'conditions.locations'
        $c.ChkLoc.IsChecked = [bool]$lo
        $c.LstLInc.Items.Clear(); $c.LstLExc.Items.Clear()
        $li = Get-Arr $lo 'includeLocations'
        if (-not $li.Count -or $li -contains 'All') { $c.RbLAll.IsChecked = $true } else { $c.RbLSel.IsChecked = $true }
        Add-IdItems $c.LstLInc @($li | Where-Object { $_ -ne 'All' } | ForEach-Object { @{ id = $_; kind = 'location' } })
        Add-IdItems $c.LstLExc @(Get-Arr $lo 'excludeLocations' | ForEach-Object { @{ id = $_; kind = 'location' } })
        foreach ($k in 'high', 'medium', 'low') {
            $chk["ur_$k"].IsChecked = ((Get-Arr $P 'conditions.userRiskLevels') -contains $k)
            $chk["sr_$k"].IsChecked = ((Get-Arr $P 'conditions.signInRiskLevels') -contains $k)
        }
        $df = Get-Val $P 'conditions.devices.deviceFilter'
        $c.CmbDevMode.SelectedItem = $(if ($df -and $df.mode) { [string]$df.mode } else { 'exclude' })
        $c.TxtDevRule.Text = $(if ($df) { [string]$df.rule } else { '' })
        $tm = [string](Get-Val $P 'conditions.authenticationFlows.transferMethods')
        $c.ChkFlowDc.IsChecked = ($tm -match 'deviceCodeFlow')
        $c.ChkFlowAt.IsChecked = ($tm -match 'authenticationTransfer')
        # Toekennen
        $gc = Get-Val $P 'grantControls'
        $bc = Get-Arr $gc 'builtInControls'
        $asId = [string](Get-Val $gc 'authenticationStrength.id')
        if ($bc -contains 'block') { $c.RbBlock.IsChecked = $true }
        elseif ($bc.Count -or $asId -or (Get-Arr $gc 'termsOfUse').Count) { $c.RbGrant.IsChecked = $true }
        elseif ($P.Contains('displayName') -or $P.Contains('id')) { $c.RbNoGrant.IsChecked = $true }
        else { $c.RbGrant.IsChecked = $true }
        foreach ($k in $script:GrantLabels.Keys) { $chk["g_$k"].IsChecked = ($bc -contains $k) }
        $c.CmbStrength.SelectedValue = $asId
        if ($null -eq $c.CmbStrength.SelectedItem) { $c.CmbStrength.SelectedIndex = 0 }
        if ($gc -and $gc.operator -eq 'AND') { $c.RbOpAnd.IsChecked = $true } else { $c.RbOpOr.IsChecked = $true }
        # Sessie
        $sc = Get-Val $P 'sessionControls'
        $sif = Get-Val $sc 'signInFrequency'
        $c.ChkSif.IsChecked = [bool](Get-Val $sif 'isEnabled')
        $c.RbSifEvery.IsChecked = ($sif -and $sif.frequencyInterval -eq 'everyTime')
        $c.RbSifTime.IsChecked = -not $c.RbSifEvery.IsChecked
        $c.TxtSifVal.Text = $(if ($sif -and $sif.value) { [string]$sif.value } else { '12' })
        $c.CmbSifType.SelectedItem = $(if ($sif -and $sif.type) { [string]$sif.type } else { 'hours' })
        $pb = Get-Val $sc 'persistentBrowser'
        $c.ChkPb.IsChecked = [bool](Get-Val $pb 'isEnabled')
        $c.CmbPb.SelectedItem = $(if ($pb -and $pb.mode) { [string]$pb.mode } else { 'never' })
        $c.ChkAer.IsChecked = [bool](Get-Val $sc 'applicationEnforcedRestrictions.isEnabled')
        $c.ChkRes.IsChecked = [bool](Get-Val $sc 'disableResilienceDefaults')
        $other = @()
        if (Get-Val $sc 'cloudAppSecurity.isEnabled') { $other += "Defender for Cloud Apps ($($sc.cloudAppSecurity.cloudAppSecurityType))" }
        if ((Get-Arr $gc 'termsOfUse').Count) { $other += 'gebruiksvoorwaarden' }
        $c.LblSessOther.Text = "Behouden (alleen via JSON aan te passen): $($other -join ', ')"
        $c.LblSessOther.Visibility = $(if ($other.Count) { 'Visible' } else { 'Collapsed' })
        & $refresh
    }

    # --- Zichtbaarheid ---
    $refresh = {
        $c.PnlUSel.IsEnabled = [bool]$c.RbUSel.IsChecked
        $c.GbASel.Visibility = $(if ($c.RbASel.IsChecked) { 'Visible' } else { 'Collapsed' })
        $c.GbAExc.Visibility = $(if ($c.RbAAll.IsChecked -or $c.RbASel.IsChecked) { 'Visible' } else { 'Collapsed' })
        $c.GbAAction.Visibility = $(if ($c.RbAAction.IsChecked) { 'Visible' } else { 'Collapsed' })
        $c.GbACtx.Visibility = $(if ($c.RbACtx.IsChecked) { 'Visible' } else { 'Collapsed' })
        $c.PnlPlat.IsEnabled = [bool]$c.ChkPlat.IsChecked
        $c.PnlLoc.IsEnabled = [bool]$c.ChkLoc.IsChecked
        $c.LstLInc.IsEnabled = [bool]$c.RbLSel.IsChecked
        $c.PnlGrant.IsEnabled = [bool]$c.RbGrant.IsChecked
    }
    foreach ($n in 'RbUNone', 'RbUAll', 'RbUSel', 'RbAAll', 'RbASel', 'RbAAction', 'RbACtx', 'RbANone', 'RbLAll', 'RbLSel', 'RbBlock', 'RbGrant', 'RbNoGrant') {
        $c[$n].Add_Checked({ & $refresh })
    }
    foreach ($n in 'ChkPlat', 'ChkLoc') { $c[$n].Add_Click({ & $refresh }) }

    # --- Formulier -> policy ---
    $build = {
        $p = Copy-Deep $script:EdBase
        if (-not $p) { $p = [ordered]@{} }
        $p.displayName = $c.TxtName.Text.Trim()
        $p.state = [string]$c.CmbState.SelectedValue
        if (-not $p.Contains('conditions') -or $null -eq $p.conditions) { $p.conditions = [ordered]@{} }
        $cond = $p.conditions
        if (-not $cond.Contains('users') -or $null -eq $cond.users) { $cond.users = [ordered]@{} }
        $u = $cond.users
        $incU = @(); $incG = @(); $incR = @()
        if ($c.RbUAll.IsChecked) { $incU = @('All') }
        elseif ($c.RbUNone.IsChecked) { $incU = @('None') }
        else {
            $incU = @(Get-IdItems $c.LstUInc 'user')
            if ($c.ChkUGuests.IsChecked) { $incU += 'GuestsOrExternalUsers' }
            $incG = @(Get-IdItems $c.LstUInc 'group')
            $incR = @(Get-IdItems $c.LstUInc 'role')
        }
        $u.includeUsers = @($incU); $u.includeGroups = @($incG); $u.includeRoles = @($incR)
        $exU = @(Get-IdItems $c.LstUExc 'user')
        if ($c.ChkXGuests.IsChecked) { $exU += 'GuestsOrExternalUsers' }
        $u.excludeUsers = @($exU)
        $u.excludeGroups = @(Get-IdItems $c.LstUExc 'group')
        $u.excludeRoles = @(Get-IdItems $c.LstUExc 'role')
        if (-not $c.RbUSel.IsChecked) { $u.includeGuestsOrExternalUsers = $null }

        if (-not $cond.Contains('applications') -or $null -eq $cond.applications) { $cond.applications = [ordered]@{} }
        $a = $cond.applications
        $a.includeUserActions = @(); $a.includeAuthenticationContextClassReferences = @()
        if ($c.RbAAll.IsChecked) { $a.includeApplications = @('All') }
        elseif ($c.RbASel.IsChecked) { $a.includeApplications = @(Get-IdItems $c.LstAInc) }
        elseif ($c.RbAAction.IsChecked) {
            $a.includeApplications = @()
            $a.includeUserActions = @($(if ($c.RbActDev.IsChecked) { 'urn:user:registerdevice' } else { 'urn:user:registersecurityinfo' }))
        }
        elseif ($c.RbACtx.IsChecked) {
            $a.includeApplications = @()
            $a.includeAuthenticationContextClassReferences = @($chk.Keys | Where-Object { $_ -like 'ctx_*' -and $chk[$_].IsChecked } | ForEach-Object { $chk[$_].Tag } | Sort-Object)
        }
        else { $a.includeApplications = @('None') }
        $a.excludeApplications = $(if ($c.RbAAll.IsChecked -or $c.RbASel.IsChecked) { @(Get-IdItems $c.LstAExc) } else { @() })

        $cat = @($script:ClientAppTypes.Keys | Where-Object { $chk["ca_$_"].IsChecked })
        $cond.clientAppTypes = $(if ($cat.Count) { @($cat) } else { @('all') })
        if ($c.ChkPlat.IsChecked) {
            $pi = @($script:Platforms | Where-Object { $chk["pi_$_"].IsChecked })
            $cond.platforms = [ordered]@{
                includePlatforms = $(if ($pi.Count) { @($pi) } else { @('all') })
                excludePlatforms = @($script:Platforms | Where-Object { $chk["px_$_"].IsChecked })
            }
        }
        else { $cond.platforms = $null }
        if ($c.ChkLoc.IsChecked) {
            $cond.locations = [ordered]@{
                includeLocations = $(if ($c.RbLAll.IsChecked) { @('All') } else { @(Get-IdItems $c.LstLInc) })
                excludeLocations = @(Get-IdItems $c.LstLExc)
            }
        }
        else { $cond.locations = $null }
        $cond.userRiskLevels = @('high', 'medium', 'low' | Where-Object { $chk["ur_$_"].IsChecked })
        $cond.signInRiskLevels = @('high', 'medium', 'low' | Where-Object { $chk["sr_$_"].IsChecked })
        $rule = $c.TxtDevRule.Text.Trim()
        if ($rule) {
            if (-not $cond.Contains('devices') -or $null -eq $cond.devices) { $cond.devices = [ordered]@{} }
            $cond.devices.deviceFilter = [ordered]@{ mode = [string]$c.CmbDevMode.SelectedItem; rule = $rule }
        }
        elseif ($cond.Contains('devices') -and $cond.devices) { $cond.devices = $null }
        $tm = @(); if ($c.ChkFlowDc.IsChecked) { $tm += 'deviceCodeFlow' }; if ($c.ChkFlowAt.IsChecked) { $tm += 'authenticationTransfer' }
        $cond.authenticationFlows = $(if ($tm.Count) { [ordered]@{ transferMethods = ($tm -join ',') } } else { $null })

        $oldGc = $p['grantControls']
        if ($c.RbBlock.IsChecked) {
            $p.grantControls = [ordered]@{ operator = 'OR'; builtInControls = @('block') }
        }
        elseif ($c.RbGrant.IsChecked) {
            $g = [ordered]@{ operator = $(if ($c.RbOpAnd.IsChecked) { 'AND' } else { 'OR' }) }
            $g.builtInControls = @($script:GrantLabels.Keys | Where-Object { $chk["g_$_"].IsChecked })
            if ($oldGc -and (Get-Arr $oldGc 'termsOfUse').Count) { $g.termsOfUse = @(Get-Arr $oldGc 'termsOfUse') }
            if ($oldGc -and (Get-Arr $oldGc 'customAuthenticationFactors').Count) { $g.customAuthenticationFactors = @(Get-Arr $oldGc 'customAuthenticationFactors') }
            $sid = [string]$c.CmbStrength.SelectedValue
            $g.authenticationStrength = $(if ($sid) { [ordered]@{ id = $sid } } else { $null })
            $p.grantControls = $g
        }
        else { $p.grantControls = $null }

        $sc = if ($p['sessionControls']) { $p.sessionControls } else { [ordered]@{} }
        if ($c.ChkSif.IsChecked) {
            if ($c.RbSifEvery.IsChecked) {
                $sc.signInFrequency = [ordered]@{ isEnabled = $true; frequencyInterval = 'everyTime'; authenticationType = 'primaryAndSecondaryAuthentication'; value = $null; type = $null }
            }
            else {
                $v = 0; [void][int]::TryParse($c.TxtSifVal.Text.Trim(), [ref]$v)
                $sc.signInFrequency = [ordered]@{ isEnabled = $true; frequencyInterval = 'timeBased'; authenticationType = 'primaryAndSecondaryAuthentication'; value = $v; type = [string]$c.CmbSifType.SelectedItem }
            }
        }
        else { $sc.signInFrequency = $null }
        $sc.persistentBrowser = $(if ($c.ChkPb.IsChecked) { [ordered]@{ isEnabled = $true; mode = [string]$c.CmbPb.SelectedItem } } else { $null })
        $sc.applicationEnforcedRestrictions = $(if ($c.ChkAer.IsChecked) { [ordered]@{ isEnabled = $true } } else { $null })
        $sc.disableResilienceDefaults = $(if ($c.ChkRes.IsChecked) { $true } else { $null })
        $used = @($sc.Keys | Where-Object { $null -ne $sc[$_] -and $_ -notlike '*@odata*' })
        $p.sessionControls = $(if ($used.Count) { $sc } else { $null })
        return $p
    }

    $validate = {
        param($P)
        $err = @()
        if (-not $P.displayName) { $err += 'Naam is verplicht.' }
        $iu = Get-Arr $P 'conditions.users.includeUsers'
        if (-not $iu.Count -and -not (Get-Arr $P 'conditions.users.includeGroups').Count -and -not (Get-Arr $P 'conditions.users.includeRoles').Count -and -not (Get-Val $P 'conditions.users.includeGuestsOrExternalUsers') -and -not (Get-Val $P 'conditions.clientApplications')) {
            $err += 'Kies bij Gebruikers minstens 1 gebruiker, groep, rol of gasten (of Geen).'
        }
        $a = Get-Val $P 'conditions.applications'
        if (-not (Get-Arr $a 'includeApplications').Count -and -not (Get-Arr $a 'includeUserActions').Count -and -not (Get-Arr $a 'includeAuthenticationContextClassReferences').Count) {
            $err += 'Kies bij Doelresources minstens 1 app, een gebruikersactie of een authenticatiecontext.'
        }
        $gc = Get-Val $P 'grantControls'
        if ($gc) {
            $bc = Get-Arr $gc 'builtInControls'
            if (-not $bc.Count -and -not (Get-Val $gc 'authenticationStrength.id') -and -not (Get-Arr $gc 'termsOfUse').Count) { $err += "Kies bij 'Toegang verlenen met' minstens 1 vereiste." }
            if ($bc -contains 'mfa' -and (Get-Val $gc 'authenticationStrength.id')) { $err += "'MFA vereisen' en een authenticatiesterkte kunnen niet samen; kies een van beide." }
        }
        elseif (-not (Get-Val $P 'sessionControls')) { $err += 'Kies een toekenning of minstens 1 sessie-instelling.' }
        $sif = Get-Val $P 'sessionControls.signInFrequency'
        if ($sif -and $sif.frequencyInterval -eq 'timeBased' -and [int]$sif.value -le 0) { $err += 'Vul een geldige aanmeldfrequentie in.' }
        if ((Get-Val $P 'sessionControls.persistentBrowser') -and (Get-Arr $P 'conditions.applications.includeApplications') -notcontains 'All') {
            $err += "'Permanente browsersessie' werkt alleen met 'Alle resources'."
        }
        return $err
    }

    # --- JSON-tab synchroniseren ---
    $script:EdJsonShown = $null
    $c.Tabs.Add_SelectionChanged({
        if ($_.OriginalSource -ne $c.Tabs) { return }
        $toJson = ($c.Tabs.SelectedItem -eq $c.TabJson)
        $fromJson = (@($_.RemovedItems) -contains $c.TabJson)
        if ($toJson) {
            $c.TxtJson.Text = ConvertTo-JsonText (& $build)
            $script:EdJsonShown = $c.TxtJson.Text
        }
        elseif ($fromJson -and $c.TxtJson.Text -ne $script:EdJsonShown) {
            try {
                $parsed = ConvertTo-Hashtable ($c.TxtJson.Text | ConvertFrom-Json)
                Resolve-PolicyNames @($parsed)
                $script:EdBase = $parsed
                & $load $parsed
            }
            catch {
                [System.Windows.MessageBox]::Show($w, "De JSON is ongeldig en is niet overgenomen:`n$_", 'JSON', 'OK', 'Warning') | Out-Null
            }
        }
    })

    # --- Knoppen ---
    $c.BtnUIncAdd.Add_Click({ Add-IdItems $c.LstUInc (Show-ObjectPicker -Kinds user, group, role -Title 'Opnemen' -Owner $w) })
    $c.BtnUIncAdmins.Add_Click({ Add-IdItems $c.LstUInc @($script:AdminRoles.Keys | ForEach-Object { $script:Names[$_] = $script:AdminRoles[$_]; @{ id = $_; kind = 'role' } }) })
    $c.BtnUIncDel.Add_Click({ Remove-SelectedIdItems $c.LstUInc })
    $c.BtnUExcAdd.Add_Click({ Add-IdItems $c.LstUExc (Show-ObjectPicker -Kinds user, group, role -Title 'Uitsluiten' -Owner $w) })
    $c.BtnUExcBg.Add_Click({
        $bg = @(Get-BreakGlass)
        if (-not $bg.Count) { [System.Windows.MessageBox]::Show($w, 'Er zijn voor deze tenant nog geen noodaccounts ingesteld (tab Back-up / import).', 'Noodaccounts', 'OK', 'Information') | Out-Null; return }
        foreach ($b in $bg) { if (-not $script:Names.ContainsKey($b.id)) { $script:Names[$b.id] = $b.upn } }
        Add-IdItems $c.LstUExc @($bg | ForEach-Object { @{ id = $_.id; kind = 'user' } })
    })
    $c.BtnUExcMe.Add_Click({
        if (-not $script:Me) { return }
        $script:Names[$script:Me.id] = "$($script:Me.displayName) ($($script:Me.upn))"
        Add-IdItems $c.LstUExc @(@{ id = $script:Me.id; kind = 'user' })
    })
    $c.BtnUExcDel.Add_Click({ Remove-SelectedIdItems $c.LstUExc })
    $c.BtnAIncAdd.Add_Click({ Add-IdItems $c.LstAInc (Show-ObjectPicker -Kinds app -Title 'Apps opnemen' -Owner $w) })
    $c.BtnAIncDel.Add_Click({ Remove-SelectedIdItems $c.LstAInc })
    $c.BtnAExcAdd.Add_Click({ Add-IdItems $c.LstAExc (Show-ObjectPicker -Kinds app -Title 'Apps uitsluiten' -Owner $w) })
    $c.BtnAExcDel.Add_Click({ Remove-SelectedIdItems $c.LstAExc })
    $c.BtnLIncAdd.Add_Click({ Add-IdItems $c.LstLInc (Show-ObjectPicker -Kinds location -Title 'Locaties opnemen' -Owner $w) })
    $c.BtnLIncDel.Add_Click({ Remove-SelectedIdItems $c.LstLInc })
    $c.BtnLExcAdd.Add_Click({ Add-IdItems $c.LstLExc (Show-ObjectPicker -Kinds location -Title 'Locaties uitsluiten' -Owner $w) })
    $c.BtnLExcDel.Add_Click({ Remove-SelectedIdItems $c.LstLExc })

    $script:EdResult = $null
    $c.BtnSave.Add_Click({
        try {
            if ($c.Tabs.SelectedItem -eq $c.TabJson -and $c.TxtJson.Text -ne $script:EdJsonShown) {
                $p = ConvertTo-Hashtable ($c.TxtJson.Text | ConvertFrom-Json)
            }
            else { $p = & $build }
        }
        catch { [System.Windows.MessageBox]::Show($w, "Ongeldige invoer:`n$_", 'Fout', 'OK', 'Error') | Out-Null; return }
        $err = @(& $validate $p)
        if ($err.Count) { [System.Windows.MessageBox]::Show($w, ($err -join "`n"), 'Controleer de policy', 'OK', 'Warning') | Out-Null; return }
        $script:EdResult = $p
        $w.DialogResult = $true
    })

    & $load $base
    if (-not $Policy) { $c.RbUSel.IsChecked = $true; $c.RbAAll.IsChecked = $true; $c.RbGrant.IsChecked = $true }
    if ($w.ShowDialog()) { return $script:EdResult }
    return $null
}

#endregion


#region --- Sjablonen ---

function New-PolicySkeleton {
    param([string]$Name)
    $p = [ordered]@{
        displayName = $Name
        state       = 'enabledForReportingButNotEnforced'
        conditions  = [ordered]@{
            clientAppTypes = @('all')
            users          = [ordered]@{ includeUsers = @('All'); excludeUsers = @(); includeGroups = @(); excludeGroups = @(); includeRoles = @(); excludeRoles = @() }
            applications   = [ordered]@{ includeApplications = @('All'); excludeApplications = @(); includeUserActions = @() }
        }
        grantControls   = [ordered]@{ operator = 'OR'; builtInControls = @('mfa') }
        sessionControls = $null
    }
    # Noodaccounts standaard uitsluiten
    $p.conditions.users.excludeUsers = @(Get-BreakGlass | ForEach-Object { $_.id })
    return $p
}

$script:Templates = @(
    [pscustomobject]@{ Name = 'CA001 - MFA voor alle gebruikers'; Info = 'Alle gebruikers, alle resources: MFA vereisen.'
        Make = { New-PolicySkeleton 'CA001 - MFA voor alle gebruikers' } }
    [pscustomobject]@{ Name = 'CA002 - Legacy authenticatie blokkeren'; Info = 'Blokkeert Exchange ActiveSync en andere legacy clients.'
        Make = { $p = New-PolicySkeleton 'CA002 - Legacy authenticatie blokkeren'
                 $p.conditions.clientAppTypes = @('exchangeActiveSync', 'other')
                 $p.grantControls = [ordered]@{ operator = 'OR'; builtInControls = @('block') }; $p } }
    [pscustomobject]@{ Name = 'CA003 - Phishing-resistente MFA voor beheerders'; Info = 'Beheerrollen: authenticatiesterkte Phishing-resistente MFA.'
        Make = { $p = New-PolicySkeleton 'CA003 - Phishing-resistente MFA voor beheerders'
                 $p.conditions.users.includeUsers = @(); $p.conditions.users.includeRoles = @($script:AdminRoles.Keys)
                 $p.grantControls = [ordered]@{ operator = 'OR'; builtInControls = @(); authenticationStrength = [ordered]@{ id = '00000000-0000-0000-0000-000000000004' } }; $p } }
    [pscustomobject]@{ Name = 'CA004 - MFA voor beheerders'; Info = 'Beheerrollen: MFA vereisen (als phishing-resistent nog niet kan).'
        Make = { $p = New-PolicySkeleton 'CA004 - MFA voor beheerders'
                 $p.conditions.users.includeUsers = @(); $p.conditions.users.includeRoles = @($script:AdminRoles.Keys); $p } }
    [pscustomobject]@{ Name = 'CA005 - MFA voor Azure-beheer'; Info = 'Windows Azure Service Management API (portal, CLI, PowerShell): MFA.'
        Make = { $p = New-PolicySkeleton 'CA005 - MFA voor Azure-beheer'
                 $p.conditions.applications.includeApplications = @('797f4846-ba00-4fd7-ba43-dac1f8f63013')
                 $script:Names['797f4846-ba00-4fd7-ba43-dac1f8f63013'] = 'Windows Azure Service Management API'; $p } }
    [pscustomobject]@{ Name = 'CA006 - MFA voor beheerportals'; Info = 'Microsoft-beheerportals (Entra, Intune, M365 admin e.a.): MFA voor iedereen.'
        Make = { $p = New-PolicySkeleton 'CA006 - MFA voor beheerportals'
                 $p.conditions.applications.includeApplications = @('MicrosoftAdminPortals'); $p } }
    [pscustomobject]@{ Name = 'CA007 - MFA voor gasten'; Info = 'Gasten en externe gebruikers: MFA vereisen.'
        Make = { $p = New-PolicySkeleton 'CA007 - MFA voor gasten'
                 $p.conditions.users.includeUsers = @('GuestsOrExternalUsers'); $p } }
    [pscustomobject]@{ Name = 'CA008 - Beveiligingsgegevens registreren buiten vertrouwde locatie'; Info = 'Registratie van MFA/SSPR buiten vertrouwde locaties: MFA (nieuwe medewerkers: TAP).'
        Make = { $p = New-PolicySkeleton 'CA008 - Beveiligingsgegevens registreren buiten vertrouwde locatie'
                 $p.conditions.applications.includeApplications = @(); $p.conditions.applications.includeUserActions = @('urn:user:registersecurityinfo')
                 $p.conditions.locations = [ordered]@{ includeLocations = @('All'); excludeLocations = @('AllTrusted') }; $p } }
    [pscustomobject]@{ Name = 'CA009 - Toegang buiten toegestane landen blokkeren'; Info = 'Blokkeert alles behalve de gekozen landen-locatie (je kiest die zo).'; NeedsLocation = $true
        Make = { $p = New-PolicySkeleton 'CA009 - Toegang buiten toegestane landen blokkeren'
                 $p.conditions.locations = [ordered]@{ includeLocations = @('All'); excludeLocations = @() }
                 $p.grantControls = [ordered]@{ operator = 'OR'; builtInControls = @('block') }; $p } }
    [pscustomobject]@{ Name = 'CA010 - Compliant of hybrid joined apparaat (Windows/macOS)'; Info = 'Windows en macOS: apparaat compliant OF hybrid joined.'
        Make = { $p = New-PolicySkeleton 'CA010 - Compliant of hybrid joined apparaat (Windows/macOS)'
                 $p.conditions.platforms = [ordered]@{ includePlatforms = @('windows', 'macOS'); excludePlatforms = @() }
                 $p.grantControls = [ordered]@{ operator = 'OR'; builtInControls = @('compliantDevice', 'domainJoinedDevice') }; $p } }
    [pscustomobject]@{ Name = 'CA011 - App-beveiligingsbeleid op iOS/Android'; Info = 'Mobiel: goedgekeurde app OF app-beveiligingsbeleid (MAM).'
        Make = { $p = New-PolicySkeleton 'CA011 - App-beveiligingsbeleid op iOS/Android'
                 $p.conditions.applications.includeApplications = @('Office365')
                 $p.conditions.platforms = [ordered]@{ includePlatforms = @('android', 'iOS'); excludePlatforms = @() }
                 $p.grantControls = [ordered]@{ operator = 'OR'; builtInControls = @('compliantApplication') }; $p } }
    [pscustomobject]@{ Name = 'CA012 - Sessielimiet op onbeheerde apparaten'; Info = 'Niet-compliante apparaten: elke 12 uur aanmelden, geen permanente browser.'
        Make = { $p = New-PolicySkeleton 'CA012 - Sessielimiet op onbeheerde apparaten'
                 $p.conditions.devices = [ordered]@{ deviceFilter = [ordered]@{ mode = 'include'; rule = 'device.isCompliant -ne True -and device.trustType -ne "ServerAD"' } }
                 $p.grantControls = $null
                 $p.sessionControls = [ordered]@{
                     signInFrequency   = [ordered]@{ isEnabled = $true; value = 12; type = 'hours'; frequencyInterval = 'timeBased'; authenticationType = 'primaryAndSecondaryAuthentication' }
                     persistentBrowser = [ordered]@{ isEnabled = $true; mode = 'never' } }; $p } }
    [pscustomobject]@{ Name = 'CA013 - Device code flow blokkeren'; Info = 'Blokkeert aanmelden via device code (veelgebruikt bij phishing).'
        Make = { $p = New-PolicySkeleton 'CA013 - Device code flow blokkeren'
                 $p.conditions.authenticationFlows = [ordered]@{ transferMethods = 'deviceCodeFlow' }
                 $p.grantControls = [ordered]@{ operator = 'OR'; builtInControls = @('block') }; $p } }
    [pscustomobject]@{ Name = 'CA014 - Hoog aanmeldrisico: MFA (P2)'; Info = 'Aanmeldrisico hoog/gemiddeld: MFA en elke keer aanmelden.'
        Make = { $p = New-PolicySkeleton 'CA014 - Hoog aanmeldrisico: MFA (P2)'
                 $p.conditions.signInRiskLevels = @('high', 'medium')
                 $p.sessionControls = [ordered]@{ signInFrequency = [ordered]@{ isEnabled = $true; frequencyInterval = 'everyTime'; authenticationType = 'primaryAndSecondaryAuthentication'; value = $null; type = $null } }; $p } }
    [pscustomobject]@{ Name = 'CA015 - Hoog gebruikersrisico: wachtwoord wijzigen (P2)'; Info = 'Gebruikersrisico hoog: MFA EN wachtwoord wijzigen.'
        Make = { $p = New-PolicySkeleton 'CA015 - Hoog gebruikersrisico: wachtwoord wijzigen (P2)'
                 $p.conditions.userRiskLevels = @('high')
                 $p.grantControls = [ordered]@{ operator = 'AND'; builtInControls = @('mfa', 'passwordChange') }
                 $p.sessionControls = [ordered]@{ signInFrequency = [ordered]@{ isEnabled = $true; frequencyInterval = 'everyTime'; authenticationType = 'primaryAndSecondaryAuthentication'; value = $null; type = $null } }; $p } }
)

function Show-TemplatePicker {
    $w = New-Object System.Windows.Window
    $w.Title = 'Nieuwe policy uit sjabloon'; $w.Width = 720; $w.Height = 520; $w.WindowStartupLocation = 'CenterOwner'; $w.Owner = $script:Window
    $w.FontFamily = 'Segoe UI'; $w.FontSize = 12
    $dp = New-Object System.Windows.Controls.DockPanel
    $dp.Margin = '12'
    $hint = New-Object System.Windows.Controls.TextBlock
    $hint.Text = "Sjablonen worden in 'Alleen rapport' aangemaakt; noodaccounts worden automatisch uitgesloten. Je kunt alles nog aanpassen voor het opslaan."
    $hint.TextWrapping = 'Wrap'; $hint.Margin = '0,0,0,8'; $hint.Foreground = New-Brush '#555555'
    [System.Windows.Controls.DockPanel]::SetDock($hint, 'Top')
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = 'Openen in editor'; $ok.Padding = '14,6'; $ok.HorizontalAlignment = 'Right'; $ok.Margin = '0,8,0,0'
    [System.Windows.Controls.DockPanel]::SetDock($ok, 'Bottom')
    $lst = New-Object System.Windows.Controls.ListView
    $gv = New-Object System.Windows.Controls.GridView
    foreach ($col in @(@('Sjabloon', 'Name', 300), @('Omschrijving', 'Info', 370))) {
        $gc = New-Object System.Windows.Controls.GridViewColumn
        $gc.Header = $col[0]; $gc.DisplayMemberBinding = New-Object System.Windows.Data.Binding($col[1]); $gc.Width = $col[2]
        $gv.Columns.Add($gc)
    }
    $lst.View = $gv
    $lst.ItemsSource = $script:Templates
    [void]$dp.Children.Add($hint); [void]$dp.Children.Add($ok); [void]$dp.Children.Add($lst)
    $w.Content = $dp
    $ok.Add_Click({ if ($lst.SelectedItem) { $w.DialogResult = $true } })
    $lst.Add_MouseDoubleClick({ if ($lst.SelectedItem) { $w.DialogResult = $true } })
    if ($w.ShowDialog()) { return $lst.SelectedItem }
    return $null
}

#endregion

#region --- Policy-acties ---

function Save-Policy {
    # Nieuw (zonder id) of bijwerken. Geeft $true bij succes.
    param($Policy, [string]$Id)
    if ($Policy.state -eq 'enabled' -and -not (Test-SelfLockout $Policy)) { return $false }
    $body = Get-PolicyBody $Policy
    try {
        if ($Id) {
            $old = $script:Policies | Where-Object { $_.id -eq $Id } | Select-Object -First 1
            Save-Backup $old 'policy' 'voor-wijziging'
            [void](Invoke-Graph "$($script:CaRoot)/policies/$Id" -Method PATCH -Body $body)
            Write-Log "Policy bijgewerkt: $($Policy.displayName)" "success"
        }
        else {
            $r = Invoke-Graph "$($script:CaRoot)/policies" -Method POST -Body $body
            Write-Log "Policy aangemaakt: $($Policy.displayName) ($($r.id))" "success"
        }
        return $true
    }
    catch {
        Show-Msg "Opslaan mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
        return $false
    }
}

function Get-SelectedPolicies {
    $ids = @(Get-SelectedRows $script:PolGrid | ForEach-Object { $_['Id'] })
    return @($script:Policies | Where-Object { $ids -contains $_.id })
}

function Invoke-PolicyEdit {
    param($Policy, [switch]$Duplicate)
    if ($Duplicate) {
        $copy = Copy-Deep $Policy
        $copy.displayName = "$($Policy.displayName) (kopie)"
        $copy.state = 'disabled'
        foreach ($k in 'id', 'createdDateTime', 'modifiedDateTime', 'templateId') { if ($copy.Contains($k)) { $copy.Remove($k) } }
        $res = Show-PolicyEditor $copy 'Policy dupliceren'
        if ($res -and (Save-Policy $res)) { Update-AllData }
        return
    }
    $res = Show-PolicyEditor $Policy "Policy bewerken - $($Policy.displayName)"
    if ($res -and (Save-Policy $res $Policy.id)) { Update-AllData }
}

function Invoke-NewPolicy {
    param([switch]$FromTemplate)
    $p = $null
    if ($FromTemplate) {
        $t = Show-TemplatePicker
        if (-not $t) { return }
        $p = & $t.Make
        if ($t.NeedsLocation) {
            $countries = @($script:Locations | Where-Object { $_['@odata.type'] -like '*countryNamedLocation' })
            if (-not $countries.Count) {
                Show-Msg "Er is nog geen landen-locatie. Maak eerst op tab Locaties een locatie met de toegestane landen." "Sjabloon" | Out-Null
                return
            }
            $pick = @(Show-ObjectPicker -Kinds location -Title 'Kies de locatie met de TOEGESTANE landen' -Single)
            if (-not $pick.Count) { return }
            $p.conditions.locations.excludeLocations = @($pick[0].id)
        }
    }
    $res = Show-PolicyEditor $p 'Nieuwe policy'
    if ($res -and (Save-Policy $res)) { Update-AllData }
}

function Set-PolicyState {
    param([string]$State)
    $sel = @(Get-SelectedPolicies)
    if (-not $sel.Count) { return }
    $label = $script:StateLabels[$State]
    $r = Show-Msg "$($sel.Count) policy/policies op '$label' zetten?`n`n$(($sel | ForEach-Object { '- ' + $_.displayName }) -join "`n")" "Status wijzigen" "YesNo" "Question"
    if ($r -ne 'Yes') { return }
    Start-LongOp "Status wijzigen..."
    try {
        foreach ($p in $sel) {
            if ($script:StopRequested) { break }
            if ($p.state -eq $State) { continue }
            $test = Copy-Deep $p; $test.state = $State
            if ($State -eq 'enabled' -and -not (Test-SelfLockout $test)) { Write-Log "Overgeslagen: $($p.displayName)" "warning"; continue }
            Save-Backup $p 'policy' 'voor-status'
            try {
                [void](Invoke-Graph "$($script:CaRoot)/policies/$($p.id)" -Method PATCH -Body @{ state = $State })
                Write-Log "$($p.displayName) -> $label" "success"
            }
            catch { }
        }
    }
    finally { Stop-LongOp }
    Update-AllData
}

function Remove-Policies {
    $sel = @(Get-SelectedPolicies)
    if (-not $sel.Count) { return }
    $r = Show-Msg "$($sel.Count) policy/policies DEFINITIEF verwijderen? (er wordt eerst een JSON-back-up gemaakt)`n`n$(($sel | ForEach-Object { '- ' + $_.displayName }) -join "`n")" "Verwijderen" "YesNo" "Warning"
    if ($r -ne 'Yes') { return }
    Start-LongOp "Verwijderen..."
    try {
        foreach ($p in $sel) {
            Save-Backup $p 'policy' 'verwijderd'
            try { [void](Invoke-Graph "$($script:CaRoot)/policies/$($p.id)" -Method DELETE); Write-Log "Verwijderd: $($p.displayName)" "success" } catch { }
        }
    }
    finally { Stop-LongOp }
    Update-AllData
}

function Update-PolicyExclusion {
    # Bulk: objecten toevoegen aan of verwijderen uit de uitsluitingen
    param([ValidateSet('Add', 'Remove')][string]$Mode, $Objects)
    $sel = @(Get-SelectedPolicies)
    if (-not $sel.Count -or -not @($Objects).Count) { return }
    $names = ($Objects | ForEach-Object { $_.name }) -join ', '
    $verb = if ($Mode -eq 'Add') { 'uitsluiten in' } else { 'weghalen uit de uitsluitingen van' }
    $r = Show-Msg "$names $verb $($sel.Count) policy/policies?" "Uitsluitingen" "YesNo" "Question"
    if ($r -ne 'Yes') { return }
    $map = @{ user = 'excludeUsers'; group = 'excludeGroups'; role = 'excludeRoles' }
    Start-LongOp "Uitsluitingen bijwerken..."
    try {
        foreach ($p in $sel) {
            if ($script:StopRequested) { break }
            $users = [ordered]@{}
            $changed = $false
            foreach ($prop in $map.Values) { $users[$prop] = @(Get-Arr $p "conditions.users.$prop") }
            foreach ($o in $Objects) {
                $prop = $map[$o.kind]
                if (-not $prop) { continue }
                if ($Mode -eq 'Add' -and $users[$prop] -notcontains $o.id) { $users[$prop] = @($users[$prop]) + $o.id; $changed = $true }
                if ($Mode -eq 'Remove' -and $users[$prop] -contains $o.id) { $users[$prop] = @($users[$prop] | Where-Object { $_ -ne $o.id }); $changed = $true }
            }
            if (-not $changed) { continue }
            $test = Copy-Deep $p
            foreach ($k in $users.Keys) { $test.conditions.users[$k] = @($users[$k]) }
            if ($Mode -eq 'Remove' -and $p.state -eq 'enabled' -and -not (Test-SelfLockout $test)) { continue }
            Save-Backup $p 'policy' 'voor-uitsluiting'
            # Volledige users-conditie meesturen, anders gaan andere velden verloren
            $body = @{ conditions = @{ users = (Get-CleanObject $test.conditions.users -Drop @()) } }
            try {
                [void](Invoke-Graph "$($script:CaRoot)/policies/$($p.id)" -Method PATCH -Body $body)
                Write-Log "Uitsluitingen bijgewerkt: $($p.displayName)" "success"
            }
            catch { }
        }
    }
    finally { Stop-LongOp }
    Update-AllData
}

#endregion


#region --- Toewijzingen: overzicht en bulk ---

# Velden in een policy waarin gebruikers, groepen, rollen en apps staan
$script:AsgFields = @(
    @{ Section = 'users';        Prop = 'includeUsers';        Kind = 'user';  Dir = 'in' }
    @{ Section = 'users';        Prop = 'excludeUsers';        Kind = 'user';  Dir = 'ex' }
    @{ Section = 'users';        Prop = 'includeGroups';       Kind = 'group'; Dir = 'in' }
    @{ Section = 'users';        Prop = 'excludeGroups';       Kind = 'group'; Dir = 'ex' }
    @{ Section = 'users';        Prop = 'includeRoles';        Kind = 'role';  Dir = 'in' }
    @{ Section = 'users';        Prop = 'excludeRoles';        Kind = 'role';  Dir = 'ex' }
    @{ Section = 'applications'; Prop = 'includeApplications'; Kind = 'app';   Dir = 'in' }
    @{ Section = 'applications'; Prop = 'excludeApplications'; Kind = 'app';   Dir = 'ex' }
)
$script:AsgKindLabels = @{ user = 'Gebruiker'; group = 'Groep'; role = 'Rol'; app = 'App' }
$script:AsgKindOrder  = @{ user = 0; group = 1; role = 2; app = 3 }
$script:AsgDirLabels  = @{ in = 'Opgenomen'; ex = 'Uitgesloten' }
$script:AsgRows       = @()
$script:AsgUpdating   = $false

function Get-AsgName {
    param([string]$Id, [string]$Kind)
    if ($Id -eq 'All') { return $(if ($Kind -eq 'app') { 'Alle resources' } else { 'Alle gebruikers' }) }
    return (Get-Name $Id)
}

function New-AsgChange {
    param([string]$PolicyId, [string]$Kind, [string]$Dir, [string]$Id, [ValidateSet('Add', 'Remove')][string]$Op)
    $f = $script:AsgFields | Where-Object { $_.Kind -eq $Kind -and $_.Dir -eq $Dir } | Select-Object -First 1
    return @{ PolicyId = $PolicyId; Section = $f.Section; Prop = $f.Prop; Kind = $Kind; Dir = $Dir; Id = $Id; Op = $Op }
}

function Get-AssignmentRows {
    # Een regel per (policy, veld, object)
    $rows = New-Object System.Collections.ArrayList
    foreach ($p in $script:Policies) {
        foreach ($f in $script:AsgFields) {
            foreach ($id in (Get-Arr $p "conditions.$($f.Section).$($f.Prop)")) {
                [void]$rows.Add(@{
                    Key = "$($f.Kind)|$id"; Kind = $f.Kind; Id = [string]$id; Name = Get-AsgName $id $f.Kind; Dir = $f.Dir
                    PolicyId = [string]$p.id; PolicyName = [string]$p.displayName; State = [string]$p.state
                })
            }
        }
    }
    return $rows.ToArray()
}

function Get-AsgFiltered {
    $kind = [string]$script:CmbAsgKind.SelectedValue
    $dir  = [string]$script:CmbAsgDir.SelectedValue
    $f    = $script:TxtAsgFilter.Text.Trim()
    return @($script:AsgRows | Where-Object {
        (-not $kind -or $_.Kind -eq $kind) -and (-not $dir -or $_.Dir -eq $dir) -and
        (-not $f -or "$($_.Name) $($_.Id) $($_.PolicyName)" -like "*$f*")
    })
}

function Update-AssignmentView {
    # Linkerlijst: een regel per object met het aantal policies
    param([switch]$Reload)
    if ($Reload) { $script:AsgRows = @(Get-AssignmentRows) }
    $keep = @(Get-SelectedRows $script:AsgObjGrid | ForEach-Object { $_['Key'] })
    $list = @(Get-AsgFiltered)
    $rows = foreach ($g in @($list | Group-Object { $_.Key })) {
        $first = $g.Group[0]
        @{
            Key = $g.Name; Soort = $script:AsgKindLabels[$first.Kind]; Naam = $first.Name; Order = $script:AsgKindOrder[$first.Kind]
            Opgenomen   = @($g.Group | Where-Object { $_.Dir -eq 'in' }).Count
            Uitgesloten = @($g.Group | Where-Object { $_.Dir -eq 'ex' }).Count
        }
    }
    $rows = @($rows | Sort-Object { $_.Order }, { $_.Naam })
    $script:AsgUpdating = $true
    try {
        Set-GridRows $script:AsgObjGrid @('Soort', 'Naam', 'Opgenomen', 'Uitgesloten', 'Key') $rows
        foreach ($item in @($script:AsgObjGrid.Items)) { if ($keep -contains $item.Row['Key']) { [void]$script:AsgObjGrid.SelectedItems.Add($item) } }
    }
    finally { $script:AsgUpdating = $false }
    $script:LblAsgCount.Text = "$($rows.Count) objecten, $($list.Count) toewijzingen"
    Update-AssignmentDetail
}

function Update-AssignmentDetail {
    # Rechterlijst: de toewijzingen van de geselecteerde objecten (of alles)
    if ($script:AsgUpdating) { return }
    $keys = @(Get-SelectedRows $script:AsgObjGrid | ForEach-Object { $_['Key'] })
    $list = @(Get-AsgFiltered)
    if ($keys.Count) { $list = @($list | Where-Object { $keys -contains $_.Key }) }
    $rows = foreach ($r in @($list | Sort-Object { $script:AsgKindOrder[$_.Kind] }, { $_.Name }, { $_.Dir }, { $_.PolicyName })) {
        @{
            Object = $r.Name; Soort = $script:AsgKindLabels[$r.Kind]; Richting = $script:AsgDirLabels[$r.Dir]
            Policy = $r.PolicyName; Status = $script:StateLabels[$r.State]
            Kleur = $(if ($r.Dir -eq 'in') { '#107C10' } else { '#B26200' })
            PolicyId = $r.PolicyId; Kind = $r.Kind; Dir = $r.Dir; Id = $r.Id
        }
    }
    Set-GridRows $script:AsgGrid @('Object', 'Soort', 'Richting', 'Policy', 'Status', 'Kleur', 'PolicyId', 'Kind', 'Dir', 'Id') @($rows)
}

function Get-AsgSelection {
    return @(Get-SelectedRows $script:AsgGrid | ForEach-Object {
        @{ PolicyId = [string]$_['PolicyId']; Kind = [string]$_['Kind']; Dir = [string]$_['Dir']; Id = [string]$_['Id']; Name = [string]$_['Object'] }
    })
}

function Show-PolicyPicker {
    # Modale lijst met vinkjes. Geeft @{ Ids; Dir } terug of $null.
    param([string]$Title, [string]$Hint = '', [switch]$AskDirection)
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="660" Height="620" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <DockPanel Margin="12">
    <TextBlock x:Name="LblHint" DockPanel.Dock="Top" Style="{StaticResource Hint}" TextWrapping="Wrap" Margin="0,0,0,8"/>
    <StackPanel x:Name="PnlDir" DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Toevoegen als:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,10,0"/>
      <RadioButton x:Name="RbIn" Content="Opnemen" IsChecked="True" VerticalAlignment="Center" Margin="0,0,14,0"/>
      <RadioButton x:Name="RbEx" Content="Uitsluiten" VerticalAlignment="Center"/>
    </StackPanel>
    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
      <Button x:Name="BtnNone" DockPanel.Dock="Right" Content="Niets" Style="{StaticResource BtnGray}" Margin="6,0,0,0"/>
      <Button x:Name="BtnAll" DockPanel.Dock="Right" Content="Zichtbare aanvinken" Margin="6,0,0,0"/>
      <ComboBox x:Name="CmbState" DockPanel.Dock="Right" Width="130" Height="26" Margin="6,0,0,0"/>
      <TextBox x:Name="TxtFind" Height="26" VerticalContentAlignment="Center" ToolTip="Filter op naam"/>
    </DockPanel>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="BtnOk" Content="Doorgaan" Style="{StaticResource BtnGreen}"/>
        <Button x:Name="BtnCancel" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True"/>
      </StackPanel>
      <TextBlock x:Name="LblCount" VerticalAlignment="Center" Foreground="Gray"/>
    </DockPanel>
    <ListBox x:Name="LstPol"/>
  </DockPanel>
</Window>
'@
    $w = ConvertFrom-Xaml $xaml.Replace('{{STYLES}}', $script:StylesXaml)
    $w.Title = $Title
    if ($script:Window) { $w.Owner = $script:Window }
    $lst = $w.FindName('LstPol'); $txt = $w.FindName('TxtFind'); $cmb = $w.FindName('CmbState'); $lbl = $w.FindName('LblCount')
    $w.FindName('LblHint').Text = $Hint
    if (-not $Hint) { $w.FindName('LblHint').Visibility = 'Collapsed' }
    if (-not $AskDirection) { $w.FindName('PnlDir').Visibility = 'Collapsed' }
    $cmb.ItemsSource = @('Alle statussen', 'Aan', 'Alleen rapport', 'Uit')
    $cmb.SelectedIndex = 0
    $count = { $lbl.Text = "$(@($boxes | Where-Object { $_.IsChecked }).Count) policy/policies aangevinkt" }
    $boxes = @(foreach ($p in @($script:Policies | Sort-Object { $_.displayName })) {
        $state = $script:StateLabels[[string]$p.state]
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = "$($p.displayName)   [$state]"
        $cb.Tag = @{ Id = [string]$p.id; Name = [string]$p.displayName; State = $state }
        $cb.Margin = '2,3'
        $cb.Add_Click($count)
        $cb
    })
    foreach ($cb in $boxes) { [void]$lst.Items.Add($cb) }
    $apply = {
        $f = $txt.Text.Trim(); $s = [string]$cmb.SelectedItem
        foreach ($b in $boxes) {
            $ok = (-not $f -or $b.Tag.Name -like "*$f*") -and ($s -eq 'Alle statussen' -or $b.Tag.State -eq $s)
            $b.Visibility = $(if ($ok) { 'Visible' } else { 'Collapsed' })
        }
    }
    $txt.Add_TextChanged($apply)
    $cmb.Add_SelectionChanged($apply)
    $w.FindName('BtnAll').Add_Click({ foreach ($b in $boxes) { if ($b.Visibility -eq 'Visible') { $b.IsChecked = $true } }; & $count })
    $w.FindName('BtnNone').Add_Click({ foreach ($b in $boxes) { $b.IsChecked = $false }; & $count })
    $w.FindName('BtnOk').Add_Click({ if (@($boxes | Where-Object { $_.IsChecked }).Count) { $w.DialogResult = $true } })
    $w.Add_ContentRendered({ $txt.Focus(); & $count })
    if (-not $w.ShowDialog()) { return $null }
    return @{
        Ids = @($boxes | Where-Object { $_.IsChecked } | ForEach-Object { $_.Tag.Id })
        Dir = $(if ($w.FindName('RbEx').IsChecked) { 'ex' } else { 'in' })
    }
}

function Invoke-AssignmentChanges {
    # Voert wijzigingen uit (lijst van New-AsgChange). Per policy: toepassen op
    # een kopie, controleren, samenvatting + buitensluit-check, back-up, PATCH.
    param($Changes, [string]$Title = 'Toewijzingen bijwerken')
    $plans = New-Object System.Collections.ArrayList
    $skipped = New-Object System.Collections.ArrayList
    foreach ($g in @(@($Changes) | Group-Object { $_.PolicyId })) {
        $p = $script:Policies | Where-Object { $_.id -eq $g.Name } | Select-Object -First 1
        if (-not $p) { continue }
        $test = Copy-Deep $p
        $lines = New-Object System.Collections.ArrayList
        $why = New-Object System.Collections.ArrayList
        foreach ($c in $g.Group) {
            $sec = Get-Val $test "conditions.$($c.Section)"
            $lbl = "$($script:AsgKindLabels[$c.Kind]) $(Get-AsgName $c.Id $c.Kind)"
            if ($null -eq $sec) {
                [void]$why.Add("${lbl}: policy heeft geen $(if ($c.Section -eq 'users') { 'gebruikersvoorwaarde (workload identities)' } else { 'app-voorwaarde' })")
                continue
            }
            $cur = @(Get-Arr $sec $c.Prop)
            if ($c.Op -eq 'Remove') {
                if ($cur -notcontains $c.Id) { continue }
                $sec[$c.Prop] = @($cur | Where-Object { $_ -ne $c.Id })
                [void]$lines.Add("  - $lbl weghalen ($($script:AsgDirLabels[$c.Dir].ToLower()))")
                continue
            }
            if ($cur -contains $c.Id) { continue }
            if ($c.Dir -eq 'ex' -and ('All', 'None' -contains $c.Id)) { [void]$why.Add("${lbl}: kan niet worden uitgesloten"); continue }
            if ($c.Dir -eq 'in') {
                $mainProp = if ($c.Section -eq 'users') { 'includeUsers' } else { 'includeApplications' }
                $main = @(Get-Arr $sec $mainProp)
                if ($c.Section -eq 'applications' -and ((Get-Arr $sec 'includeUserActions').Count -or (Get-Arr $sec 'includeAuthenticationContextClassReferences').Count)) {
                    [void]$why.Add("${lbl}: policy is gericht op een gebruikersactie of authenticatiecontext"); continue
                }
                if ($c.Id -eq 'All') {
                    # 'Alle' vervangt de huidige opname
                    $sec[$mainProp] = @('All')
                    if ($c.Section -eq 'users') { $sec['includeGroups'] = @(); $sec['includeRoles'] = @() }
                    [void]$lines.Add("  + $lbl opnemen (vervangt de huidige opname)")
                    continue
                }
                if ($main -contains 'All') { [void]$why.Add("${lbl}: valt al onder $(Get-AsgName 'All' $c.Kind), wordt niet apart opgenomen"); continue }
                if ($main -contains 'None') { $sec[$mainProp] = @($main | Where-Object { $_ -ne 'None' }) }
                $cur = @(Get-Arr $sec $c.Prop)
            }
            $sec[$c.Prop] = @($cur) + $c.Id
            [void]$lines.Add("  + $lbl $(if ($c.Dir -eq 'in') { 'opnemen' } else { 'uitsluiten' })")
        }
        foreach ($m in $why) { [void]$skipped.Add("$($p.displayName) - $m") }
        if (-not $lines.Count) { continue }
        # Een policy moet altijd iets opnemen
        $u = Get-Val $test 'conditions.users'
        if ($u -and -not ((Get-Arr $u 'includeUsers').Count + (Get-Arr $u 'includeGroups').Count + (Get-Arr $u 'includeRoles').Count) -and -not (Get-Val $u 'includeGuestsOrExternalUsers')) {
            [void]$skipped.Add("$($p.displayName) - zou geen opgenomen gebruikers meer hebben (zet de policy liever uit)"); continue
        }
        $a = Get-Val $test 'conditions.applications'
        if ($a -and -not ((Get-Arr $a 'includeApplications').Count + (Get-Arr $a 'includeUserActions').Count + (Get-Arr $a 'includeAuthenticationContextClassReferences').Count)) {
            [void]$skipped.Add("$($p.displayName) - zou geen opgenomen resources meer hebben"); continue
        }
        [void]$plans.Add(@{ Policy = $p; Test = $test; Lines = $lines; Risk = $false })
    }

    $skipTxt = if ($skipped.Count) { "`n`nOvergeslagen:`n" + ((@($skipped) | Select-Object -First 15 | ForEach-Object { "- $_" }) -join "`n") + $(if ($skipped.Count -gt 15) { "`n... en nog $($skipped.Count - 15)" }) } else { '' }
    if (-not $plans.Count) { Show-Msg "Er hoeft niets gewijzigd te worden.$skipTxt" $Title | Out-Null; return }
    $total = 0; foreach ($pl in $plans) { $total += $pl.Lines.Count }
    $out = New-Object System.Collections.ArrayList
    foreach ($pl in $plans) {
        [void]$out.Add([string]$pl.Policy.displayName)
        foreach ($l in $pl.Lines) { [void]$out.Add($l) }
    }
    $shown = (@($out) | Select-Object -First 30) -join "`n"
    if ($out.Count -gt 30) { $shown += "`n... en nog $($out.Count - 30) regel(s)" }
    $r = Show-Msg "$total wijziging(en) in $($plans.Count) policy/policies:`n`n$shown$skipTxt`n`nVan elke policy wordt eerst een JSON-back-up gemaakt. Doorgaan?" $Title "YesNo" "Question"
    if ($r -ne 'Yes') { return }

    # Alleen waarschuwen voor NIEUWE buitensluit-risico's door deze wijziging
    $risky = @()
    foreach ($pl in $plans) {
        if ($pl.Policy.state -ne 'enabled') { continue }
        $old = @(Get-LockoutWarnings $pl.Policy)
        $added = @(Get-LockoutWarnings $pl.Test | Where-Object { $old -notcontains $_ })
        if ($added.Count) { $pl.Risk = $true; $risky += "$($pl.Policy.displayName):`n  $($added -join "`n  ")" }
    }
    if ($risky.Count) {
        $r = Show-Msg "Pas op - na deze wijziging:`n`n$($risky -join "`n`n")`n`nJa = toch alles doorvoeren`nNee = deze policies overslaan, de rest wel`nAnnuleren = niets doen" "Pas op: buitensluiten" "YesNoCancel" "Warning"
        if ($r -eq 'Cancel') { return }
        if ($r -eq 'No') { $plans = @($plans | Where-Object { -not $_.Risk }) }
    }

    Start-LongOp "$Title..."
    $ok = 0
    try {
        foreach ($pl in @($plans)) {
            if ($script:StopRequested) { Write-Log "Gestopt op verzoek" "warning"; break }
            Save-Backup $pl.Policy 'policy' 'voor-toewijzing'
            # Hele conditions-blok meesturen, zodat er niets verloren gaat
            $body = @{ conditions = (Get-CleanObject $pl.Test.conditions -Drop @()) }
            try {
                [void](Invoke-Graph "$($script:CaRoot)/policies/$($pl.Policy.id)" -Method PATCH -Body $body)
                Write-Log "Toewijzingen bijgewerkt: $($pl.Policy.displayName) ($($pl.Lines.Count) wijziging(en))" "success"
                $ok++
            }
            catch { }
        }
    }
    finally { Stop-LongOp }
    Write-Log "$Title klaar: $ok van $(@($plans).Count) policy/policies bijgewerkt" $(if ($ok -eq @($plans).Count) { 'success' } else { 'warning' })
    Update-AllData
}

function Invoke-AsgAdd {
    $objs = @()
    $selObj = @(Get-SelectedRows $script:AsgObjGrid)
    if ($selObj.Count) {
        $r = Show-Msg "De $($selObj.Count) links geselecteerde object(en) gebruiken?`n`nJa = de geselecteerde objecten`nNee = andere objecten zoeken" "Toevoegen aan policies" "YesNoCancel" "Question"
        if ($r -eq 'Cancel') { return }
        if ($r -eq 'Yes') {
            $objs = @($selObj | ForEach-Object { $k = ([string]$_['Key']).Split('|', 2); @{ kind = $k[0]; id = $k[1]; name = [string]$_['Naam'] } })
        }
    }
    if (-not $objs.Count) { $objs = @(Show-ObjectPicker -Kinds user, group, role, app -Title 'Objecten kiezen om aan policies toe te voegen') }
    if (-not $objs.Count) { return }
    $names = (@($objs | Select-Object -First 10 | ForEach-Object { "$($script:AsgKindLabels[$_.kind]) $($_.name)" }) -join "`n")
    if ($objs.Count -gt 10) { $names += "`n... en nog $($objs.Count - 10)" }
    $pick = Show-PolicyPicker -Title 'Aan welke policies toevoegen?' -Hint "Toevoegen:`n$names" -AskDirection
    if (-not $pick) { return }
    $changes = foreach ($polId in $pick.Ids) { foreach ($o in $objs) { New-AsgChange $polId $o.kind $pick.Dir $o.id 'Add' } }
    Invoke-AssignmentChanges @($changes) 'Toevoegen aan policies'
}

function Invoke-AsgCopy {
    $sel = @(Get-AsgSelection)
    if (-not $sel.Count) { Show-Msg "Selecteer rechts de toewijzingen die je wilt kopieren." "Kopieren" | Out-Null; return }
    $uniq = [ordered]@{}
    foreach ($s in $sel) { $uniq["$($s.Kind)|$($s.Dir)|$($s.Id)"] = $s }
    $names = (@($uniq.Values | Select-Object -First 10 | ForEach-Object { "$($script:AsgDirLabels[$_.Dir]): $($script:AsgKindLabels[$_.Kind]) $($_.Name)" }) -join "`n")
    if ($uniq.Count -gt 10) { $names += "`n... en nog $($uniq.Count - 10)" }
    $pick = Show-PolicyPicker -Title 'Naar welke policies kopieren?' -Hint "Wordt toegevoegd op dezelfde manier (opgenomen/uitgesloten):`n$names"
    if (-not $pick) { return }
    $changes = foreach ($polId in $pick.Ids) { foreach ($u in $uniq.Values) { New-AsgChange $polId $u.Kind $u.Dir $u.Id 'Add' } }
    Invoke-AssignmentChanges @($changes) 'Kopieren naar policies'
}

function Invoke-AsgFlip {
    $sel = @(Get-AsgSelection | Where-Object { 'All', 'None' -notcontains $_.Id })
    if (-not $sel.Count) { Show-Msg "Selecteer rechts de toewijzingen die je wilt omdraaien ('Alle' kan niet worden uitgesloten)." "Omdraaien" | Out-Null; return }
    $changes = foreach ($s in $sel) {
        $other = if ($s.Dir -eq 'in') { 'ex' } else { 'in' }
        New-AsgChange $s.PolicyId $s.Kind $s.Dir $s.Id 'Remove'
        New-AsgChange $s.PolicyId $s.Kind $other $s.Id 'Add'
    }
    Invoke-AssignmentChanges @($changes) 'Omdraaien (opnemen/uitsluiten)'
}

function Invoke-AsgReplace {
    $sel = @(Get-AsgSelection | Where-Object { 'All', 'None' -notcontains $_.Id })
    if (-not $sel.Count) { Show-Msg "Selecteer rechts de toewijzingen die je wilt vervangen." "Vervangen" | Out-Null; return }
    $kinds = @($sel | ForEach-Object { $_.Kind } | Select-Object -Unique)
    if ($kinds.Count -ne 1) { Show-Msg "Selecteer regels van een soort (bijv. alleen groepen)." "Vervangen" | Out-Null; return }
    $old = @($sel | ForEach-Object { $_.Name } | Select-Object -Unique)
    $new = @(Show-ObjectPicker -Kinds $kinds[0] -Title "Vervangen door ($($old.Count) object(en) worden vervangen)" -Single)
    if (-not $new.Count) { return }
    $changes = foreach ($s in $sel) {
        if ($s.Id -eq $new[0].id) { continue }
        New-AsgChange $s.PolicyId $s.Kind $s.Dir $s.Id 'Remove'
        New-AsgChange $s.PolicyId $s.Kind $s.Dir $new[0].id 'Add'
    }
    Invoke-AssignmentChanges @($changes) "Vervangen door $($new[0].name)"
}

function Invoke-AsgRemove {
    $sel = @(Get-AsgSelection)
    if (-not $sel.Count) { Show-Msg "Selecteer rechts de toewijzingen die je wilt weghalen." "Weghalen" | Out-Null; return }
    $changes = foreach ($s in $sel) { New-AsgChange $s.PolicyId $s.Kind $s.Dir $s.Id 'Remove' }
    Invoke-AssignmentChanges @($changes) 'Toewijzingen weghalen'
}

function Export-Assignments {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV-bestand (*.csv)|*.csv'
    $dlg.FileName = "CA-toewijzingen $(Get-Date -Format 'yyyy-MM-dd').csv"
    if ($dlg.ShowDialog() -ne 'OK') { return }
    @(Get-AsgFiltered | Sort-Object { $script:AsgKindOrder[$_.Kind] }, { $_.Name }, { $_.Dir }, { $_.PolicyName }) | ForEach-Object {
        [pscustomobject]@{
            Soort = $script:AsgKindLabels[$_.Kind]; Object = $_.Name; ObjectId = $_.Id; Richting = $script:AsgDirLabels[$_.Dir]
            Policy = $_.PolicyName; Status = $script:StateLabels[$_.State]; PolicyId = $_.PolicyId
        }
    } | Export-Csv -Path $dlg.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Log "Toewijzingen geexporteerd: $($dlg.FileName)" "success"
}

#endregion


#region --- Locaties, sterktes en contexten ---

function Get-UsedIn {
    # Namen van de policies die een ID gebruiken (locatie, sterkte, context)
    param([string]$Id, [ValidateSet('location', 'strength', 'context')][string]$Kind)
    $hits = foreach ($p in $script:Policies) {
        $used = switch ($Kind) {
            'location' { ((Get-Arr $p 'conditions.locations.includeLocations') + (Get-Arr $p 'conditions.locations.excludeLocations')) -contains $Id }
            'strength' { [string](Get-Val $p 'grantControls.authenticationStrength.id') -eq $Id }
            'context'  { (Get-Arr $p 'conditions.applications.includeAuthenticationContextClassReferences') -contains $Id }
        }
        if ($used) { [string]$p.displayName }
    }
    return @($hits)
}

function Edit-NamedLocation {
    param($Location, [ValidateSet('ip', 'country')][string]$Type)
    if ($Location) { $Type = if ($Location['@odata.type'] -like '*country*') { 'country' } else { 'ip' } }
    if ($Type -eq 'ip') {
        $ranges = @($Location.ipRanges | ForEach-Object { $_.cidrAddress }) -join "`r`n"
        $res = Show-InputDialog -Title $(if ($Location) { 'IP-locatie bewerken' } else { 'Nieuwe IP-locatie' }) -Fields @(
            @{ Key = 'name'; Label = 'Naam'; Value = $Location.displayName }
            @{ Key = 'ranges'; Label = 'IP-bereiken in CIDR, 1 per regel (bijv. 203.0.113.10/32 of 2001:db8::/48)'; Value = $ranges; Multi = $true }
            @{ Key = 'trusted'; Label = 'Vertrouwen'; Check = 'Markeren als vertrouwde locatie'; Value = [bool]$Location.isTrusted }
        )
        if (-not $res) { return }
        $list = @($res.ranges -split '[\r\n,; ]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $bad = @()
        $ipRanges = foreach ($r in $list) {
            if ($r -notmatch '/') { $r = if ($r -match ':') { "$r/128" } else { "$r/32" } }
            $ip = $null
            if (-not [System.Net.IPAddress]::TryParse(($r -split '/')[0], [ref]$ip)) { $bad += $r; continue }
            $t = if ($ip.AddressFamily -eq 'InterNetworkV6') { '#microsoft.graph.iPv6CidrRange' } else { '#microsoft.graph.iPv4CidrRange' }
            [ordered]@{ '@odata.type' = $t; cidrAddress = $r }
        }
        if ($bad.Count) { Show-Msg "Ongeldige adressen: $($bad -join ', ')" "Locatie" "OK" "Warning" | Out-Null; return }
        if (-not $res.name -or -not @($ipRanges).Count) { Show-Msg "Naam en minstens 1 IP-bereik zijn verplicht." "Locatie" "OK" "Warning" | Out-Null; return }
        $body = [ordered]@{ '@odata.type' = '#microsoft.graph.ipNamedLocation'; displayName = $res.name.Trim(); isTrusted = [bool]$res.trusted; ipRanges = @($ipRanges) }
    }
    else {
        $codes = @($Location.countriesAndRegions) -join ', '
        $res = Show-InputDialog -Title $(if ($Location) { 'Landen-locatie bewerken' } else { 'Nieuwe landen-locatie' }) -Fields @(
            @{ Key = 'name'; Label = 'Naam'; Value = $Location.displayName }
            @{ Key = 'codes'; Label = 'Landcodes (ISO 3166, 2 letters), gescheiden door komma of spatie - bijv. NL, BE, DE'; Value = $codes }
            @{ Key = 'lookup'; Label = 'Land bepalen op basis van'; Options = @('clientIpAddress', 'authenticatorAppGps'); Value = $(if ($Location.countryLookupMethod) { [string]$Location.countryLookupMethod } else { 'clientIpAddress' }) }
            @{ Key = 'unknown'; Label = 'Onbekend'; Check = 'Ook onbekende landen/regio''s opnemen'; Value = [bool]$Location.includeUnknownCountriesAndRegions }
        ) -Hint 'Tip: voor een "landen blokkeren"-policy maak je hier een locatie met de TOEGESTANE landen en sluit je die uit in de policy (sjabloon CA009).'
        if (-not $res) { return }
        $list = @($res.codes.ToUpper() -split '[\s,;]+' | Where-Object { $_ })
        $bad = @($list | Where-Object { $_ -notmatch '^[A-Z]{2}$' })
        if ($bad.Count -or -not $list.Count -or -not $res.name) { Show-Msg "Naam en geldige landcodes (2 letters) zijn verplicht. Ongeldig: $($bad -join ', ')" "Locatie" "OK" "Warning" | Out-Null; return }
        $body = [ordered]@{ '@odata.type' = '#microsoft.graph.countryNamedLocation'; displayName = $res.name.Trim()
            countriesAndRegions = @($list | Select-Object -Unique); includeUnknownCountriesAndRegions = [bool]$res.unknown; countryLookupMethod = $res.lookup }
    }
    try {
        if ($Location) {
            Save-Backup $Location 'locatie' 'voor-wijziging'
            [void](Invoke-Graph "$($script:CaRoot)/namedLocations/$($Location.id)" -Method PATCH -Body $body)
        }
        else { [void](Invoke-Graph "$($script:CaRoot)/namedLocations" -Method POST -Body $body) }
        Write-Log "Locatie opgeslagen: $($body.displayName)" "success"
        Update-AllData
    }
    catch { Show-Msg "Opslaan mislukt:`n$_" "Fout" "OK" "Error" | Out-Null }
}

function Remove-NamedLocations {
    $ids = @(Get-SelectedRows $script:LocGrid | ForEach-Object { $_['Id'] })
    $sel = @($script:Locations | Where-Object { $ids -contains $_.id })
    if (-not $sel.Count) { return }
    $inUse = @($sel | Where-Object { (Get-UsedIn $_.id 'location').Count })
    if ($inUse.Count) {
        Show-Msg "Nog in gebruik, eerst uit de policies halen:`n$(($inUse | ForEach-Object { "- $($_.displayName): $((Get-UsedIn $_.id 'location') -join ', ')" }) -join "`n")" "Locatie" "OK" "Warning" | Out-Null
        return
    }
    if ((Show-Msg "$($sel.Count) locatie(s) verwijderen?`n$(($sel | ForEach-Object { '- ' + $_.displayName }) -join "`n")" "Verwijderen" "YesNo" "Warning") -ne 'Yes') { return }
    foreach ($l in $sel) {
        Save-Backup $l 'locatie' 'verwijderd'
        # Een vertrouwde IP-locatie moet eerst op niet-vertrouwd voordat hij weg mag
        if ($l.isTrusted) { try { [void](Invoke-Graph "$($script:CaRoot)/namedLocations/$($l.id)" -Method PATCH -Body @{ '@odata.type' = $l['@odata.type']; isTrusted = $false }) } catch { } }
        try { [void](Invoke-Graph "$($script:CaRoot)/namedLocations/$($l.id)" -Method DELETE); Write-Log "Locatie verwijderd: $($l.displayName)" "success" } catch { }
    }
    Update-AllData
}

function Get-Combinations {
    if ($script:Combinations) { return $script:Combinations }
    try {
        $r = Invoke-Graph "$($script:CaRoot)/authenticationStrength/combinations" -Quiet
        $script:Combinations = @($r.value | ForEach-Object { [string]$_ } | Where-Object { $_ })
    }
    catch { }
    if (-not $script:Combinations) { $script:Combinations = $script:DefaultCombinations }
    return $script:Combinations
}

function Edit-AuthStrength {
    param($Strength)
    if ($Strength -and $Strength.policyType -eq 'builtIn') { Show-Msg "Ingebouwde sterktes kun je niet wijzigen. Gebruik 'Dupliceren'." "Sterkte" | Out-Null; return }
    $w = New-Object System.Windows.Window
    $w.Title = $(if ($Strength -and $Strength.id) { 'Authenticatiesterkte bewerken' } else { 'Nieuwe authenticatiesterkte' })
    $w.Width = 560; $w.Height = 640; $w.WindowStartupLocation = 'CenterOwner'; $w.Owner = $script:Window; $w.FontFamily = 'Segoe UI'; $w.FontSize = 12
    $dp = New-Object System.Windows.Controls.DockPanel
    $dp.Margin = '12'
    $top = New-Object System.Windows.Controls.StackPanel
    [System.Windows.Controls.DockPanel]::SetDock($top, 'Top')
    $l1 = New-Object System.Windows.Controls.TextBlock -Property @{ Text = 'Naam'; Margin = '0,0,0,3' }
    $tn = New-Object System.Windows.Controls.TextBox -Property @{ Text = [string]$Strength.displayName; Margin = '0,0,0,8'; Padding = '4,3' }
    $l2 = New-Object System.Windows.Controls.TextBlock -Property @{ Text = 'Omschrijving'; Margin = '0,0,0,3' }
    $td = New-Object System.Windows.Controls.TextBox -Property @{ Text = [string]$Strength.description; Margin = '0,0,0,8'; Padding = '4,3' }
    $l3 = New-Object System.Windows.Controls.TextBlock -Property @{ Text = 'Toegestane combinaties:'; Margin = '0,0,0,3'; FontWeight = 'SemiBold' }
    foreach ($x in $l1, $tn, $l2, $td, $l3) { [void]$top.Children.Add($x) }
    $ok = New-Object System.Windows.Controls.Button -Property @{ Content = 'Opslaan'; Padding = '14,6'; HorizontalAlignment = 'Right'; Margin = '0,8,0,0' }
    [System.Windows.Controls.DockPanel]::SetDock($ok, 'Bottom')
    $sv = New-Object System.Windows.Controls.ScrollViewer
    $sp = New-Object System.Windows.Controls.StackPanel
    $sv.Content = $sp
    $boxes = @()
    $cur = @($Strength.allowedCombinations)
    foreach ($comb in (Get-Combinations)) {
        $cb = New-Object System.Windows.Controls.CheckBox -Property @{ Content = $comb; Tag = $comb; IsChecked = ($cur -contains $comb); Margin = '0,0,0,4' }
        [void]$sp.Children.Add($cb); $boxes += $cb
    }
    foreach ($x in $top, $ok, $sv) { [void]$dp.Children.Add($x) }
    $w.Content = $dp
    $ok.Add_Click({ $w.DialogResult = $true })
    if (-not $w.ShowDialog()) { return }
    $sel = @($boxes | Where-Object { $_.IsChecked } | ForEach-Object { [string]$_.Tag })
    if (-not $tn.Text.Trim() -or -not $sel.Count) { Show-Msg "Naam en minstens 1 combinatie zijn verplicht." "Sterkte" "OK" "Warning" | Out-Null; return }
    try {
        if ($Strength -and $Strength.id) {
            Save-Backup $Strength 'sterkte' 'voor-wijziging'
            [void](Invoke-Graph "$($script:CaRoot)/authenticationStrength/policies/$($Strength.id)" -Method PATCH -Body @{ displayName = $tn.Text.Trim(); description = $td.Text })
            $oldC = (@($Strength.allowedCombinations) | Sort-Object) -join '|'
            $newC = (@($sel) | Sort-Object) -join '|'
            if ($oldC -ne $newC) {
                [void](Invoke-Graph "$($script:CaRoot)/authenticationStrength/policies/$($Strength.id)/updateAllowedCombinations" -Method POST -Body @{ allowedCombinations = $sel })
            }
        }
        else {
            [void](Invoke-Graph "$($script:CaRoot)/authenticationStrength/policies" -Method POST -Body @{ displayName = $tn.Text.Trim(); description = $td.Text; allowedCombinations = $sel })
        }
        Write-Log "Authenticatiesterkte opgeslagen: $($tn.Text.Trim())" "success"
        Update-AllData
    }
    catch { Show-Msg "Opslaan mislukt:`n$_" "Fout" "OK" "Error" | Out-Null }
}

function Remove-AuthStrengths {
    $ids = @(Get-SelectedRows $script:StrGrid | ForEach-Object { $_['Id'] })
    $sel = @($script:Strengths | Where-Object { $ids -contains $_.id -and $_.policyType -ne 'builtIn' })
    if (-not $sel.Count) { Show-Msg "Selecteer een of meer eigen (niet ingebouwde) sterktes." "Sterkte" | Out-Null; return }
    $inUse = @($sel | Where-Object { (Get-UsedIn $_.id 'strength').Count })
    if ($inUse.Count) { Show-Msg "Nog in gebruik:`n$(($inUse | ForEach-Object { "- $($_.displayName): $((Get-UsedIn $_.id 'strength') -join ', ')" }) -join "`n")" "Sterkte" "OK" "Warning" | Out-Null; return }
    if ((Show-Msg "$($sel.Count) sterkte(s) verwijderen?" "Verwijderen" "YesNo" "Warning") -ne 'Yes') { return }
    foreach ($s in $sel) {
        Save-Backup $s 'sterkte' 'verwijderd'
        try { [void](Invoke-Graph "$($script:CaRoot)/authenticationStrength/policies/$($s.id)" -Method DELETE); Write-Log "Sterkte verwijderd: $($s.displayName)" "success" } catch { }
    }
    Update-AllData
}

function Edit-AuthContext {
    param($Context)
    $free = @(1..99 | ForEach-Object { "c$_" } | Where-Object { @($script:Contexts | ForEach-Object { $_.id }) -notcontains $_ })
    $fields = @()
    if (-not $Context) { $fields += @{ Key = 'id'; Label = 'ID'; Options = $free } }
    $fields += @{ Key = 'name'; Label = 'Naam'; Value = $Context.displayName }
    $fields += @{ Key = 'desc'; Label = 'Omschrijving'; Value = $Context.description }
    $fields += @{ Key = 'avail'; Label = 'Publiceren'; Check = 'Beschikbaar maken voor apps (gepubliceerd)'; Value = $(if ($Context) { [bool]$Context.isAvailable } else { $true }) }
    $res = Show-InputDialog -Title $(if ($Context) { "Context $($Context.id) bewerken" } else { 'Nieuwe authenticatiecontext' }) -Fields $fields `
        -Hint 'Een authenticatiecontext koppel je in een policy (Doelresources > Authenticatiecontext) en gebruik je in bijv. gevoeligheidslabels, SharePoint-sites of PIM.'
    if (-not $res) { return }
    if (-not $res.name) { Show-Msg "Naam is verplicht." "Context" "OK" "Warning" | Out-Null; return }
    $id = if ($Context) { $Context.id } else { $res.id }
    if ($Context) { Save-Backup $Context 'context' 'voor-wijziging' }
    try {
        # PATCH maakt de context aan als hij nog niet bestaat
        [void](Invoke-Graph "$($script:CaRoot)/authenticationContextClassReferences/$id" -Method PATCH -Body @{ displayName = $res.name; description = $res.desc; isAvailable = [bool]$res.avail })
        Write-Log "Authenticatiecontext $id opgeslagen" "success"
        Update-AllData
    }
    catch { Show-Msg "Opslaan mislukt:`n$_" "Fout" "OK" "Error" | Out-Null }
}

function Remove-AuthContexts {
    $ids = @(Get-SelectedRows $script:CtxGrid | ForEach-Object { $_['Id'] })
    $sel = @($script:Contexts | Where-Object { $ids -contains $_.id })
    if (-not $sel.Count) { return }
    $inUse = @($sel | Where-Object { (Get-UsedIn $_.id 'context').Count })
    if ($inUse.Count) { Show-Msg "Nog in gebruik in policies:`n$(($inUse | ForEach-Object { "- $($_.id): $((Get-UsedIn $_.id 'context') -join ', ')" }) -join "`n")" "Context" "OK" "Warning" | Out-Null; return }
    if ((Show-Msg "$($sel.Count) context(en) verwijderen? Gepubliceerde contexten worden eerst op 'niet beschikbaar' gezet." "Verwijderen" "YesNo" "Warning") -ne 'Yes') { return }
    foreach ($x in $sel) {
        Save-Backup $x 'context' 'verwijderd'
        try {
            if ($x.isAvailable) { [void](Invoke-Graph "$($script:CaRoot)/authenticationContextClassReferences/$($x.id)" -Method PATCH -Body @{ isAvailable = $false }) }
            [void](Invoke-Graph "$($script:CaRoot)/authenticationContextClassReferences/$($x.id)" -Method DELETE)
            Write-Log "Context verwijderd: $($x.id)" "success"
        }
        catch { }
    }
    Update-AllData
}

#endregion


#region --- Gebruiker-impact en aanmeldingen ---

function Invoke-ImpactPick {
    $pick = @(Show-ObjectPicker -Kinds user -Title 'Gebruiker kiezen' -Single)
    if (-not $pick.Count) { return }
    Enter-Busy "Groepen en rollen van de gebruiker ophalen..."
    try {
        $script:ImpactUser = Get-UserContext $pick[0].id
        $u = $script:ImpactUser
        $script:LblImpactUser.Text = "$($u.displayName) ($($u.upn))$(if ($u.isGuest) { ' - gast' }) - $($u.groups.Count) groep(en), $($u.roles.Count) actieve rol(len)"
        Show-Impact
        $script:SignGrid.ItemsSource = $null
        $script:TxtSignDetail.Text = ''
        Set-Status "Klaar" "success"
    }
    catch { Set-Status "Ophalen mislukt" "error" }
    finally { Exit-Busy }
}

function Show-Impact {
    $u = $script:ImpactUser
    if (-not $u) { return }
    $rows = foreach ($p in $script:Policies) {
        $s = Test-InScope $p $u
        if ($script:ChkImpactAll.IsChecked -ne $true -and $s.Result -eq 'Niet van toepassing') { continue }
        @{
            Resultaat = $s.Result; Reden = $s.Reason; Naam = [string]$p.displayName; Status = $script:StateLabels[[string]$p.state]
            Apps = Get-AppsShort $p; Toegang = Get-GrantShort $p; Id = [string]$p.id
            Kleur = switch ($s.Result) { 'Van toepassing' { '#107C10' } 'Uitgesloten' { '#B26200' } default { '#8A8A8A' } }
        }
    }
    $rows = @($rows | Sort-Object { switch ($_.Resultaat) { 'Van toepassing' { 0 } 'Uitgesloten' { 1 } default { 2 } } }, { $_.Naam })
    Set-GridRows $script:ImpactGrid @('Resultaat', 'Reden', 'Naam', 'Status', 'Apps', 'Toegang', 'Id', 'Kleur') $rows
    $n = @($rows | Where-Object { $_.Resultaat -eq 'Van toepassing' }).Count
    $script:LblImpactInfo.Text = "$n policy/policies van toepassing op basis van gebruikers/groepen/rollen. Andere voorwaarden (apps, locatie, platform, risico) bepalen bij het aanmelden of ze echt gelden."
}

function Invoke-LoadSignIns {
    $u = $script:ImpactUser
    if (-not $u) { Show-Msg "Kies eerst een gebruiker." "Aanmeldingen" | Out-Null; return }
    Enter-Busy "Aanmeldingen ophalen (kan even duren)..."
    try {
        $since = (Get-Date).ToUniversalTime().AddDays(-[int]$script:CmbSignDays.SelectedItem).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $uri = "$($script:GraphRoot)/auditLogs/signIns?`$filter=userId eq '$($u.id)' and createdDateTime ge $since&`$top=100"
        $list = Get-GraphPaged $uri -Max 200
        $script:SignIns = @($list)
        $rows = foreach ($s in $list) {
            $applied = @($s.appliedConditionalAccessPolicies | Where-Object { $_.result -in 'success', 'failure', 'reportOnlyFailure', 'reportOnlySuccess', 'reportOnlyInterrupted' })
            @{
                Tijd = ([datetime]$s.createdDateTime).ToLocalTime().ToString('dd-MM HH:mm:ss'); App = [string]$s.appDisplayName
                IP = [string]$s.ipAddress; Locatie = "$($s.location.city) $($s.location.countryOrRegion)".Trim()
                Resultaat = $(if ([int]$s.status.errorCode -eq 0) { 'Gelukt' } else { "Fout $($s.status.errorCode)" })
                CA = [string]$s.conditionalAccessStatus; Policies = "$($applied.Count)"; Id = [string]$s.id
                Kleur = $(if ([int]$s.status.errorCode -eq 0) { '#107C10' } else { '#C42B1C' })
            }
        }
        Set-GridRows $script:SignGrid @('Tijd', 'App', 'IP', 'Locatie', 'Resultaat', 'CA', 'Policies', 'Id', 'Kleur') @($rows)
        Write-Log "$(@($list).Count) aanmelding(en) geladen voor $($u.upn)" "success"
        Set-Status "Klaar" "success"
    }
    catch { Set-Status "Aanmeldingen ophalen mislukt (rechten AuditLog.Read.All / licentie Entra ID P1 nodig)" "error" }
    finally { Exit-Busy }
}

function Show-SignInDetail {
    $row = $script:SignGrid.SelectedItem
    if (-not $row) { $script:TxtSignDetail.Text = ''; return }
    $s = $script:SignIns | Where-Object { $_.id -eq $row.Row['Id'] } | Select-Object -First 1
    if (-not $s) { return }
    $map = @{ success = 'Toegepast - voldaan'; failure = 'Toegepast - NIET voldaan'; notApplied = 'Niet van toepassing'; notEnabled = 'Uitgeschakeld'
              reportOnlySuccess = 'Rapport: zou voldaan zijn'; reportOnlyFailure = 'Rapport: zou NIET voldaan zijn'; reportOnlyNotApplied = 'Rapport: niet van toepassing'; reportOnlyInterrupted = 'Rapport: zou onderbreken' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$($s.appDisplayName) - $(([datetime]$s.createdDateTime).ToLocalTime().ToString('dd-MM-yyyy HH:mm:ss'))")
    [void]$sb.AppendLine("Client: $($s.clientAppUsed)   Apparaat: $($s.deviceDetail.displayName) $($s.deviceDetail.operatingSystem) (compliant: $($s.deviceDetail.isCompliant), beheerd: $($s.deviceDetail.isManaged))")
    [void]$sb.AppendLine("Status: $($s.status.errorCode) $($s.status.failureReason)")
    if ($s.status.additionalDetails) { [void]$sb.AppendLine("        $($s.status.additionalDetails)") }
    [void]$sb.AppendLine("")
    foreach ($g in @($s.appliedConditionalAccessPolicies | Sort-Object { if ($_.result -match 'notApplied|notEnabled') { 1 } else { 0 } }, { $_.displayName })) {
        $res = if ($map.ContainsKey([string]$g.result)) { $map[[string]$g.result] } else { [string]$g.result }
        $ctl = @($g.enforcedGrantControls) + @($g.enforcedSessionControls) | Where-Object { $_ }
        [void]$sb.AppendLine(("{0,-34} {1}{2}" -f $res, $g.displayName, $(if ($ctl) { "  [$($ctl -join ', ')]" })))
    }
    $script:TxtSignDetail.Text = $sb.ToString()
}

#endregion

#region --- Back-up, import en noodaccounts ---

function Export-CaConfig {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Map voor de export (er wordt een submap met datum aangemaakt)'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $org = try { (Invoke-Graph "$($script:GraphRoot)/organization?`$select=displayName" -Quiet).value[0].displayName } catch { $script:TenantId }
    $dir = Join-Path $dlg.SelectedPath ("CA-export {0} {1}" -f (($org -replace '[\\/:*?"<>|]', '_')), (Get-Date -Format 'yyyy-MM-dd HHmm'))
    [void](New-Item -ItemType Directory -Path $dir -Force)
    # Namen meesturen, zodat bij import groepen/locaties op naam gekoppeld kunnen worden
    $refs = @{}
    foreach ($p in $script:Policies) {
        foreach ($path in 'conditions.users.includeUsers', 'conditions.users.excludeUsers', 'conditions.users.includeGroups', 'conditions.users.excludeGroups') {
            foreach ($id in (Get-Arr $p $path)) { if (Test-Guid $id) { $refs[$id] = Get-Name $id } }
        }
    }
    $all = [ordered]@{
        exportedAt = (Get-Date).ToString('s'); tenantId = $script:TenantId; tenantName = $org; tool = 'CA Beheer Tool'
        policies = @($script:Policies); namedLocations = @($script:Locations)
        authenticationStrengths = @($script:Strengths | Where-Object { $_.policyType -ne 'builtIn' })
        authenticationContexts = @($script:Contexts); references = $refs
    }
    ConvertTo-JsonText $all | Set-Content -Path (Join-Path $dir 'CA-export.json') -Encoding UTF8
    $pd = Join-Path $dir 'Policies'
    [void](New-Item -ItemType Directory -Path $pd -Force)
    foreach ($p in $script:Policies) {
        $n = ([string]$p.displayName) -replace '[\\/:*?"<>|]', '_'
        ConvertTo-JsonText $p | Set-Content -Path (Join-Path $pd "$n.json") -Encoding UTF8
    }
    # Leesbaar overzicht als CSV
    $script:Policies | ForEach-Object {
        [pscustomobject]@{ Naam = $_.displayName; Status = $script:StateLabels[[string]$_.state]; Gebruikers = Get-UsersShort $_; Resources = Get-AppsShort $_; Toegang = Get-GrantShort $_; Id = $_.id }
    } | Export-Csv -Path (Join-Path $dir 'Policies-overzicht.csv') -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Log "Export opgeslagen in $dir" "success"
    if ((Show-Msg "Export klaar:`n$dir`n`nMap openen?" "Export" "YesNo") -eq 'Yes') { Start-Process explorer.exe "`"$dir`"" }
}

function Read-ImportFile {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'JSON (*.json)|*.json'
    $dlg.Multiselect = $true
    $dlg.Title = 'CA-export.json of losse policy-bestanden kiezen'
    if (-not $dlg.ShowDialog($script:Window)) { return }
    $data = [ordered]@{ policies = @(); namedLocations = @(); authenticationStrengths = @(); authenticationContexts = @(); references = @{} }
    foreach ($f in $dlg.FileNames) {
        try { $j = ConvertTo-Hashtable (Get-Content $f -Raw -Encoding UTF8 | ConvertFrom-Json) }
        catch { Show-Msg "Kan $f niet lezen:`n$_" "Import" "OK" "Error" | Out-Null; continue }
        if ($j.Contains('policies')) {
            foreach ($k in 'policies', 'namedLocations', 'authenticationStrengths', 'authenticationContexts') { if ($j[$k]) { $data[$k] += @($j[$k]) } }
            if ($j.references) { foreach ($k in $j.references.Keys) { $data.references[$k] = $j.references[$k] } }
            if ($j.tenantId) { $data.tenantId = [string]$j.tenantId }
        }
        elseif ($j.Contains('object')) { $data.policies += @($j.object) }      # back-up uit deze tool
        elseif ($j.Contains('conditions')) { $data.policies += @($j) }
    }
    $script:ImportData = $data
    Show-ImportPreview
}

function Show-ImportPreview {
    $d = $script:ImportData
    if (-not $d) { return }
    $existing = @($script:Policies | ForEach-Object { [string]$_.displayName })
    $locNames = @($script:Locations | ForEach-Object { [string]$_.displayName })
    $rows = @()
    foreach ($l in $d.namedLocations) {
        $rows += @{ Soort = 'Locatie'; Naam = [string]$l.displayName; Actie = $(if ($locNames -contains $l.displayName) { 'Bestaat al - wordt gekoppeld' } else { 'Aanmaken' }); Key = "loc:$($l.id)" }
    }
    foreach ($s in $d.authenticationStrengths) {
        $ex = @($script:Strengths | Where-Object { $_.displayName -eq $s.displayName }).Count
        $rows += @{ Soort = 'Sterkte'; Naam = [string]$s.displayName; Actie = $(if ($ex) { 'Bestaat al - wordt gekoppeld' } else { 'Aanmaken' }); Key = "str:$($s.id)" }
    }
    foreach ($x in $d.authenticationContexts) {
        $ex = @($script:Contexts | Where-Object { $_.id -eq $x.id }).Count
        $rows += @{ Soort = 'Context'; Naam = "$($x.id) - $($x.displayName)"; Actie = $(if ($ex) { 'Bestaat al' } else { 'Aanmaken' }); Key = "ctx:$($x.id)" }
    }
    foreach ($p in $d.policies) {
        $rows += @{ Soort = 'Policy'; Naam = [string]$p.displayName; Actie = $(if ($existing -contains $p.displayName) { 'Bestaat al - overslaan' } else { 'Aanmaken' }); Key = "pol:$($p.displayName)" }
    }
    Set-GridRows $script:ImpGrid @('Soort', 'Naam', 'Actie', 'Key') $rows
    $script:ImpGrid.SelectAll()
    $script:LblImport.Text = "$($d.policies.Count) policies, $($d.namedLocations.Count) locaties, $($d.authenticationStrengths.Count) sterktes, $($d.authenticationContexts.Count) contexten ingelezen. Selecteer wat je wilt importeren."
}

function Invoke-Import {
    $d = $script:ImportData
    if (-not $d) { return }
    $keys = @(Get-SelectedRows $script:ImpGrid | Where-Object { $_['Actie'] -notlike '*overslaan*' } | ForEach-Object { $_['Key'] })
    if (-not $keys.Count) { return }
    $state = if ($script:RbImpReport.IsChecked) { 'enabledForReportingButNotEnforced' } else { 'disabled' }
    $r = Show-Msg "$($keys.Count) onderdeel/onderdelen importeren in tenant $($script:TenantId)?`nPolicies worden aangemaakt als '$($script:StateLabels[$state])'." "Importeren" "YesNo" "Question"
    if ($r -ne 'Yes') { return }
    Start-LongOp "Importeren..."
    $idMap = @{}   # oud ID -> nieuw ID
    $warn = New-Object System.Collections.ArrayList
    try {
        # 1. Locaties
        foreach ($l in $d.namedLocations) {
            $hit = $script:Locations | Where-Object { $_.displayName -eq $l.displayName } | Select-Object -First 1
            if ($hit) { $idMap[[string]$l.id] = [string]$hit.id; continue }
            if ($keys -notcontains "loc:$($l.id)") { continue }
            try {
                $n = Invoke-Graph "$($script:CaRoot)/namedLocations" -Method POST -Body (Get-CleanObject $l)
                $idMap[[string]$l.id] = [string]$n.id
                Write-Log "Locatie aangemaakt: $($l.displayName)" "success"
            }
            catch { [void]$warn.Add("Locatie $($l.displayName): $_") }
        }
        # 2. Sterktes (ingebouwde hebben in elke tenant hetzelfde ID)
        foreach ($s in $d.authenticationStrengths) {
            $hit = $script:Strengths | Where-Object { $_.displayName -eq $s.displayName } | Select-Object -First 1
            if ($hit) { $idMap[[string]$s.id] = [string]$hit.id; continue }
            if ($keys -notcontains "str:$($s.id)") { continue }
            try {
                $n = Invoke-Graph "$($script:CaRoot)/authenticationStrength/policies" -Method POST -Body @{ displayName = $s.displayName; description = $s.description; allowedCombinations = @($s.allowedCombinations) }
                $idMap[[string]$s.id] = [string]$n.id
                Write-Log "Sterkte aangemaakt: $($s.displayName)" "success"
            }
            catch { [void]$warn.Add("Sterkte $($s.displayName): $_") }
        }
        # 3. Contexten
        foreach ($x in $d.authenticationContexts) {
            if ($keys -notcontains "ctx:$($x.id)" -or @($script:Contexts | Where-Object { $_.id -eq $x.id }).Count) { continue }
            try { [void](Invoke-Graph "$($script:CaRoot)/authenticationContextClassReferences/$($x.id)" -Method PATCH -Body @{ displayName = $x.displayName; description = $x.description; isAvailable = [bool]$x.isAvailable }) }
            catch { [void]$warn.Add("Context $($x.id): $_") }
        }
        # 4. Groepen op naam koppelen
        $groupCache = @{}
        foreach ($p in $d.policies) {
            if ($script:StopRequested) { break }
            if ($keys -notcontains "pol:$($p.displayName)") { continue }
            $np = Copy-Deep $p
            $np.state = $state
            $u = $np.conditions.users
            $drop = @()
            foreach ($prop in 'includeGroups', 'excludeGroups') {
                $new = @()
                foreach ($gid in (Get-Arr $u $prop)) {
                    $gname = $d.references[$gid]
                    if (-not $gname) { $drop += "groep $gid"; continue }
                    if (-not $groupCache.ContainsKey($gname)) {
                        $flt = "displayName eq '$($gname.Replace("'", "''"))'"
                        $res = try { Invoke-Graph "$($script:GraphRoot)/groups?`$filter=$([uri]::EscapeDataString($flt))&`$select=id" -Quiet } catch { $null }
                        $groupCache[$gname] = if ($res -and @($res.value).Count -eq 1) { [string]$res.value[0].id } else { $null }
                    }
                    if ($groupCache[$gname]) { $new += $groupCache[$gname] } else { $drop += "groep '$gname'" }
                }
                $u[$prop] = @($new)
            }
            # Gebruikers bestaan niet in een andere tenant: weghalen, noodaccounts erbij
            foreach ($prop in 'includeUsers', 'excludeUsers') {
                $keep = @(Get-Arr $u $prop | Where-Object { -not (Test-Guid $_) -or ($d.tenantId -eq $script:TenantId) })
                $gone = @(Get-Arr $u $prop).Count - $keep.Count
                if ($gone) { $drop += "$gone gebruiker(s) in $prop" }
                $u[$prop] = $keep
            }
            $u.excludeUsers = @(@($u.excludeUsers) + @(Get-BreakGlass | ForEach-Object { $_.id }) | Select-Object -Unique)
            if (-not (Get-Arr $u 'includeUsers').Count -and -not (Get-Arr $u 'includeGroups').Count -and -not (Get-Arr $u 'includeRoles').Count -and -not $u.includeGuestsOrExternalUsers) {
                $u.includeUsers = @('None')
                $drop += 'geen gebruikers meer over - opnemen staat op Geen'
            }
            $lo = $np.conditions.locations
            if ($lo) {
                foreach ($prop in 'includeLocations', 'excludeLocations') {
                    $locIds = @($script:Locations | ForEach-Object { [string]$_.id })
                    $newLoc = @()
                    foreach ($lid in (Get-Arr $lo $prop)) {
                        if ($idMap.ContainsKey($lid)) { $newLoc += $idMap[$lid] }
                        elseif (-not (Test-Guid $lid) -or $locIds -contains $lid -or $lid -eq '00000000-0000-0000-0000-000000000000') { $newLoc += $lid }
                        else { $drop += "locatie $lid" }
                    }
                    $lo[$prop] = @($newLoc)
                }
            }
            $asId = [string](Get-Val $np 'grantControls.authenticationStrength.id')
            if ($asId -and $idMap.ContainsKey($asId)) { $np.grantControls.authenticationStrength = [ordered]@{ id = $idMap[$asId] } }
            if ($np.grantControls -and $np.grantControls.termsOfUse) { $np.grantControls.termsOfUse = @(); $drop += 'gebruiksvoorwaarden' }
            try {
                [void](Invoke-Graph "$($script:CaRoot)/policies" -Method POST -Body (Get-PolicyBody $np))
                Write-Log "Policy aangemaakt: $($np.displayName)$(if ($drop.Count) { ' - LET OP: ' + ($drop -join '; ') })" $(if ($drop.Count) { 'warning' } else { 'success' })
                if ($drop.Count) { [void]$warn.Add("$($np.displayName): niet overgenomen: $($drop -join '; ')") }
            }
            catch { [void]$warn.Add("Policy $($np.displayName): $_") }
        }
    }
    finally { Stop-LongOp }
    Update-AllData
    Show-ImportPreview
    if ($warn.Count) { Show-TextWindow 'Import - aandachtspunten' ($warn -join "`r`n`r`n") }
    else { Show-Msg "Import klaar." "Import" | Out-Null }
}

function Show-BreakGlass {
    $rows = @(Get-BreakGlass | ForEach-Object { @{ UPN = $_.upn; Id = $_.id } })
    Set-GridRows $script:BgGrid @('UPN', 'Id') $rows
}

function Add-BreakGlass {
    $pick = @(Show-ObjectPicker -Kinds user -Title 'Noodaccount(s) kiezen')
    if (-not $pick.Count) { return }
    $list = @(Get-BreakGlass)
    foreach ($p in $pick) {
        if (@($list | Where-Object { $_.id -eq $p.id }).Count) { continue }
        $upn = if ($p.name -match '\(([^)]+)\)$') { $Matches[1] } else { $p.name }
        $list += [ordered]@{ id = $p.id; upn = $upn }
    }
    Set-BreakGlass $list
    Show-BreakGlass
    Update-PolicyGrid
}

function Remove-BreakGlass {
    $ids = @(Get-SelectedRows $script:BgGrid | ForEach-Object { $_['Id'] })
    Set-BreakGlass @(Get-BreakGlass | Where-Object { $ids -notcontains $_.id })
    Show-BreakGlass
    Update-PolicyGrid
}

#endregion


#region --- XAML ---

$script:StylesXaml = @'
<Style TargetType="Button">
  <Setter Property="Foreground" Value="White"/>
  <Setter Property="Background" Value="#0078D4"/>
  <Setter Property="BorderThickness" Value="0"/>
  <Setter Property="Padding" Value="12,6"/>
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
<Style x:Key="BtnBlue"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"/>
<Style x:Key="BtnGreen" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#107C10"/></Style>
<Style x:Key="BtnRed"   TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#C42B1C"/></Style>
<Style x:Key="BtnGray"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#7A7A7A"/></Style>
<Style x:Key="BtnDark"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="#005A9E"/><Setter Property="Padding" Value="14,7"/><Setter Property="Margin" Value="8,0,0,0"/>
</Style>
<Style x:Key="BtnLog"   TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="#464646"/><Setter Property="Foreground" Value="#B4B4B4"/>
  <Setter Property="Padding" Value="10,2"/><Setter Property="Margin" Value="0,2,4,2"/>
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
<Style x:Key="CellQueue" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="Foreground" Value="#0063B1"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
</Style>
<Style x:Key="CellStatus" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="Foreground" Value="{Binding Kleur}"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
</Style>
<Style TargetType="GroupBox">
  <Setter Property="Margin" Value="0,0,0,8"/>
  <Setter Property="Padding" Value="8,6"/>
  <Setter Property="BorderBrush" Value="#CCCCCC"/>
  <Setter Property="Background" Value="White"/>
</Style>
<Style TargetType="RadioButton">
  <Setter Property="Margin" Value="0,0,0,6"/>
</Style>
<Style x:Key="FieldLabel" TargetType="TextBlock">
  <Setter Property="VerticalAlignment" Value="Center"/>
  <Setter Property="Margin" Value="0,0,8,6"/>
  <Setter Property="Foreground" Value="#333333"/>
</Style>
<Style x:Key="FieldBox" TargetType="TextBox">
  <Setter Property="Height" Value="26"/>
  <Setter Property="VerticalContentAlignment" Value="Center"/>
  <Setter Property="Margin" Value="0,0,0,6"/>
  <Setter Property="Padding" Value="4,0"/>
</Style>
<Style x:Key="Hint" TargetType="TextBlock">
  <Setter Property="Foreground" Value="Gray"/>
  <Setter Property="TextWrapping" Value="Wrap"/>
  <Setter Property="Margin" Value="0,0,0,8"/>
</Style>
'@



$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Conditional Access Beheer" Width="1500" Height="940" MinWidth="1100" MinHeight="680"
        WindowStartupLocation="CenterScreen" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
<Style x:Key="Tool" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Padding" Value="10,4"/><Setter Property="Margin" Value="0,0,6,6"/>
</Style>
<Style x:Key="ToolGray" TargetType="Button" BasedOn="{StaticResource Tool}"><Setter Property="Background" Value="#7A7A7A"/></Style>
<Style x:Key="ToolGreen" TargetType="Button" BasedOn="{StaticResource Tool}"><Setter Property="Background" Value="#107C10"/></Style>
<Style x:Key="ToolRed" TargetType="Button" BasedOn="{StaticResource Tool}"><Setter Property="Background" Value="#C42B1C"/></Style>
<Style x:Key="ToolOrange" TargetType="Button" BasedOn="{StaticResource Tool}"><Setter Property="Background" Value="#B26200"/></Style>
<Style x:Key="Sep" TargetType="Border"><Setter Property="Width" Value="1"/><Setter Property="Background" Value="#CCCCCC"/><Setter Property="Margin" Value="4,2,10,8"/></Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="*" MinHeight="380"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="170" MinHeight="70"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0078D4">
      <DockPanel Margin="16,0,12,0">
        <StackPanel x:Name="HeaderPanel" DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Admin-account:" Foreground="White" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="TxtAdmin" Width="210" Height="26" VerticalContentAlignment="Center"
                   ToolTip="Wordt als hint in het inlogscherm gezet; je kiest altijd zelf het account."/>
          <TextBlock Text="Tenant (optioneel, GDAP):" Foreground="White" VerticalAlignment="Center" Margin="12,0,6,0"/>
          <TextBox x:Name="TxtTenant" Width="170" Height="26" VerticalContentAlignment="Center"
                   ToolTip="Leeg = thuis-tenant van het admin-account. Voor een klant via GDAP: domein (klant.onmicrosoft.com) of tenant-ID."/>
          <TextBlock x:Name="LblConn" Text="Niet verbonden" Foreground="#FFC864" VerticalAlignment="Center" Margin="14,0,0,0" MaxWidth="260" TextTrimming="CharacterEllipsis"/>
          <Button x:Name="BtnConnect" Content="Inloggen" Style="{StaticResource BtnDark}"/>
          <Button x:Name="BtnDisconnect" Content="Uitloggen" Style="{StaticResource BtnDark}" IsEnabled="False"/>
        </StackPanel>
        <TextBlock Text="Conditional Access Beheer" Foreground="White" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </DockPanel>
    </Border>

    <TabControl x:Name="MainArea" Grid.Row="1" Margin="6,6,6,0" IsEnabled="False">

      <TabItem Header="Policies">
        <DockPanel Margin="6">
          <WrapPanel DockPanel.Dock="Top">
            <Button x:Name="BtnRefresh" Content="Vernieuwen" Style="{StaticResource ToolGray}"/>
            <Border Style="{StaticResource Sep}"/>
            <Button x:Name="BtnPolNew" Content="Nieuw" Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnPolTpl" Content="Nieuw uit sjabloon" Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnPolEdit" Content="Bewerken" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnPolDup" Content="Dupliceren" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnPolJson" Content="JSON bekijken" Style="{StaticResource ToolGray}"/>
            <Border Style="{StaticResource Sep}"/>
            <Button x:Name="BtnStOn" Content="Aan" Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnStReport" Content="Alleen rapport" Style="{StaticResource ToolOrange}"/>
            <Button x:Name="BtnStOff" Content="Uit" Style="{StaticResource ToolGray}"/>
            <Border Style="{StaticResource Sep}"/>
            <Button x:Name="BtnExAdd" Content="Uitsluiten in selectie..." Style="{StaticResource Tool}"/>
            <Button x:Name="BtnExDel" Content="Uitsluiting weghalen..." Style="{StaticResource Tool}"/>
            <Button x:Name="BtnExBg" Content="Noodaccounts uitsluiten" Style="{StaticResource Tool}"/>
            <Border Style="{StaticResource Sep}"/>
            <Button x:Name="BtnPolDel" Content="Verwijderen" Style="{StaticResource ToolRed}"/>
          </WrapPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
            <ComboBox x:Name="CmbStateFilter" DockPanel.Dock="Right" Width="150" Height="26" Margin="8,0,0,0"/>
            <TextBlock x:Name="LblPolCount" DockPanel.Dock="Right" VerticalAlignment="Center" Foreground="Gray" Margin="8,0,0,0"/>
            <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="TxtPolFilter" Height="26" VerticalContentAlignment="Center" ToolTip="Zoekt in naam, gebruikers, resources en toegang"/>
          </DockPanel>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="3*"/><ColumnDefinition Width="5"/><ColumnDefinition Width="2*" MinWidth="300"/></Grid.ColumnDefinitions>
            <DataGrid x:Name="PolGrid">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Naam"       Binding="{Binding Naam}"       Width="2*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Status"     Binding="{Binding Status}"     Width="95" ElementStyle="{StaticResource CellStatus}"/>
                <DataGridTextColumn Header="Gebruikers" Binding="{Binding Gebruikers}" Width="*"  ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Resources"  Binding="{Binding Resources}"  Width="*"  ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Toegang"    Binding="{Binding Toegang}"    Width="1.3*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Nood"       Binding="{Binding Nood}"       Width="50" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Gewijzigd"  Binding="{Binding Gewijzigd}"  Width="105" ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
            <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="Transparent"/>
            <TextBox x:Name="TxtPolDetail" Grid.Column="2" IsReadOnly="True" FontFamily="Consolas" FontSize="12" Padding="6"
                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Background="White"/>
          </Grid>
        </DockPanel>
      </TabItem>

      <TabItem Header="Toewijzingen">
        <DockPanel Margin="6">
          <WrapPanel DockPanel.Dock="Top">
            <Button x:Name="BtnAsgAdd" Content="Toevoegen aan policies..." Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnAsgCopy" Content="Kopieren naar policies..." Style="{StaticResource Tool}"/>
            <Button x:Name="BtnAsgFlip" Content="Omdraaien (opnemen/uitsluiten)" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnAsgReplace" Content="Vervangen door..." Style="{StaticResource Tool}"/>
            <Button x:Name="BtnAsgDel" Content="Weghalen" Style="{StaticResource ToolRed}"/>
            <Border Style="{StaticResource Sep}"/>
            <Button x:Name="BtnAsgCsv" Content="Exporteren (CSV)" Style="{StaticResource ToolGray}"/>
          </WrapPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
            <ComboBox x:Name="CmbAsgDir" DockPanel.Dock="Right" Width="180" Height="26" Margin="8,0,0,0" DisplayMemberPath="Label" SelectedValuePath="Value"/>
            <ComboBox x:Name="CmbAsgKind" DockPanel.Dock="Right" Width="130" Height="26" Margin="8,0,0,0" DisplayMemberPath="Label" SelectedValuePath="Value"/>
            <TextBlock x:Name="LblAsgCount" DockPanel.Dock="Right" VerticalAlignment="Center" Foreground="Gray" Margin="8,0,0,0"/>
            <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="TxtAsgFilter" Height="26" VerticalContentAlignment="Center" ToolTip="Zoekt in objectnaam, object-ID en policynaam"/>
          </DockPanel>
          <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,4"
                     Text="Links alle gebruikers, groepen, rollen en apps uit de policies; selecteer er een of meer om rechts te zien waar ze staan. Selecteer rechts regels (Ctrl/Shift) voor Kopieren, Omdraaien, Vervangen of Weghalen. Dubbelklik rechts opent de policy."/>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="2*"/><ColumnDefinition Width="5"/><ColumnDefinition Width="3*"/></Grid.ColumnDefinitions>
            <DataGrid x:Name="AsgObjGrid">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Soort"          Binding="{Binding Soort}"       Width="75" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Naam"           Binding="{Binding Naam}"        Width="*"  ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Opgenomen in"   Binding="{Binding Opgenomen}"   Width="90" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Uitgesloten in" Binding="{Binding Uitgesloten}" Width="95" ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
            <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="Transparent"/>
            <DataGrid x:Name="AsgGrid" Grid.Column="2">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Object"   Binding="{Binding Object}"   Width="*"    ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Soort"    Binding="{Binding Soort}"    Width="75"   ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Richting" Binding="{Binding Richting}" Width="90"   ElementStyle="{StaticResource CellStatus}"/>
                <DataGridTextColumn Header="Policy"   Binding="{Binding Policy}"   Width="1.4*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Status"   Binding="{Binding Status}"   Width="95"   ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
          </Grid>
        </DockPanel>
      </TabItem>

      <TabItem Header="Locaties">
        <DockPanel Margin="6">
          <WrapPanel DockPanel.Dock="Top">
            <Button x:Name="BtnLocIp" Content="Nieuwe IP-locatie" Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnLocCountry" Content="Nieuwe landen-locatie" Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnLocEdit" Content="Bewerken" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnLocDel" Content="Verwijderen" Style="{StaticResource ToolRed}"/>
          </WrapPanel>
          <DataGrid x:Name="LocGrid">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Naam"        Binding="{Binding Naam}"        Width="*"   ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Type"        Binding="{Binding Type}"        Width="80"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Vertrouwd"   Binding="{Binding Vertrouwd}"   Width="80"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Waarden"     Binding="{Binding Waarden}"     Width="2*"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Gebruikt in" Binding="{Binding Gebruikt}"    Width="2*"  ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </TabItem>

      <TabItem Header="Authenticatiesterktes">
        <DockPanel Margin="6">
          <WrapPanel DockPanel.Dock="Top">
            <Button x:Name="BtnStrNew" Content="Nieuw" Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnStrEdit" Content="Bewerken" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnStrDup" Content="Dupliceren" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnStrDel" Content="Verwijderen" Style="{StaticResource ToolRed}"/>
          </WrapPanel>
          <DataGrid x:Name="StrGrid">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Naam"         Binding="{Binding Naam}"         Width="*"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Type"         Binding="{Binding Type}"         Width="90" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Combinaties"  Binding="{Binding Combinaties}"  Width="3*" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Gebruikt in"  Binding="{Binding Gebruikt}"     Width="2*" ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </TabItem>

      <TabItem Header="Authenticatiecontexten">
        <DockPanel Margin="6">
          <WrapPanel DockPanel.Dock="Top">
            <Button x:Name="BtnCtxNew" Content="Nieuw" Style="{StaticResource ToolGreen}"/>
            <Button x:Name="BtnCtxEdit" Content="Bewerken" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnCtxDel" Content="Verwijderen" Style="{StaticResource ToolRed}"/>
          </WrapPanel>
          <DataGrid x:Name="CtxGrid">
            <DataGrid.Columns>
              <DataGridTextColumn Header="ID"           Binding="{Binding Id}"           Width="50"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Naam"         Binding="{Binding Naam}"         Width="*"   ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Omschrijving" Binding="{Binding Omschrijving}" Width="2*"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Gepubliceerd" Binding="{Binding Beschikbaar}"  Width="90"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Gebruikt in"  Binding="{Binding Gebruikt}"     Width="2*"  ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </TabItem>

      <TabItem Header="Gebruiker-impact">
        <Grid Margin="6">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="5"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <StackPanel>
            <WrapPanel>
              <Button x:Name="BtnImpactPick" Content="Gebruiker kiezen..." Style="{StaticResource Tool}"/>
              <CheckBox x:Name="ChkImpactAll" Content="Ook niet-van-toepassing tonen" VerticalAlignment="Center" Margin="6,0,0,6"/>
              <TextBlock x:Name="LblImpactUser" FontWeight="SemiBold" VerticalAlignment="Center" Margin="14,0,0,6"/>
            </WrapPanel>
            <TextBlock x:Name="LblImpactInfo" Style="{StaticResource Hint}" Margin="0,0,0,4"/>
          </StackPanel>
          <DataGrid x:Name="ImpactGrid" Grid.Row="1">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Resultaat" Binding="{Binding Resultaat}" Width="115" ElementStyle="{StaticResource CellStatus}"/>
              <DataGridTextColumn Header="Waarom"    Binding="{Binding Reden}"     Width="*"   ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Policy"    Binding="{Binding Naam}"      Width="1.5*" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Status"    Binding="{Binding Status}"    Width="95"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Resources" Binding="{Binding Apps}"      Width="*"   ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Toegang"   Binding="{Binding Toegang}"   Width="1.3*" ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
          <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" Background="Transparent" ResizeDirection="Rows"/>
          <DockPanel Grid.Row="3" Margin="0,4,0,0">
            <WrapPanel DockPanel.Dock="Top">
              <TextBlock Text="Aanmeldingen van de afgelopen" VerticalAlignment="Center" Margin="0,0,6,6"/>
              <ComboBox x:Name="CmbSignDays" Width="55" Height="24" Margin="0,0,6,6"/>
              <TextBlock Text="dagen" VerticalAlignment="Center" Margin="0,0,10,6"/>
              <Button x:Name="BtnSignLoad" Content="Ophalen" Style="{StaticResource Tool}"/>
            </WrapPanel>
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="5"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <DataGrid x:Name="SignGrid" SelectionMode="Single">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Tijd"      Binding="{Binding Tijd}"      Width="100" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="App"       Binding="{Binding App}"       Width="*"   ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Locatie"   Binding="{Binding Locatie}"   Width="90"  ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Resultaat" Binding="{Binding Resultaat}" Width="80"  ElementStyle="{StaticResource CellStatus}"/>
                  <DataGridTextColumn Header="CA"        Binding="{Binding CA}"        Width="80"  ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="IP"        Binding="{Binding IP}"        Width="100" ElementStyle="{StaticResource CellText}"/>
                </DataGrid.Columns>
              </DataGrid>
              <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="Transparent"/>
              <TextBox x:Name="TxtSignDetail" Grid.Column="2" IsReadOnly="True" FontFamily="Consolas" FontSize="11" Padding="6"
                       VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Background="White"/>
            </Grid>
          </DockPanel>
        </Grid>
      </TabItem>

      <TabItem Header="Back-up / import">
        <Grid Margin="6">
          <Grid.ColumnDefinitions><ColumnDefinition Width="2*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*" MinWidth="320"/></Grid.ColumnDefinitions>
          <DockPanel>
            <GroupBox DockPanel.Dock="Top" Header="Exporteren">
              <StackPanel>
                <TextBlock Style="{StaticResource Hint}" Text="Alle policies, locaties, eigen sterktes en contexten als JSON (1 totaalbestand + 1 bestand per policy) en een CSV-overzicht. Groepsnamen gaan mee, zodat import in een andere tenant groepen op naam kan koppelen."/>
                <WrapPanel>
                  <Button x:Name="BtnExport" Content="Alles exporteren..." Style="{StaticResource Tool}"/>
                  <Button x:Name="BtnBackupDir" Content="Map met automatische back-ups openen" Style="{StaticResource ToolGray}"/>
                </WrapPanel>
              </StackPanel>
            </GroupBox>
            <GroupBox Header="Importeren (bijv. baseline naar een klant-tenant)">
              <DockPanel>
                <StackPanel DockPanel.Dock="Top">
                  <TextBlock Style="{StaticResource Hint}" Text="Kies een CA-export.json, losse policy-bestanden of back-ups uit deze tool. Locaties en sterktes worden op naam gekoppeld of aangemaakt; groepen op naam gekoppeld; gebruikers uit een andere tenant vervallen; noodaccounts van deze tenant worden uitgesloten. Bestaande policies (zelfde naam) worden overgeslagen."/>
                  <WrapPanel>
                    <Button x:Name="BtnImpRead" Content="Bestand(en) inlezen..." Style="{StaticResource Tool}"/>
                    <RadioButton x:Name="RbImpReport" Content="Aanmaken als 'Alleen rapport'" IsChecked="True" VerticalAlignment="Center" Margin="8,0,12,6"/>
                    <RadioButton x:Name="RbImpOff" Content="Aanmaken als 'Uit'" VerticalAlignment="Center" Margin="0,0,12,6"/>
                    <Button x:Name="BtnImpRun" Content="Geselecteerde importeren" Style="{StaticResource ToolGreen}"/>
                  </WrapPanel>
                  <TextBlock x:Name="LblImport" Style="{StaticResource Hint}"/>
                </StackPanel>
                <DataGrid x:Name="ImpGrid">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Soort" Binding="{Binding Soort}" Width="80"  ElementStyle="{StaticResource CellText}"/>
                    <DataGridTextColumn Header="Naam"  Binding="{Binding Naam}"  Width="2*"  ElementStyle="{StaticResource CellText}"/>
                    <DataGridTextColumn Header="Actie" Binding="{Binding Actie}" Width="*"   ElementStyle="{StaticResource CellText}"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </GroupBox>
          </DockPanel>
          <GroupBox Grid.Column="2" Header="Noodaccounts (break-glass) van deze tenant">
            <DockPanel>
              <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}" Text="Worden bij sjablonen en import automatisch uitgesloten, en de kolom 'Nood' laat per policy zien of ze uitgesloten zijn. Opgeslagen per tenant op deze pc."/>
              <WrapPanel DockPanel.Dock="Top">
                <Button x:Name="BtnBgAdd" Content="Toevoegen..." Style="{StaticResource Tool}"/>
                <Button x:Name="BtnBgDel" Content="Verwijderen" Style="{StaticResource ToolGray}"/>
              </WrapPanel>
              <DataGrid x:Name="BgGrid">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Account" Binding="{Binding UPN}" Width="*" ElementStyle="{StaticResource CellText}"/>
                </DataGrid.Columns>
              </DataGrid>
            </DockPanel>
          </GroupBox>
        </Grid>
      </TabItem>
    </TabControl>

    <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" Background="#3C3C3C" ResizeDirection="Rows"/>

    <DockPanel Grid.Row="3" Background="#1E1E1E">
      <Border DockPanel.Dock="Top" Background="#2D2D2D" Height="30">
        <DockPanel>
          <Button x:Name="BtnClearLog" DockPanel.Dock="Right" Content="Leegmaken" Style="{StaticResource BtnLog}"/>
          <TextBlock Text="  Log" Foreground="#B4B4B4" FontWeight="SemiBold" VerticalAlignment="Center"/>
        </DockPanel>
      </Border>
      <RichTextBox x:Name="LogBox" IsReadOnly="True" Background="#141414" Foreground="#C8C8C8" BorderThickness="0"
                   FontFamily="Consolas" FontSize="12" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
        <FlowDocument PageWidth="5000">
          <Paragraph x:Name="LogPara" Margin="4,2"/>
        </FlowDocument>
      </RichTextBox>
    </DockPanel>

    <Border x:Name="StatusBar" Grid.Row="4" Background="#0078D4">
      <DockPanel>
        <Button x:Name="BtnStop" DockPanel.Dock="Right" Content="Stoppen" Style="{StaticResource BtnRed}" Padding="14,2" Margin="0,3,8,3" Visibility="Collapsed"/>
        <TextBlock x:Name="LblStatus" Foreground="White" VerticalAlignment="Center" Margin="12,0" Text="Welkom - log in met je admin-account"/>
      </DockPanel>
    </Border>
  </Grid>
</Window>
'@

function ConvertFrom-Xaml {
    param([string]$Text)
    [xml]$x = $Text
    return [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
}

#endregion


#region --- Overzichten vullen ---

function Update-AllData {
    Enter-Busy "Gegevens laden..."
    try {
        Get-AllCaData
        Update-PolicyGrid
        Update-AssignmentView -Reload
        Update-LocationGrid
        Update-StrengthGrid
        Update-ContextGrid
        Show-Impact
        Set-Status "Klaar - $($script:Policies.Count) policies" "success"
    }
    catch { Set-Status "Laden mislukt: $_" "error" }
    finally { Exit-Busy }
}

function Update-PolicyGrid {
    $f = $script:TxtPolFilter.Text.Trim()
    $sf = [string]$script:CmbStateFilter.SelectedItem
    $bg = @(Get-BreakGlass | ForEach-Object { $_.id })
    $keep = @(Get-SelectedRows $script:PolGrid | ForEach-Object { $_['Id'] })
    $rows = foreach ($p in @($script:Policies | Sort-Object { $_.displayName })) {
        $state = $script:StateLabels[[string]$p.state]
        if ($sf -and $sf -ne 'Alle statussen' -and $sf -ne $state) { continue }
        $r = @{
            Naam = [string]$p.displayName; Status = $state; Gebruikers = Get-UsersShort $p; Resources = Get-AppsShort $p
            Toegang = Get-GrantShort $p; Id = [string]$p.id
            Gewijzigd = $(if ($p.modifiedDateTime) { ([datetime]$p.modifiedDateTime).ToLocalTime().ToString('dd-MM-yy HH:mm') } elseif ($p.createdDateTime) { ([datetime]$p.createdDateTime).ToLocalTime().ToString('dd-MM-yy HH:mm') } else { '' })
            Kleur = switch ([string]$p.state) { 'enabled' { '#107C10' } 'disabled' { '#8A8A8A' } default { '#B26200' } }
        }
        $ex = @(Get-Arr $p 'conditions.users.excludeUsers')
        $r.Nood = if (-not $bg.Count) { '' } elseif (@($bg | Where-Object { $ex -notcontains $_ }).Count) { 'NEE' } else { 'ja' }
        if ($f) {
            $hay = "$($r.Naam) $($r.Gebruikers) $($r.Resources) $($r.Toegang)"
            if ($hay -notlike "*$f*") { continue }
        }
        $r
    }
    $rows = @($rows)
    Set-GridRows $script:PolGrid @('Naam', 'Status', 'Gebruikers', 'Resources', 'Toegang', 'Nood', 'Gewijzigd', 'Id', 'Kleur') $rows
    $on = @($script:Policies | Where-Object { $_.state -eq 'enabled' }).Count
    $rep = @($script:Policies | Where-Object { $_.state -eq 'enabledForReportingButNotEnforced' }).Count
    $script:LblPolCount.Text = "$($rows.Count) van $($script:Policies.Count)  (aan $on, rapport $rep)"
    foreach ($item in @($script:PolGrid.Items)) { if ($keep -contains $item.Row['Id']) { [void]$script:PolGrid.SelectedItems.Add($item) } }
}

function Show-PolicyDetail {
    $sel = @(Get-SelectedPolicies)
    if ($sel.Count -eq 1) { $script:TxtPolDetail.Text = Get-PolicySummary $sel[0] }
    elseif ($sel.Count -gt 1) { $script:TxtPolDetail.Text = "$($sel.Count) policies geselecteerd:`r`n`r`n$(($sel | ForEach-Object { '- ' + $_.displayName }) -join "`r`n")" }
    else { $script:TxtPolDetail.Text = '' }
}

function Update-LocationGrid {
    $rows = foreach ($l in @($script:Locations | Sort-Object { $_.displayName })) {
        $isCountry = $l['@odata.type'] -like '*country*'
        @{
            Naam = [string]$l.displayName; Id = [string]$l.id
            Type = $(if ($isCountry) { 'Landen' } else { 'IP' })
            Vertrouwd = $(if ($isCountry) { '-' } elseif ($l.isTrusted) { 'Ja' } else { 'Nee' })
            Waarden = $(if ($isCountry) { (@($l.countriesAndRegions) -join ', ') + $(if ($l.includeUnknownCountriesAndRegions) { ' + onbekend' }) } else { @($l.ipRanges | ForEach-Object { $_.cidrAddress }) -join ', ' })
            Gebruikt = (Get-UsedIn $l.id 'location') -join ', '
        }
    }
    Set-GridRows $script:LocGrid @('Naam', 'Type', 'Vertrouwd', 'Waarden', 'Gebruikt', 'Id') @($rows)
}

function Update-StrengthGrid {
    $rows = foreach ($s in @($script:Strengths | Sort-Object { $_.policyType -ne 'builtIn' }, { $_.displayName })) {
        @{
            Naam = [string]$s.displayName; Id = [string]$s.id
            Type = $(if ($s.policyType -eq 'builtIn') { 'Ingebouwd' } else { 'Eigen' })
            Combinaties = @($s.allowedCombinations) -join ' | '
            Gebruikt = (Get-UsedIn $s.id 'strength') -join ', '
        }
    }
    Set-GridRows $script:StrGrid @('Naam', 'Type', 'Combinaties', 'Gebruikt', 'Id') @($rows)
}

function Update-ContextGrid {
    $rows = foreach ($c in @($script:Contexts | Sort-Object { [int]($_.id -replace '\D', '') })) {
        @{
            Id = [string]$c.id; Naam = [string]$c.displayName; Omschrijving = [string]$c.description
            Beschikbaar = $(if ($c.isAvailable) { 'Ja' } else { 'Nee' })
            Gebruikt = (Get-UsedIn $c.id 'context') -join ', '
        }
    }
    Set-GridRows $script:CtxGrid @('Id', 'Naam', 'Omschrijving', 'Beschikbaar', 'Gebruikt') @($rows)
}

function Get-GridSelection {
    param($Grid, $List)
    $ids = @(Get-SelectedRows $Grid | ForEach-Object { $_['Id'] })
    return @($List | Where-Object { $ids -contains $_.id })
}

#endregion

#region --- Hoofdvenster ---

function Invoke-Connect {
    $tenant = $script:TxtTenant.Text.Trim()
    $hint = $script:TxtAdmin.Text.Trim()
    Set-Status "Inloggen - kies in de browser je admin-account..."
    $script:HeaderPanel.IsEnabled = $false
    $script:Window.WindowState = 'Minimized'
    $upn = $null
    try { $upn = Connect-ToGraph -Tenant $tenant -LoginHint $hint }
    catch {
        $script:Window.WindowState = 'Normal'
        Show-Msg "Inloggen mislukt:`n$_" "Inloggen" "OK" "Error" | Out-Null
        Set-Status "Inloggen mislukt" "error"
    }
    finally {
        $script:Window.WindowState = 'Normal'
        [void]$script:Window.Activate()
        $script:HeaderPanel.IsEnabled = $true
    }
    if (-not $upn) { return }
    $ctx = Get-MgContext
    $script:TenantId = [string]$ctx.TenantId
    if (-not $script:TenantId -and $script:Auth) { $script:TenantId = [string]$script:Auth.Account.HomeAccountId.TenantId }
    $script:Connected = $true
    $s = Get-Settings
    $s.LastAdmin = $upn; $s.LastTenant = $tenant
    Save-Settings
    $script:TxtAdmin.Text = $upn
    $script:Names = @{}
    $script:RolesLoaded = $false
    $script:Combinations = $null
    $script:ImportData = $null
    $script:ImpactUser = $null
    $org = ''
    try {
        # Tenant-ID uit de organisatie: bij GDAP is dat de klant, niet de thuis-tenant van het account
        $o = (Invoke-Graph "$($script:GraphRoot)/organization?`$select=displayName,id" -Quiet).value[0]
        $org = [string]$o.displayName
        if ($o.id) { $script:TenantId = [string]$o.id }
    }
    catch { }
    $script:LblConn.Text = "$upn @ $(if ($org) { $org } else { $script:TenantId })"
    $script:LblConn.ToolTip = "Account: $upn`nTenant: $org ($($script:TenantId))"
    $script:LblConn.Foreground = New-Brush "#90EE90"
    $script:BtnConnect.Content = 'Opnieuw inloggen'
    $script:BtnDisconnect.IsEnabled = $true
    $script:MainArea.IsEnabled = $true
    Write-Log "Ingelogd als $upn - tenant $org ($($script:TenantId))" "success"
    if ($script:Auth) {
        $missing = @($script:GraphScopes | Where-Object { $script:Auth.Scopes -notcontains $_ })
        if ($missing.Count) { Write-Log "LET OP: niet toegekend (admin consent?): $($missing -join ', ')" "warning" }
    }
    try {
        $me = Invoke-Graph "$($script:GraphRoot)/me?`$select=id" -Quiet
        $script:Me = Get-UserContext $me.id
    }
    catch { Write-Log "Eigen account niet opgehaald - buitensluit-controle werkt beperkt (bij GDAP normaal)" "warning" }
    Show-BreakGlass
    if (-not @(Get-BreakGlass).Count) { Write-Log "Nog geen noodaccounts ingesteld voor deze tenant (tab Back-up / import)" "warning" }
    Update-AllData
}

function New-MainWindow {
    $script:Window = ConvertFrom-Xaml $script:MainXaml.Replace('{{STYLES}}', $script:StylesXaml)
    $names = @(
        'HeaderPanel', 'TxtAdmin', 'TxtTenant', 'LblConn', 'BtnConnect', 'BtnDisconnect', 'MainArea',
        'BtnRefresh', 'BtnPolNew', 'BtnPolTpl', 'BtnPolEdit', 'BtnPolDup', 'BtnPolJson', 'BtnStOn', 'BtnStReport', 'BtnStOff',
        'BtnExAdd', 'BtnExDel', 'BtnExBg', 'BtnPolDel', 'CmbStateFilter', 'LblPolCount', 'TxtPolFilter', 'PolGrid', 'TxtPolDetail',
        'BtnAsgAdd', 'BtnAsgCopy', 'BtnAsgFlip', 'BtnAsgReplace', 'BtnAsgDel', 'BtnAsgCsv', 'CmbAsgDir', 'CmbAsgKind', 'LblAsgCount', 'TxtAsgFilter', 'AsgObjGrid', 'AsgGrid',
        'BtnLocIp', 'BtnLocCountry', 'BtnLocEdit', 'BtnLocDel', 'LocGrid',
        'BtnStrNew', 'BtnStrEdit', 'BtnStrDup', 'BtnStrDel', 'StrGrid',
        'BtnCtxNew', 'BtnCtxEdit', 'BtnCtxDel', 'CtxGrid',
        'BtnImpactPick', 'ChkImpactAll', 'LblImpactUser', 'LblImpactInfo', 'ImpactGrid', 'CmbSignDays', 'BtnSignLoad', 'SignGrid', 'TxtSignDetail',
        'BtnExport', 'BtnBackupDir', 'BtnImpRead', 'RbImpReport', 'RbImpOff', 'BtnImpRun', 'LblImport', 'ImpGrid', 'BtnBgAdd', 'BtnBgDel', 'BgGrid',
        'BtnClearLog', 'LogBox', 'LogPara', 'StatusBar', 'LblStatus', 'BtnStop'
    )
    foreach ($n in $names) {
        $ctl = $script:Window.FindName($n)
        if ($null -eq $ctl) { throw "XAML-element '$n' niet gevonden" }
        Set-Variable -Scope Script -Name $n -Value $ctl
    }
    $s = Get-Settings
    $script:TxtAdmin.Text = [string]$s.LastAdmin
    $script:TxtTenant.Text = [string]$s.LastTenant
    $script:CmbStateFilter.ItemsSource = @('Alle statussen', 'Aan', 'Alleen rapport', 'Uit')
    $script:CmbStateFilter.SelectedIndex = 0
    $script:CmbAsgKind.ItemsSource = @(
        [pscustomobject]@{ Label = 'Alle soorten'; Value = '' }, [pscustomobject]@{ Label = 'Gebruikers'; Value = 'user' }
        [pscustomobject]@{ Label = 'Groepen'; Value = 'group' }, [pscustomobject]@{ Label = 'Rollen'; Value = 'role' }
        [pscustomobject]@{ Label = 'Apps'; Value = 'app' }
    )
    $script:CmbAsgKind.SelectedIndex = 0
    $script:CmbAsgDir.ItemsSource = @(
        [pscustomobject]@{ Label = 'Opgenomen en uitgesloten'; Value = '' }
        [pscustomobject]@{ Label = 'Alleen opgenomen'; Value = 'in' }, [pscustomobject]@{ Label = 'Alleen uitgesloten'; Value = 'ex' }
    )
    $script:CmbAsgDir.SelectedIndex = 0
    $script:CmbSignDays.ItemsSource = @(1, 3, 7, 14, 30)
    $script:CmbSignDays.SelectedItem = 7

    # --- Verbinden ---
    $script:BtnConnect.Add_Click({ Invoke-Connect })
    $script:BtnDisconnect.Add_Click({
        Disconnect-All
        $script:MainArea.IsEnabled = $false
        $script:BtnDisconnect.IsEnabled = $false
        $script:BtnConnect.Content = 'Inloggen'
        $script:LblConn.Text = 'Niet verbonden'
        $script:LblConn.Foreground = New-Brush "#FFC864"
        $script:AsgRows = @()
        foreach ($g in $script:PolGrid, $script:AsgObjGrid, $script:AsgGrid, $script:LocGrid, $script:StrGrid, $script:CtxGrid, $script:ImpactGrid, $script:SignGrid, $script:ImpGrid, $script:BgGrid) { $g.ItemsSource = $null }
        $script:TxtPolDetail.Text = ''; $script:TxtSignDetail.Text = ''; $script:LblImpactUser.Text = ''
        $script:Policies = @(); $script:Locations = @(); $script:Strengths = @(); $script:Contexts = @()
        Write-Log "Uitgelogd" "info"
        Set-Status "Uitgelogd - log in met je admin-account"
    })

    # --- Policies ---
    $script:BtnRefresh.Add_Click({ Update-AllData })
    $script:TxtPolFilter.Add_TextChanged({ Update-PolicyGrid })
    $script:CmbStateFilter.Add_SelectionChanged({ if ($script:Connected) { Update-PolicyGrid } })
    $script:PolGrid.Add_SelectionChanged({ Show-PolicyDetail })
    $script:PolGrid.Add_MouseDoubleClick({ $p = @(Get-SelectedPolicies); if ($p.Count -eq 1) { Invoke-PolicyEdit $p[0] } })
    $script:BtnPolNew.Add_Click({ Invoke-NewPolicy })
    $script:BtnPolTpl.Add_Click({ Invoke-NewPolicy -FromTemplate })
    $script:BtnPolEdit.Add_Click({
        $p = @(Get-SelectedPolicies)
        if ($p.Count -ne 1) { Show-Msg "Selecteer 1 policy." "Bewerken" | Out-Null; return }
        Invoke-PolicyEdit $p[0]
    })
    $script:BtnPolDup.Add_Click({
        $p = @(Get-SelectedPolicies)
        if ($p.Count -ne 1) { Show-Msg "Selecteer 1 policy." "Dupliceren" | Out-Null; return }
        Invoke-PolicyEdit $p[0] -Duplicate
    })
    $script:BtnPolJson.Add_Click({
        $p = @(Get-SelectedPolicies)
        if (-not $p.Count) { return }
        Show-TextWindow "JSON - $(if ($p.Count -eq 1) { $p[0].displayName } else { "$($p.Count) policies" })" (ConvertTo-JsonText $(if ($p.Count -eq 1) { $p[0] } else { $p }))
    })
    $script:BtnStOn.Add_Click({ Set-PolicyState 'enabled' })
    $script:BtnStReport.Add_Click({ Set-PolicyState 'enabledForReportingButNotEnforced' })
    $script:BtnStOff.Add_Click({ Set-PolicyState 'disabled' })
    $script:BtnExAdd.Add_Click({
        if (-not @(Get-SelectedPolicies).Count) { Show-Msg "Selecteer eerst de policies." "Uitsluiten" | Out-Null; return }
        $o = @(Show-ObjectPicker -Kinds user, group, role -Title 'Uitsluiten in de geselecteerde policies')
        if ($o.Count) { Update-PolicyExclusion -Mode Add -Objects $o }
    })
    $script:BtnExDel.Add_Click({
        $sel = @(Get-SelectedPolicies)
        if (-not $sel.Count) { Show-Msg "Selecteer eerst de policies." "Uitsluiten" | Out-Null; return }
        # Alleen kiezen uit wat nu in de selectie is uitgesloten
        $opts = @{}
        foreach ($p in $sel) {
            foreach ($pair in @(@('excludeUsers', 'user'), @('excludeGroups', 'group'), @('excludeRoles', 'role'))) {
                foreach ($id in (Get-Arr $p "conditions.users.$($pair[0])")) { $opts[$id] = @{ id = $id; kind = $pair[1]; name = "[$($script:KindLabels[$pair[1]])] $(Get-Name $id)" } }
            }
        }
        if (-not $opts.Count) { Show-Msg "De geselecteerde policies hebben geen uitsluitingen." "Uitsluiten" | Out-Null; return }
        $res = Show-InputDialog -Title 'Uitsluiting weghalen' -Fields @(@{ Key = 'o'; Label = 'Welke uitsluiting weghalen?'; Options = @($opts.Values | ForEach-Object { $_.name } | Sort-Object) })
        if (-not $res) { return }
        $o = @($opts.Values | Where-Object { $_.name -eq $res.o })
        Update-PolicyExclusion -Mode Remove -Objects $o
    })
    $script:BtnExBg.Add_Click({
        $bg = @(Get-BreakGlass)
        if (-not $bg.Count) { Show-Msg "Stel eerst noodaccounts in op tab Back-up / import." "Noodaccounts" | Out-Null; return }
        if (-not @(Get-SelectedPolicies).Count) { $script:PolGrid.SelectAll() }
        Update-PolicyExclusion -Mode Add -Objects @($bg | ForEach-Object { @{ id = $_.id; kind = 'user'; name = $_.upn } })
    })
    $script:BtnPolDel.Add_Click({ Remove-Policies })

    # --- Toewijzingen ---
    $script:TxtAsgFilter.Add_TextChanged({ if ($script:Connected) { Update-AssignmentView } })
    $script:CmbAsgKind.Add_SelectionChanged({ if ($script:Connected) { Update-AssignmentView } })
    $script:CmbAsgDir.Add_SelectionChanged({ if ($script:Connected) { Update-AssignmentView } })
    $script:AsgObjGrid.Add_SelectionChanged({ Update-AssignmentDetail })
    $script:AsgGrid.Add_MouseDoubleClick({
        $row = $script:AsgGrid.SelectedItem
        if (-not $row) { return }
        $p = $script:Policies | Where-Object { $_.id -eq $row.Row['PolicyId'] } | Select-Object -First 1
        if ($p) { Invoke-PolicyEdit $p }
    })
    $script:BtnAsgAdd.Add_Click({ Invoke-AsgAdd })
    $script:BtnAsgCopy.Add_Click({ Invoke-AsgCopy })
    $script:BtnAsgFlip.Add_Click({ Invoke-AsgFlip })
    $script:BtnAsgReplace.Add_Click({ Invoke-AsgReplace })
    $script:BtnAsgDel.Add_Click({ Invoke-AsgRemove })
    $script:BtnAsgCsv.Add_Click({ Export-Assignments })

    # --- Locaties ---
    $script:BtnLocIp.Add_Click({ Edit-NamedLocation -Type ip })
    $script:BtnLocCountry.Add_Click({ Edit-NamedLocation -Type country })
    $script:BtnLocEdit.Add_Click({ $l = @(Get-GridSelection $script:LocGrid $script:Locations); if ($l.Count -eq 1) { Edit-NamedLocation $l[0] } })
    $script:LocGrid.Add_MouseDoubleClick({ $l = @(Get-GridSelection $script:LocGrid $script:Locations); if ($l.Count -eq 1) { Edit-NamedLocation $l[0] } })
    $script:BtnLocDel.Add_Click({ Remove-NamedLocations })

    # --- Sterktes ---
    $script:BtnStrNew.Add_Click({ Edit-AuthStrength $null })
    $script:BtnStrEdit.Add_Click({ $s = @(Get-GridSelection $script:StrGrid $script:Strengths); if ($s.Count -eq 1) { Edit-AuthStrength $s[0] } })
    $script:StrGrid.Add_MouseDoubleClick({ $s = @(Get-GridSelection $script:StrGrid $script:Strengths); if ($s.Count -eq 1) { Edit-AuthStrength $s[0] } })
    $script:BtnStrDup.Add_Click({
        $s = @(Get-GridSelection $script:StrGrid $script:Strengths)
        if ($s.Count -ne 1) { return }
        Edit-AuthStrength @{ displayName = "$($s[0].displayName) (kopie)"; description = $s[0].description; allowedCombinations = @($s[0].allowedCombinations) }
    })
    $script:BtnStrDel.Add_Click({ Remove-AuthStrengths })

    # --- Contexten ---
    $script:BtnCtxNew.Add_Click({
        if (@($script:Contexts).Count -ge 99) { Show-Msg "Alle 99 contexten zijn in gebruik." "Context" | Out-Null; return }
        Edit-AuthContext $null
    })
    $script:BtnCtxEdit.Add_Click({ $x = @(Get-GridSelection $script:CtxGrid $script:Contexts); if ($x.Count -eq 1) { Edit-AuthContext $x[0] } })
    $script:CtxGrid.Add_MouseDoubleClick({ $x = @(Get-GridSelection $script:CtxGrid $script:Contexts); if ($x.Count -eq 1) { Edit-AuthContext $x[0] } })
    $script:BtnCtxDel.Add_Click({ Remove-AuthContexts })

    # --- Gebruiker-impact ---
    $script:BtnImpactPick.Add_Click({ Invoke-ImpactPick })
    $script:ChkImpactAll.Add_Click({ Show-Impact })
    $script:ImpactGrid.Add_MouseDoubleClick({
        $row = $script:ImpactGrid.SelectedItem
        if ($row) { $p = $script:Policies | Where-Object { $_.id -eq $row.Row['Id'] } | Select-Object -First 1; if ($p) { Show-TextWindow $p.displayName (Get-PolicySummary $p) } }
    })
    $script:BtnSignLoad.Add_Click({ Invoke-LoadSignIns })
    $script:SignGrid.Add_SelectionChanged({ Show-SignInDetail })

    # --- Back-up / import ---
    $script:BtnExport.Add_Click({ Export-CaConfig })
    $script:BtnBackupDir.Add_Click({ Start-Process explorer.exe "`"$(Get-BackupDir)`"" })
    $script:BtnImpRead.Add_Click({ Read-ImportFile })
    $script:BtnImpRun.Add_Click({ Invoke-Import })
    $script:BtnBgAdd.Add_Click({ Add-BreakGlass })
    $script:BtnBgDel.Add_Click({ Remove-BreakGlass })

    # --- Algemeen ---
    $script:BtnStop.Add_Click({
        $script:StopRequested = $true
        $script:BtnStop.IsEnabled = $false
        Set-Status "Stoppen na de huidige stap..."
    })
    $script:BtnClearLog.Add_Click({ $script:LogPara.Inlines.Clear() })
    $script:Window.Add_Closing({
        if ($script:LongOp) {
            $script:StopRequested = $true
            $_.Cancel = $true
            Set-Status "Bezig - wordt gestopt, sluit daarna opnieuw"
        }
    })
    $script:Window.Add_Closed({ Disconnect-All })
}

function Show-MainWindow {
    New-MainWindow
    Write-Log "Log in met je admin-account. Het inlogscherm vraagt altijd welk account je wilt gebruiken." "info"
    if (-not $script:MsalAvailable) { Write-Log "MSAL niet gevonden in de Graph-module - terugval op standaard login" "warning" }
    Write-Log "Automatische back-ups van gewijzigde/verwijderde objecten: $($script:AppDir)\Backups" "info"
    [void]$script:Window.ShowDialog()
}

#endregion

if (-not (Initialize-Modules)) {
    [System.Windows.MessageBox]::Show("Module Microsoft.Graph.Authentication niet gevonden.`n`nInstalleer met:`nInstall-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser", "Module ontbreekt", "OK", "Error") | Out-Null
    exit 1
}
Show-MainWindow
