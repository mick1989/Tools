#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Beheer (WPF) - overzicht van alle sites, OneDrives, bibliotheken
    en machtigingen, en bulk-wijzigingen, met een apart admin-account.
.DESCRIPTION
    Tabbladen:
      Sites              alle sites: type, M365-groep, gebruikt/quotum/%, bestanden,
                         actieve bestanden en paginaweergaven (30 dagen), laatste
                         activiteit, eigenaar, extern delen, vergrendeling.
                         'Details scannen' voegt toe: aantal bibliotheken, lijsten met
                         unieke rechten, ruimte in versies en in de prullenbak.
                         Onderin per site:
                           Machtigingen      site + alle bibliotheken/lijsten met
                                             eigen rechten (kolom Bereik), met bij
                                             elke groep statisch of dynamisch.
                           Bibliotheken      grootte, versies, versiebeheer, unieke
                                             rechten; rechten per bibliotheek inzien,
                                             overname verbreken/herstellen, toegang
                                             geven/intrekken, versielimiet, versies
                                             opschonen.
                         Bulk: toegang geven/intrekken, extern delen, quotum,
                         vergrendeling, versies opschonen, prullenbak terugzetten
                         (filter op wie/wanneer/type) of legen, machtigingenrapport,
                         groepen scannen.
      OneDrive           alle OneDrives: eigenaar, status account (actief/uit/
                         verwijderd), gebruikt/quotum/%, bestanden, activiteit,
                         prullenbak, versies. Bulk: beheerder toevoegen/verwijderen
                         (bijv. manager bij uitdienst), quotum, vergrendeling,
                         versies opschonen, prullenbak terugzetten of legen,
                         OneDrive leegmaken.
      Gebruiker-toegang  waar heeft iemand toegang (site, bibliotheek, via groep,
                         via 'Iedereen'), optioneel ook in OneDrives.
      Groepen            Entra-groepen in sites, statisch/dynamisch, groep-editor.

    Inloggen in 2 stappen, beide met account-kiezer (apart admin-account):
      1. Microsoft Graph (systeembrowser)
      2. SharePoint (eigen inlogvenster, geen browser-SSO)
    Tokens blijven alleen in dit proces.

    Geen toegang tot een site of OneDrive? De tool kan je tijdelijk
    sitebeheerder maken; die sites worden onthouden zodat je het kunt intrekken.

    Wijzigingen worden gelogd in
      %LOCALAPPDATA%\SharePointBeheerTool\Logs\<tenant>\wijzigingen-<jjjj-mm>.csv
.NOTES
    Vereist: Windows PowerShell 5.1 en een module die MSAL meelevert
      (Microsoft.Graph.Authentication, Az.Accounts of ExchangeOnlineManagement).
    Rollen: SharePoint Administrator + Groups/User Administrator (groepen) +
            Reports Reader of Global Reader (gebruiksrapporten).
    Rapporten met verborgen namen: M365-beheercentrum > Instellingen >
      Organisatie-instellingen > Rapporten > 'Verborgen namen tonen'.
#>

# Windows PowerShell 5.1 in STA nodig (WPF + ingebouwd inlogvenster van MSAL)
if ($PSVersionTable.PSEdition -eq 'Core' -or [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $self = $MyInvocation.MyCommand.Path
    if ($self) {
        Start-Process powershell.exe -ArgumentList "-STA -NoProfile -ExecutionPolicy Bypass -File `"$self`""
    }
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

#region --- Globale state ---

$script:GraphRoot     = 'https://graph.microsoft.com/v1.0'
$script:GraphClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools
$script:GraphScopes   = @(
    'https://graph.microsoft.com/User.Read.All', 'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/GroupMember.ReadWrite.All', 'https://graph.microsoft.com/Directory.Read.All',
    'https://graph.microsoft.com/Sites.Read.All', 'https://graph.microsoft.com/Reports.Read.All'
)
$script:SpClientIdDefault = '9bc3ab49-b65d-410a-85ad-de819febfddc'   # SharePoint Online Management Shell
$script:SpRedirectDefault = 'https://oauth.spops.microsoft.com/'
$script:TenantTypeId      = '{268004ae-ef6b-4e9b-8425-127220d84719}'  # CSOM Microsoft.Online.SharePoint.TenantAdministration.Tenant
$script:GroupSelect = 'id,displayName,description,groupTypes,membershipRule,membershipRuleProcessingState,securityEnabled,mailEnabled,mail,onPremisesSyncEnabled'

$script:Window        = $null
$script:Connected     = $false
$script:SpOk          = $false
$script:HasAdmin      = $false
$script:GraphApp      = $null
$script:GraphAuth     = $null   # @{ Token; Expires; Account }
$script:SpApp         = $null
$script:SpAccount     = $null
$script:SpTokens      = @{}     # resource -> @{ Token; Expires }
$script:TenantId      = $null
$script:MeUpn         = $null
$script:MeId          = $null
$script:RootUrl       = $null
$script:AdminUrl      = $null
$script:MyHost        = $null
$script:LongOp        = $false
$script:StopRequested = $false
$script:AutoElevate   = $false

$script:Sites       = New-Object System.Collections.ArrayList
$script:OneDrives   = New-Object System.Collections.ArrayList
$script:AggOneDrives = @()
$script:OdLoaded    = $false
$script:GroupCache  = @{}       # groep-id -> Graph-object (of @{ Missing = $true })
$script:Ctx         = @{}       # 'S' (sites) en 'O' (OneDrive): detailpaneel-state
$script:AccUser     = $null
$script:AccRows     = @()
$script:GroupRows   = @{}       # groep-id -> HashSet van site-titels
$script:Elevated    = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

$script:AppDir   = Join-Path $env:LOCALAPPDATA 'SharePointBeheerTool'
$script:Settings = $null

$script:KindLabels = @{
    user = 'Gebruiker'; guest = 'Gast'; secgroup = 'Entra-groep'; m365members = 'M365-groep (leden)'
    m365owners = 'M365-groep (eigenaars)'; everyone = 'Iedereen'; everyoneext = 'Iedereen behalve externe gebruikers'
    other = 'Overig'
}
$script:SharingLabels = @{
    0 = 'Alleen intern'; 3 = 'Bestaande gasten'; 1 = 'Nieuwe en bestaande gasten'; 2 = 'Iedereen (anonieme links)'
}
$script:LockLabels = [ordered]@{ Unlock = 'Niet vergrendeld'; ReadOnly = 'Alleen-lezen'; NoAccess = 'Geen toegang' }
$script:RoleLabels = [ordered]@{
    Owner   = 'Eigenaar (Volledig beheer)'
    Member  = 'Lid (Bewerken)'
    Visitor = 'Bezoeker (Lezen)'
    Admin   = 'Sitecollectiebeheerder'
}
# RoleTypeKind voor rechten op een bibliotheek
$script:LibRoles = @(
    [pscustomobject]@{ Label = 'Volledig beheer'; Value = 5 }
    [pscustomobject]@{ Label = 'Bewerken'; Value = 6 }
    [pscustomobject]@{ Label = 'Bijdragen'; Value = 3 }
    [pscustomobject]@{ Label = 'Lezen'; Value = 2 }
)
$script:RuleExamples = @(
    '(user.department -eq "Verkoop")'
    '(user.companyName -eq "Bedrijfsnaam")'
    '(user.city -eq "Amsterdam")'
    '(user.jobTitle -contains "Manager")'
    '(user.accountEnabled -eq true) -and (user.userType -eq "Member")'
    '(user.userType -eq "Guest")'
    '(user.mail -endsWith "@domein.nl")'
    '(user.extensionAttribute1 -eq "waarde")'
)

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

function Wait-Seconds {
    param([double]$Seconds)
    $until = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $until) { Update-UI; Start-Sleep -Milliseconds 150 }
}

function Show-Msg {
    param([string]$Text, [string]$Title = 'Info', [string]$Button = 'OK', [string]$Icon = 'Information', $Owner = $null)
    $o = if ($Owner) { $Owner } elseif ($script:Window -and $script:Window.IsLoaded) { $script:Window } else { $null }
    if ($o) { return [System.Windows.MessageBox]::Show($o, $Text, $Title, $Button, $Icon) }
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
    if ($null -eq $script:LblStatus) { return }
    $script:LblStatus.Text = $Text
    $script:StatusBar.Background = switch ($Type) {
        'success' { New-Brush '#107C10' }
        'error'   { New-Brush '#C42B1C' }
        default   { New-Brush '#0078D4' }
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
    $script:MainArea.IsEnabled = $script:Connected
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
    # Kolommen als 'Naam' (tekst) of 'Naam:int' / 'Naam:double' (sorteert numeriek)
    param([string[]]$Columns)
    $t = New-Object System.Data.DataTable
    foreach ($c in $Columns) {
        $parts = $c.Split(':')
        $type = switch ($parts[1]) { 'int' { [int] } 'double' { [double] } default { [string] } }
        [void]$t.Columns.Add($parts[0], $type)
    }
    return ,$t
}

function Set-GridRows {
    param($Grid, [string[]]$Columns, $Rows)
    $t = New-Table $Columns
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $dr = $t.NewRow()
        foreach ($c in $Columns) {
            $n = $c.Split(':')[0]
            $v = $r[$n]
            if ($null -eq $v -or "$v" -eq '') { $dr[$n] = [DBNull]::Value } else { $dr[$n] = $v }
        }
        [void]$t.Rows.Add($dr)
    }
    $Grid.ItemsSource = $t.DefaultView
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

function Get-P {
    # Eerste gevulde eigenschap uit een lijst kandidaat-namen
    param($Object, [string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($n in $Names) {
        $p = $Object.PSObject.Properties[$n]
        if ($p -and $null -ne $p.Value -and "$($p.Value)" -ne '') { return $p.Value }
    }
    return $null
}

function ConvertTo-Num {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    try { return [double]$Value } catch { return $null }
}

function ConvertTo-GB {
    param($MB)
    if ($null -eq $MB) { return $null }
    return [Math]::Round([double]$MB / 1024, 2)
}

function ConvertTo-DateText {
    param($Value)
    if (-not $Value) { return '' }
    try { return ([datetime]$Value).ToString('yyyy-MM-dd') } catch { return [string]$Value }
}

function ConvertFrom-Xaml {
    param([string]$Text)
    [xml]$x = $Text
    return [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
}

function Get-XamlControls {
    # Alle x:Name-elementen uit de XAML als hashtable
    param($Root, [string]$Xaml)
    $c = @{}
    foreach ($m in [regex]::Matches($Xaml, 'x:Name="(\w+)"')) {
        $n = $m.Groups[1].Value
        $c[$n] = $Root.FindName($n)
    }
    return $c
}

function Export-GridCsv {
    # Exporteert de zichtbare kolommen en (gefilterde) rijen van een DataGrid
    param($Grid, [string]$DefaultName)
    $view = $Grid.ItemsSource
    if ($null -eq $view -or $view.Count -eq 0) { Show-Msg 'Er is niets om te exporteren.' 'Export' | Out-Null; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (puntkomma)|*.csv'
    $dlg.FileName = "$DefaultName-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
    if (-not $dlg.ShowDialog($script:Window)) { return }
    $cols = @(foreach ($col in $Grid.Columns) { @{ Header = [string]$col.Header; Path = $col.Binding.Path.Path } })
    $out = foreach ($rv in $view) {
        $o = [ordered]@{}
        foreach ($c in $cols) { $o[$c.Header] = [string]$rv.Row[$c.Path] }
        [pscustomobject]$o
    }
    $out | Export-Csv -Path $dlg.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Log "Geexporteerd: $($dlg.FileName)" 'success'
}

#endregion

#region --- Instellingen en wijzigingslog ---

function Get-Settings {
    if ($script:Settings) { return $script:Settings }
    $file = Join-Path $script:AppDir 'settings.json'
    $s = @{}
    if (Test-Path $file) {
        try {
            $j = Get-Content $file -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $j.PSObject.Properties) { $s[$p.Name] = $p.Value }
        } catch { }
    }
    foreach ($k in 'LastAdmin', 'LastTenant', 'SpClientId', 'SpRedirectUri') { if (-not $s.ContainsKey($k) -or $null -eq $s[$k]) { $s[$k] = '' } }
    # Elevated: tenant-id -> lijst van site-URL's waar de tool je tijdelijk beheerder maakte
    $el = @{}
    if ($s.Elevated) { foreach ($p in $s.Elevated.PSObject.Properties) { $el[$p.Name] = @($p.Value) } }
    $s.Elevated = $el
    $script:Settings = $s
    return $s
}

function Save-Settings {
    try {
        if (-not (Test-Path $script:AppDir)) { [void](New-Item -ItemType Directory -Path $script:AppDir -Force) }
        $file = Join-Path $script:AppDir 'settings.json'
        $script:Settings | ConvertTo-Json -Depth 6 | Set-Content -Path $file -Encoding UTF8
    }
    catch { Write-Log "Instellingen opslaan mislukt: $_" 'warning' }
}

function Save-Elevated {
    $s = Get-Settings
    $s.Elevated[$script:TenantId] = @($script:Elevated)
    Save-Settings
    Update-ElevatedButton
}

function Update-ElevatedButton {
    $n = $script:Elevated.Count
    foreach ($b in @($script:BtnElevRevoke, $script:BtnOdElevRevoke)) {
        if (-not $b) { continue }
        $b.Content = "Tijdelijke beheerrechten intrekken ($n)"
        $b.IsEnabled = ($n -gt 0)
    }
}

function Get-LogDir {
    $d = Join-Path $script:AppDir "Logs\$($script:TenantId)"
    if (-not (Test-Path $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    return $d
}

function Write-ChangeLog {
    param([string]$Site, [string]$Action, [string]$Principal, [string]$Result)
    try {
        $file = Join-Path (Get-LogDir) ("wijzigingen-{0}.csv" -f (Get-Date -Format 'yyyy-MM'))
        [pscustomobject][ordered]@{
            Tijd = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Door = $script:MeUpn; Site = $Site
            Actie = $Action; Wie = $Principal; Resultaat = $Result
        } | Export-Csv -Path $file -Append -NoTypeInformation -Delimiter ';' -Encoding UTF8
    } catch { }
}

#endregion

#region --- MSAL en inloggen ---

function Initialize-Msal {
    if ('Microsoft.Identity.Client.PublicClientApplicationBuilder' -as [type]) { return $true }
    $docs  = [Environment]::GetFolderPath('MyDocuments')
    $roots = @($env:PSModulePath -split ';') + @(
        (Join-Path $docs 'WindowsPowerShell\Modules'), (Join-Path $docs 'PowerShell\Modules'),
        (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'), (Join-Path $env:ProgramFiles 'PowerShell\Modules')
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
    $cands = foreach ($r in $roots) {
        foreach ($m in 'Microsoft.Graph.Authentication', 'Az.Accounts', 'ExchangeOnlineManagement', 'MicrosoftTeams') {
            $p = Join-Path $r $m
            if (Test-Path $p) { Get-ChildItem $p -Recurse -Filter 'Microsoft.Identity.Client.dll' -ErrorAction SilentlyContinue }
        }
    }
    # .NET Framework-build nodig (Desktop / net4x), niet de Core-build. Voorkeur: Graph.Authentication.
    $cands = @($cands | Where-Object { $_.FullName -notmatch '\\(Core|netcore\w*|net[5-9]\.\d+\w*|netstandard[\d.]*)\\' })
    $best = $cands | Sort-Object @{ Expression = { $_.FullName -match 'Microsoft\.Graph\.Authentication' }; Descending = $true },
                                 @{ Expression = { try { [version]$_.VersionInfo.FileVersion } catch { [version]'0.0' } }; Descending = $true } |
            Select-Object -First 1
    if (-not $best) { return $false }
    # Afhankelijkheden (Microsoft.IdentityModel.*) zoeken in de map en de bovenliggende map,
    # ongeacht versie (modules leunen anders op een binding redirect)
    if (-not ('SpTool.AsmResolver' -as [type])) {
        Add-Type -TypeDefinition @'
using System; using System.IO; using System.Reflection;
namespace SpTool {
  public static class AsmResolver {
    static string[] _dirs = new string[0];
    public static void Register(string[] dirs) {
      _dirs = dirs;
      AppDomain.CurrentDomain.AssemblyResolve += delegate(object s, ResolveEventArgs e) {
        string name = new AssemblyName(e.Name).Name;
        foreach (Assembly a in AppDomain.CurrentDomain.GetAssemblies()) {
          if (string.Equals(a.GetName().Name, name, StringComparison.OrdinalIgnoreCase)) return a;
        }
        foreach (string d in _dirs) {
          string p = Path.Combine(d, name + ".dll");
          if (File.Exists(p)) return Assembly.LoadFrom(p);
        }
        return null;
      };
    }
  }
}
'@
    }
    [SpTool.AsmResolver]::Register([string[]]@($best.DirectoryName, $best.Directory.Parent.FullName))
    try { Add-Type -Path $best.FullName -ErrorAction Stop } catch { }
    return [bool]('Microsoft.Identity.Client.PublicClientApplicationBuilder' -as [type])
}

function Wait-MsalTask {
    param($Task)
    while (-not $Task.IsCompleted) { Update-UI; Start-Sleep -Milliseconds 100 }
    if ($Task.IsCanceled) { throw 'Inloggen geannuleerd of verlopen' }
    if ($Task.IsFaulted) { throw $Task.Exception.GetBaseException().Message }
    return $Task.Result
}

function New-MsalApp {
    param([string]$ClientId, [string]$RedirectUri, [string]$Tenant)
    $authority = "https://login.microsoftonline.com/$(if ($Tenant) { $Tenant } else { 'organizations' })"
    return [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).
        WithAuthority($authority).WithRedirectUri($RedirectUri).Build()
}

function Get-GraphToken {
    if (-not $script:GraphAuth) { throw 'Niet ingelogd' }
    if ((Get-Date) -lt $script:GraphAuth.Expires.AddMinutes(-5)) { return $script:GraphAuth.Token }
    $r = Wait-MsalTask $script:GraphApp.AcquireTokenSilent([string[]]$script:GraphScopes, $script:GraphAuth.Account).ExecuteAsync()
    $script:GraphAuth = @{ Token = $r.AccessToken; Expires = $r.ExpiresOn.LocalDateTime; Account = $r.Account }
    return $r.AccessToken
}

function Get-SpToken {
    param([string]$Resource)
    if (-not $script:SpAccount) { throw 'Niet ingelogd bij SharePoint' }
    $key = $Resource.TrimEnd('/').ToLower()
    $t = $script:SpTokens[$key]
    if ($t -and (Get-Date) -lt $t.Expires.AddMinutes(-5)) { return $t.Token }
    $r = Wait-MsalTask $script:SpApp.AcquireTokenSilent([string[]]@("$key/.default"), $script:SpAccount).ExecuteAsync()
    $script:SpTokens[$key] = @{ Token = $r.AccessToken; Expires = $r.ExpiresOn.LocalDateTime }
    return $r.AccessToken
}

function Connect-All {
    param([string]$Tenant, [string]$LoginHint)
    Disconnect-All

    # 1. Graph - systeembrowser met account-kiezer
    Set-Status 'Inloggen bij Microsoft Graph (kies je admin-account in de browser)...'
    $script:GraphApp = New-MsalApp $script:GraphClientId 'http://localhost' $Tenant
    $b = $script:GraphApp.AcquireTokenInteractive([string[]]$script:GraphScopes).
        WithUseEmbeddedWebView($false).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)
    if ($LoginHint) { $b = $b.WithLoginHint($LoginHint) }
    $cts = New-Object System.Threading.CancellationTokenSource ([TimeSpan]::FromMinutes(5))
    $r = Wait-MsalTask $b.ExecuteAsync($cts.Token)
    $script:GraphAuth = @{ Token = $r.AccessToken; Expires = $r.ExpiresOn.LocalDateTime; Account = $r.Account }
    $script:TenantId = $r.TenantId
    $me = Invoke-Graph 'me?$select=id,displayName,userPrincipalName' -Quiet
    $script:MeUpn = [string]$me.userPrincipalName
    $script:MeId  = [string]$me.id
    Write-Log "Graph: ingelogd als $($script:MeUpn) (tenant $($script:TenantId))" 'success'

    $root = Invoke-Graph 'sites/root?$select=webUrl' -Quiet
    $script:RootUrl = ([uri]$root.webUrl).GetLeftPart([UriPartial]::Authority).ToLower()
    $h = ([uri]$script:RootUrl).Host
    $label = $h.Split('.')[0]
    $script:AdminUrl = 'https://' + $label + '-admin' + $h.Substring($label.Length)
    $script:MyHost   = 'https://' + $label + '-my' + $h.Substring($label.Length)
    Write-Log "SharePoint: $($script:RootUrl)  |  beheer: $($script:AdminUrl)  |  OneDrive: $($script:MyHost)" 'info'

    # 2. SharePoint - eigen inlogvenster (geen SSO met de browser)
    $s = Get-Settings
    $spClient = if ($s.SpClientId) { [string]$s.SpClientId } else { $script:SpClientIdDefault }
    $spRedir  = if ($s.SpRedirectUri) { [string]$s.SpRedirectUri } else { $script:SpRedirectDefault }
    try {
        Set-Status 'Inloggen bij SharePoint (kies hetzelfde admin-account)...'
        $script:SpApp = New-MsalApp $spClient $spRedir $script:TenantId
        $embedded = $spRedir -notmatch '^http://localhost'
        $b2 = $script:SpApp.AcquireTokenInteractive([string[]]@("$($script:RootUrl)/.default")).
            WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount).WithLoginHint($script:MeUpn).
            WithUseEmbeddedWebView($embedded)
        if ($embedded) {
            try { $b2 = $b2.WithParentActivityOrWindow((New-Object System.Windows.Interop.WindowInteropHelper $script:Window).Handle) } catch { }
        }
        $cts2 = New-Object System.Threading.CancellationTokenSource ([TimeSpan]::FromMinutes(5))
        $r2 = Wait-MsalTask $b2.ExecuteAsync($cts2.Token)
        $script:SpAccount = $r2.Account
        $script:SpTokens[$script:RootUrl] = @{ Token = $r2.AccessToken; Expires = $r2.ExpiresOn.LocalDateTime }
        if ($r2.Account.Username -ne $script:MeUpn) {
            Write-Log "LET OP: SharePoint-account ($($r2.Account.Username)) is anders dan het Graph-account ($($script:MeUpn))" 'warning'
        }
        $script:SpOk = $true
        Write-Log "SharePoint: ingelogd als $($r2.Account.Username)" 'success'
        try {
            [void](Invoke-SP "$($script:AdminUrl)/_api/web?`$select=Title" -Quiet)
            $script:HasAdmin = $true
            Write-Log 'SharePoint-beheercentrum bereikbaar (SharePoint Administrator)' 'success'
        }
        catch { Write-Log "Geen toegang tot het SharePoint-beheercentrum: $($_.Exception.Message). Sitelijst via Graph, geen quotum/delen/tijdelijk-beheerder." 'warning' }
    }
    catch {
        Write-Log "SharePoint-login mislukt: $($_.Exception.Message). Alleen Graph-functies beschikbaar." 'error'
    }

    $script:Connected = $true
    $script:Elevated.Clear()
    foreach ($u in @($s.Elevated[$script:TenantId])) { if ($u) { [void]$script:Elevated.Add([string]$u) } }
    Update-ElevatedButton
    if ($script:Elevated.Count) {
        Write-Log "Je bent nog tijdelijk sitebeheerder op $($script:Elevated.Count) site(s) uit een vorige sessie - intrekken kan via de knop op tab Sites of OneDrive." 'warning'
    }
}

function Disconnect-All {
    foreach ($app in @($script:GraphApp, $script:SpApp)) {
        if ($app) { try { foreach ($a in @((Wait-MsalTask $app.GetAccountsAsync()))) { [void](Wait-MsalTask $app.RemoveAsync($a)) } } catch { } }
    }
    $script:GraphApp = $null; $script:GraphAuth = $null
    $script:SpApp = $null; $script:SpAccount = $null; $script:SpTokens = @{}
    $script:Connected = $false; $script:SpOk = $false; $script:HasAdmin = $false
    $script:GroupCache = @{}
}

#endregion

#region --- HTTP helpers (Graph, SharePoint REST, CSOM) ---

function Invoke-Api {
    param([string]$Uri, [string]$Method = 'GET', $Body = $null, [hashtable]$Headers, [string]$ContentType = 'application/json; charset=utf-8',
          [switch]$Quiet, [switch]$Raw)
    $short = $Uri -replace '^https://[^/]+', ''
    if ($short.Length -gt 150) { $short = $short.Substring(0, 150) + '...' }
    if (-not $Quiet) { Write-Log "$Method $short" 'api' }
    for ($attempt = 1; ; $attempt++) {
        $p = @{ Uri = $Uri; Method = $Method; Headers = $Headers; UseBasicParsing = $true; ErrorAction = 'Stop' }
        if ($null -ne $Body) {
            $txt = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 -Compress }
            $p.Body = [Text.Encoding]::UTF8.GetBytes($txt)
            $p.ContentType = $ContentType
        }
        elseif ($Method -in 'POST', 'PATCH', 'MERGE') {
            $p.Body = [byte[]]@(); $p.ContentType = $ContentType
        }
        try {
            if ($Raw) {
                $c = (Invoke-WebRequest @p).Content
                if ($c -is [byte[]]) { $c = [Text.Encoding]::UTF8.GetString($c) }
                return $c
            }
            return Invoke-RestMethod @p
        }
        catch {
            $code = 0; $retryAfter = 0
            $resp = $_.Exception.Response
            if ($resp) {
                try { $code = [int]$resp.StatusCode } catch { }
                try { $ra = $resp.Headers['Retry-After']; if ($ra) { $retryAfter = [int]$ra } } catch { }
            }
            if ($code -in 429, 503, 504 -and $attempt -lt 5) {
                $wait = if ($retryAfter -gt 0) { [Math]::Min($retryAfter, 60) } else { 5 * $attempt }
                Write-Log "Throttling ($code) - $wait s wachten..." 'warning'
                Wait-Seconds $wait
                continue
            }
            $msg = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                try {
                    $d = $_.ErrorDetails.Message | ConvertFrom-Json
                    if ($d.'odata.error'.message.value) { $msg = [string]$d.'odata.error'.message.value }
                    elseif ($d.error.message.value) { $msg = [string]$d.error.message.value }
                    elseif ($d.error.message) { $msg = [string]$d.error.message }
                }
                catch { $msg = "$msg - $($_.ErrorDetails.Message)" }
            }
            $ex = New-Object System.Exception ("[$code] $msg")
            $ex.Data['Status'] = $code
            throw $ex
        }
    }
}

function Test-AccessDenied {
    param($ErrorRecord)
    $e = $ErrorRecord.Exception
    $code = 0
    if ($e.Data.Contains('Status')) { $code = [int]$e.Data['Status'] }
    return ($code -in 401, 403 -or $e.Message -match 'Access denied|Toegang geweigerd|UnauthorizedAccess|do not have permission')
}

function Invoke-Graph {
    param([string]$Uri, [string]$Method = 'GET', $Body = $null, [hashtable]$Headers, [switch]$Quiet, [switch]$Raw)
    if ($Uri -notmatch '^https://') { $Uri = "$($script:GraphRoot)/$($Uri.TrimStart('/'))" }
    $h = @{ Authorization = "Bearer $(Get-GraphToken)" }
    if ($Headers) { foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] } }
    return Invoke-Api -Uri $Uri -Method $Method -Body $Body -Headers $h -Quiet:$Quiet -Raw:$Raw
}

function Get-GraphPaged {
    param([string]$Uri, [hashtable]$Headers, [int]$Max = 0, [switch]$Quiet)
    $all = New-Object System.Collections.ArrayList
    do {
        $r = Invoke-Graph -Uri $Uri -Headers $Headers -Quiet:$Quiet
        foreach ($v in @($r.value)) { [void]$all.Add($v) }
        $Uri = $r.'@odata.nextLink'
        if ($Max -and $all.Count -ge $Max) { break }
    } while ($Uri -and -not $script:StopRequested)
    return $all.ToArray()
}

function Invoke-GraphBatch {
    # GET-requests via /$batch (max 20 per aanroep). Requests: @{ id; url }. Geeft id -> response.
    param([System.Collections.IList]$Requests)
    $results = @{}
    $total = $Requests.Count
    for ($i = 0; $i -lt $total; $i += 20) {
        $todo = @($Requests[$i..([Math]::Min($i + 19, $total - 1))])
        for ($attempt = 1; $attempt -le 4 -and $todo.Count -gt 0; $attempt++) {
            $payload = @{ requests = @(foreach ($r in $todo) { [ordered]@{ id = [string]$r.id; method = 'GET'; url = $r.url } }) }
            $resp = Invoke-Graph "`$batch" -Method POST -Body $payload -Quiet
            $retry = @()
            foreach ($x in @($resp.responses)) {
                if ([int]$x.status -eq 429 -and $attempt -lt 4) { $retry += @($todo | Where-Object { [string]$_.id -eq [string]$x.id }) }
                else { $results[[string]$x.id] = $x }
            }
            $todo = $retry
            if ($todo.Count -gt 0) { Wait-Seconds 3 }
        }
    }
    return $results
}

function Get-UsageReport {
    # Gebruiksrapport (CSV) als objecten; $null als het niet lukt
    param([string]$Name)
    try {
        $txt = Invoke-Graph "reports/$Name(period='D30')" -Raw -Quiet
        $txt = ([string]$txt).TrimStart([char]0xFEFF)
        return @($txt -split "`r?`n" | Where-Object { $_.Trim() } | ConvertFrom-Csv)
    }
    catch {
        Write-Log "Rapport $Name niet beschikbaar: $($_.Exception.Message)" 'warning'
        return $null
    }
}

function Invoke-SP {
    param([string]$Url, [string]$Method = 'GET', $Body = $null, [switch]$VerboseOData, [hashtable]$Headers, [switch]$Quiet, [switch]$Raw)
    $res = ([uri]$Url).GetLeftPart([UriPartial]::Authority)
    $fmt = if ($VerboseOData) { 'application/json;odata=verbose' } else { 'application/json;odata=nometadata' }
    $h = @{ Authorization = "Bearer $(Get-SpToken $res)"; Accept = $fmt }
    if ($Headers) { foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] } }
    $r = Invoke-Api -Uri $Url -Method $Method -Body $Body -Headers $h -ContentType "$fmt;charset=utf-8" -Quiet:$Quiet -Raw:$Raw
    if ($VerboseOData -and $r -and $r.PSObject.Properties['d']) { return $r.d }
    return $r
}

function Get-SpItemsPage {
    # Lijstitems ophalen (raw: PS 5.1 kan 'Id' en 'ID' in hetzelfde object niet aan)
    param([string]$Url)
    $txt = Invoke-SP $Url -Raw -Quiet
    $txt = $txt -creplace '"ID":', '"ID_dup":'
    return ($txt | ConvertFrom-Json)
}

function Invoke-Csom {
    # CSOM-aanroep op het beheercentrum (voor wat de REST-API niet biedt)
    param([string]$Actions, [string]$ObjectPaths)
    $xml = '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="SharePointBeheerTool">' +
           "<Actions>$Actions</Actions><ObjectPaths>$ObjectPaths</ObjectPaths></Request>"
    $h = @{ Authorization = "Bearer $(Get-SpToken $script:AdminUrl)"; Accept = 'application/json' }
    $r = Invoke-Api -Uri "$($script:AdminUrl)/_vti_bin/client.svc/ProcessQuery" -Method POST -Body $xml -Headers $h -ContentType 'text/xml; charset=utf-8' -Quiet
    if ($r -is [string]) { $r = $r | ConvertFrom-Json }
    $arr = @($r)
    if ($arr.Count -and $arr[0].ErrorInfo) { throw "CSOM: $($arr[0].ErrorInfo.ErrorMessage)" }
    return ,$arr
}

function ConvertTo-XmlText { param([string]$Text) return [System.Security.SecurityElement]::Escape($Text) }

function Get-CsomSitePath {
    param([string]$SiteUrl)
    return "<Constructor Id=`"1`" TypeId=`"$($script:TenantTypeId)`" />" +
           "<Method Id=`"3`" ParentId=`"1`" Name=`"GetSitePropertiesByUrl`"><Parameters><Parameter Type=`"String`">$(ConvertTo-XmlText $SiteUrl)</Parameter><Parameter Type=`"Boolean`">false</Parameter></Parameters></Method>"
}

function Set-TenantSiteAdmin {
    param([string]$SiteUrl, [string]$Login, [bool]$IsAdmin)
    $a = '<ObjectPath Id="2" ObjectPathId="1" /><Method Name="SetSiteAdmin" Id="3" ObjectPathId="1"><Parameters>' +
         "<Parameter Type=`"String`">$(ConvertTo-XmlText $SiteUrl)</Parameter>" +
         "<Parameter Type=`"String`">$(ConvertTo-XmlText $Login)</Parameter>" +
         "<Parameter Type=`"Boolean`">$($IsAdmin.ToString().ToLower())</Parameter></Parameters></Method>"
    $o = "<Constructor Id=`"1`" TypeId=`"$($script:TenantTypeId)`" />"
    [void](Invoke-Csom $a $o)
}

function Get-TenantSiteProps {
    param([string]$SiteUrl)
    $props = 'SharingCapability', 'LockState', 'StorageMaximumLevel', 'StorageUsage', 'Owner', 'Title'
    $a = '<ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties>' +
         (($props | ForEach-Object { "<Property Name=`"$_`" ScalarProperty=`"true`" />" }) -join '') + '</Properties></Query></Query>'
    $arr = Invoke-Csom $a (Get-CsomSitePath $SiteUrl)
    return ($arr | Where-Object { $_ -and $_.PSObject -and $_.PSObject.Properties['SharingCapability'] } | Select-Object -First 1)
}

function Set-TenantSiteProps {
    # Props: naam -> @(CSOM-type, waarde), bijv. @{ StorageMaximumLevel = @('Int64', 10240) }
    param([string]$SiteUrl, [System.Collections.IDictionary]$Props)
    $a = '<ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" />'
    $id = 10
    foreach ($k in $Props.Keys) {
        $t = $Props[$k][0]; $v = [string]$Props[$k][1]
        $a += "<SetProperty Id=`"$id`" ObjectPathId=`"3`" Name=`"$k`"><Parameter Type=`"$t`">$(ConvertTo-XmlText $v)</Parameter></SetProperty>"
        $id++
    }
    $a += "<ObjectPath Id=`"$id`" ObjectPathId=`"6`" />"
    $o = (Get-CsomSitePath $SiteUrl) + '<Method Id="6" ParentId="3" Name="Update" />'
    [void](Invoke-Csom $a $o)
}

#endregion

#region --- Groepen (Graph) ---

function Resolve-Groups {
    param([string[]]$Ids, [switch]$Force)
    $todo = @($Ids | Where-Object { $_ } | Select-Object -Unique | Where-Object { $Force -or -not $script:GroupCache.ContainsKey($_) })
    if (-not $todo.Count) { return }
    $req = New-Object System.Collections.ArrayList
    foreach ($id in $todo) { [void]$req.Add(@{ id = $id; url = "/groups/$id`?`$select=$($script:GroupSelect)" }) }
    $res = Invoke-GraphBatch $req
    foreach ($id in $todo) {
        $x = $res[$id]
        if ($x -and [int]$x.status -eq 200) { $script:GroupCache[$id] = $x.body }
        else { $script:GroupCache[$id] = [pscustomobject]@{ id = $id; Missing = $true; displayName = '(niet gevonden)' } }
    }
}

function Test-GroupDynamic { param($G) return ($G -and -not $G.Missing -and @($G.groupTypes) -contains 'DynamicMembership') }

function Get-GroupKindLabel {
    param($G)
    if (-not $G -or $G.Missing) { return '?' }
    $k = if (@($G.groupTypes) -contains 'Unified') { 'Microsoft 365' }
         elseif ($G.mailEnabled -and $G.securityEnabled) { 'E-mail beveiliging' }
         elseif ($G.securityEnabled) { 'Beveiliging' }
         else { 'Distributie' }
    if ($G.onPremisesSyncEnabled -eq $true) { $k += ' (on-prem)' }
    return $k
}

function Get-GroupMembershipLabel {
    param($G)
    if (-not $G) { return '' }
    if ($G.Missing) { return 'Niet gevonden' }
    if (Test-GroupDynamic $G) {
        if ($G.membershipRuleProcessingState -eq 'Paused') { return 'Dynamisch (gepauzeerd)' }
        return 'Dynamisch'
    }
    return 'Statisch'
}

function Get-GroupLabelById {
    param([string]$Id)
    if (-not $Id) { return '' }
    return (Get-GroupMembershipLabel $script:GroupCache[$Id])
}

#endregion

#region --- Sites en OneDrives laden ---

function Get-SiteTypeLabel {
    param([string]$Template, [bool]$HasGroup, [bool]$Teams)
    if ($Template -match '^TEAMCHANNEL') { return 'Teams-kanaal' }
    if ($Teams) { return 'Teams' }
    if ($Template -match '^GROUP#' -or $HasGroup) { return 'M365-groep' }
    switch -regex ($Template) {
        '^SITEPAGEPUBLISHING' { return 'Communicatie' }
        '^STS#(3|0)'          { return 'Teamsite (zonder groep)' }
        '^APPCATALOG'         { return 'App-catalogus' }
        '^SRCHCEN'            { return 'Zoekcentrum' }
        '^EHS#1'              { return 'Teamsite (zonder groep)' }
        '^$'                  { return 'Onbekend' }
    }
    return $Template
}

function New-SiteObject {
    param([hashtable]$Values)
    $o = @{
        Url = ''; Title = ''; Template = ''; GroupId = ''; Teams = $false; Type = 'Onbekend'; IsOneDrive = $false
        UsedMB = $null; QuotaMB = $null; Files = $null; ActiveFiles = $null; PageViews = $null
        LastActivity = ''; Created = ''; Owner = ''; OwnerUpn = ''; Status = ''
        Sharing = ''; Lock = ''; Libraries = $null; Unique = $null; VersionsMB = $null; RecycleMB = $null
        Meta = $null; Lists = $null; LibInfo = $null
    }
    foreach ($k in $Values.Keys) { $o[$k] = $Values[$k] }
    return $o
}

function ConvertFrom-AggItem {
    param($It)
    $url = [string](Get-P $It 'SiteUrl', 'Url')
    if (-not $url) { return $null }
    $tpl = [string](Get-P $It 'TemplateName', 'Template')
    if ($tpl -match '^(REDIRECTSITE|SPSMSITEHOST|POINTPUBLISHINGPERSONAL)') { return $null }
    if (Get-P $It 'TimeDeleted') { return $null }
    $gid = [string](Get-P $it 'GroupId')
    if ($gid -match '^0{8}-0{4}') { $gid = '' }
    $teams = "$(Get-P $It 'IsTeamsConnected')" -eq 'True'
    $used = ConvertTo-Num (Get-P $It 'StorageUsed')
    $quota = ConvertTo-Num (Get-P $It 'StorageQuota')
    return (New-SiteObject @{
        Url = $url.TrimEnd('/'); Title = [string](Get-P $It 'Title'); Template = $tpl; GroupId = $gid.ToLower(); Teams = $teams
        Type = (Get-SiteTypeLabel $tpl ([bool]$gid) $teams); IsOneDrive = ($tpl -match '^SPSPERS' -or $url -match '-my\.sharepoint\.')
        UsedMB = $(if ($null -ne $used) { $used / 1MB } else { $null }); QuotaMB = $(if ($quota) { $quota / 1MB } else { $null })
        Files = (ConvertTo-Num (Get-P $It 'NumOfFiles')); PageViews = (ConvertTo-Num (Get-P $It 'PageViews'))
        LastActivity = (ConvertTo-DateText (Get-P $It 'LastActivityOn')); Created = (ConvertTo-DateText (Get-P $It 'TimeCreated'))
        Owner = [string](Get-P $It 'SiteOwnerEmail', 'SiteOwnerName', 'CreatedBy')
        OwnerUpn = [string](Get-P $It 'SiteOwnerEmail')
    })
}

function Import-Sites {
    if (-not $script:Connected) { return }
    Start-LongOp 'Sites laden...'
    try {
        $list = New-Object System.Collections.ArrayList
        $ods = New-Object System.Collections.ArrayList
        $ok = $false
        if ($script:HasAdmin) {
            try {
                $url = "$($script:AdminUrl)/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?`$top=500"
                while ($url -and -not $script:StopRequested) {
                    $j = Get-SpItemsPage $url
                    foreach ($it in @($j.value)) {
                        $s = ConvertFrom-AggItem $it
                        if (-not $s) { continue }
                        if ($s.IsOneDrive) { [void]$ods.Add($s) } else { [void]$list.Add($s) }
                    }
                    $url = $j.'odata.nextLink'
                    Set-Status "$($list.Count) sites geladen..."
                }
                $ok = $true
            }
            catch { Write-Log "Sitelijst uit beheercentrum niet beschikbaar ($($_.Exception.Message)) - terugval op Graph" 'warning' }
        }
        $script:AggOneDrives = $ods.ToArray()
        if (-not $ok) {
            $raw = @(Get-GraphPaged 'sites?search=*&$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime&$top=500')
            foreach ($x in $raw) {
                if ([string]$x.webUrl -match '-my\.sharepoint\.') { continue }
                [void]$list.Add((New-SiteObject @{
                    Url = ([string]$x.webUrl).TrimEnd('/'); Title = [string]$x.displayName
                    LastActivity = (ConvertTo-DateText $x.lastModifiedDateTime); Created = (ConvertTo-DateText $x.createdDateTime)
                }))
            }
        }

        # Aanvullen met het gebruiksrapport (quotum, actieve bestanden, paginaweergaven)
        Set-Status 'Gebruiksrapport SharePoint ophalen...'
        $rep = Get-UsageReport 'getSharePointSiteUsageDetail'
        if ($rep) {
            $byUrl = @{}
            foreach ($s in $list) { $byUrl[$s.Url.ToLower()] = $s }
            $hidden = 0
            foreach ($r in $rep) {
                if ("$($r.'Is Deleted')" -eq 'True') { continue }
                $u = ([string]$r.'Site URL').TrimEnd('/').ToLower()
                if (-not $u) { $hidden++; continue }
                $s = $byUrl[$u]
                if (-not $s) { continue }
                $q = ConvertTo-Num $r.'Storage Allocated (Byte)'
                if ($q) { $s.QuotaMB = $q / 1MB }
                $s.ActiveFiles = ConvertTo-Num $r.'Active File Count'
                if ($null -eq $s.PageViews) { $s.PageViews = ConvertTo-Num $r.'Page View Count' }
                if ($null -eq $s.UsedMB) { $v = ConvertTo-Num $r.'Storage Used (Byte)'; if ($null -ne $v) { $s.UsedMB = $v / 1MB } }
                if ($null -eq $s.Files) { $s.Files = ConvertTo-Num $r.'File Count' }
                if (-not $s.LastActivity) { $s.LastActivity = ConvertTo-DateText $r.'Last Activity Date' }
                if (-not $s.Owner) { $s.Owner = [string]$r.'Owner Principal Name' }
                if ($s.Type -eq 'Onbekend' -and $r.'Root Web Template') { $s.Type = [string]$r.'Root Web Template' }
            }
            if ($hidden) { Write-Log 'Het gebruiksrapport heeft verborgen site-URL''s: zet in het M365-beheercentrum (Org-instellingen > Rapporten) de verborgen namen uit voor quotum/activiteit.' 'warning' }
        }

        $script:Sites = New-Object System.Collections.ArrayList
        foreach ($s in @($list | Sort-Object { $_.Title })) { [void]$script:Sites.Add($s) }
        $types = @('Alle typen') + @($script:Sites | ForEach-Object { $_.Type } | Sort-Object -Unique)
        $script:CmbSiteType.ItemsSource = $types
        $script:CmbSiteType.SelectedIndex = 0
        Update-SiteGrid
        Write-Log "$($script:Sites.Count) sites geladen" 'success'
        Set-Status "$($script:Sites.Count) sites geladen" 'success'
    }
    catch {
        Write-Log "Sites laden mislukt: $($_.Exception.Message)" 'error'
        Set-Status 'Sites laden mislukt' 'error'
    }
    finally { Stop-LongOp }
}

function Import-OneDrives {
    if (-not $script:Connected) { return }
    Start-LongOp 'OneDrives laden...'
    try {
        # Gebruikers voor de account-status (actief / uit / verwijderd)
        Set-Status 'Gebruikers ophalen...'
        $users = @{}
        foreach ($u in @(Get-GraphPaged 'users?$select=id,displayName,userPrincipalName,accountEnabled,assignedLicenses,userType&$top=999' -Quiet)) {
            $users[([string]$u.userPrincipalName).ToLower()] = $u
        }
        $map = @{}
        Set-Status 'Gebruiksrapport OneDrive ophalen...'
        $rep = Get-UsageReport 'getOneDriveUsageAccountDetail'
        $concealed = $false
        foreach ($r in @($rep)) {
            if (-not $r) { continue }
            $url = ([string]$r.'Site URL').TrimEnd('/')
            $upn = [string]$r.'Owner Principal Name'
            if (-not $url -or $upn -notmatch '@') { $concealed = $true; continue }
            $used = ConvertTo-Num $r.'Storage Used (Byte)'; $q = ConvertTo-Num $r.'Storage Allocated (Byte)'
            $map[$url.ToLower()] = New-SiteObject @{
                Url = $url; Title = [string]$r.'Owner Display Name'; Owner = [string]$r.'Owner Display Name'; OwnerUpn = $upn
                IsOneDrive = $true; Type = 'OneDrive'
                UsedMB = $(if ($null -ne $used) { $used / 1MB } else { $null }); QuotaMB = $(if ($q) { $q / 1MB } else { $null })
                Files = (ConvertTo-Num $r.'File Count'); ActiveFiles = (ConvertTo-Num $r.'Active File Count')
                LastActivity = (ConvertTo-DateText $r.'Last Activity Date')
                Status = $(if ("$($r.'Is Deleted')" -eq 'True') { 'OneDrive verwijderd' } else { '' })
            }
        }
        # Aanvullen met de beheercentrum-lijst (als die OneDrives bevat)
        foreach ($a in @($script:AggOneDrives)) {
            $k = $a.Url.ToLower()
            if ($map.ContainsKey($k)) {
                $m = $map[$k]
                if ($null -eq $m.UsedMB) { $m.UsedMB = $a.UsedMB }
                if (-not $m.Created) { $m.Created = $a.Created }
            }
            else {
                $a.Type = 'OneDrive'
                if (-not $a.OwnerUpn) { $a.OwnerUpn = $a.Owner }
                $map[$k] = $a
            }
        }
        # Laatste terugval: URL afleiden van gelicentieerde gebruikers
        if (-not $map.Count -and $users.Count) {
            Write-Log 'Geen OneDrive-gegevens uit rapport of beheercentrum (verborgen namen?). OneDrive-URL''s worden afgeleid van gelicentieerde gebruikers; gebruik Details scannen voor de grootte.' 'warning'
            foreach ($u in $users.Values) {
                if (-not @($u.assignedLicenses).Count -or [string]$u.userPrincipalName -match '#EXT#') { continue }
                $url = "$($script:MyHost)/personal/" + (([string]$u.userPrincipalName -replace '[^a-zA-Z0-9]', '_').ToLower())
                $map[$url.ToLower()] = New-SiteObject @{ Url = $url; Title = [string]$u.displayName; Owner = [string]$u.displayName; OwnerUpn = [string]$u.userPrincipalName; IsOneDrive = $true; Type = 'OneDrive'; Status = 'Afgeleid (niet gecontroleerd)' }
            }
        }
        elseif ($concealed) {
            Write-Log 'Het OneDrive-rapport heeft verborgen namen: zet in het M365-beheercentrum (Org-instellingen > Rapporten) de verborgen namen uit.' 'warning'
        }
        foreach ($o in $map.Values) {
            $u = $users[([string]$o.OwnerUpn).ToLower()]
            $st = @()
            if ($o.Status) { $st += $o.Status }
            if (-not $u) { if ($o.OwnerUpn) { $st += 'Gebruiker verwijderd' } }
            elseif ($u.accountEnabled -eq $false) { $st += 'Account uitgeschakeld' }
            elseif (-not @($u.assignedLicenses).Count) { $st += 'Geen licentie' }
            if (-not $st.Count) { $st += 'Actief' }
            $o.Status = $st -join ', '
            if ($u -and -not $o.Owner) { $o.Owner = [string]$u.displayName; $o.Title = $o.Owner }
        }
        $script:OneDrives = New-Object System.Collections.ArrayList
        foreach ($o in @($map.Values | Sort-Object { $_.Owner })) { [void]$script:OneDrives.Add($o) }
        $script:OdLoaded = $true
        Update-OdGrid
        Write-Log "$($script:OneDrives.Count) OneDrives geladen" 'success'
        Set-Status "$($script:OneDrives.Count) OneDrives geladen" 'success'
    }
    catch {
        Write-Log "OneDrives laden mislukt: $($_.Exception.Message)" 'error'
        Set-Status 'OneDrives laden mislukt' 'error'
    }
    finally { Stop-LongOp }
}

function Get-SizeCols {
    param($S)
    $pct = if ($S.QuotaMB -and $null -ne $S.UsedMB) { [Math]::Round(100 * $S.UsedMB / $S.QuotaMB, 1) } else { $null }
    $save = $null
    if ($null -ne $S.VersionsMB -or $null -ne $S.RecycleMB) { $save = ConvertTo-GB ([double]$S.VersionsMB + [double]$S.RecycleMB) }
    return @{
        GebruiktGB = (ConvertTo-GB $S.UsedMB); QuotumGB = (ConvertTo-GB $S.QuotaMB); Pct = $pct
        VersiesGB = (ConvertTo-GB $S.VersionsMB); PrullenbakGB = (ConvertTo-GB $S.RecycleMB); TeWinnenGB = $save
    }
}

$script:SiteCols = @('Key:int', 'Titel', 'Url', 'Type', 'Groep', 'GebruiktGB:double', 'QuotumGB:double', 'Pct:double', 'Bestanden:double', 'Actief:double',
                     'Weergaven:double', 'Activiteit', 'Eigenaar', 'Delen', 'Lock', 'Bibliotheken:int', 'Uniek:int', 'VersiesGB:double', 'PrullenbakGB:double',
                     'TeWinnenGB:double', 'Aangemaakt')

function Update-SiteGrid {
    $rows = for ($i = 0; $i -lt $script:Sites.Count; $i++) {
        $s = $script:Sites[$i]
        $h = Get-SizeCols $s
        $h.Key = $i; $h.Titel = $s.Title; $h.Url = $s.Url; $h.Type = $s.Type; $h.Groep = $(if ($s.GroupId) { 'Ja' } else { '' })
        $h.Bestanden = $s.Files; $h.Actief = $s.ActiveFiles; $h.Weergaven = $s.PageViews; $h.Activiteit = $s.LastActivity
        $h.Eigenaar = $s.Owner; $h.Delen = $s.Sharing; $h.Lock = $s.Lock; $h.Bibliotheken = $s.Libraries; $h.Uniek = $s.Unique
        $h.Aangemaakt = $s.Created
        $h
    }
    Set-GridRows $script:SiteGrid $script:SiteCols $rows
    Update-SiteFilter
}

function Update-SiteFilter {
    $view = $script:SiteGrid.ItemsSource
    if ($null -eq $view) { return }
    $parts = @()
    $t = $script:TxtSiteFilter.Text.Trim()
    if ($t) {
        $l = ConvertTo-LikeLiteral $t
        $parts += "(Titel LIKE '%$l%' OR Url LIKE '%$l%' OR Eigenaar LIKE '%$l%')"
    }
    $ty = [string]$script:CmbSiteType.SelectedItem
    if ($ty -and $ty -ne 'Alle typen') { $parts += "Type = '$(($ty -replace "'", "''"))'" }
    $view.RowFilter = ($parts -join ' AND ')
    $script:LblSiteCount.Text = Get-TotalsText $view $script:Sites.Count 'sites'
}

function Get-TotalsText {
    param($View, [int]$Total, [string]$Noun)
    $gb = 0.0; $save = 0.0
    foreach ($rv in $View) {
        if ($rv.Row['GebruiktGB'] -isnot [DBNull]) { $gb += [double]$rv.Row['GebruiktGB'] }
        if ($rv.Row['TeWinnenGB'] -isnot [DBNull]) { $save += [double]$rv.Row['TeWinnenGB'] }
    }
    $txt = "$($View.Count) van $Total $Noun  |  gebruikt: $([Math]::Round($gb, 1)) GB"
    if ($save -gt 0) { $txt += "  |  te winnen (versies + prullenbak): $([Math]::Round($save, 1)) GB" }
    return $txt
}

$script:OdCols = @('Key:int', 'Eigenaar', 'Upn', 'Status', 'GebruiktGB:double', 'QuotumGB:double', 'Pct:double', 'Bestanden:double', 'Actief:double',
                   'Activiteit', 'VersiesGB:double', 'PrullenbakGB:double', 'TeWinnenGB:double', 'Lock', 'Delen', 'Url')

function Update-OdGrid {
    $rows = for ($i = 0; $i -lt $script:OneDrives.Count; $i++) {
        $o = $script:OneDrives[$i]
        $h = Get-SizeCols $o
        $h.Key = $i; $h.Eigenaar = $o.Owner; $h.Upn = $o.OwnerUpn; $h.Status = $o.Status; $h.Bestanden = $o.Files
        $h.Actief = $o.ActiveFiles; $h.Activiteit = $o.LastActivity; $h.Lock = $o.Lock; $h.Delen = $o.Sharing; $h.Url = $o.Url
        $h
    }
    Set-GridRows $script:OdGrid $script:OdCols $rows
    Update-OdFilter
}

function Update-OdFilter {
    $view = $script:OdGrid.ItemsSource
    if ($null -eq $view) { return }
    $parts = @()
    $t = $script:TxtOdFilter.Text.Trim()
    if ($t) {
        $l = ConvertTo-LikeLiteral $t
        $parts += "(Eigenaar LIKE '%$l%' OR Upn LIKE '%$l%' OR Url LIKE '%$l%')"
    }
    $old = (Get-Date).AddDays(-90).ToString('yyyy-MM-dd')
    switch ([string]$script:CmbOdFilter.SelectedItem) {
        'Account uit of verwijderd' { $parts += "(Status LIKE '%uitgeschakeld%' OR Status LIKE '%verwijderd%' OR Status LIKE '%Geen licentie%')" }
        'Groter dan 10 GB'          { $parts += 'GebruiktGB > 10' }
        'Groter dan 50 GB'          { $parts += 'GebruiktGB > 50' }
        'Meer dan 80% vol'          { $parts += 'Pct > 80' }
        'Inactief > 90 dagen'       { $parts += "(Activiteit < '$old' OR Activiteit IS NULL)" }
    }
    $view.RowFilter = ($parts -join ' AND ')
    $script:LblOdCount.Text = Get-TotalsText $view $script:OneDrives.Count 'OneDrives'
}

function Get-GridTargets {
    # Geselecteerde rijen, of (na bevestiging) alle zichtbare rijen
    param($Grid, $List, [string]$What)
    $sel = @(Get-SelectedRows $Grid | ForEach-Object { $List[[int]$_['Key']] })
    if ($sel.Count) { return $sel }
    $view = $Grid.ItemsSource
    if ($null -eq $view -or $view.Count -eq 0) { Show-Msg 'Laad eerst de lijst.' $What | Out-Null; return @() }
    if ((Show-Msg "Niets geselecteerd. $What voor alle $($view.Count) zichtbare regels?" $What 'YesNo' 'Question') -ne 'Yes') { return @() }
    return @(foreach ($rv in $view) { $List[[int]$rv.Row['Key']] })
}

function Get-SelectedSites { return @(Get-SelectedRows $script:SiteGrid | ForEach-Object { $script:Sites[[int]$_['Key']] }) }
function Get-SelectedOds   { return @(Get-SelectedRows $script:OdGrid | ForEach-Object { $script:OneDrives[[int]$_['Key']] }) }

function Get-Targets {
    param([string]$Kind, [string]$What)
    if ($Kind -eq 'O') { return @(Get-GridTargets $script:OdGrid $script:OneDrives $What) }
    return @(Get-GridTargets $script:SiteGrid $script:Sites $What)
}

function Update-KindGrid {
    param([string]$Kind)
    if ($Kind -eq 'O') { Update-OdGrid } else { Update-SiteGrid }
}

#endregion

#region --- Machtigingen en bibliotheken per site ---

function Invoke-OnSite {
    # Voert een actie uit op een site; bij 'geen toegang' optioneel jezelf
    # tijdelijk sitebeheerder maken en opnieuw proberen.
    param($OnSite, [scriptblock]$Do)
    try { return (& $Do) }
    catch {
        if ((Test-AccessDenied $_) -and $script:HasAdmin -and $script:AutoElevate -and -not $script:Elevated.Contains($OnSite.Url)) {
            Write-Log "Geen toegang tot $($OnSite.Url) - jezelf tijdelijk sitebeheerder maken" 'warning'
            Set-TenantSiteAdmin -SiteUrl $OnSite.Url -Login "i:0#.f|membership|$($script:MeUpn)" -IsAdmin $true
            [void]$script:Elevated.Add($OnSite.Url)
            Save-Elevated
            Write-ChangeLog $OnSite.Url 'Tijdelijk sitebeheerder (zelf)' $script:MeUpn 'OK'
            Wait-Seconds 3
            return (& $Do)
        }
        throw
    }
}

function Get-PrincipalInfo {
    param([string]$Login)
    if ($Login -match 'SHAREPOINT\\system|app@sharepoint|^i:0i\.t\|') { return @{ Kind = 'skip' } }
    if ($Login -match '^i:0#\.f\|membership\|(.+)$') {
        $upn = $Matches[1]
        return @{ Kind = $(if ($upn -match '#ext#') { 'guest' } else { 'user' }); Upn = $upn; ObjectId = '' }
    }
    if ($Login -match '^c:0t\.c\|tenant\|([0-9a-fA-F-]{36})') { return @{ Kind = 'secgroup'; ObjectId = $Matches[1].ToLower(); Upn = '' } }
    if ($Login -match '^c:0o\.c\|federateddirectoryclaimprovider\|([0-9a-fA-F-]{36})(_o)?') {
        return @{ Kind = $(if ($Matches[2]) { 'm365owners' } else { 'm365members' }); ObjectId = $Matches[1].ToLower(); Upn = '' }
    }
    if ($Login -match 'spo-grid-all-users') { return @{ Kind = 'everyoneext'; ObjectId = ''; Upn = '' } }
    if ($Login -match '^c:0\(\.s\|true$') { return @{ Kind = 'everyone'; ObjectId = ''; Upn = '' } }
    return @{ Kind = 'other'; ObjectId = ''; Upn = '' }
}

function New-PermRow {
    param($Site, [string]$Scope, [string]$ListId, [string]$Source, [string]$Roles, $User, [string]$Mode, $SpGroupId, $SpGroupTitle)
    $login = [string]$User.LoginName
    if ([string]$User.Title -eq 'System Account') { return $null }
    $info = Get-PrincipalInfo $login
    if ($info.Kind -eq 'skip') { return $null }
    return @{
        SiteUrl = $Site.Url; SiteTitle = $Site.Title; Scope = $Scope; ListId = $ListId
        Source = $Source; Roles = $Roles; Name = [string]$User.Title
        Kind = $info.Kind; ObjectId = $info.ObjectId; Upn = $info.Upn; Login = $login; Email = [string]$User.Email
        Mode = $Mode; SpGroupId = $SpGroupId; SpGroupTitle = $SpGroupTitle; SpUserId = $User.Id
    }
}

function Get-Coll {
    # Uitgeklapte collectie: in nometadata een array, in verbose { results = [...] }
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value.PSObject.Properties['results']) { return @($Value.results) }
    return @($Value)
}

function Get-SiteMeta {
    param($Site)
    if ($Site.Meta) { return $Site.Meta }
    $base = $Site.Url
    $w = Invoke-SP ("$base/_api/web?`$select=Title,AssociatedOwnerGroup/Id,AssociatedOwnerGroup/Title,AssociatedMemberGroup/Id,AssociatedMemberGroup/Title," +
                    "AssociatedVisitorGroup/Id,AssociatedVisitorGroup/Title&`$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup") -Quiet
    try {
        $si = Invoke-SP "$base/_api/site?`$select=GroupId" -Quiet
        $gid = [string]$si.GroupId
        if ($gid -and $gid -notmatch '^0{8}-0{4}') { $Site.GroupId = $gid.ToLower() }
    } catch { }
    $Site.Meta = @{
        Owner = $w.AssociatedOwnerGroup.Id; OwnerTitle = [string]$w.AssociatedOwnerGroup.Title
        Member = $w.AssociatedMemberGroup.Id; MemberTitle = [string]$w.AssociatedMemberGroup.Title
        Visitor = $w.AssociatedVisitorGroup.Id; VisitorTitle = [string]$w.AssociatedVisitorGroup.Title
    }
    return $Site.Meta
}

function Get-SiteLists {
    # Zichtbare lijsten en bibliotheken van de hoofdweb (cache op de site)
    param($Site, [switch]$Force)
    if ($Site.Lists -and -not $Force) { return $Site.Lists }
    $r = Invoke-SP ("$($Site.Url)/_api/web/lists?`$filter=Hidden eq false&`$select=Id,Title,BaseType,BaseTemplate,ItemCount,HasUniqueRoleAssignments," +
                    "LastItemModifiedDate,EnableVersioning,MajorVersionLimit,RootFolder/ServerRelativeUrl&`$expand=RootFolder") -Quiet
    $Site.Lists = @($r.value)
    return $Site.Lists
}

function Add-RoleRows {
    param($Rows, $Site, $Assignments, $GroupsById, $Meta, [string]$Scope, [string]$ListId)
    foreach ($a in @($Assignments)) {
        $m = $a.Member
        if (-not $m) { continue }
        $roles = @(Get-Coll $a.RoleDefinitionBindings | Where-Object { -not $_.Hidden -and $_.Name -notmatch '^(Limited Access|Beperkte toegang)' } | ForEach-Object { $_.Name })
        if (-not $roles.Count) { continue }
        $rolesTxt = $roles -join ', '
        if ([int]$m.PrincipalType -eq 8) {
            $tag = if ($m.Id -eq $Meta.Owner) { ' [eigenaars]' } elseif ($m.Id -eq $Meta.Member) { ' [leden]' } elseif ($m.Id -eq $Meta.Visitor) { ' [bezoekers]' } else { '' }
            $src = if ([string]$m.Title -match '^SharingLinks\.') { 'Deellink' } else { "SP-groep: $($m.Title)$tag" }
            $g = $GroupsById[[int]$m.Id]
            $users = @(if ($g) { Get-Coll $g.Users })
            if (-not $users.Count) {
                # Lege groep toch tonen (zodat de groep van de bibliotheek te verwijderen is)
                [void]$Rows.Add(@{ SiteUrl = $Site.Url; SiteTitle = $Site.Title; Scope = $Scope; ListId = $ListId; Source = $src; Roles = $rolesTxt
                                   Name = '(lege groep)'; Kind = 'other'; ObjectId = ''; Upn = ''; Login = ''; Email = ''; Mode = 'group'
                                   SpGroupId = $m.Id; SpGroupTitle = [string]$m.Title; SpUserId = $null; Empty = $true })
            }
            foreach ($u in $users) {
                $r = New-PermRow -Site $Site -Scope $Scope -ListId $ListId -Source $src -Roles $rolesTxt -User $u -Mode 'group' -SpGroupId $m.Id -SpGroupTitle ([string]$m.Title)
                if ($r) { [void]$Rows.Add($r) }
            }
        }
        else {
            $r = New-PermRow -Site $Site -Scope $Scope -ListId $ListId -Source 'Directe rechten' -Roles $rolesTxt -User $m -Mode 'direct'
            if ($r) { [void]$Rows.Add($r) }
        }
    }
}

function Get-SitePermissionsCore {
    param($Site, [switch]$NoLists)
    $base = $Site.Url
    $meta = Get-SiteMeta $Site
    $ra = @((Invoke-SP "$base/_api/web/roleassignments?`$expand=Member,RoleDefinitionBindings" -Quiet).value)
    $sg = @((Invoke-SP "$base/_api/web/sitegroups?`$expand=Users" -Quiet).value)
    $byId = @{}
    foreach ($g in $sg) { $byId[[int]$g.Id] = $g }
    $rows = New-Object System.Collections.ArrayList
    Add-RoleRows $rows $Site $ra $byId $meta 'Site' ''
    $admins = @((Invoke-SP "$base/_api/web/siteusers?`$filter=IsSiteAdmin eq true" -Quiet).value)
    foreach ($u in $admins) {
        $r = New-PermRow -Site $Site -Scope 'Site' -ListId '' -Source 'Sitebeheerder' -Roles 'Sitecollectiebeheerder' -User $u -Mode 'admin'
        if ($r) { [void]$rows.Add($r) }
    }
    # Bibliotheken en lijsten met eigen (unieke) rechten
    if (-not $NoLists) {
        foreach ($l in @(Get-SiteLists $Site -Force | Where-Object { $_.HasUniqueRoleAssignments })) {
            $kind = if ([int]$l.BaseType -eq 1) { 'Bibliotheek' } else { 'Lijst' }
            try {
                $lra = @((Invoke-SP "$base/_api/web/lists(guid'$($l.Id)')/roleassignments?`$expand=Member,RoleDefinitionBindings" -Quiet).value)
                Add-RoleRows $rows $Site $lra $byId $meta "$($kind): $($l.Title)" ([string]$l.Id)
            }
            catch { Write-Log "Rechten van $kind '$($l.Title)' niet gelezen: $($_.Exception.Message)" 'warning' }
        }
    }
    # Groeptypes (statisch/dynamisch) ophalen
    $gids = @($rows | Where-Object { $_.ObjectId } | ForEach-Object { $_.ObjectId })
    if ($Site.GroupId) { $gids += $Site.GroupId }
    if ($gids.Count) { Resolve-Groups $gids }
    return $rows.ToArray()
}

function Get-SitePermissions {
    param($Site)
    return @(Invoke-OnSite $Site { Get-SitePermissionsCore $Site })
}

function Get-LibraryInfoCore {
    # Lijsten + grootte (StorageMetrics) per bibliotheek
    param($Site)
    $out = New-Object System.Collections.ArrayList
    foreach ($l in @(Get-SiteLists $Site)) {
        $isLib = [int]$l.BaseType -eq 1
        $size = $null; $files = $null
        if ($isLib) {
            try {
                $m = Invoke-SP "$($Site.Url)/_api/web/lists(guid'$($l.Id)')/RootFolder/StorageMetrics" -Quiet
                $size = ConvertTo-Num $m.TotalSize; $files = ConvertTo-Num $m.TotalFileStreamSize
            } catch { }
        }
        $ver = if ($null -ne $size -and $null -ne $files) { [Math]::Max(0, $size - $files) } else { $null }
        [void]$out.Add(@{
            Id = [string]$l.Id; Title = [string]$l.Title; IsLib = $isLib; Kind = $(if ($isLib) { 'Bibliotheek' } else { 'Lijst' })
            Template = [int]$l.BaseTemplate; Items = [int]$l.ItemCount; Unique = [bool]$l.HasUniqueRoleAssignments
            SizeMB = $(if ($null -ne $size) { $size / 1MB } else { $null }); FilesMB = $(if ($null -ne $files) { $files / 1MB } else { $null })
            VersionsMB = $(if ($null -ne $ver) { $ver / 1MB } else { $null })
            Versioning = $(if ($l.EnableVersioning) { "Aan (max $($l.MajorVersionLimit))" } else { 'Uit' })
            Modified = (ConvertTo-DateText $l.LastItemModifiedDate); Path = [string]$l.RootFolder.ServerRelativeUrl
        })
    }
    $Site.LibInfo = $out.ToArray()
    return $Site.LibInfo
}

function Get-RecycleInfoCore {
    param($Site)
    $r = Invoke-SP "$($Site.Url)/_api/site/RecycleBin?`$select=Size,ItemState&`$top=5000" -Quiet
    $items = @($r.value)
    $total = 0.0
    foreach ($i in $items) { $v = ConvertTo-Num $i.Size; if ($v) { $total += $v } }
    return @{ Count = $items.Count; MB = $total / 1MB; Capped = ($items.Count -ge 5000) }
}

# --- Detailpaneel (S = sites, O = OneDrive) ---

function ConvertTo-PermGridRow {
    param($R, [int]$Idx, [switch]$WithSite)
    $g = if ($R.ObjectId) { $script:GroupCache[$R.ObjectId] } else { $null }
    $soort = $script:KindLabels[$R.Kind]
    if ($g -and -not $g.Missing) { $soort = "$soort - $(Get-GroupKindLabel $g)" }
    $h = @{
        Idx = $Idx; Bereik = $R.Scope; Bron = $R.Source; Rechten = $R.Roles; Naam = $R.Name; Soort = $soort
        Lidmaatschap = $(if ($R.ObjectId) { Get-GroupMembershipLabel $g } else { '' })
        Login = $(if ($R.Upn) { $R.Upn } else { $R.Login })
    }
    if ($WithSite) { $h.Site = $R.SiteTitle; $h.SiteUrl = $R.SiteUrl }
    return $h
}

$script:PermCols = @('Idx:int', 'Bereik', 'Bron', 'Rechten', 'Naam', 'Soort', 'Lidmaatschap', 'Login')

function Update-PermGrid {
    param([string]$P)
    $x = $script:Ctx[$P]
    $i = 0
    $rows = foreach ($r in $x.PermRows) { ConvertTo-PermGridRow $r ($i++) }
    Set-GridRows $x.PermGrid $script:PermCols $rows
    $s = $x.Site
    if ($P -eq 'S') {
        if ($s -and $s.GroupId) {
            $g = $script:GroupCache[$s.GroupId]
            $x.LblGroup.Text = "Microsoft 365-groep: $($g.displayName)  -  $(Get-GroupMembershipLabel $g)"
            $x.BtnM365Edit.IsEnabled = $true
        }
        else {
            $x.LblGroup.Text = $(if ($s) { 'Geen Microsoft 365-groep gekoppeld' } else { '' })
            $x.BtnM365Edit.IsEnabled = $false
        }
    }
    else {
        $admins = @($x.PermRows | Where-Object { $_.Mode -eq 'admin' } | ForEach-Object { if ($_.Upn) { $_.Upn } else { $_.Name } })
        $x.LblGroup.Text = $(if ($s) { "Beheerders: $($admins -join ', ')" } else { '' })
    }
}

function Update-LibGrid {
    param([string]$P)
    $x = $script:Ctx[$P]
    $i = 0
    $rows = foreach ($l in @($x.Libs)) {
        @{
            Idx = $i++; Titel = $l.Title; Soort = $l.Kind; Items = $l.Items; GrootteMB = $(if ($null -ne $l.SizeMB) { [Math]::Round($l.SizeMB, 1) })
            BestandenMB = $(if ($null -ne $l.FilesMB) { [Math]::Round($l.FilesMB, 1) }); VersiesMB = $(if ($null -ne $l.VersionsMB) { [Math]::Round($l.VersionsMB, 1) })
            Versiebeheer = $l.Versioning; Uniek = $(if ($l.Unique) { 'Ja - eigen rechten' } else { 'Nee (van site)' }); Gewijzigd = $l.Modified; Pad = $l.Path
        }
    }
    Set-GridRows $x.LibGrid @('Idx:int', 'Titel', 'Soort', 'Items:int', 'GrootteMB:double', 'BestandenMB:double', 'VersiesMB:double', 'Versiebeheer', 'Uniek', 'Gewijzigd', 'Pad') $rows
    Update-LibPermGrid $P
}

function Get-SelectedLibs {
    param([string]$P)
    $x = $script:Ctx[$P]
    return @(Get-SelectedRows $x.LibGrid | ForEach-Object { $x.Libs[[int]$_['Idx']] })
}

function Update-LibPermGrid {
    param([string]$P)
    $x = $script:Ctx[$P]
    $lib = @(Get-SelectedLibs $P) | Select-Object -First 1
    if (-not $lib) {
        $x.LblLibPerm.Text = 'Selecteer een bibliotheek of lijst om de rechten te zien.'
        $x.LibPermGrid.ItemsSource = $null
        return
    }
    $scopeRows = if ($lib.Unique) { @($x.PermRows | Where-Object { $_.ListId -eq $lib.Id }) } else { @($x.PermRows | Where-Object { $_.Scope -eq 'Site' -and $_.Mode -ne 'admin' }) }
    $x.LblLibPerm.Text = if ($lib.Unique) { "Rechten van '$($lib.Title)' (eigen rechten, overname verbroken):" } else { "'$($lib.Title)' neemt de rechten van de site over:" }
    $rows = foreach ($r in $scopeRows) { ConvertTo-PermGridRow $r ([array]::IndexOf($x.PermRows, $r)) }
    Set-GridRows $x.LibPermGrid $script:PermCols $rows
}

function Show-SiteDetail {
    param([string]$P, $Site)
    $x = $script:Ctx[$P]
    if (-not $script:SpOk -or -not $Site) { return }
    $x.Site = $Site
    $x.LblTitle.Text = "$($Site.Title)  -  $($Site.Url)"
    Enter-Busy "Details laden: $($Site.Title)..."
    try {
        $Site.Lists = $null
        $x.PermRows = @(Get-SitePermissions $Site)
        $x.Libs = @(Invoke-OnSite $Site { Get-LibraryInfoCore $Site })
        $Site.Libraries = @($x.Libs | Where-Object { $_.IsLib }).Count
        $Site.Unique = @($x.Libs | Where-Object { $_.Unique }).Count
        $v = 0.0; foreach ($l in $x.Libs) { if ($l.VersionsMB) { $v += $l.VersionsMB } }
        $Site.VersionsMB = $v
        Update-PermGrid $P
        Update-LibGrid $P
        $nU = @($x.PermRows | Where-Object { $_.Scope -ne 'Site' } | ForEach-Object { $_.Scope } | Select-Object -Unique).Count
        Set-Status "$($Site.Title): $($x.PermRows.Count) machtigingen, $($x.Libs.Count) lijsten/bibliotheken ($nU met eigen rechten)" 'success'
    }
    catch {
        $x.PermRows = @(); $x.Libs = @()
        Update-PermGrid $P; Update-LibGrid $P
        $msg = $_.Exception.Message
        if (Test-AccessDenied $_) { $msg += '  (zet "Mezelf automatisch sitebeheerder maken" aan om toch te kijken)' }
        $x.LblGroup.Text = "Niet geladen: $msg"
        Write-Log "Details $($Site.Url): $msg" 'error'
        Set-Status 'Details laden mislukt' 'error'
    }
    finally { Exit-Busy }
}

function Invoke-DetailReload {
    param([string]$P)
    $x = $script:Ctx[$P]
    if ($x.Site) { $x.Site.Meta = $null; Show-SiteDetail $P $x.Site }
}

#endregion

#region --- Wijzigen op sites ---

function Get-LoginName {
    param($P)
    switch ($P.Kind) {
        'user'        { return "i:0#.f|membership|$($P.Upn)" }
        'group'       { if ($P.Unified) { return "c:0o.c|federateddirectoryclaimprovider|$($P.Id)" } else { return "c:0t.c|tenant|$($P.Id)" } }
        'everyoneext' { return "c:0-.f|rolemanager|spo-grid-all-users/$($script:TenantId)" }
        'everyone'    { return 'c:0(.s|true' }
    }
    throw "Onbekend type $($P.Kind)"
}

function Resolve-SpUser {
    param([string]$Base, [string]$Login)
    return (Invoke-SP "$Base/_api/web/ensureuser" -Method POST -VerboseOData -Body @{ logonName = $Login } -Quiet)
}

function Set-SpSiteAdmin {
    param([string]$Base, $UserId, [bool]$Value)
    [void](Invoke-SP "$Base/_api/web/getuserbyid($UserId)" -Method POST -VerboseOData -Quiet `
        -Headers @{ 'X-HTTP-Method' = 'MERGE'; 'IF-MATCH' = '*' } -Body @{ __metadata = @{ type = 'SP.User' }; IsSiteAdmin = $Value })
}

function Add-GraphRef {
    param([string]$GroupId, [string]$ObjectId, [ValidateSet('members', 'owners')][string]$Rel)
    try {
        [void](Invoke-Graph "groups/$GroupId/$Rel/`$ref" -Method POST -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ObjectId" } -Quiet)
        return $true
    }
    catch {
        if ($_.Exception.Message -match 'already exist') { return $false }
        throw
    }
}

function Remove-GraphRef {
    param([string]$GroupId, [string]$ObjectId, [ValidateSet('members', 'owners')][string]$Rel)
    try { [void](Invoke-Graph "groups/$GroupId/$Rel/$ObjectId/`$ref" -Method DELETE -Quiet); return $true }
    catch {
        if ($_.Exception.Data['Status'] -eq 404 -or $_.Exception.Message -match 'does not exist|not found') { return $false }
        throw
    }
}

function Add-SiteAccess {
    param($Site, $P, [string]$Role, [bool]$ViaM365)
    return (Invoke-OnSite $Site {
        $base = $Site.Url
        $meta = Get-SiteMeta $Site
        if ($ViaM365 -and $Site.GroupId -and $P.Kind -eq 'user' -and $Role -in 'Owner', 'Member') {
            Resolve-Groups @($Site.GroupId)
            $g = $script:GroupCache[$Site.GroupId]
            if ((Test-GroupDynamic $g) -and $Role -eq 'Member') {
                throw "De M365-groep '$($g.displayName)' is dynamisch - pas de regel aan (tab Groepen) of zet 'via M365-groep' uit."
            }
            $txt = @()
            if (-not (Test-GroupDynamic $g)) { $txt += $(if (Add-GraphRef $Site.GroupId $P.Id 'members') { 'lid M365-groep' } else { 'was al lid' }) }
            if ($Role -eq 'Owner') { $txt += $(if (Add-GraphRef $Site.GroupId $P.Id 'owners') { 'eigenaar M365-groep' } else { 'was al eigenaar' }) }
            return ($txt -join ', ')
        }
        $login = Get-LoginName $P
        if ($Role -eq 'Admin') {
            $u = Resolve-SpUser $base $login
            Set-SpSiteAdmin $base $u.Id $true
            return 'sitebeheerder'
        }
        $gid = $meta[$Role]
        if (-not $gid) { throw "Site heeft geen standaard SharePoint-groep voor rol '$Role'" }
        [void](Resolve-SpUser $base $login)
        [void](Invoke-SP "$base/_api/web/sitegroups/GetById($gid)/users" -Method POST -VerboseOData -Quiet `
            -Body @{ __metadata = @{ type = 'SP.User' }; LoginName = $login })
        return "SP-groep '$($meta["$($Role)Title"])'"
    })
}

function Remove-PermRow {
    # Mode group: gebruiker uit SP-groep (geldt voor de hele site). Met -GroupFromList:
    # de SP-groep zelf van de bibliotheek/lijst weghalen.
    param($Row, [switch]$GroupFromList)
    $base = $Row.SiteUrl
    if ($Row.Mode -eq 'group' -and ($GroupFromList -or $Row.Empty)) {
        if (-not $Row.ListId) { throw 'Een SP-groep van de site zelf weghalen kan hier niet; haal de gebruikers uit de groep.' }
        [void](Invoke-SP "$base/_api/web/lists(guid'$($Row.ListId)')/roleassignments/getbyprincipalid($($Row.SpGroupId))" -Method DELETE -Quiet)
        return
    }
    switch ($Row.Mode) {
        'group'  { [void](Invoke-SP "$base/_api/web/sitegroups/GetById($($Row.SpGroupId))/users/removebyid($($Row.SpUserId))" -Method POST -Quiet) }
        'direct' {
            if ($Row.ListId) { [void](Invoke-SP "$base/_api/web/lists(guid'$($Row.ListId)')/roleassignments/getbyprincipalid($($Row.SpUserId))" -Method DELETE -Quiet) }
            else { [void](Invoke-SP "$base/_api/web/roleassignments/getbyprincipalid($($Row.SpUserId))" -Method DELETE -Quiet) }
        }
        'admin'  { Set-SpSiteAdmin $base $Row.SpUserId $false }
    }
}

function Get-RowKey {
    param($Row)
    switch ($Row.Mode) {
        'group'  { return "g|$($Row.SpGroupId)|$($Row.SpUserId)" }
        'direct' { return "d|$($Row.ListId)|$($Row.SpUserId)" }
        default  { return "a|$($Row.SpUserId)" }
    }
}

function Test-RowMatch {
    param($Row, $P)
    switch ($P.Kind) {
        'user'        { return ($Row.Upn -and $Row.Upn -eq $P.Upn) }
        'group'       { return ($Row.ObjectId -and $Row.ObjectId -eq ([string]$P.Id).ToLower()) }
        'everyoneext' { return ($Row.Kind -eq 'everyoneext') }
        'everyone'    { return ($Row.Kind -eq 'everyone') }
    }
    return $false
}

function Remove-SiteAccess {
    param($Site, $P, $Opt)
    return (Invoke-OnSite $Site {
        $rows = @(Get-SitePermissionsCore $Site)
        $done = @()
        $seen = @{}
        foreach ($r in @($rows | Where-Object { Test-RowMatch $_ $P })) {
            $k = Get-RowKey $r
            if ($seen.ContainsKey($k)) { continue }
            $seen[$k] = $true
            if (($r.Mode -in 'group', 'direct' -and $Opt.Sp) -or ($r.Mode -eq 'admin' -and $Opt.Admin)) {
                Remove-PermRow $r
                $done += $(if ($r.Scope -ne 'Site') { "$($r.Source) ($($r.Scope))" } else { $r.Source })
            }
        }
        if ($Opt.M365 -and $Site.GroupId -and $P.Kind -eq 'user') {
            $g = $script:GroupCache[$Site.GroupId]
            if (Remove-GraphRef $Site.GroupId $P.Id 'owners') { $done += 'eigenaar M365-groep' }
            if (Test-GroupDynamic $g) { $done += '(M365-groep is dynamisch: lidmaatschap via regel)' }
            elseif (Remove-GraphRef $Site.GroupId $P.Id 'members') { $done += 'lid M365-groep' }
        }
        if (-not $done.Count) { return 'geen toegang gevonden' }
        return ($done -join '; ')
    })
}

function Get-PrincipalText {
    param($P)
    switch ($P.Kind) {
        'user'        { return $P.Upn }
        'group'       { return "$($P.Name) (groep)" }
        'everyoneext' { return 'Iedereen behalve externe gebruikers' }
        'everyone'    { return 'Iedereen' }
    }
}

function Invoke-BulkAccess {
    param([ValidateSet('Add', 'Remove')][string]$Mode)
    $what = if ($Mode -eq 'Add') { 'Toegang geven' } else { 'Toegang intrekken' }
    $sites = @(Get-Targets 'S' $what)
    if (-not $sites.Count) { return }
    $cfg = Show-BulkDialog -Mode $Mode -SiteCount $sites.Count
    if (-not $cfg) { return }
    $ps = @($cfg.Principals)
    $txt = if ($Mode -eq 'Add') { "$($ps.Count) gebruiker(s)/groep(en) toevoegen als $($script:RoleLabels[$cfg.Role])" } else { "$($ps.Count) gebruiker(s)/groep(en) verwijderen (site en bibliotheken)" }
    if ((Show-Msg "$txt op $($sites.Count) site(s)?`n`nElke wijziging wordt gelogd." $what 'YesNo' 'Warning') -ne 'Yes') { return }

    Start-LongOp "$what..."
    $ok = 0; $fail = 0; $n = 0
    try {
        foreach ($s in $sites) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            $n++
            Set-Status "$what - site $n van $($sites.Count): $($s.Title)"
            foreach ($p in $ps) {
                $who = Get-PrincipalText $p
                try {
                    $res = if ($Mode -eq 'Add') { Add-SiteAccess $s $p $cfg.Role $cfg.ViaM365 } else { Remove-SiteAccess $s $p $cfg }
                    Write-Log "$($s.Title): $who -> $res" 'success'
                    Write-ChangeLog $s.Url $what $who $res
                    $ok++
                }
                catch {
                    Write-Log "$($s.Title): $who -> $($_.Exception.Message)" 'error'
                    Write-ChangeLog $s.Url $what $who "FOUT: $($_.Exception.Message)"
                    $fail++
                }
            }
        }
    }
    finally { Stop-LongOp }
    Set-Status "$what klaar: $ok gelukt, $fail mislukt" $(if ($fail) { 'error' } else { 'success' })
    $cur = $script:Ctx['S'].Site
    if ($cur -and @($sites | Where-Object { $_.Url -eq $cur.Url }).Count) { Show-SiteDetail 'S' $cur }
}

function Invoke-PermRows {
    # Verwijdert de gekozen machtigingsregels van het detailpaneel
    param([string]$P, $Rows, [switch]$FromLibraryTab)
    $x = $script:Ctx[$P]
    $sel = @($Rows)
    if (-not $sel.Count) { Show-Msg 'Selecteer eerst een of meer regels.' 'Verwijderen' | Out-Null; return }
    $groupFromList = $false
    $grpRows = @($sel | Where-Object { $_.Mode -eq 'group' -and $_.ListId -and -not $_.Empty })
    if ($grpRows.Count) {
        $choice = Show-ChoiceDialog -Title 'SP-groep op bibliotheek' -Text "Er zitten $($grpRows.Count) regel(s) bij die via een SharePoint-groep op een bibliotheek/lijst staan. Wat wil je doen?" -Items @(
            [pscustomobject]@{ Label = 'De SP-groep van deze bibliotheek/lijst weghalen'; Value = 'list' }
            [pscustomobject]@{ Label = 'De gebruiker uit de SP-groep halen (geldt voor de hele site)'; Value = 'user' })
        if (-not $choice) { return }
        $groupFromList = ($choice.Value -eq 'list')
    }
    $lines = ($sel | Select-Object -First 15 | ForEach-Object { "- $($_.Name)  ($($_.Source), $($_.Scope))" }) -join "`n"
    if ($sel.Count -gt 15) { $lines += "`n... en nog $($sel.Count - 15)" }
    if ((Show-Msg "Deze $($sel.Count) toegang(en) verwijderen op $($x.Site.Title)?`n`n$lines" 'Verwijderen' 'YesNo' 'Warning') -ne 'Yes') { return }
    Enter-Busy 'Verwijderen...'
    $seen = @{}
    try {
        foreach ($r in $sel) {
            $gfl = $groupFromList -and $r.Mode -eq 'group' -and $r.ListId
            $k = if ($gfl) { "gl|$($r.ListId)|$($r.SpGroupId)" } else { Get-RowKey $r }
            if ($seen.ContainsKey($k)) { continue }
            $seen[$k] = $true
            $who = if ($gfl) { "SP-groep $($r.SpGroupTitle)" } elseif ($r.Upn) { $r.Upn } else { $r.Name }
            try {
                Invoke-OnSite $x.Site { Remove-PermRow $r -GroupFromList:$gfl }
                Write-Log "$($r.SiteTitle): $who verwijderd ($($r.Source), $($r.Scope))" 'success'
                Write-ChangeLog $r.SiteUrl "Verwijderen ($($r.Source), $($r.Scope))" $who 'OK'
            }
            catch {
                Write-Log "$who ($($r.Scope)): $($_.Exception.Message)" 'error'
                Write-ChangeLog $r.SiteUrl "Verwijderen ($($r.Source), $($r.Scope))" $who "FOUT: $($_.Exception.Message)"
            }
        }
    }
    finally { Exit-Busy }
    Show-SiteDetail $P $x.Site
}

# --- Bibliotheek-acties ---

function Get-OneLib {
    param([string]$P, [string]$What)
    $lib = @(Get-SelectedLibs $P) | Select-Object -First 1
    if (-not $lib) { Show-Msg 'Selecteer eerst een bibliotheek of lijst.' $What | Out-Null }
    return $lib
}

function Invoke-LibBreak {
    param([string]$P)
    $x = $script:Ctx[$P]; $lib = Get-OneLib $P 'Eigen rechten'
    if (-not $lib) { return }
    if ($lib.Unique) { Show-Msg "'$($lib.Title)' heeft al eigen rechten." 'Eigen rechten' | Out-Null; return }
    if ((Show-Msg "Overname van rechten verbreken voor '$($lib.Title)'?`n`nDe huidige site-rechten worden gekopieerd; daarna kun je ze per bibliotheek aanpassen." 'Eigen rechten' 'YesNo' 'Question') -ne 'Yes') { return }
    try {
        Invoke-OnSite $x.Site { [void](Invoke-SP "$($x.Site.Url)/_api/web/lists(guid'$($lib.Id)')/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)" -Method POST) }
        Write-ChangeLog $x.Site.Url "Overname verbroken: $($lib.Title)" '' 'OK'
        Write-Log "$($lib.Title): eigen rechten ingesteld (kopie van site)" 'success'
    }
    catch { Write-Log "Overname verbreken mislukt: $($_.Exception.Message)" 'error' }
    Show-SiteDetail $P $x.Site
}

function Invoke-LibReset {
    param([string]$P)
    $x = $script:Ctx[$P]; $lib = Get-OneLib $P 'Overname herstellen'
    if (-not $lib) { return }
    if (-not $lib.Unique) { Show-Msg "'$($lib.Title)' neemt de site-rechten al over." 'Overname herstellen' | Out-Null; return }
    if ((Show-Msg "Rechten van '$($lib.Title)' weer gelijk maken aan de site?`n`nAlle eigen rechten van deze bibliotheek/lijst vervallen." 'Overname herstellen' 'YesNo' 'Warning') -ne 'Yes') { return }
    try {
        Invoke-OnSite $x.Site { [void](Invoke-SP "$($x.Site.Url)/_api/web/lists(guid'$($lib.Id)')/resetroleinheritance" -Method POST) }
        Write-ChangeLog $x.Site.Url "Overname hersteld: $($lib.Title)" '' 'OK'
        Write-Log "$($lib.Title): rechten weer van de site" 'success'
    }
    catch { Write-Log "Overname herstellen mislukt: $($_.Exception.Message)" 'error' }
    Show-SiteDetail $P $x.Site
}

function Invoke-LibGrant {
    param([string]$P)
    $x = $script:Ctx[$P]; $lib = Get-OneLib $P 'Toegang geven'
    if (-not $lib) { return }
    if (-not $lib.Unique) {
        if ((Show-Msg "'$($lib.Title)' neemt nu de site-rechten over. Eerst eigen rechten instellen (kopie van de site) en dan de toegang toevoegen?" 'Toegang geven' 'YesNo' 'Question') -ne 'Yes') { return }
    }
    $ps = @(Show-PrincipalPicker -Title "Toegang geven tot '$($lib.Title)'")
    if (-not $ps.Count) { return }
    $role = Show-ChoiceDialog -Title 'Rechtenniveau' -Text "Welk niveau krijgen $($ps.Count) gebruiker(s)/groep(en) op '$($lib.Title)'?" -Items $script:LibRoles
    if (-not $role) { return }
    Enter-Busy 'Toegang geven...'
    try {
        Invoke-OnSite $x.Site {
            $base = $x.Site.Url
            if (-not $lib.Unique) { [void](Invoke-SP "$base/_api/web/lists(guid'$($lib.Id)')/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)" -Method POST -Quiet) }
            $rd = Invoke-SP "$base/_api/web/roledefinitions/getbytype($($role.Value))" -Quiet
            foreach ($p in $ps) {
                $who = Get-PrincipalText $p
                try {
                    $u = Resolve-SpUser $base (Get-LoginName $p)
                    [void](Invoke-SP "$base/_api/web/lists(guid'$($lib.Id)')/roleassignments/addroleassignment(principalid=$($u.Id),roledefid=$($rd.Id))" -Method POST -Quiet)
                    Write-Log "$($lib.Title): $who -> $($role.Label)" 'success'
                    Write-ChangeLog $base "Toegang $($lib.Title): $($role.Label)" $who 'OK'
                }
                catch {
                    Write-Log "$($lib.Title): $who -> $($_.Exception.Message)" 'error'
                    Write-ChangeLog $base "Toegang $($lib.Title): $($role.Label)" $who "FOUT: $($_.Exception.Message)"
                }
            }
        }
    }
    catch { Write-Log "Toegang geven mislukt: $($_.Exception.Message)" 'error' }
    finally { Exit-Busy }
    Show-SiteDetail $P $x.Site
}

function Invoke-LibPermRemove {
    param([string]$P)
    $x = $script:Ctx[$P]
    $lib = @(Get-SelectedLibs $P) | Select-Object -First 1
    if ($lib -and -not $lib.Unique) {
        Show-Msg "'$($lib.Title)' neemt de site-rechten over. Verwijder op tab Machtigingen, of stel eerst eigen rechten in." 'Verwijderen' | Out-Null
        return
    }
    $rows = @(Get-SelectedRows $x.LibPermGrid | ForEach-Object { $x.PermRows[[int]$_['Idx']] })
    Invoke-PermRows $P $rows -FromLibraryTab
}

function Invoke-LibVersionLimit {
    param([string]$P)
    $x = $script:Ctx[$P]
    $libs = @(Get-SelectedLibs $P | Where-Object { $_.IsLib })
    if (-not $libs.Count) { Show-Msg 'Selecteer een of meer bibliotheken.' 'Versielimiet' | Out-Null; return }
    $v = Show-InputDialog -Title 'Versielimiet' -Prompt "Maximaal aantal hoofdversies voor $($libs.Count) bibliotheek/bibliotheken (1-50000, SharePoint-standaard 500):" -Default '100'
    if (-not $v) { return }
    $n = 0
    if (-not [int]::TryParse($v, [ref]$n) -or $n -lt 1 -or $n -gt 50000) { Show-Msg 'Ongeldig getal.' 'Versielimiet' | Out-Null; return }
    Enter-Busy 'Versielimiet instellen...'
    try {
        foreach ($l in $libs) {
            try {
                Invoke-OnSite $x.Site { Set-ListVersionLimit $x.Site.Url $l.Id $n }
                Write-Log "$($l.Title): max $n versies" 'success'
                Write-ChangeLog $x.Site.Url "Versielimiet $($l.Title)" '' "$n"
            }
            catch { Write-Log "$($l.Title): $($_.Exception.Message)" 'error' }
        }
    }
    finally { Exit-Busy }
    Show-SiteDetail $P $x.Site
}

function Set-ListVersionLimit {
    param([string]$Base, [string]$ListId, [int]$Limit)
    [void](Invoke-SP "$Base/_api/web/lists(guid'$ListId')" -Method POST -VerboseOData -Quiet `
        -Headers @{ 'X-HTTP-Method' = 'MERGE'; 'IF-MATCH' = '*' } -Body @{ __metadata = @{ type = 'SP.List' }; EnableVersioning = $true; MajorVersionLimit = $Limit })
}

function Invoke-LibVersionClean {
    param([string]$P)
    $x = $script:Ctx[$P]
    $libs = @(Get-SelectedLibs $P | Where-Object { $_.IsLib })
    if (-not $libs.Count) { Show-Msg 'Selecteer een of meer bibliotheken.' 'Versies opschonen' | Out-Null; return }
    $cfg = Show-VersionDialog -Text "Versies opschonen in $($libs.Count) bibliotheek/bibliotheken van $($x.Site.Title)."
    if (-not $cfg) { return }
    Start-LongOp 'Versies opschonen...'
    try {
        foreach ($l in $libs) {
            if ($script:StopRequested) { break }
            try {
                $res = Invoke-OnSite $x.Site { Clear-ListVersions $x.Site $l $cfg.Keep }
                if ($cfg.SetLimit) { Invoke-OnSite $x.Site { Set-ListVersionLimit $x.Site.Url $l.Id $cfg.Limit } }
                Write-Log "$($l.Title): $($res.Files) bestanden met versies, $($res.Deleted) opgeschoond" 'success'
                Write-ChangeLog $x.Site.Url "Versies opschonen $($l.Title) (bewaar $($cfg.Keep))" '' "$($res.Deleted) bestanden"
            }
            catch { Write-Log "$($l.Title): $($_.Exception.Message)" 'error' }
        }
    }
    finally { Stop-LongOp }
    Show-SiteDetail $P $x.Site
}

# --- Opslag: versies, prullenbak, leegmaken ---

function Clear-ListVersions {
    # Oude versies van alle bestanden in een bibliotheek verwijderen; bewaar de laatste $Keep vorige versies
    param($Site, $Lib, [int]$Keep)
    $base = $Site.Url
    $files = 0; $deleted = 0
    $url = "$base/_api/web/lists(guid'$($Lib.Id)')/items?`$select=Id,FSObjType,UniqueId,OData__UIVersionString&`$top=2000"
    while ($url -and -not $script:StopRequested) {
        $j = Get-SpItemsPage $url
        foreach ($it in @($j.value)) {
            if ($script:StopRequested) { break }
            if ([int]$it.FSObjType -ne 0) { continue }
            $vs = [string]$it.OData__UIVersionString
            if ($vs -and $vs -match '^(1\.0|0\.1)$') { continue }   # eerste versie: geen oudere versies
            $files++
            if ($files % 25 -eq 0) { Set-Status "$($Site.Title) / $($Lib.Title): $files bestanden met versies, $deleted opgeschoond..." }
            try {
                $fileApi = "$base/_api/web/GetFileById('$($it.UniqueId)')/versions"
                if ($Keep -le 0) {
                    [void](Invoke-SP "$fileApi/DeleteAll()" -Method POST -Quiet)
                    $deleted++
                }
                else {
                    $vers = @((Invoke-SP "$fileApi`?`$select=ID" -Quiet).value | ForEach-Object { [int]$_.ID } | Sort-Object)
                    $drop = $vers.Count - $Keep
                    if ($drop -gt 0) {
                        foreach ($vid in ($vers | Select-Object -First $drop)) { [void](Invoke-SP "$fileApi/DeleteByID($vid)" -Method POST -Quiet) }
                        $deleted++
                    }
                }
            }
            catch { Write-Log "Versies van item $($it.Id) in $($Lib.Title): $($_.Exception.Message)" 'warning' }
        }
        $url = $j.'odata.nextLink'
    }
    return @{ Files = $files; Deleted = $deleted }
}

function Clear-SiteRecycleCore {
    # 1e fase leegmaken (gaat naar 2e fase) en daarna 2e fase leegmaken
    param($Site)
    [void](Invoke-SP "$($Site.Url)/_api/web/RecycleBin/DeleteAll()" -Method POST -Quiet)
    # Ook items van subsites (sitecollectie-niveau, 1e fase)
    try { [void](Invoke-SP "$($Site.Url)/_api/site/RecycleBin/DeleteAll()" -Method POST -Quiet) } catch { Write-Log "$($Site.Title): prullenbak sitecollectie: $($_.Exception.Message)" 'warning' }
    [void](Invoke-SP "$($Site.Url)/_api/site/RecycleBin/DeleteAllSecondStageItems()" -Method POST -Quiet)
}

function Get-ContentLibs {
    # Bibliotheken met gebruikersbestanden (OneDrive: Documenten = sjabloon 700)
    param($Site)
    $libs = @(Get-SiteLists $Site -Force | Where-Object { [int]$_.BaseType -eq 1 -and [string]$_.Title -notmatch 'Preservation Hold|Bewaarbibliotheek|Form Templates|Formuliersjablonen|Style Library|Stijlbibliotheek|Site Assets|Siteactiva|Site Pages|Sitepagina' })
    $od = @($libs | Where-Object { [int]$_.BaseTemplate -eq 700 })
    if ($Site.IsOneDrive -and $od.Count) { return $od }
    return $libs
}

function Clear-OneDriveContentCore {
    param($Site)
    $base = $Site.Url
    $n = 0
    foreach ($l in @(Get-ContentLibs $Site)) {
        $root = "$base/_api/web/lists(guid'$($l.Id)')/RootFolder"
        foreach ($f in @((Invoke-SP "$root/Folders?`$select=UniqueId,Name" -Quiet).value)) {
            if ($script:StopRequested) { break }
            if ([string]$f.Name -eq 'Forms') { continue }
            [void](Invoke-SP "$base/_api/web/GetFolderById('$($f.UniqueId)')/recycle()" -Method POST -Quiet)
            $n++
        }
        foreach ($f in @((Invoke-SP "$root/Files?`$select=UniqueId,Name" -Quiet).value)) {
            if ($script:StopRequested) { break }
            [void](Invoke-SP "$base/_api/web/GetFileById('$($f.UniqueId)')/recycle()" -Method POST -Quiet)
            $n++
        }
    }
    return $n
}

function Invoke-BulkDetails {
    param([string]$Kind)
    $items = @(Get-Targets $Kind 'Details scannen')
    if (-not $items.Count) { return }
    if ($items.Count -gt 25 -and (Show-Msg "Details scannen voor $($items.Count) items kan even duren (meerdere aanroepen per site). Doorgaan?" 'Details' 'YesNo' 'Question') -ne 'Yes') { return }
    Start-LongOp 'Details scannen...'
    $n = 0
    try {
        foreach ($s in $items) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            $n++
            Set-Status "Details $n van $($items.Count): $($s.Title)"
            if ($script:HasAdmin) {
                try {
                    $p = Get-TenantSiteProps $s.Url
                    if ($p) {
                        $s.Sharing = $script:SharingLabels[[int]$p.SharingCapability]
                        $s.Lock = $(if ([string]$p.LockState -and [string]$p.LockState -ne 'Unlock') { $script:LockLabels[[string]$p.LockState] } else { '' })
                        if ($null -ne $p.StorageUsage) { $s.UsedMB = [double]$p.StorageUsage }
                        if ($p.StorageMaximumLevel) { $s.QuotaMB = [double]$p.StorageMaximumLevel }
                        if (-not $s.Owner -and $p.Owner) { $s.Owner = [string]$p.Owner }
                        if ($s.Status -match 'Afgeleid') { $s.Status = $s.Status -replace 'Afgeleid \(niet gecontroleerd\)', 'Gecontroleerd' }
                    }
                }
                catch { Write-Log "$($s.Url) (eigenschappen): $($_.Exception.Message)" 'warning' }
            }
            try {
                Invoke-OnSite $s {
                    $s.Lists = $null
                    $libs = @(Get-LibraryInfoCore $s)
                    $s.Libraries = @($libs | Where-Object { $_.IsLib }).Count
                    $s.Unique = @($libs | Where-Object { $_.Unique }).Count
                    $v = 0.0; foreach ($l in $libs) { if ($l.VersionsMB) { $v += $l.VersionsMB } }
                    $s.VersionsMB = $v
                    $rb = Get-RecycleInfoCore $s
                    $s.RecycleMB = $rb.MB
                    if ($rb.Capped) { Write-Log "$($s.Title): prullenbak heeft 5000+ items, grootte is een ondergrens" 'warning' }
                }
            }
            catch { Write-Log "$($s.Url) (bibliotheken/prullenbak): $($_.Exception.Message)" 'error' }
        }
    }
    finally { Stop-LongOp }
    Update-KindGrid $Kind
    Set-Status "Details van $n item(s) bijgewerkt" 'success'
}

function Invoke-BulkSharing {
    param([string]$Kind)
    if (-not $script:HasAdmin) { Show-Msg 'Hiervoor is de rol SharePoint Administrator nodig.' 'Extern delen' | Out-Null; return }
    $items = @(Get-Targets $Kind 'Extern delen instellen')
    if (-not $items.Count) { return }
    $opts = @(0, 3, 1, 2 | ForEach-Object { [pscustomobject]@{ Value = $_; Label = $script:SharingLabels[$_] } })
    $choice = Show-ChoiceDialog -Title 'Extern delen instellen' -Text "Nieuw niveau voor $($items.Count) item(s). Een site kan niet ruimer staan dan de tenant-instelling." -Items $opts
    if ($null -eq $choice) { return }
    if ((Show-Msg "Extern delen op '$($choice.Label)' zetten voor $($items.Count) item(s)?" 'Extern delen' 'YesNo' 'Warning') -ne 'Yes') { return }
    Invoke-BulkCsom $Kind $items 'Extern delen' @{ SharingCapability = @('Enum', [int]$choice.Value) } { param($s) $s.Sharing = $choice.Label } $choice.Label
}

function Invoke-BulkQuota {
    param([string]$Kind)
    if (-not $script:HasAdmin) { Show-Msg 'Hiervoor is de rol SharePoint Administrator nodig.' 'Quotum' | Out-Null; return }
    $items = @(Get-Targets $Kind 'Quotum instellen')
    if (-not $items.Count) { return }
    $v = Show-InputDialog -Title 'Opslagquotum' -Prompt "Nieuw quotum in GB voor $($items.Count) item(s). Waarschuwing wordt op 90% gezet.`n(Bij automatisch sitequotum in de tenant heeft dit voor sites geen effect.)" -Default '1024'
    if (-not $v) { return }
    $gb = 0
    if (-not [int]::TryParse($v, [ref]$gb) -or $gb -lt 1) { Show-Msg 'Ongeldig getal.' 'Quotum' | Out-Null; return }
    if ((Show-Msg "Quotum op $gb GB zetten voor $($items.Count) item(s)?" 'Quotum' 'YesNo' 'Warning') -ne 'Yes') { return }
    $mb = [int64]$gb * 1024
    Invoke-BulkCsom $Kind $items 'Quotum' ([ordered]@{ StorageMaximumLevel = @('Int64', $mb); StorageWarningLevel = @('Int64', [int64]($mb * 0.9)) }) { param($s) $s.QuotaMB = $mb } "$gb GB"
}

function Invoke-BulkLock {
    param([string]$Kind)
    if (-not $script:HasAdmin) { Show-Msg 'Hiervoor is de rol SharePoint Administrator nodig.' 'Vergrendeling' | Out-Null; return }
    $items = @(Get-Targets $Kind 'Vergrendeling instellen')
    if (-not $items.Count) { return }
    $opts = @($script:LockLabels.Keys | ForEach-Object { [pscustomobject]@{ Value = $_; Label = $script:LockLabels[$_] } })
    $choice = Show-ChoiceDialog -Title 'Vergrendeling' -Text "Vergrendeling voor $($items.Count) item(s). 'Geen toegang' blokkeert iedereen, ook de eigenaar." -Items $opts
    if ($null -eq $choice) { return }
    if ((Show-Msg "Vergrendeling '$($choice.Label)' instellen voor $($items.Count) item(s)?" 'Vergrendeling' 'YesNo' 'Warning') -ne 'Yes') { return }
    Invoke-BulkCsom $Kind $items 'Vergrendeling' @{ LockState = @('String', [string]$choice.Value) } { param($s) $s.Lock = $(if ($choice.Value -eq 'Unlock') { '' } else { $choice.Label }) } $choice.Label
}

function Invoke-BulkCsom {
    param([string]$Kind, $Items, [string]$What, $Props, [scriptblock]$OnOk, [string]$ResultText)
    Start-LongOp "$What instellen..."
    $ok = 0; $fail = 0
    try {
        foreach ($s in $Items) {
            if ($script:StopRequested) { break }
            Set-Status "$($What): $($s.Title)"
            try {
                Set-TenantSiteProps $s.Url $Props
                & $OnOk $s
                Write-Log "$($s.Title): $What -> $ResultText" 'success'
                Write-ChangeLog $s.Url $What '' $ResultText
                $ok++
            }
            catch {
                Write-Log "$($s.Title): $($_.Exception.Message)" 'error'
                Write-ChangeLog $s.Url $What '' "FOUT: $($_.Exception.Message)"
                $fail++
            }
        }
    }
    finally { Stop-LongOp }
    Update-KindGrid $Kind
    Set-Status "$($What): $ok gelukt, $fail mislukt" $(if ($fail) { 'error' } else { 'success' })
}

function Invoke-BulkVersions {
    param([string]$Kind)
    $items = @(Get-Targets $Kind 'Versies opschonen')
    if (-not $items.Count) { return }
    $cfg = Show-VersionDialog -Text "Versies opschonen in alle documentbibliotheken van $($items.Count) item(s). Dit kan lang duren bij veel bestanden (1 aanroep per bestand met versies)."
    if (-not $cfg) { return }
    Start-LongOp 'Versies opschonen...'
    $n = 0
    try {
        foreach ($s in $items) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            $n++
            try {
                Invoke-OnSite $s {
                    $tot = 0
                    foreach ($l in @(Get-ContentLibs $s)) {
                        if ($script:StopRequested) { break }
                        Set-Status "Versies $n van $($items.Count): $($s.Title) / $($l.Title)"
                        $res = Clear-ListVersions $s @{ Id = [string]$l.Id; Title = [string]$l.Title } $cfg.Keep
                        if ($cfg.SetLimit) { Set-ListVersionLimit $s.Url ([string]$l.Id) $cfg.Limit }
                        $tot += $res.Deleted
                    }
                    Write-Log "$($s.Title): versies opgeschoond van $tot bestanden" 'success'
                    Write-ChangeLog $s.Url "Versies opschonen (bewaar $($cfg.Keep)$(if ($cfg.SetLimit) { ", limiet $($cfg.Limit)" }))" '' "$tot bestanden"
                    $s.VersionsMB = $null
                }
            }
            catch {
                Write-Log "$($s.Title): $($_.Exception.Message)" 'error'
                Write-ChangeLog $s.Url 'Versies opschonen' '' "FOUT: $($_.Exception.Message)"
            }
        }
    }
    finally { Stop-LongOp }
    Update-KindGrid $Kind
    Set-Status "Versies opgeschoond op $n item(s). Ruimte komt vrij via de prullenbak (2e fase) - leeg die zo nodig." 'success'
}

function Invoke-BulkRecycle {
    param([string]$Kind)
    $items = @(Get-Targets $Kind 'Prullenbak legen')
    if (-not $items.Count) { return }
    if ((Show-Msg "De prullenbak (1e EN 2e fase) DEFINITIEF legen voor $($items.Count) item(s)?`n`nVerwijderde bestanden zijn daarna niet meer terug te zetten." 'Prullenbak legen' 'YesNo' 'Warning') -ne 'Yes') { return }
    Start-LongOp 'Prullenbak legen...'
    $ok = 0; $fail = 0
    try {
        foreach ($s in $items) {
            if ($script:StopRequested) { break }
            Set-Status "Prullenbak legen: $($s.Title)"
            try {
                Invoke-OnSite $s { Clear-SiteRecycleCore $s }
                $s.RecycleMB = 0
                Write-Log "$($s.Title): prullenbak geleegd" 'success'
                Write-ChangeLog $s.Url 'Prullenbak legen' '' 'OK'
                $ok++
            }
            catch {
                Write-Log "$($s.Title): $($_.Exception.Message)" 'error'
                Write-ChangeLog $s.Url 'Prullenbak legen' '' "FOUT: $($_.Exception.Message)"
                $fail++
            }
        }
    }
    finally { Stop-LongOp }
    Update-KindGrid $Kind
    Set-Status "Prullenbak: $ok geleegd, $fail mislukt" $(if ($fail) { 'error' } else { 'success' })
}

# --- Prullenbak terugzetten ---

$script:RecycleTypeLabels = @{
    1 = 'Bestand'; 2 = 'Bestandsversie'; 3 = 'Lijstitem'; 4 = 'Lijst/bibliotheek'; 5 = 'Map'; 6 = 'Map'
    7 = 'Bijlage'; 8 = 'Itemversie'; 9 = 'Overig'; 10 = 'Subsite'; 11 = 'App'
}

function Get-RecycleItemsCore {
    # Alle items uit de prullenbak van de sitecollectie (1e fase van alle webs + 2e fase).
    # Pagineert via GetRecycleBinItems; lukt dat niet, dan terugval op de eerste 5000.
    param($Site, [int]$Max = 100000)
    $sel = 'Id,Title,LeafName,DirName,ItemType,ItemState,DeletedDate,DeletedByName,DeletedByEmail,AuthorName,Size'
    $all = New-Object System.Collections.ArrayList
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $paging = ''
    $paged = $true
    $capped = $false
    while (-not $script:StopRequested) {
        $page = @()
        if ($paged) {
            try {
                $pi = [uri]::EscapeDataString($paging.Replace("'", "''"))
                $page = @((Invoke-SP "$($Site.Url)/_api/site/GetRecycleBinItems(pagingInfo='$pi',rowLimit=5000,isAscending=true,orderBy=1,itemState=0)?`$select=$sel" -Quiet).value)
            }
            catch {
                if ($paging) { Write-Log "$($Site.Title): volgende pagina prullenbak mislukt ($($_.Exception.Message)) - lijst is onvolledig" 'warning'; $capped = $true; break }
                if (Test-AccessDenied $_) { throw }
                $paged = $false
            }
        }
        if (-not $paged) { $page = @((Invoke-SP "$($Site.Url)/_api/site/RecycleBin?`$select=$sel&`$top=5000" -Quiet).value) }
        $new = 0
        foreach ($i in $page) { if ($i -and $seen.Add([string]$i.Id)) { [void]$all.Add($i); $new++ } }
        if ($page.Count -lt 5000) { break }
        if (-not $paged -or $new -eq 0 -or $all.Count -ge $Max) { $capped = $true; break }
        Set-Status "Prullenbak $($Site.Title): $($all.Count) items..."
        $last = $page[$page.Count - 1]
        $paging = "id=$($last.Id)&title=$([uri]::EscapeDataString([string]$last.Title))"
    }
    if ($capped) { Write-Log "$($Site.Title): prullenbak heeft meer dan $($all.Count) items - niet alles is opgehaald" 'warning' }
    return ,$all.ToArray()
}

function ConvertFrom-RecycleItem {
    param($Item, $Site, [int]$SiteKey)
    $dt = [datetime]::MinValue
    if ($Item.DeletedDate) {
        try { $dt = ([datetime]::Parse([string]$Item.DeletedDate, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)).ToLocalTime() } catch { }
    }
    $type = [int]$Item.ItemType
    $name = [string]$Item.LeafName
    if (-not $name) { $name = [string]$Item.Title }
    return @{
        Id = [string]$Item.Id; SiteKey = $SiteKey; Site = [string]$Site.Title; SiteUrl = $Site.Url
        Name = $name; Dir = [string]$Item.DirName; ItemType = $type
        Type = $(if ($script:RecycleTypeLabels.ContainsKey($type)) { $script:RecycleTypeLabels[$type] } else { "Type $type" })
        Stage = $(if ([int]$Item.ItemState -eq 2) { '2e fase' } else { '1e fase' })
        By = $(if ($Item.DeletedByName) { [string]$Item.DeletedByName } else { [string]$Item.DeletedByEmail })
        ByEmail = [string]$Item.DeletedByEmail; Author = [string]$Item.AuthorName
        Deleted = $dt; SizeMB = [Math]::Round(([double](ConvertTo-Num $Item.Size)) / 1MB, 2)
    }
}

function Restore-RecycleItemsCore {
    # Zet items terug; containers (subsites, lijsten, mappen) eerst, ondiepe paden eerst.
    param($Site, $Items)
    $ok = 0; $failed = New-Object System.Collections.ArrayList
    $order = @($Items | Sort-Object @{ Expression = { if ($_.ItemType -in 10, 4, 5, 6) { 0 } else { 1 } } }, @{ Expression = { $_.Dir.Length } }, @{ Expression = { $_.Name } })
    $base = "$($Site.Url)/_api/site/RecycleBin"
    $useBatch = $true
    for ($i = 0; $i -lt $order.Count; $i += 50) {
        if ($script:StopRequested) { break }
        $chunk = @($order[$i..([Math]::Min($i + 49, $order.Count - 1))])
        Set-Status "Terugzetten $($Site.Title): $([Math]::Min($i + 50, $order.Count)) van $($order.Count)..."
        if ($useBatch -and $chunk.Count -gt 1) {
            try {
                [void](Invoke-SP "$base/RestoreByIds" -Method POST -Body @{ ids = @($chunk | ForEach-Object { $_.Id }) } -Quiet)
                $ok += $chunk.Count
                continue
            }
            catch {
                if (Test-AccessDenied $_) { throw }
                if ($_.Exception.Message -match 'RestoreByIds|does not exist on type|bestaat niet op type|\[404\]') { $useBatch = $false }
                # Anders: per item opnieuw om te zien welke items het probleem zijn
            }
        }
        foreach ($it in $chunk) {
            if ($script:StopRequested) { break }
            try {
                [void](Invoke-SP "$base('$($it.Id)')/restore()" -Method POST -Quiet)
                $ok++
            }
            catch {
                if (Test-AccessDenied $_) { throw }
                $msg = $_.Exception.Message
                # Al teruggezet door een eerdere (deels gelukte) batch of via de bovenliggende map
                if ($msg -match '\[404\]|does not exist|bestaat niet|niet gevonden|not found') {
                    $chk = $null
                    try { $chk = Invoke-SP "$base('$($it.Id)')?`$select=Id" -Quiet } catch { }
                    if (-not $chk) { $ok++; continue }
                }
                [void]$failed.Add(@{ Item = $it; Error = $msg })
            }
        }
    }
    return @{ Ok = $ok; Failed = $failed.ToArray() }
}

function Invoke-BulkRestore {
    param([string]$Kind)
    $items = @(Get-Targets $Kind 'Prullenbak terugzetten')
    if (-not $items.Count) { return }
    if ($items.Count -gt 25 -and (Show-Msg "De prullenbak van $($items.Count) item(s) ophalen kan even duren. Doorgaan?" 'Prullenbak terugzetten' 'YesNo' 'Question') -ne 'Yes') { return }
    # 1. Inhoud van de prullenbakken ophalen
    Start-LongOp 'Prullenbak ophalen...'
    $rows = New-Object System.Collections.ArrayList
    $n = 0
    try {
        foreach ($s in $items) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            $n++
            Set-Status "Prullenbak ophalen $n van $($items.Count): $($s.Title)"
            try {
                $key = $(if ($Kind -eq 'O') { $script:OneDrives.IndexOf($s) } else { $script:Sites.IndexOf($s) })
                $raw = @(Invoke-OnSite $s { Get-RecycleItemsCore $s })
                foreach ($r in $raw) { [void]$rows.Add((ConvertFrom-RecycleItem $r $s $key)) }
                Write-Log "$($s.Title): $($raw.Count) item(s) in de prullenbak" 'info'
            }
            catch { Write-Log "$($s.Title) (prullenbak): $($_.Exception.Message)" 'error' }
        }
    }
    finally { Stop-LongOp }
    if (-not $rows.Count) { Set-Status 'Geen items in de prullenbak gevonden'; Show-Msg 'Er staan geen (leesbare) items in de prullenbak van de gekozen site(s).' 'Prullenbak terugzetten' | Out-Null; return }
    Set-Status "$($rows.Count) item(s) in de prullenbak gevonden"
    # 2. Kiezen wat terug moet
    $chosen = @(Show-RestoreDialog -Rows $rows.ToArray() -SiteCount $n)
    if (-not $chosen.Count) { return }
    # 3. Terugzetten per site
    Start-LongOp 'Terugzetten...'
    $ok = 0; $fail = 0
    $list = if ($Kind -eq 'O') { $script:OneDrives } else { $script:Sites }
    try {
        foreach ($grp in @($chosen | Group-Object { $_.SiteKey })) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            $s = $list[[int]$grp.Name]
            try {
                $res = Invoke-OnSite $s { Restore-RecycleItemsCore $s @($grp.Group) }
                $ok += $res.Ok; $fail += $res.Failed.Count
                foreach ($f in $res.Failed) {
                    Write-Log "$($s.Title): '$($f.Item.Dir)/$($f.Item.Name)' niet teruggezet: $($f.Error)" 'error'
                    Write-ChangeLog $s.Url 'Prullenbak terugzetten' "$($f.Item.Dir)/$($f.Item.Name)" "FOUT: $($f.Error)"
                }
                $txt = "$($res.Ok) van $($grp.Count) item(s) teruggezet"
                Write-Log "$($s.Title): $txt" $(if ($res.Failed.Count) { 'warning' } else { 'success' })
                Write-ChangeLog $s.Url 'Prullenbak terugzetten' '' $txt
                $s.RecycleMB = $null
            }
            catch {
                $fail += $grp.Count
                Write-Log "$($s.Title): $($_.Exception.Message)" 'error'
                Write-ChangeLog $s.Url 'Prullenbak terugzetten' '' "FOUT: $($_.Exception.Message)"
            }
        }
    }
    finally { Stop-LongOp }
    Update-KindGrid $Kind
    $msg = "Prullenbak terugzetten: $ok teruggezet, $fail mislukt"
    Set-Status $msg $(if ($fail) { 'error' } else { 'success' })
    if ($fail) { Show-Msg "$msg.`n`nVeelvoorkomende oorzaak: er staat op de oorspronkelijke plek al een bestand met dezelfde naam. Zie de log voor details." 'Prullenbak terugzetten' 'OK' 'Warning' | Out-Null }
}

function Invoke-OdEmpty {
    $items = @(Get-SelectedOds)
    if (-not $items.Count) { Show-Msg 'Selecteer eerst de OneDrive(s) die je wilt leegmaken.' 'OneDrive leegmaken' | Out-Null; return }
    $active = @($items | Where-Object { $_.Status -eq 'Actief' })
    $cfg = Show-EmptyDialog -Items $items -ActiveCount $active.Count
    if (-not $cfg) { return }
    Start-LongOp 'OneDrive leegmaken...'
    $ok = 0; $fail = 0
    try {
        foreach ($s in $items) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            Set-Status "Leegmaken: $($s.Owner)"
            try {
                $n = Invoke-OnSite $s { Clear-OneDriveContentCore $s }
                $txt = "$n mappen/bestanden naar de prullenbak"
                if ($cfg.EmptyBins) {
                    Invoke-OnSite $s { Clear-SiteRecycleCore $s }
                    $txt += ', prullenbak definitief geleegd'
                    $s.RecycleMB = 0
                }
                Write-Log "$($s.Owner): $txt" 'success'
                Write-ChangeLog $s.Url 'OneDrive leegmaken' $s.OwnerUpn $txt
                $ok++
            }
            catch {
                Write-Log "$($s.Owner): $($_.Exception.Message)" 'error'
                Write-ChangeLog $s.Url 'OneDrive leegmaken' $s.OwnerUpn "FOUT: $($_.Exception.Message)"
                $fail++
            }
        }
    }
    finally { Stop-LongOp }
    Update-OdGrid
    Set-Status "OneDrive leegmaken: $ok gelukt, $fail mislukt. De gebruikte opslag in het overzicht wordt pas na een tijd bijgewerkt." $(if ($fail) { 'error' } else { 'success' })
}

function Invoke-OdAdmin {
    param([switch]$Remove)
    if (-not $script:HasAdmin) { Show-Msg 'Hiervoor is de rol SharePoint Administrator nodig.' 'Beheerder' | Out-Null; return }
    $what = if ($Remove) { 'Beheerder verwijderen' } else { 'Beheerder toevoegen' }
    $items = @(Get-Targets 'O' $what)
    if (-not $items.Count) { return }
    $ps = @(Show-PrincipalPicker -Kinds user -Title "$what ($($items.Count) OneDrive(s))")
    if (-not $ps.Count) { return }
    $names = ($ps | ForEach-Object { $_.Upn }) -join ', '
    $msg = if ($Remove) { "$names als beheerder VERWIJDEREN van $($items.Count) OneDrive(s)?" } else { "$names als beheerder TOEVOEGEN aan $($items.Count) OneDrive(s)?`n`nZij krijgen volledige toegang tot alle bestanden (bijv. manager bij uitdienst)." }
    if ((Show-Msg $msg $what 'YesNo' 'Warning') -ne 'Yes') { return }
    Start-LongOp "$what..."
    try {
        foreach ($s in $items) {
            if ($script:StopRequested) { break }
            foreach ($p in $ps) {
                if ($Remove -and $p.Upn -eq $s.OwnerUpn) { Write-Log "$($s.Owner): eigenaar wordt niet verwijderd" 'warning'; continue }
                try {
                    Set-TenantSiteAdmin -SiteUrl $s.Url -Login "i:0#.f|membership|$($p.Upn)" -IsAdmin (-not $Remove)
                    Write-Log "$($s.Owner): $($p.Upn) $(if ($Remove) { 'verwijderd als' } else { 'toegevoegd als' }) beheerder" 'success'
                    Write-ChangeLog $s.Url $what $p.Upn 'OK'
                }
                catch {
                    Write-Log "$($s.Owner): $($p.Upn) -> $($_.Exception.Message)" 'error'
                    Write-ChangeLog $s.Url $what $p.Upn "FOUT: $($_.Exception.Message)"
                }
            }
        }
    }
    finally { Stop-LongOp }
    $cur = $script:Ctx['O'].Site
    if ($cur -and @($items | Where-Object { $_.Url -eq $cur.Url }).Count) { Show-SiteDetail 'O' $cur }
}

function Invoke-PermReport {
    $sites = @(Get-Targets 'S' 'Machtigingenrapport')
    if (-not $sites.Count) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (puntkomma)|*.csv'
    $dlg.FileName = "SharePoint-machtigingen-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
    if (-not $dlg.ShowDialog($script:Window)) { return }
    Start-LongOp 'Machtigingenrapport maken...'
    $out = New-Object System.Collections.ArrayList
    $n = 0
    try {
        foreach ($s in $sites) {
            if ($script:StopRequested) { break }
            $n++
            Set-Status "Rapport - site $n van $($sites.Count): $($s.Title)"
            try {
                foreach ($r in @(Get-SitePermissions $s)) {
                    $g = ConvertTo-PermGridRow $r 0 -WithSite
                    [void]$out.Add([pscustomobject][ordered]@{
                        Site = $s.Title; Url = $s.Url; SiteType = $s.Type; Bereik = $g.Bereik; Bron = $g.Bron; Rechten = $g.Rechten
                        Naam = $g.Naam; Soort = $g.Soort; Lidmaatschap = $g.Lidmaatschap; Login = $g.Login
                    })
                }
            }
            catch {
                [void]$out.Add([pscustomobject][ordered]@{ Site = $s.Title; Url = $s.Url; SiteType = $s.Type; Bereik = ''; Bron = 'FOUT'; Rechten = ''; Naam = $_.Exception.Message; Soort = ''; Lidmaatschap = ''; Login = '' })
                Write-Log "$($s.Url): $($_.Exception.Message)" 'error'
            }
        }
        $out | Export-Csv -Path $dlg.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Write-Log "Rapport opgeslagen: $($dlg.FileName) ($($out.Count) regels)" 'success'
        Set-Status "Rapport opgeslagen ($($out.Count) regels)" 'success'
    }
    finally { Stop-LongOp }
}

function Invoke-ElevateRevoke {
    param([switch]$NoConfirm)
    $list = @($script:Elevated)
    if (-not $list.Count) { return }
    if (-not $NoConfirm -and (Show-Msg "Je tijdelijke sitebeheerrechten intrekken op $($list.Count) site(s)/OneDrive(s)?" 'Intrekken' 'YesNo' 'Question') -ne 'Yes') { return }
    Start-LongOp 'Tijdelijke beheerrechten intrekken...'
    try {
        foreach ($u in $list) {
            try {
                Set-TenantSiteAdmin -SiteUrl $u -Login "i:0#.f|membership|$($script:MeUpn)" -IsAdmin $false
                [void]$script:Elevated.Remove($u)
                Write-Log "Beheerrechten ingetrokken: $u" 'success'
                Write-ChangeLog $u 'Tijdelijk sitebeheerder ingetrokken (zelf)' $script:MeUpn 'OK'
            }
            catch { Write-Log "Intrekken op $u mislukt: $($_.Exception.Message)" 'error' }
        }
    }
    finally { Save-Elevated; Stop-LongOp }
}

#endregion

#region --- Gebruiker-toegang en groepen-overzicht ---

function Invoke-AccessScan {
    if (-not $script:AccUser) { Show-Msg 'Kies eerst een gebruiker.' 'Gebruiker-toegang' | Out-Null; return }
    $sites = @(if ($script:ChkAccSelected.IsChecked) { Get-SelectedSites } else { $script:Sites })
    if ($script:ChkAccOneDrive.IsChecked) { $sites += @($script:OneDrives | Where-Object { $_.OwnerUpn -ne $script:AccUser.Upn }) }
    if (-not $sites.Count) { Show-Msg 'Geen sites (laad de sites of selecteer er een paar).' 'Gebruiker-toegang' | Out-Null; return }
    $u = $script:AccUser
    Start-LongOp "Toegang van $($u.Upn) zoeken..."
    $result = New-Object System.Collections.ArrayList
    try {
        $member = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($g in @(Get-GraphPaged "users/$($u.Id)/transitiveMemberOf/microsoft.graph.group?`$select=id&`$top=999" -Quiet)) { [void]$member.Add([string]$g.id) }
        $owned = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($g in @(Get-GraphPaged "users/$($u.Id)/ownedObjects/microsoft.graph.group?`$select=id&`$top=999" -Quiet)) { [void]$owned.Add([string]$g.id) }
        $isGuest = ($u.UserType -eq 'Guest' -or $u.Upn -match '#ext#')
        $n = 0
        foreach ($s in $sites) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            $n++
            Set-Status "Site $n van $($sites.Count): $($s.Title)  ($($result.Count) treffers)"
            try {
                foreach ($r in @(Get-SitePermissions $s)) {
                    $via = $null
                    if ($r.Upn -and $r.Upn -eq $u.Upn) { $via = 'Direct' }
                    elseif ($r.Kind -in 'secgroup', 'm365members' -and $member.Contains($r.ObjectId)) { $via = "Via groep: $($r.Name)" }
                    elseif ($r.Kind -eq 'm365owners' -and $owned.Contains($r.ObjectId)) { $via = "Via groep-eigenaar: $($r.Name)" }
                    elseif ($r.Kind -eq 'everyoneext' -and -not $isGuest) { $via = 'Iedereen behalve externe gebruikers' }
                    elseif ($r.Kind -eq 'everyone') { $via = 'Iedereen' }
                    if ($via) { [void]$result.Add(@{ Row = $r; Via = $via; IsOneDrive = [bool]$s.IsOneDrive }) }
                }
            }
            catch { Write-Log "$($s.Url): $($_.Exception.Message)" 'error' }
        }
    }
    catch { Write-Log "Toegang zoeken mislukt: $($_.Exception.Message)" 'error' }
    finally { Stop-LongOp }
    $script:AccRows = $result.ToArray()
    Update-AccessGrid
    Set-Status "$($result.Count) toegang(en) gevonden voor $($u.Upn)" 'success'
}

function Update-AccessGrid {
    $i = 0
    $rows = foreach ($a in $script:AccRows) {
        $r = $a.Row
        @{
            Idx = $i++; Site = $(if ($a.IsOneDrive) { "OneDrive: $($r.SiteTitle)" } else { $r.SiteTitle }); Url = $r.SiteUrl; Bereik = $r.Scope
            Via = $a.Via; Bron = $r.Source; Rechten = $r.Roles; Lidmaatschap = $(Get-GroupLabelById $r.ObjectId)
        }
    }
    Set-GridRows $script:AccGrid @('Idx:int', 'Site', 'Url', 'Bereik', 'Via', 'Bron', 'Rechten', 'Lidmaatschap') $rows
}

function Invoke-AccessRemove {
    $sel = @(Get-SelectedRows $script:AccGrid | ForEach-Object { $script:AccRows[[int]$_['Idx']] })
    $direct = @($sel | Where-Object { $_.Via -eq 'Direct' })
    if (-not $direct.Count) {
        Show-Msg "Selecteer regels met 'Direct'. Toegang via een groep pas je aan met 'Groep bewerken'; 'Iedereen' haal je weg via bulk 'Toegang intrekken' op tab Sites." 'Intrekken' | Out-Null
        return
    }
    $u = $script:AccUser
    if ((Show-Msg "Directe toegang van $($u.Upn) intrekken op $($direct.Count) plek(ken)?" 'Intrekken' 'YesNo' 'Warning') -ne 'Yes') { return }
    Start-LongOp 'Toegang intrekken...'
    $seen = @{}
    try {
        foreach ($a in $direct) {
            $r = $a.Row
            $k = "$($r.SiteUrl)|$(Get-RowKey $r)"
            if ($seen.ContainsKey($k)) { $a.Via = 'Verwijderd'; continue }
            $seen[$k] = $true
            $site = @($script:Sites) + @($script:OneDrives) | Where-Object { $_.Url -eq $r.SiteUrl } | Select-Object -First 1
            if (-not $site) { $site = @{ Url = $r.SiteUrl; Title = $r.SiteTitle } }
            try {
                Invoke-OnSite $site { Remove-PermRow $r }
                Write-Log "$($r.SiteTitle): $($u.Upn) verwijderd ($($r.Source), $($r.Scope))" 'success'
                Write-ChangeLog $r.SiteUrl "Verwijderen ($($r.Source), $($r.Scope))" $u.Upn 'OK'
                $a.Via = 'Verwijderd'
            }
            catch {
                Write-Log "$($r.SiteTitle): $($_.Exception.Message)" 'error'
                Write-ChangeLog $r.SiteUrl "Verwijderen ($($r.Source), $($r.Scope))" $u.Upn "FOUT: $($_.Exception.Message)"
            }
        }
    }
    finally { Stop-LongOp }
    Update-AccessGrid
}

function Invoke-GroupScan {
    $sites = @(Get-Targets 'S' 'Groepen scannen')
    if (-not $sites.Count) { return }
    Start-LongOp 'Groepen in sites zoeken...'
    $map = @{}
    $n = 0
    try {
        foreach ($s in $sites) {
            if ($script:StopRequested) { Write-Log 'Gestopt door gebruiker' 'warning'; break }
            $n++
            Set-Status "Site $n van $($sites.Count): $($s.Title)  ($($map.Count) groepen)"
            try {
                $rows = @(Get-SitePermissions $s)
                $pairs = @($rows | Where-Object { $_.ObjectId } | ForEach-Object { @{ Id = $_.ObjectId; Where = $(if ($_.Scope -eq 'Site') { $s.Title } else { "$($s.Title) ($($_.Scope))" }) } })
                if ($s.GroupId) { $pairs += @{ Id = $s.GroupId; Where = $s.Title } }
                foreach ($pr in $pairs) {
                    if (-not $map.ContainsKey($pr.Id)) { $map[$pr.Id] = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase) }
                    [void]$map[$pr.Id].Add($pr.Where)
                }
            }
            catch { Write-Log "$($s.Url): $($_.Exception.Message)" 'error' }
        }
    }
    finally { Stop-LongOp }
    $script:GroupRows = $map
    Update-GroupGrid
    Set-Status "$($map.Count) groepen gevonden in $n site(s)" 'success'
}

function Invoke-GroupSearch {
    $q = $script:TxtGrpSearch.Text.Trim() -replace '"', ''
    if (-not $q) { return }
    Enter-Busy 'Groepen zoeken...'
    try {
        $res = @(Get-GraphPaged "groups?`$search=`"displayName:$([uri]::EscapeDataString($q))`"&`$select=$($script:GroupSelect)&`$top=100" -Headers @{ ConsistencyLevel = 'eventual' } -Max 200)
        $map = @{}
        foreach ($g in $res) {
            $id = ([string]$g.id).ToLower()
            $script:GroupCache[$id] = $g
            $map[$id] = $(if ($script:GroupRows.ContainsKey($id)) { $script:GroupRows[$id] } else { $null })
        }
        $script:GroupRows = $map
        Update-GroupGrid
        Set-Status "$($res.Count) groepen gevonden" 'success'
    }
    catch { Write-Log "Zoeken mislukt: $($_.Exception.Message)" 'error' }
    finally { Exit-Busy }
}

function Update-GroupGrid {
    $rows = foreach ($id in $script:GroupRows.Keys) {
        $g = $script:GroupCache[$id]
        $sites = $script:GroupRows[$id]
        @{
            Id = $id; Naam = [string]$g.displayName; Soort = (Get-GroupKindLabel $g); Lidmaatschap = (Get-GroupMembershipLabel $g)
            Regel = [string]$g.membershipRule; Sites = $(if ($sites) { $sites.Count } else { $null })
            SiteLijst = $(if ($sites) { (@($sites) | Sort-Object) -join '; ' } else { '' })
        }
    }
    $rows = @($rows | Sort-Object { $_.Naam })
    Set-GridRows $script:GrpGrid @('Id', 'Naam', 'Soort', 'Lidmaatschap', 'Regel', 'Sites:int', 'SiteLijst') $rows
    Update-GroupFilter
}

function Update-GroupFilter {
    $view = $script:GrpGrid.ItemsSource
    if ($null -eq $view) { return }
    $f = [string]$script:CmbGrpFilter.SelectedItem
    $view.RowFilter = switch ($f) {
        'Alleen statisch'  { "Lidmaatschap = 'Statisch'" }
        'Alleen dynamisch' { "Lidmaatschap LIKE 'Dynamisch*'" }
        default            { '' }
    }
    $script:LblGrpCount.Text = "$($view.Count) groepen"
}

function Open-GroupEditor {
    param([string]$GroupId)
    if (-not $GroupId) { Show-Msg 'Selecteer een regel met een Entra- of M365-groep.' 'Groep bewerken' | Out-Null; return }
    $changed = Show-GroupEditor $GroupId
    if ($changed) {
        try { Resolve-Groups @($GroupId) -Force } catch { }
        foreach ($P in 'S', 'O') { if ($script:Ctx[$P].PermRows.Count) { Update-PermGrid $P; Update-LibPermGrid $P } }
        if ($script:AccRows.Count) { Update-AccessGrid }
        if ($script:GroupRows.Count) { Update-GroupGrid }
    }
}

#endregion

#region --- XAML: stijlen en dialogen ---

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
<Style x:Key="BtnGreen" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#107C10"/></Style>
<Style x:Key="BtnRed"   TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#C42B1C"/></Style>
<Style x:Key="BtnGray"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#7A7A7A"/></Style>
<Style x:Key="BtnOrange" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#B26200"/></Style>
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
<Style x:Key="CellNum" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="HorizontalAlignment" Value="Right"/>
</Style>
<Style x:Key="CellMember" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Style.Triggers>
    <Trigger Property="Text" Value="Statisch"><Setter Property="Foreground" Value="#0063B1"/></Trigger>
    <Trigger Property="Text" Value="Dynamisch"><Setter Property="Foreground" Value="#8A2BE2"/></Trigger>
    <Trigger Property="Text" Value="Dynamisch (gepauzeerd)"><Setter Property="Foreground" Value="#B26200"/></Trigger>
    <Trigger Property="Text" Value="Niet gevonden"><Setter Property="Foreground" Value="#C42B1C"/></Trigger>
  </Style.Triggers>
</Style>
<Style x:Key="CellScope" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="Foreground" Value="#B26200"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Style.Triggers>
    <Trigger Property="Text" Value="Site"><Setter Property="Foreground" Value="#555555"/><Setter Property="FontWeight" Value="Normal"/></Trigger>
  </Style.Triggers>
</Style>
<Style x:Key="CellStatus" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="Foreground" Value="#C42B1C"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Style.Triggers>
    <Trigger Property="Text" Value="Actief"><Setter Property="Foreground" Value="#107C10"/><Setter Property="FontWeight" Value="Normal"/></Trigger>
  </Style.Triggers>
</Style>
<Style TargetType="GroupBox">
  <Setter Property="Margin" Value="0,0,0,8"/>
  <Setter Property="Padding" Value="8,6"/>
  <Setter Property="BorderBrush" Value="#CCCCCC"/>
  <Setter Property="Background" Value="White"/>
</Style>
<Style TargetType="CheckBox"><Setter Property="Margin" Value="0,0,0,6"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
<Style TargetType="RadioButton"><Setter Property="Margin" Value="0,0,16,6"/></Style>
<Style x:Key="FieldLabel" TargetType="TextBlock">
  <Setter Property="VerticalAlignment" Value="Center"/>
  <Setter Property="Margin" Value="0,0,8,6"/>
  <Setter Property="Foreground" Value="#333333"/>
</Style>
<Style x:Key="FieldBox" TargetType="TextBox">
  <Setter Property="Height" Value="26"/>
  <Setter Property="VerticalContentAlignment" Value="Center"/>
  <Setter Property="Margin" Value="0,0,8,6"/>
  <Setter Property="Padding" Value="4,0"/>
</Style>
<Style x:Key="Hint" TargetType="TextBlock">
  <Setter Property="Foreground" Value="Gray"/>
  <Setter Property="TextWrapping" Value="Wrap"/>
  <Setter Property="Margin" Value="0,0,0,8"/>
</Style>
'@

function New-Dialog {
    param([string]$Xaml)
    $full = $Xaml.Replace('{{STYLES}}', $script:StylesXaml)
    $w = ConvertFrom-Xaml $full
    if ($script:Window -and $script:Window.IsLoaded) { $w.Owner = $script:Window }
    return @{ W = $w; C = (Get-XamlControls $w $Xaml) }
}

function Show-InputDialog {
    param([string]$Title, [string]$Prompt, [string]$Default = '')
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="480" SizeToContent="Height" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5" ResizeMode="NoResize">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <StackPanel Margin="14">
    <TextBlock x:Name="LblText" Style="{StaticResource Hint}" Foreground="#333333"/>
    <TextBox x:Name="TxtValue" Style="{StaticResource FieldBox}" Margin="0,0,0,12"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnOk" Content="OK" Style="{StaticResource BtnGreen}" IsDefault="True"/>
      <Button Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </StackPanel>
  </StackPanel>
</Window>
'@
    $w = $d.W; $c = $d.C
    $w.Title = $Title
    $c.LblText.Text = $Prompt
    $c.TxtValue.Text = $Default
    $c.BtnOk.Add_Click({ $w.DialogResult = $true })
    $w.Add_ContentRendered({ [void]$c.TxtValue.Focus(); $c.TxtValue.SelectAll() })
    if ($w.ShowDialog()) { return $c.TxtValue.Text.Trim() }
    return $null
}

function Show-VersionDialog {
    param([string]$Text)
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Versies opschonen" Width="520" SizeToContent="Height" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5" ResizeMode="NoResize">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <StackPanel Margin="14">
    <TextBlock x:Name="LblText" Style="{StaticResource Hint}" Foreground="#333333"/>
    <StackPanel Orientation="Horizontal">
      <TextBlock Text="Bewaar de laatste" Style="{StaticResource FieldLabel}"/>
      <TextBox x:Name="TxtKeep" Width="60" Text="0" Style="{StaticResource FieldBox}"/>
      <TextBlock Text="vorige versies per bestand (0 = alle oude versies weg)" Style="{StaticResource FieldLabel}"/>
    </StackPanel>
    <StackPanel Orientation="Horizontal">
      <CheckBox x:Name="ChkLimit" Content="Daarna versielimiet instellen op" IsChecked="True" Margin="0,0,8,6"/>
      <TextBox x:Name="TxtLimit" Width="60" Text="100" Style="{StaticResource FieldBox}"/>
      <TextBlock Text="hoofdversies (voorkomt dat het weer groeit)" Style="{StaticResource FieldLabel}"/>
    </StackPanel>
    <TextBlock Style="{StaticResource Hint}" Text="De huidige versie van elk bestand blijft altijd staan. Verwijderde versies gaan naar de prullenbak (2e fase); leeg die om de ruimte direct vrij te maken. Bij een bewaarbeleid (retention) kunnen versies niet worden verwijderd."/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnOk" Content="Opschonen" Style="{StaticResource BtnOrange}"/>
      <Button Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </StackPanel>
  </StackPanel>
</Window>
'@
    $w = $d.W; $c = $d.C
    $c.LblText.Text = $Text
    $st = @{ Keep = 0; Limit = 100 }
    $c.BtnOk.Add_Click({
        $k = 0; $l = 0
        if (-not [int]::TryParse($c.TxtKeep.Text, [ref]$k) -or $k -lt 0 -or $k -gt 500) { Show-Msg 'Aantal te bewaren versies: 0 - 500.' 'Versies' -Owner $w | Out-Null; return }
        if ($c.ChkLimit.IsChecked -and (-not [int]::TryParse($c.TxtLimit.Text, [ref]$l) -or $l -lt 1 -or $l -gt 50000)) { Show-Msg 'Versielimiet: 1 - 50000.' 'Versies' -Owner $w | Out-Null; return }
        $st.Keep = $k; $st.Limit = $l
        $w.DialogResult = $true
    })
    if (-not $w.ShowDialog()) { return $null }
    return @{ Keep = $st.Keep; SetLimit = [bool]$c.ChkLimit.IsChecked; Limit = $st.Limit }
}

function Show-EmptyDialog {
    param($Items, [int]$ActiveCount)
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="OneDrive leegmaken" Width="560" SizeToContent="Height" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5" ResizeMode="NoResize">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <StackPanel Margin="14">
    <TextBlock Text="OneDrive leegmaken" FontSize="16" FontWeight="SemiBold" Foreground="#C42B1C" Margin="0,0,0,6"/>
    <TextBlock x:Name="LblText" Style="{StaticResource Hint}" Foreground="#333333"/>
    <TextBox x:Name="TxtList" Height="110" IsReadOnly="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" Margin="0,0,0,8"/>
    <TextBlock x:Name="LblActive" Style="{StaticResource Hint}" Foreground="#C42B1C" FontWeight="SemiBold"/>
    <CheckBox x:Name="ChkBins" Content="Daarna de prullenbak (1e en 2e fase) DEFINITIEF legen - pas dan komt de ruimte vrij"/>
    <TextBlock Style="{StaticResource Hint}" Text="Alle mappen en bestanden in de documentbibliotheek gaan naar de prullenbak. Zonder het vinkje hierboven zijn ze nog 93 dagen terug te zetten. Een bewaarbeleid (retention/hold) kan verwijderen blokkeren."/>
    <TextBlock Text="Typ LEEGMAKEN om te bevestigen:" Style="{StaticResource FieldLabel}"/>
    <TextBox x:Name="TxtConfirm" Style="{StaticResource FieldBox}"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,6,0,0">
      <Button x:Name="BtnOk" Content="Leegmaken" Style="{StaticResource BtnRed}" IsEnabled="False"/>
      <Button Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </StackPanel>
  </StackPanel>
</Window>
'@
    $w = $d.W; $c = $d.C
    $c.LblText.Text = "De inhoud van $(@($Items).Count) OneDrive(s) wordt verwijderd:"
    $c.TxtList.Text = (@($Items) | ForEach-Object { "$($_.Owner)  ($($_.OwnerUpn))  -  $(ConvertTo-GB $_.UsedMB) GB  -  $($_.Status)" }) -join "`r`n"
    if ($ActiveCount) { $c.LblActive.Text = "LET OP: $ActiveCount van deze OneDrive(s) hoort bij een ACTIEF account!" } else { $c.LblActive.Visibility = 'Collapsed' }
    $c.TxtConfirm.Add_TextChanged({ $c.BtnOk.IsEnabled = ($c.TxtConfirm.Text.Trim() -ceq 'LEEGMAKEN') })
    $c.BtnOk.Add_Click({ $w.DialogResult = $true })
    if (-not $w.ShowDialog()) { return $null }
    return @{ EmptyBins = [bool]$c.ChkBins.IsChecked }
}

function Show-RestoreDialog {
    # Toont de prullenbak-items met filters; geeft de gekozen items terug. Handlers vuren tijdens ShowDialog().
    param($Rows, [int]$SiteCount)
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Prullenbak terugzetten" Width="1280" Height="760" MinWidth="900" MinHeight="500" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="LblIntro" Style="{StaticResource Hint}" Foreground="#333333"/>
    <WrapPanel Grid.Row="1">
      <TextBlock Text="Zoeken:" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,6,6"/>
      <TextBox x:Name="TxtFind" Width="200" Style="{StaticResource FieldBox}" Margin="0,0,12,6" ToolTip="Zoekt in naam, locatie en site"/>
      <TextBlock Text="Verwijderd door:" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,6,6"/>
      <ComboBox x:Name="CmbBy" Width="200" Height="26" Margin="0,0,12,6"/>
      <TextBlock Text="Verwijderd van:" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,6,6"/>
      <DatePicker x:Name="DpFrom" Width="120" Height="26" Margin="0,0,6,6"/>
      <TextBlock Text="t/m:" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,6,6"/>
      <DatePicker x:Name="DpTo" Width="120" Height="26" Margin="0,0,12,6"/>
    </WrapPanel>
    <WrapPanel Grid.Row="2">
      <TextBlock Text="Type:" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,6,6"/>
      <ComboBox x:Name="CmbType" Width="160" Height="26" Margin="0,0,12,6"/>
      <TextBlock Text="Fase:" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,6,6"/>
      <ComboBox x:Name="CmbStage" Width="130" Height="26" Margin="0,0,12,6"/>
      <TextBlock Text="Site:" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,6,6"/>
      <ComboBox x:Name="CmbSite" Width="220" Height="26" Margin="0,0,12,6"/>
      <Button x:Name="BtnToday" Content="Vandaag" Style="{StaticResource BtnGray}" Margin="0,0,6,6"/>
      <Button x:Name="BtnWeek" Content="Laatste 7 dagen" Style="{StaticResource BtnGray}" Margin="0,0,6,6"/>
      <Button x:Name="BtnReset" Content="Filters wissen" Style="{StaticResource BtnGray}" Margin="0,0,6,6"/>
    </WrapPanel>
    <DataGrid Grid.Row="3" x:Name="Grid" SelectionMode="Extended" Margin="0,4,0,8">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="230" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Oorspronkelijke locatie" Binding="{Binding Locatie}" Width="300" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="100" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Verwijderd op" Binding="{Binding Verwijderd, StringFormat=yyyy-MM-dd HH:mm}" Width="115" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Verwijderd door" Binding="{Binding Door}" Width="160" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Grootte (MB)" Binding="{Binding GrootteMB}" Width="85" ElementStyle="{StaticResource CellNum}"/>
        <DataGridTextColumn Header="Fase" Binding="{Binding Fase}" Width="60" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Gemaakt door" Binding="{Binding Auteur}" Width="140" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Site" Binding="{Binding Site}" Width="180" ElementStyle="{StaticResource CellText}"/>
      </DataGrid.Columns>
    </DataGrid>
    <TextBlock Grid.Row="4" x:Name="LblCount" FontWeight="SemiBold" Margin="0,0,0,6"/>
    <DockPanel Grid.Row="5">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="BtnSel" Content="Geselecteerde terugzetten" Style="{StaticResource BtnGreen}"/>
        <Button x:Name="BtnAll" Content="Alles in filter terugzetten" Style="{StaticResource BtnOrange}"/>
        <Button Content="Sluiten" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal">
        <Button x:Name="BtnSelAll" Content="Alles selecteren" Style="{StaticResource BtnGray}"/>
        <Button x:Name="BtnCsv" Content="Export CSV" Style="{StaticResource BtnGray}"/>
      </StackPanel>
    </DockPanel>
  </Grid>
</Window>
'@
    $w = $d.W; $c = $d.C
    $all = @($Rows)
    $t = New-Object System.Data.DataTable
    foreach ($col in @(@('Idx', [int]), @('Naam', [string]), @('Locatie', [string]), @('Type', [string]), @('Verwijderd', [datetime]),
                       @('Door', [string]), @('GrootteMB', [double]), @('Fase', [string]), @('Auteur', [string]), @('Site', [string]))) {
        [void]$t.Columns.Add($col[0], $col[1])
    }
    for ($i = 0; $i -lt $all.Count; $i++) {
        $r = $all[$i]
        $dr = $t.NewRow()
        $dr['Idx'] = $i; $dr['Naam'] = $r.Name; $dr['Locatie'] = $r.Dir; $dr['Type'] = $r.Type
        if ($r.Deleted -gt [datetime]::MinValue) { $dr['Verwijderd'] = $r.Deleted } else { $dr['Verwijderd'] = [DBNull]::Value }
        $dr['Door'] = $r.By; $dr['GrootteMB'] = $r.SizeMB; $dr['Fase'] = $r.Stage; $dr['Auteur'] = $r.Author; $dr['Site'] = $r.Site
        [void]$t.Rows.Add($dr)
    }
    $view = $t.DefaultView
    $view.Sort = 'Verwijderd DESC'
    $c.Grid.ItemsSource = $view
    $c.LblIntro.Text = "$($all.Count) item(s) in de prullenbak van $SiteCount site(s). Filter en selecteer wat terug moet. Items gaan terug naar hun oorspronkelijke locatie; verwijderde mappen worden eerst teruggezet. Staat er al een bestand met dezelfde naam, dan mislukt dat item."
    $c.CmbBy.ItemsSource = @('(iedereen)') + @($all | ForEach-Object { $_.By } | Where-Object { $_ } | Sort-Object -Unique)
    $c.CmbType.ItemsSource = @('(alle typen)') + @($all | ForEach-Object { $_.Type } | Sort-Object -Unique)
    $c.CmbStage.ItemsSource = @('(beide)', '1e fase', '2e fase')
    $c.CmbSite.ItemsSource = @('(alle sites)') + @($all | ForEach-Object { $_.Site } | Sort-Object -Unique)
    foreach ($cb in @($c.CmbBy, $c.CmbType, $c.CmbStage, $c.CmbSite)) { $cb.SelectedIndex = 0 }
    if (@($c.CmbSite.ItemsSource).Count -le 2) { $c.CmbSite.IsEnabled = $false }

    $st = @{ Result = @(); Busy = $false }
    $q = { param([string]$v) return $v.Replace("'", "''") }
    $inv = [cultureinfo]::InvariantCulture
    $size = { param([double]$MB) if ($MB -ge 1024) { "$([Math]::Round($MB / 1024, 2)) GB" } else { "$([Math]::Round($MB, 1)) MB" } }
    $count = {
        $n = 0; $mb = 0.0
        foreach ($rv in $view) { $n++; if ($rv.Row['GrootteMB'] -isnot [DBNull]) { $mb += [double]$rv.Row['GrootteMB'] } }
        $sel = $c.Grid.SelectedItems.Count
        $c.LblCount.Text = "$n van $($all.Count) item(s) in filter ($(& $size $mb))$(if ($sel) { "  |  $sel geselecteerd" })"
    }
    $apply = {
        if ($st.Busy) { return }
        $parts = @()
        $txt = $c.TxtFind.Text.Trim()
        if ($txt) { $l = ConvertTo-LikeLiteral $txt; $parts += "(Naam LIKE '%$l%' OR Locatie LIKE '%$l%' OR Site LIKE '%$l%')" }
        if ($c.CmbBy.SelectedIndex -gt 0) { $parts += "Door = '$(& $q ([string]$c.CmbBy.SelectedItem))'" }
        if ($c.CmbType.SelectedIndex -gt 0) { $parts += "Type = '$(& $q ([string]$c.CmbType.SelectedItem))'" }
        if ($c.CmbStage.SelectedIndex -gt 0) { $parts += "Fase = '$(& $q ([string]$c.CmbStage.SelectedItem))'" }
        if ($c.CmbSite.SelectedIndex -gt 0) { $parts += "Site = '$(& $q ([string]$c.CmbSite.SelectedItem))'" }
        if ($c.DpFrom.SelectedDate) { $parts += "Verwijderd >= #$(([datetime]$c.DpFrom.SelectedDate).Date.ToString('MM/dd/yyyy', $inv))#" }
        if ($c.DpTo.SelectedDate) { $parts += "Verwijderd < #$(([datetime]$c.DpTo.SelectedDate).Date.AddDays(1).ToString('MM/dd/yyyy', $inv))#" }
        $view.RowFilter = $parts -join ' AND '
        & $count
    }
    $c.TxtFind.Add_TextChanged($apply)
    foreach ($cb in @($c.CmbBy, $c.CmbType, $c.CmbStage, $c.CmbSite)) { $cb.Add_SelectionChanged($apply) }
    $c.DpFrom.Add_SelectedDateChanged($apply)
    $c.DpTo.Add_SelectedDateChanged($apply)
    $c.Grid.Add_SelectionChanged($count)
    $c.BtnToday.Add_Click({ $st.Busy = $true; $c.DpFrom.SelectedDate = (Get-Date).Date; $c.DpTo.SelectedDate = (Get-Date).Date; $st.Busy = $false; & $apply })
    $c.BtnWeek.Add_Click({ $st.Busy = $true; $c.DpFrom.SelectedDate = (Get-Date).Date.AddDays(-6); $c.DpTo.SelectedDate = (Get-Date).Date; $st.Busy = $false; & $apply })
    $c.BtnReset.Add_Click({
        $st.Busy = $true
        $c.TxtFind.Text = ''; $c.DpFrom.SelectedDate = $null; $c.DpTo.SelectedDate = $null
        foreach ($cb in @($c.CmbBy, $c.CmbType, $c.CmbStage, $c.CmbSite)) { $cb.SelectedIndex = 0 }
        $st.Busy = $false
        & $apply
    })
    $c.BtnSelAll.Add_Click({ $c.Grid.SelectAll(); [void]$c.Grid.Focus() })
    $c.BtnCsv.Add_Click({ Export-GridCsv $c.Grid 'Prullenbak' })
    $finish = {
        param([object[]]$Picked, [string]$What)
        if (-not $Picked.Count) { Show-Msg 'Er zijn geen items om terug te zetten.' 'Terugzetten' -Owner $w | Out-Null; return }
        $mb = 0.0; foreach ($p in $Picked) { $mb += [double]$p.SizeMB }
        $sites = @($Picked | ForEach-Object { $_.SiteKey } | Sort-Object -Unique).Count
        if ((Show-Msg "$What`: $($Picked.Count) item(s) ($(& $size $mb)) op $sites site(s) terugzetten naar de oorspronkelijke locatie?" 'Terugzetten' 'YesNo' 'Question' -Owner $w) -ne 'Yes') { return }
        $st.Result = $Picked
        $w.DialogResult = $true
    }
    $c.BtnSel.Add_Click({
        $picked = @($c.Grid.SelectedItems | ForEach-Object { $all[[int]$_.Row['Idx']] })
        if (-not $picked.Count) { Show-Msg 'Selecteer eerst items (Ctrl/Shift-klik of Alles selecteren), of gebruik Alles in filter terugzetten.' 'Terugzetten' -Owner $w | Out-Null; return }
        & $finish $picked 'Geselecteerde items'
    })
    $c.BtnAll.Add_Click({ & $finish @(foreach ($rv in $view) { $all[[int]$rv.Row['Idx']] }) 'Alles in het filter' })
    & $count
    if (-not $w.ShowDialog()) { return @() }
    return $st.Result
}

function Show-PrincipalPicker {
    # Zoekt gebruikers en/of groepen. Handlers vuren tijdens ShowDialog(), dus lokale variabelen werken.
    param([string[]]$Kinds = @('user', 'group'), [string]$Title = 'Kiezen', [switch]$Single, $Owner)
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="640" Height="540" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <DockPanel Grid.Row="0">
      <Button x:Name="BtnFind" DockPanel.Dock="Right" Content="Zoeken" IsDefault="True" Margin="0,0,0,6"/>
      <ComboBox x:Name="CmbKind" DockPanel.Dock="Right" Width="150" Height="26" Margin="0,0,8,6"/>
      <TextBox x:Name="TxtFind" Style="{StaticResource FieldBox}"/>
    </DockPanel>
    <TextBlock Grid.Row="1" x:Name="LblHint" Style="{StaticResource Hint}" Text="Zoek op (begin van) naam, UPN of e-mail."/>
    <ListBox Grid.Row="2" x:Name="LstRes" Margin="0,0,0,10"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnOk" Content="Kiezen" Style="{StaticResource BtnGreen}"/>
      <Button x:Name="BtnCancel" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </StackPanel>
  </Grid>
</Window>
'@
    $w = $d.W; $c = $d.C
    if ($Owner) { $w.Owner = $Owner }
    $w.Title = $Title
    $c.LstRes.SelectionMode = $(if ($Single) { 'Single' } else { 'Extended' })
    $kindOpts = @()
    if ($Kinds -contains 'user' -and $Kinds -contains 'group') { $kindOpts += 'Gebruikers en groepen' }
    if ($Kinds -contains 'user') { $kindOpts += 'Alleen gebruikers' }
    if ($Kinds -contains 'group') { $kindOpts += 'Alleen groepen' }
    $c.CmbKind.ItemsSource = $kindOpts
    $c.CmbKind.SelectedIndex = 0
    $find = {
        $q = $c.TxtFind.Text.Trim() -replace '"', ''
        if (-not $q) { return }
        $c.LstRes.Items.Clear()
        $k = [string]$c.CmbKind.SelectedItem
        $e = [uri]::EscapeDataString($q)
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            if ($k -ne 'Alleen groepen') {
                $res = Invoke-Graph "users?`$search=`"displayName:$e`" OR `"userPrincipalName:$e`" OR `"mail:$e`"&`$select=id,displayName,userPrincipalName,mail,userType&`$top=30" -Headers @{ ConsistencyLevel = 'eventual' } -Quiet
                foreach ($u in @($res.value)) {
                    $it = New-Object System.Windows.Controls.ListBoxItem
                    $guest = if ($u.userType -eq 'Guest') { ' [gast]' } else { '' }
                    $it.Content = "[Gebruiker] $($u.displayName)  -  $($u.userPrincipalName)$guest"
                    $it.Tag = @{ Kind = 'user'; Id = [string]$u.id; Name = [string]$u.displayName; Upn = [string]$u.userPrincipalName; UserType = [string]$u.userType }
                    [void]$c.LstRes.Items.Add($it)
                }
            }
            if ($k -ne 'Alleen gebruikers') {
                $res = Invoke-Graph "groups?`$search=`"displayName:$e`"&`$select=$($script:GroupSelect)&`$top=30" -Headers @{ ConsistencyLevel = 'eventual' } -Quiet
                foreach ($g in @($res.value)) {
                    $id = ([string]$g.id).ToLower()
                    $script:GroupCache[$id] = $g
                    $it = New-Object System.Windows.Controls.ListBoxItem
                    $it.Content = "[Groep - $(Get-GroupKindLabel $g), $(Get-GroupMembershipLabel $g)] $($g.displayName)"
                    $it.Tag = @{ Kind = 'group'; Id = $id; Name = [string]$g.displayName; Unified = (@($g.groupTypes) -contains 'Unified') }
                    [void]$c.LstRes.Items.Add($it)
                }
            }
            if (-not $c.LstRes.Items.Count) { $c.LblHint.Text = "Niets gevonden voor '$q'." } else { $c.LblHint.Text = "$($c.LstRes.Items.Count) resultaten." }
        }
        catch { $c.LblHint.Text = "Zoeken mislukt: $($_.Exception.Message)" }
        finally { $w.Cursor = $null }
    }
    $c.BtnFind.Add_Click($find)
    $c.LstRes.Add_MouseDoubleClick({ if ($c.LstRes.SelectedItems.Count) { $w.DialogResult = $true } })
    $c.BtnOk.Add_Click({ if ($c.LstRes.SelectedItems.Count) { $w.DialogResult = $true } })
    [void]$c.TxtFind.Focus()
    if ($w.ShowDialog()) { return @($c.LstRes.SelectedItems | ForEach-Object { $_.Tag }) }
    return @()
}

function Show-BulkDialog {
    param([ValidateSet('Add', 'Remove')][string]$Mode, [int]$SiteCount)
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="640" Height="600" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock x:Name="LblIntro" Grid.Row="0" Style="{StaticResource Hint}" Foreground="#333333"/>
    <GroupBox Grid.Row="1" Header="Wie">
      <DockPanel>
        <StackPanel DockPanel.Dock="Top" Orientation="Horizontal">
          <Button x:Name="BtnPick" Content="Gebruiker / groep zoeken..."/>
          <Button x:Name="BtnDel" Content="Weghalen" Style="{StaticResource BtnGray}"/>
        </StackPanel>
        <StackPanel x:Name="PnlSpecial" DockPanel.Dock="Bottom" Margin="0,8,0,0">
          <CheckBox x:Name="ChkEveryoneExt" Content="Iedereen behalve externe gebruikers"/>
          <CheckBox x:Name="ChkEveryone" Content="Iedereen (incl. gasten)"/>
        </StackPanel>
        <ListBox x:Name="LstWho" SelectionMode="Extended" MinHeight="120"/>
      </DockPanel>
    </GroupBox>
    <GroupBox Grid.Row="2" x:Name="GrpOptions" Header="Opties">
      <StackPanel>
        <StackPanel x:Name="PnlAdd">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="Rol:" Style="{StaticResource FieldLabel}"/>
            <ComboBox x:Name="CmbRole" Width="240" Height="26" Margin="0,0,0,6"/>
          </StackPanel>
          <CheckBox x:Name="ChkViaM365" IsChecked="True" Content="Bij M365-groepssites eigenaars en leden via de Microsoft 365-groep toevoegen (aanbevolen)"/>
          <TextBlock Style="{StaticResource Hint}" Text="Via de M365-groep krijgt iemand ook toegang tot Teams, Planner en de groepsmailbox. Groepen worden altijd aan de SharePoint-groep toegevoegd. Rechten op losse bibliotheken geef je via het tabblad Bibliotheken."/>
        </StackPanel>
        <StackPanel x:Name="PnlRemove">
          <CheckBox x:Name="ChkRemSp" IsChecked="True" Content="Uit SharePoint-groepen en directe rechten (site en alle bibliotheken/lijsten)"/>
          <CheckBox x:Name="ChkRemAdmin" IsChecked="True" Content="Als sitecollectiebeheerder"/>
          <CheckBox x:Name="ChkRemM365" IsChecked="True" Content="Uit de Microsoft 365-groep van de site (lid en eigenaar)"/>
        </StackPanel>
      </StackPanel>
    </GroupBox>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,6,0,0">
      <Button x:Name="BtnOk" Content="Uitvoeren" Style="{StaticResource BtnGreen}"/>
      <Button x:Name="BtnCancel" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </StackPanel>
  </Grid>
</Window>
'@
    $w = $d.W; $c = $d.C
    if ($Mode -eq 'Add') {
        $w.Title = 'Toegang geven'
        $c.LblIntro.Text = "Toegang geven op $SiteCount site(s)."
        $c.PnlRemove.Visibility = 'Collapsed'
        $c.PnlSpecial.Visibility = 'Collapsed'
        $c.CmbRole.ItemsSource = @($script:RoleLabels.Values)
        $c.CmbRole.SelectedIndex = 1
    }
    else {
        $w.Title = 'Toegang intrekken'
        $c.LblIntro.Text = "Toegang intrekken op $SiteCount site(s). Alleen de gekozen soorten toegang worden verwijderd."
        $c.PnlAdd.Visibility = 'Collapsed'
        $c.BtnOk.Style = $w.FindResource('BtnRed')
    }
    $c.BtnPick.Add_Click({
        foreach ($p in @(Show-PrincipalPicker -Title 'Gebruikers of groepen kiezen' -Owner $w)) {
            $dup = @($c.LstWho.Items | Where-Object { $_.Tag.Id -eq $p.Id }).Count
            if ($dup) { continue }
            $it = New-Object System.Windows.Controls.ListBoxItem
            $it.Content = $(if ($p.Kind -eq 'user') { "[Gebruiker] $($p.Name) - $($p.Upn)" } else { "[Groep - $(Get-GroupMembershipLabel $script:GroupCache[$p.Id])] $($p.Name)" })
            $it.Tag = $p
            [void]$c.LstWho.Items.Add($it)
        }
    })
    $c.BtnDel.Add_Click({ foreach ($it in @($c.LstWho.SelectedItems)) { $c.LstWho.Items.Remove($it) } })
    $c.BtnOk.Add_Click({
        $n = $c.LstWho.Items.Count + [int][bool]$c.ChkEveryoneExt.IsChecked + [int][bool]$c.ChkEveryone.IsChecked
        if (-not $n) { Show-Msg 'Kies minstens een gebruiker of groep.' 'Bulk' -Owner $w | Out-Null; return }
        $w.DialogResult = $true
    })
    if (-not $w.ShowDialog()) { return $null }
    $ps = @($c.LstWho.Items | ForEach-Object { $_.Tag })
    if ($Mode -eq 'Remove') {
        if ($c.ChkEveryoneExt.IsChecked) { $ps += @{ Kind = 'everyoneext'; Id = ''; Name = 'Iedereen behalve externe gebruikers' } }
        if ($c.ChkEveryone.IsChecked) { $ps += @{ Kind = 'everyone'; Id = ''; Name = 'Iedereen' } }
    }
    $role = @($script:RoleLabels.Keys)[[Math]::Max(0, $c.CmbRole.SelectedIndex)]
    return @{
        Principals = $ps; Role = $role; ViaM365 = [bool]$c.ChkViaM365.IsChecked
        Sp = [bool]$c.ChkRemSp.IsChecked; Admin = [bool]$c.ChkRemAdmin.IsChecked; M365 = [bool]$c.ChkRemM365.IsChecked
    }
}

function Show-ChoiceDialog {
    param([string]$Title, [string]$Text, $Items)
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="500" SizeToContent="Height" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5" ResizeMode="NoResize">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <StackPanel Margin="14">
    <TextBlock x:Name="LblText" Style="{StaticResource Hint}" Foreground="#333333"/>
    <ComboBox x:Name="CmbChoice" Height="26" Margin="0,0,0,12" DisplayMemberPath="Label"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnOk" Content="OK" Style="{StaticResource BtnGreen}" IsDefault="True"/>
      <Button Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </StackPanel>
  </StackPanel>
</Window>
'@
    $w = $d.W; $c = $d.C
    $w.Title = $Title
    $c.LblText.Text = $Text
    $c.CmbChoice.ItemsSource = @($Items)
    $c.CmbChoice.SelectedIndex = 0
    $c.BtnOk.Add_Click({ $w.DialogResult = $true })
    if ($w.ShowDialog()) { return $c.CmbChoice.SelectedItem }
    return $null
}

function Show-SettingsDialog {
    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Instellingen" Width="560" SizeToContent="Height" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5" ResizeMode="NoResize">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <StackPanel Margin="14">
    <TextBlock Style="{StaticResource Hint}" Text="SharePoint-login. Standaard de 'SharePoint Online Management Shell'-app met eigen inlogvenster. Lukt dat niet in een tenant, vul dan een eigen app-registratie in (public client, redirect http://localhost, gedelegeerde SharePoint-rechten AllSites.FullControl). Wordt actief bij de volgende login."/>
    <TextBlock Text="Client-ID (leeg = standaard)" Style="{StaticResource FieldLabel}"/>
    <TextBox x:Name="TxtClient" Style="{StaticResource FieldBox}"/>
    <TextBlock Text="Redirect-URI (leeg = standaard)" Style="{StaticResource FieldLabel}"/>
    <TextBox x:Name="TxtRedirect" Style="{StaticResource FieldBox}"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
      <Button x:Name="BtnOk" Content="Opslaan" Style="{StaticResource BtnGreen}" IsDefault="True"/>
      <Button Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </StackPanel>
  </StackPanel>
</Window>
'@
    $w = $d.W; $c = $d.C
    $s = Get-Settings
    $c.TxtClient.Text = [string]$s.SpClientId
    $c.TxtRedirect.Text = [string]$s.SpRedirectUri
    $c.BtnOk.Add_Click({ $w.DialogResult = $true })
    if ($w.ShowDialog()) {
        $s.SpClientId = $c.TxtClient.Text.Trim()
        $s.SpRedirectUri = $c.TxtRedirect.Text.Trim()
        Save-Settings
        Write-Log 'Instellingen opgeslagen (actief bij volgende login)' 'success'
    }
}

function Show-GroupEditor {
    # Geeft $true terug als er iets gewijzigd is
    param([string]$GroupId)
    try {
        $g = Invoke-Graph "groups/$GroupId`?`$select=$($script:GroupSelect)"
    }
    catch { Show-Msg "Groep ophalen mislukt: $($_.Exception.Message)" 'Groep' 'OK' 'Error' | Out-Null; return $false }

    $d = New-Dialog @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="860" Height="720" MinWidth="700" MinHeight="560" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="12" Background="#F5F5F5">
  <Window.Resources>{{STYLES}}</Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <StackPanel Grid.Row="0">
      <TextBlock x:Name="LblName" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,4"/>
      <TextBlock x:Name="LblInfo" Style="{StaticResource Hint}"/>
      <TextBlock x:Name="LblNote" Style="{StaticResource Hint}" Foreground="#B26200" FontWeight="SemiBold"/>
    </StackPanel>
    <GroupBox Grid.Row="1" Header="Lidmaatschap">
      <StackPanel>
        <StackPanel Orientation="Horizontal">
          <RadioButton x:Name="RbStatic" GroupName="mt" Content="Statisch (leden handmatig)"/>
          <RadioButton x:Name="RbDynamic" GroupName="mt" Content="Dynamisch (leden via regel)"/>
        </StackPanel>
        <StackPanel x:Name="PnlRule">
          <DockPanel Margin="0,0,0,6">
            <TextBlock DockPanel.Dock="Left" Text="Voorbeeld invoegen:" Style="{StaticResource FieldLabel}"/>
            <ComboBox x:Name="CmbExample" Height="26"/>
          </DockPanel>
          <TextBox x:Name="TxtRule" Height="70" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" Margin="0,0,0,6"/>
          <CheckBox x:Name="ChkPaused" Content="Verwerking van de regel gepauzeerd"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal">
          <Button x:Name="BtnSaveType" Content="Lidmaatschap / regel opslaan" Style="{StaticResource BtnGreen}"/>
          <TextBlock x:Name="LblTypeHint" Style="{StaticResource Hint}" VerticalAlignment="Center" Margin="4,0,0,6"/>
        </StackPanel>
      </StackPanel>
    </GroupBox>
    <GroupBox Grid.Row="2" Header="Leden en eigenaars">
      <DockPanel>
        <WrapPanel DockPanel.Dock="Top">
          <Button x:Name="BtnAddMember" Content="Lid toevoegen..."/>
          <Button x:Name="BtnAddOwner" Content="Eigenaar toevoegen..."/>
          <Button x:Name="BtnRemove" Content="Geselecteerde verwijderen" Style="{StaticResource BtnRed}"/>
          <Button x:Name="BtnReload" Content="Vernieuwen" Style="{StaticResource BtnGray}"/>
          <TextBlock x:Name="LblCount" VerticalAlignment="Center" Margin="6,0,0,6" Foreground="Gray"/>
        </WrapPanel>
        <DataGrid x:Name="MemGrid">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Rol" Binding="{Binding Rol}" Width="90" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="230" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="UPN / e-mail" Binding="{Binding Upn}" Width="*" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="110" ElementStyle="{StaticResource CellText}"/>
          </DataGrid.Columns>
        </DataGrid>
      </DockPanel>
    </GroupBox>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,6,0,0">
      <Button Content="Sluiten" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0"/>
    </StackPanel>
  </Grid>
</Window>
'@
    $w = $d.W; $c = $d.C
    $st = @{ G = $g; Changed = $false; Members = @() }

    $refresh = {
        $gg = $st.G
        $isUnified = @($gg.groupTypes) -contains 'Unified'
        $isDyn = @($gg.groupTypes) -contains 'DynamicMembership'
        $synced = $gg.onPremisesSyncEnabled -eq $true
        $mailSec = $gg.mailEnabled -and $gg.securityEnabled -and -not $isUnified
        $dist = $gg.mailEnabled -and -not $gg.securityEnabled -and -not $isUnified
        $w.Title = "Groep bewerken - $($gg.displayName)"
        $c.LblName.Text = [string]$gg.displayName
        $info = "$(Get-GroupKindLabel $gg)  |  $(Get-GroupMembershipLabel $gg)"
        if ($gg.mail) { $info += "  |  $($gg.mail)" }
        if ($gg.description) { $info += "`n$($gg.description)" }
        $c.LblInfo.Text = $info
        $note = ''
        if ($synced) { $note = 'Deze groep komt uit de lokale AD (gesynchroniseerd) - aanpassen in de on-prem AD.' }
        elseif ($mailSec -or $dist) { $note = 'E-mail-beveiligingsgroepen en distributielijsten beheer je in Exchange (lidmaatschap kan niet via Graph).' }
        elseif ($isDyn) { $note = 'Dynamische groep: leden worden bepaald door de regel; alleen eigenaars zijn handmatig aan te passen.' }
        $c.LblNote.Text = $note
        $c.LblNote.Visibility = $(if ($note) { 'Visible' } else { 'Collapsed' })
        $canConvert = -not $synced -and -not $mailSec -and -not $dist
        $c.RbStatic.IsChecked = -not $isDyn
        $c.RbDynamic.IsChecked = $isDyn
        $c.RbStatic.IsEnabled = $canConvert; $c.RbDynamic.IsEnabled = $canConvert
        $c.TxtRule.Text = [string]$gg.membershipRule
        $c.ChkPaused.IsChecked = ($gg.membershipRuleProcessingState -eq 'Paused')
        $c.PnlRule.IsEnabled = $canConvert
        $c.BtnSaveType.IsEnabled = $canConvert
        $editMembers = -not $synced -and -not $mailSec -and -not $dist
        $c.BtnAddMember.IsEnabled = $editMembers -and -not $isDyn
        $c.BtnAddOwner.IsEnabled = $editMembers
        $c.BtnRemove.IsEnabled = $editMembers
        $c.PnlRule.Visibility = $(if ($c.RbDynamic.IsChecked) { 'Visible' } else { 'Collapsed' })
    }
    $loadMembers = {
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $list = New-Object System.Collections.ArrayList
            $sel = 'id,displayName,userPrincipalName,mail'
            foreach ($o in @(Get-GraphPaged "groups/$GroupId/owners?`$select=$sel&`$top=999" -Quiet)) { [void]$list.Add(@{ Rol = 'Eigenaar'; Obj = $o }) }
            foreach ($o in @(Get-GraphPaged "groups/$GroupId/members?`$select=$sel&`$top=999" -Quiet)) { [void]$list.Add(@{ Rol = 'Lid'; Obj = $o }) }
            $st.Members = $list.ToArray()
            $i = 0
            $rows = foreach ($m in $st.Members) {
                $o = $m.Obj
                $type = switch -regex ([string]$o.'@odata.type') { 'user$' { 'Gebruiker' } 'group$' { 'Groep' } 'device$' { 'Apparaat' } 'servicePrincipal$' { 'App' } default { '' } }
                @{ Idx = $i++; Rol = $m.Rol; Naam = [string]$o.displayName; Upn = [string]$(if ($o.userPrincipalName) { $o.userPrincipalName } else { $o.mail }); Type = $type }
            }
            Set-GridRows $c.MemGrid @('Idx:int', 'Rol', 'Naam', 'Upn', 'Type') $rows
            $nO = @($st.Members | Where-Object { $_.Rol -eq 'Eigenaar' }).Count
            $c.LblCount.Text = "$($st.Members.Count - $nO) leden, $nO eigenaars"
        }
        catch { $c.LblCount.Text = "Leden laden mislukt: $($_.Exception.Message)" }
        finally { $w.Cursor = $null }
    }
    $addRef = {
        param([string]$Rel)
        $isUnified = @($st.G.groupTypes) -contains 'Unified'
        $kinds = @(if ($Rel -eq 'members' -and -not $isUnified) { 'user', 'group' } else { 'user' })
        $picked = @(Show-PrincipalPicker -Kinds $kinds -Title $(if ($Rel -eq 'owners') { 'Eigenaar toevoegen' } else { 'Lid toevoegen' }) -Owner $w)
        if (-not $picked.Count) { return }
        foreach ($p in $picked) {
            try {
                $new = Add-GraphRef $GroupId $p.Id $Rel
                if ($Rel -eq 'owners' -and $isUnified) { [void](Add-GraphRef $GroupId $p.Id 'members') }
                $who = if ($p.Upn) { $p.Upn } else { $p.Name }
                Write-Log "$($st.G.displayName): $who toegevoegd als $(if ($Rel -eq 'owners') { 'eigenaar' } else { 'lid' })$(if (-not $new) { ' (was al)' })" 'success'
                Write-ChangeLog "groep:$($st.G.displayName)" "Toevoegen $Rel" $who 'OK'
                $st.Changed = $true
            }
            catch {
                Write-Log "Toevoegen mislukt: $($_.Exception.Message)" 'error'
                Show-Msg "Toevoegen mislukt: $($_.Exception.Message)" 'Groep' 'OK' 'Error' -Owner $w | Out-Null
            }
        }
        & $loadMembers
    }

    $c.CmbExample.ItemsSource = $script:RuleExamples
    $c.CmbExample.Add_SelectionChanged({
        $ex = [string]$c.CmbExample.SelectedItem
        if (-not $ex) { return }
        if ($c.TxtRule.Text.Trim()) { $c.TxtRule.Text = "$($c.TxtRule.Text.Trim()) -and $ex" } else { $c.TxtRule.Text = $ex }
        $c.CmbExample.SelectedIndex = -1
    })
    $c.RbStatic.Add_Checked({ $c.PnlRule.Visibility = 'Collapsed' })
    $c.RbDynamic.Add_Checked({ $c.PnlRule.Visibility = 'Visible' })
    $c.BtnAddMember.Add_Click({ & $addRef 'members' })
    $c.BtnAddOwner.Add_Click({ & $addRef 'owners' })
    $c.BtnReload.Add_Click({ & $loadMembers })
    $c.BtnRemove.Add_Click({
        $sel = @(Get-SelectedRows $c.MemGrid | ForEach-Object { $st.Members[[int]$_['Idx']] })
        if (-not $sel.Count) { return }
        $isDyn = @($st.G.groupTypes) -contains 'DynamicMembership'
        if ($isDyn -and @($sel | Where-Object { $_.Rol -eq 'Lid' }).Count) {
            Show-Msg 'Leden van een dynamische groep kun je niet handmatig verwijderen (alleen eigenaars). Pas de regel aan.' 'Groep' -Owner $w | Out-Null
            $sel = @($sel | Where-Object { $_.Rol -eq 'Eigenaar' })
            if (-not $sel.Count) { return }
        }
        if ((Show-Msg "$($sel.Count) geselecteerde verwijderen uit '$($st.G.displayName)'?" 'Verwijderen' 'YesNo' 'Warning' -Owner $w) -ne 'Yes') { return }
        foreach ($m in $sel) {
            $rel = if ($m.Rol -eq 'Eigenaar') { 'owners' } else { 'members' }
            $who = if ($m.Obj.userPrincipalName) { $m.Obj.userPrincipalName } else { $m.Obj.displayName }
            try {
                [void](Remove-GraphRef $GroupId $m.Obj.id $rel)
                Write-Log "$($st.G.displayName): $who verwijderd ($rel)" 'success'
                Write-ChangeLog "groep:$($st.G.displayName)" "Verwijderen $rel" $who 'OK'
                $st.Changed = $true
            }
            catch {
                Write-Log "Verwijderen $who mislukt: $($_.Exception.Message)" 'error'
                Show-Msg "Verwijderen van $who mislukt: $($_.Exception.Message)" 'Groep' 'OK' 'Error' -Owner $w | Out-Null
            }
        }
        & $loadMembers
    })
    $c.BtnSaveType.Add_Click({
        $gg = $st.G
        $wasDyn = @($gg.groupTypes) -contains 'DynamicMembership'
        $toDyn = [bool]$c.RbDynamic.IsChecked
        $types = @(@($gg.groupTypes) | Where-Object { $_ -and $_ -ne 'DynamicMembership' })
        $body = @{}
        if ($toDyn) {
            $rule = $c.TxtRule.Text.Trim()
            if (-not $rule) { Show-Msg 'Vul een lidmaatschapsregel in.' 'Regel' -Owner $w | Out-Null; return }
            $types += 'DynamicMembership'
            $body = @{ groupTypes = @($types); membershipRule = $rule; membershipRuleProcessingState = $(if ($c.ChkPaused.IsChecked) { 'Paused' } else { 'On' }) }
            $msg = if (-not $wasDyn) { "Groep omzetten naar DYNAMISCH?`n`nAlle huidige handmatige leden worden vervangen door de leden die aan de regel voldoen. Dit kan toegang tot sites en Teams wijzigen.`n`nRegel:`n$rule" }
                   else { "Regel / verwerking opslaan?`n`n$rule" }
        }
        else {
            if (-not $wasDyn) { Show-Msg 'De groep is al statisch.' 'Lidmaatschap' -Owner $w | Out-Null; return }
            $body = @{ groupTypes = @($types) }
            $msg = "Groep omzetten naar STATISCH?`n`nDe huidige leden blijven lid, maar worden daarna niet meer automatisch bijgewerkt."
        }
        if ((Show-Msg $msg 'Lidmaatschap' 'YesNo' 'Warning' -Owner $w) -ne 'Yes') { return }
        try {
            [void](Invoke-Graph "groups/$GroupId" -Method PATCH -Body $body)
            Write-Log "$($gg.displayName): lidmaatschap -> $(if ($toDyn) { 'dynamisch' } else { 'statisch' })" 'success'
            Write-ChangeLog "groep:$($gg.displayName)" 'Lidmaatschap wijzigen' '' $(if ($toDyn) { "Dynamisch: $($body.membershipRule) ($($body.membershipRuleProcessingState))" } else { 'Statisch' })
            $st.Changed = $true
            Wait-Seconds 1
            $st.G = Invoke-Graph "groups/$GroupId`?`$select=$($script:GroupSelect)" -Quiet
            $script:GroupCache[$GroupId] = $st.G
            & $refresh
            & $loadMembers
            if ($toDyn) { $c.LblTypeHint.Text = 'Opgeslagen. Het verwerken van de regel kan enkele minuten duren.' } else { $c.LblTypeHint.Text = 'Opgeslagen.' }
        }
        catch {
            Write-Log "Lidmaatschap wijzigen mislukt: $($_.Exception.Message)" 'error'
            Show-Msg "Opslaan mislukt:`n$($_.Exception.Message)`n`n(Dynamische groepen vereisen Entra ID P1; controleer de regel-syntaxis.)" 'Lidmaatschap' 'OK' 'Error' -Owner $w | Out-Null
        }
    })

    & $refresh
    $w.Add_ContentRendered({ & $loadMembers })
    [void]$w.ShowDialog()
    return $st.Changed
}

#endregion

#region --- Hoofdvenster ---

function Get-DetailXaml {
    # Detailpaneel (machtigingen + bibliotheken) met prefix S_ (sites) of O_ (OneDrive)
    param([string]$P)
    $x = @'
<Grid>
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
  <TextBlock x:Name="{P}_LblTitle" Grid.Row="0" FontSize="14" FontWeight="SemiBold" Margin="0,2,0,4" Text="Selecteer een regel om de details te zien" TextTrimming="CharacterEllipsis"/>
  <TabControl Grid.Row="1">
    <TabItem Header=" Machtigingen ">
      <DockPanel Margin="4">
        <WrapPanel DockPanel.Dock="Top">
          <TextBlock x:Name="{P}_LblGroup" Style="{StaticResource Lbl}" Foreground="#444444" Margin="0,0,10,6" MaxWidth="600" TextTrimming="CharacterEllipsis"/>
          <Button x:Name="{P}_BtnM365Edit" Content="M365-groep bewerken" Style="{StaticResource ToolPurple}" IsEnabled="False"/>
          <Border Style="{StaticResource Sep}"/>
          <Button x:Name="{P}_BtnPermReload" Content="Vernieuwen" Style="{StaticResource ToolGray}"/>
          <Button x:Name="{P}_BtnPermEditGroup" Content="Groep bewerken" Style="{StaticResource ToolPurple}" ToolTip="Leden of dynamische regel van de geselecteerde Entra-/M365-groep"/>
          <Button x:Name="{P}_BtnPermRemove" Content="Geselecteerde verwijderen" Style="{StaticResource ToolRed}"/>
          <Button x:Name="{P}_BtnPermCsv" Content="Export CSV" Style="{StaticResource ToolGray}"/>
        </WrapPanel>
        <DataGrid x:Name="{P}_PermGrid">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Bereik" Binding="{Binding Bereik}" Width="190" ElementStyle="{StaticResource CellScope}"/>
            <DataGridTextColumn Header="Bron" Binding="{Binding Bron}" Width="230" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Rechten" Binding="{Binding Rechten}" Width="130" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="200" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Soort" Binding="{Binding Soort}" Width="210" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Lidmaatschap" Binding="{Binding Lidmaatschap}" Width="140" ElementStyle="{StaticResource CellMember}"/>
            <DataGridTextColumn Header="Login / UPN" Binding="{Binding Login}" Width="*" ElementStyle="{StaticResource CellText}"/>
          </DataGrid.Columns>
        </DataGrid>
      </DockPanel>
    </TabItem>
    <TabItem Header=" Bibliotheken en lijsten ">
      <Grid Margin="4">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="60"/><RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="60"/></Grid.RowDefinitions>
        <WrapPanel Grid.Row="0">
          <Button x:Name="{P}_BtnLibBreak" Content="Eigen rechten instellen" Style="{StaticResource ToolOrange}" ToolTip="Overname van de site verbreken (huidige rechten worden gekopieerd)"/>
          <Button x:Name="{P}_BtnLibReset" Content="Overname herstellen" Style="{StaticResource ToolGray}" ToolTip="Weer de rechten van de site gebruiken"/>
          <Button x:Name="{P}_BtnLibGrant" Content="Toegang geven..." Style="{StaticResource ToolGreen}"/>
          <Button x:Name="{P}_BtnLibPermRemove" Content="Geselecteerde rechten verwijderen" Style="{StaticResource ToolRed}"/>
          <Border Style="{StaticResource Sep}"/>
          <Button x:Name="{P}_BtnLibVerLimit" Content="Versielimiet..." Style="{StaticResource Tool}"/>
          <Button x:Name="{P}_BtnLibVerClean" Content="Versies opschonen..." Style="{StaticResource ToolOrange}"/>
          <Button x:Name="{P}_BtnLibOpen" Content="Openen" Style="{StaticResource ToolGray}"/>
          <Button x:Name="{P}_BtnLibCsv" Content="Export CSV" Style="{StaticResource ToolGray}"/>
        </WrapPanel>
        <DataGrid x:Name="{P}_LibGrid" Grid.Row="1">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Titel" Binding="{Binding Titel}" Width="200" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Soort" Binding="{Binding Soort}" Width="85" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Items" Binding="{Binding Items}" Width="65" ElementStyle="{StaticResource CellNum}"/>
            <DataGridTextColumn Header="Totaal (MB)" Binding="{Binding GrootteMB}" Width="85" ElementStyle="{StaticResource CellNum}"/>
            <DataGridTextColumn Header="Bestanden (MB)" Binding="{Binding BestandenMB}" Width="100" ElementStyle="{StaticResource CellNum}"/>
            <DataGridTextColumn Header="Versies (MB)" Binding="{Binding VersiesMB}" Width="90" ElementStyle="{StaticResource CellNum}"/>
            <DataGridTextColumn Header="Versiebeheer" Binding="{Binding Versiebeheer}" Width="110" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Rechten" Binding="{Binding Uniek}" Width="130" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Laatst gewijzigd" Binding="{Binding Gewijzigd}" Width="105" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Pad" Binding="{Binding Pad}" Width="*" ElementStyle="{StaticResource CellText}"/>
          </DataGrid.Columns>
        </DataGrid>
        <TextBlock x:Name="{P}_LblLibPerm" Grid.Row="2" FontWeight="SemiBold" Margin="0,6,0,4"/>
        <DataGrid x:Name="{P}_LibPermGrid" Grid.Row="3">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Bron" Binding="{Binding Bron}" Width="230" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Rechten" Binding="{Binding Rechten}" Width="130" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="200" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Soort" Binding="{Binding Soort}" Width="210" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Lidmaatschap" Binding="{Binding Lidmaatschap}" Width="140" ElementStyle="{StaticResource CellMember}"/>
            <DataGridTextColumn Header="Login / UPN" Binding="{Binding Login}" Width="*" ElementStyle="{StaticResource CellText}"/>
          </DataGrid.Columns>
        </DataGrid>
      </Grid>
    </TabItem>
  </TabControl>
</Grid>
'@
    return $x.Replace('{P}', $P)
}

$script:MainXamlTemplate = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SharePoint Beheer" Width="1560" Height="980" MinWidth="1100" MinHeight="720"
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
<Style x:Key="ToolPurple" TargetType="Button" BasedOn="{StaticResource Tool}"><Setter Property="Background" Value="#6B2FA0"/></Style>
<Style x:Key="Sep" TargetType="Border"><Setter Property="Width" Value="1"/><Setter Property="Height" Value="22"/><Setter Property="Background" Value="#CCCCCC"/><Setter Property="Margin" Value="4,0,10,6"/></Style>
<Style x:Key="Lbl" TargetType="TextBlock"><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="Margin" Value="0,0,6,6"/></Style>
<Style x:Key="Bar" TargetType="Border"><Setter Property="Background" Value="#E8F1FB"/><Setter Property="CornerRadius" Value="4"/><Setter Property="Padding" Value="8,6,8,0"/><Setter Property="Margin" Value="0,0,0,6"/></Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="*" MinHeight="460"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="140" MinHeight="60"/>
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
          <TextBlock x:Name="LblConn" Text="Niet verbonden" Foreground="#FFC864" VerticalAlignment="Center" Margin="14,0,0,0" MaxWidth="280" TextTrimming="CharacterEllipsis"/>
          <Button x:Name="BtnConnect" Content="Inloggen" Style="{StaticResource BtnDark}"/>
          <Button x:Name="BtnDisconnect" Content="Uitloggen" Style="{StaticResource BtnDark}" IsEnabled="False"/>
          <Button x:Name="BtnSettings" Content="Instellingen" Style="{StaticResource BtnDark}"/>
        </StackPanel>
        <TextBlock Text="SharePoint Beheer" Foreground="White" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </DockPanel>
    </Border>

    <TabControl x:Name="MainArea" Grid.Row="1" Margin="6,6,6,0" IsEnabled="False">

      <TabItem Header="  Sites  ">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*" MinHeight="140"/>
            <RowDefinition Height="6"/>
            <RowDefinition Height="340" MinHeight="160"/>
          </Grid.RowDefinitions>
          <WrapPanel Grid.Row="0">
            <TextBlock Text="Zoeken:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtSiteFilter" Width="220" Style="{StaticResource FieldBox}" ToolTip="Zoekt in titel, URL en eigenaar"/>
            <ComboBox x:Name="CmbSiteType" Width="180" Height="26" Margin="0,0,10,6"/>
            <Button x:Name="BtnSitesLoad" Content="Sites laden" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnSiteScan" Content="Details scannen" Style="{StaticResource Tool}" ToolTip="Bibliotheken, unieke rechten, versies, prullenbak, quotum, delen en vergrendeling ophalen (selectie of alle zichtbare)"/>
            <Button x:Name="BtnSiteOpen" Content="Openen in browser" Style="{StaticResource ToolGray}"/>
            <Button x:Name="BtnSiteCsv" Content="Export CSV" Style="{StaticResource ToolGray}"/>
            <TextBlock x:Name="LblSiteCount" Style="{StaticResource Lbl}" Foreground="Gray" Margin="8,0,0,6"/>
          </WrapPanel>
          <Border Grid.Row="1" Style="{StaticResource Bar}">
            <WrapPanel>
              <TextBlock Text="Toegang:" FontWeight="SemiBold" Style="{StaticResource Lbl}"/>
              <Button x:Name="BtnBulkAdd" Content="Toegang geven..." Style="{StaticResource ToolGreen}"/>
              <Button x:Name="BtnBulkRemove" Content="Toegang intrekken..." Style="{StaticResource ToolRed}"/>
              <Button x:Name="BtnBulkSharing" Content="Extern delen..." Style="{StaticResource ToolOrange}"/>
              <Button x:Name="BtnBulkReport" Content="Machtigingenrapport (CSV)" Style="{StaticResource Tool}"/>
              <Button x:Name="BtnBulkGroups" Content="Groepen scannen" Style="{StaticResource ToolPurple}"/>
              <Border Style="{StaticResource Sep}"/>
              <TextBlock Text="Opslag:" FontWeight="SemiBold" Style="{StaticResource Lbl}"/>
              <Button x:Name="BtnSiteQuota" Content="Quotum..." Style="{StaticResource Tool}"/>
              <Button x:Name="BtnSiteLock" Content="Vergrendeling..." Style="{StaticResource ToolGray}"/>
              <Button x:Name="BtnSiteVersions" Content="Versies opschonen..." Style="{StaticResource ToolOrange}"/>
              <Button x:Name="BtnSiteRestore" Content="Prullenbak terugzetten..." Style="{StaticResource ToolGreen}" ToolTip="Items uit de prullenbak (1e en 2e fase) van de geselecteerde sites bekijken, filteren (wie, wanneer, type) en in bulk terugzetten"/>
              <Button x:Name="BtnSiteRecycle" Content="Prullenbak legen" Style="{StaticResource ToolRed}"/>
              <Border Style="{StaticResource Sep}"/>
              <CheckBox x:Name="ChkAutoElevate" Content="Mezelf automatisch sitebeheerder maken bij geen toegang" Margin="0,0,10,6"
                        ToolTip="SharePoint-beheerders zijn niet vanzelf sitebeheerder. De tool onthoudt waar dit gebeurde, zodat je het weer kunt intrekken."/>
              <Button x:Name="BtnElevRevoke" Content="Tijdelijke beheerrechten intrekken (0)" Style="{StaticResource ToolGray}" IsEnabled="False"/>
            </WrapPanel>
          </Border>
          <DataGrid x:Name="SiteGrid" Grid.Row="2" FrozenColumnCount="1">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Titel" Binding="{Binding Titel}" Width="200" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="130" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="M365" Binding="{Binding Groep}" Width="45" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Gebruikt (GB)" Binding="{Binding GebruiktGB}" Width="90" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Quotum (GB)" Binding="{Binding QuotumGB}" Width="85" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="% vol" Binding="{Binding Pct}" Width="55" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Bestanden" Binding="{Binding Bestanden}" Width="75" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Actief 30d" Binding="{Binding Actief}" Width="70" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Weergaven" Binding="{Binding Weergaven}" Width="75" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Laatste activiteit" Binding="{Binding Activiteit}" Width="105" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Eigenaar" Binding="{Binding Eigenaar}" Width="170" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Extern delen" Binding="{Binding Delen}" Width="130" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Vergrendeld" Binding="{Binding Lock}" Width="80" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Bibl." Binding="{Binding Bibliotheken}" Width="45" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Eigen rechten" Binding="{Binding Uniek}" Width="85" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Versies (GB)" Binding="{Binding VersiesGB}" Width="80" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Prullenbak (GB)" Binding="{Binding PrullenbakGB}" Width="95" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Te winnen (GB)" Binding="{Binding TeWinnenGB}" Width="95" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Aangemaakt" Binding="{Binding Aangemaakt}" Width="85" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="URL" Binding="{Binding Url}" Width="320" ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
          <GridSplitter Grid.Row="3" HorizontalAlignment="Stretch" Background="Transparent" ResizeDirection="Rows"/>
          <ContentControl Grid.Row="4">{{DETAIL_S}}</ContentControl>
        </Grid>
      </TabItem>

      <TabItem Header="  OneDrive  ">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*" MinHeight="140"/>
            <RowDefinition Height="6"/>
            <RowDefinition Height="320" MinHeight="160"/>
          </Grid.RowDefinitions>
          <WrapPanel Grid.Row="0">
            <TextBlock Text="Zoeken:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtOdFilter" Width="220" Style="{StaticResource FieldBox}" ToolTip="Zoekt in naam, UPN en URL"/>
            <ComboBox x:Name="CmbOdFilter" Width="190" Height="26" Margin="0,0,10,6"/>
            <Button x:Name="BtnOdLoad" Content="OneDrives laden" Style="{StaticResource Tool}"/>
            <Button x:Name="BtnOdScan" Content="Details scannen" Style="{StaticResource Tool}" ToolTip="Exacte grootte, quotum, versies, prullenbak en vergrendeling (selectie of alle zichtbare)"/>
            <Button x:Name="BtnOdOpen" Content="Openen in browser" Style="{StaticResource ToolGray}"/>
            <Button x:Name="BtnOdCsv" Content="Export CSV" Style="{StaticResource ToolGray}"/>
            <TextBlock x:Name="LblOdCount" Style="{StaticResource Lbl}" Foreground="Gray" Margin="8,0,0,6"/>
          </WrapPanel>
          <Border Grid.Row="1" Style="{StaticResource Bar}">
            <WrapPanel>
              <TextBlock Text="Toegang:" FontWeight="SemiBold" Style="{StaticResource Lbl}"/>
              <Button x:Name="BtnOdAdminAdd" Content="Beheerder toevoegen..." Style="{StaticResource ToolGreen}" ToolTip="Bijv. manager toegang geven tot de OneDrive van een vertrokken medewerker"/>
              <Button x:Name="BtnOdAdminRemove" Content="Beheerder verwijderen..." Style="{StaticResource ToolRed}"/>
              <Button x:Name="BtnOdSharing" Content="Extern delen..." Style="{StaticResource ToolOrange}"/>
              <Border Style="{StaticResource Sep}"/>
              <TextBlock Text="Opslag:" FontWeight="SemiBold" Style="{StaticResource Lbl}"/>
              <Button x:Name="BtnOdQuota" Content="Quotum..." Style="{StaticResource Tool}"/>
              <Button x:Name="BtnOdLock" Content="Vergrendeling..." Style="{StaticResource ToolGray}"/>
              <Button x:Name="BtnOdVersions" Content="Versies opschonen..." Style="{StaticResource ToolOrange}"/>
              <Button x:Name="BtnOdRestore" Content="Prullenbak terugzetten..." Style="{StaticResource ToolGreen}" ToolTip="Items uit de prullenbak (1e en 2e fase) van de geselecteerde OneDrives bekijken, filteren (wie, wanneer, type) en in bulk terugzetten"/>
              <Button x:Name="BtnOdRecycle" Content="Prullenbak legen" Style="{StaticResource ToolRed}"/>
              <Button x:Name="BtnOdEmpty" Content="OneDrive leegmaken..." Style="{StaticResource ToolRed}"/>
              <Border Style="{StaticResource Sep}"/>
              <CheckBox x:Name="ChkOdAutoElevate" Content="Mezelf automatisch beheerder maken bij geen toegang" Margin="0,0,10,6"/>
              <Button x:Name="BtnOdElevRevoke" Content="Tijdelijke beheerrechten intrekken (0)" Style="{StaticResource ToolGray}" IsEnabled="False"/>
            </WrapPanel>
          </Border>
          <DataGrid x:Name="OdGrid" Grid.Row="2" FrozenColumnCount="1">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Eigenaar" Binding="{Binding Eigenaar}" Width="190" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="UPN" Binding="{Binding Upn}" Width="230" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="160" ElementStyle="{StaticResource CellStatus}"/>
              <DataGridTextColumn Header="Gebruikt (GB)" Binding="{Binding GebruiktGB}" Width="90" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Quotum (GB)" Binding="{Binding QuotumGB}" Width="85" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="% vol" Binding="{Binding Pct}" Width="55" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Bestanden" Binding="{Binding Bestanden}" Width="75" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Actief 30d" Binding="{Binding Actief}" Width="70" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Laatste activiteit" Binding="{Binding Activiteit}" Width="105" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Versies (GB)" Binding="{Binding VersiesGB}" Width="80" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Prullenbak (GB)" Binding="{Binding PrullenbakGB}" Width="95" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Te winnen (GB)" Binding="{Binding TeWinnenGB}" Width="95" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Vergrendeld" Binding="{Binding Lock}" Width="80" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Extern delen" Binding="{Binding Delen}" Width="130" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="URL" Binding="{Binding Url}" Width="340" ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
          <GridSplitter Grid.Row="3" HorizontalAlignment="Stretch" Background="Transparent" ResizeDirection="Rows"/>
          <ContentControl Grid.Row="4">{{DETAIL_O}}</ContentControl>
        </Grid>
      </TabItem>

      <TabItem Header="  Gebruiker-toegang  ">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <WrapPanel Grid.Row="0">
            <Button x:Name="BtnAccPick" Content="Gebruiker kiezen..." Style="{StaticResource Tool}"/>
            <TextBlock x:Name="LblAccUser" Text="(geen gebruiker gekozen)" FontWeight="SemiBold" Style="{StaticResource Lbl}" Margin="4,0,14,6"/>
            <CheckBox x:Name="ChkAccSelected" Content="Alleen de geselecteerde sites (tab Sites)" Margin="0,0,14,6"/>
            <CheckBox x:Name="ChkAccOneDrive" Content="Ook OneDrives van anderen doorzoeken" Margin="0,0,14,6" ToolTip="Laad eerst de OneDrives. Kan lang duren bij veel gebruikers."/>
            <Button x:Name="BtnAccScan" Content="Toegang zoeken" Style="{StaticResource ToolGreen}"/>
            <Border Style="{StaticResource Sep}"/>
            <Button x:Name="BtnAccRemove" Content="Directe toegang intrekken" Style="{StaticResource ToolRed}"/>
            <Button x:Name="BtnAccGroupEdit" Content="Groep bewerken" Style="{StaticResource ToolPurple}"/>
            <Button x:Name="BtnAccCsv" Content="Export CSV" Style="{StaticResource ToolGray}"/>
          </WrapPanel>
          <TextBlock Grid.Row="1" Style="{StaticResource Hint}" Text="Zoekt per site naar directe rechten, SharePoint-groepen, sitebeheerders, Entra- en M365-groepen (ook geneste) en 'Iedereen' - op de site en op bibliotheken/lijsten met eigen rechten. Deellinks op losse bestanden/mappen worden niet getoond."/>
          <DataGrid x:Name="AccGrid" Grid.Row="2">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Site" Binding="{Binding Site}" Width="200" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Bereik" Binding="{Binding Bereik}" Width="180" ElementStyle="{StaticResource CellScope}"/>
              <DataGridTextColumn Header="Via" Binding="{Binding Via}" Width="240" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Bron" Binding="{Binding Bron}" Width="220" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Rechten" Binding="{Binding Rechten}" Width="130" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Lidmaatschap" Binding="{Binding Lidmaatschap}" Width="140" ElementStyle="{StaticResource CellMember}"/>
              <DataGridTextColumn Header="URL" Binding="{Binding Url}" Width="*" ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </TabItem>

      <TabItem Header="  Groepen  ">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <WrapPanel Grid.Row="0">
            <Button x:Name="BtnGrpScan" Content="Groepen in sites scannen" Style="{StaticResource ToolPurple}" ToolTip="Geselecteerde sites op tab Sites, of alle zichtbare sites"/>
            <Border Style="{StaticResource Sep}"/>
            <TextBlock Text="Of zoek een groep:" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtGrpSearch" Width="220" Style="{StaticResource FieldBox}"/>
            <Button x:Name="BtnGrpSearch" Content="Zoeken" Style="{StaticResource Tool}"/>
            <Border Style="{StaticResource Sep}"/>
            <ComboBox x:Name="CmbGrpFilter" Width="150" Height="26" Margin="0,0,10,6"/>
            <Button x:Name="BtnGrpEdit" Content="Groep bewerken" Style="{StaticResource ToolPurple}"/>
            <Button x:Name="BtnGrpCsv" Content="Export CSV" Style="{StaticResource ToolGray}"/>
            <TextBlock x:Name="LblGrpCount" Style="{StaticResource Lbl}" Foreground="Gray" Margin="8,0,0,6"/>
          </WrapPanel>
          <TextBlock Grid.Row="1" Style="{StaticResource Hint}" Text="Statisch = leden handmatig beheerd. Dynamisch = leden via een regel (Entra ID P1). Dubbelklik op een groep om leden, eigenaars of de regel aan te passen, of om om te zetten."/>
          <DataGrid x:Name="GrpGrid" Grid.Row="2">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="250" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Soort" Binding="{Binding Soort}" Width="160" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Lidmaatschap" Binding="{Binding Lidmaatschap}" Width="150" ElementStyle="{StaticResource CellMember}"/>
              <DataGridTextColumn Header="Regel" Binding="{Binding Regel}" Width="*" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Sites" Binding="{Binding Sites}" Width="60" ElementStyle="{StaticResource CellNum}"/>
              <DataGridTextColumn Header="Gebruikt in" Binding="{Binding SiteLijst}" Width="320" ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </TabItem>
    </TabControl>

    <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" Background="#3C3C3C" ResizeDirection="Rows"/>

    <DockPanel Grid.Row="3" Background="#1E1E1E">
      <Border DockPanel.Dock="Top" Background="#2D2D2D" Height="30">
        <DockPanel>
          <Button x:Name="BtnClearLog" DockPanel.Dock="Right" Content="Leegmaken" Style="{StaticResource BtnLog}"/>
          <Button x:Name="BtnOpenLogs" DockPanel.Dock="Right" Content="Wijzigingslog openen" Style="{StaticResource BtnLog}"/>
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

function Invoke-Connect {
    $tenant = $script:TxtTenant.Text.Trim()
    $hint = $script:TxtAdmin.Text.Trim()
    Enter-Busy 'Inloggen...'
    try {
        Connect-All -Tenant $tenant -LoginHint $hint
        $s = Get-Settings
        $s.LastAdmin = $script:MeUpn
        $s.LastTenant = $tenant
        Save-Settings
        $script:TxtAdmin.Text = $script:MeUpn
        $script:LblConn.Text = "$($script:MeUpn)$(if (-not $script:SpOk) { ' (alleen Graph)' })"
        $script:LblConn.Foreground = New-Brush $(if ($script:SpOk) { '#B4FFB4' } else { '#FFC864' })
        $script:BtnConnect.Content = 'Opnieuw inloggen'
        $script:BtnDisconnect.IsEnabled = $true
        foreach ($b in @($script:ChkAutoElevate, $script:ChkOdAutoElevate, $script:BtnBulkSharing, $script:BtnSiteQuota, $script:BtnSiteLock,
                         $script:BtnOdAdminAdd, $script:BtnOdAdminRemove, $script:BtnOdSharing, $script:BtnOdQuota, $script:BtnOdLock)) { $b.IsEnabled = $script:HasAdmin }
        foreach ($b in @($script:BtnBulkAdd, $script:BtnBulkRemove, $script:BtnBulkReport, $script:BtnBulkGroups, $script:BtnGrpScan, $script:BtnAccScan,
                         $script:BtnSiteScan, $script:BtnSiteVersions, $script:BtnSiteRecycle, $script:BtnOdScan, $script:BtnOdVersions, $script:BtnOdRecycle, $script:BtnOdEmpty,
                         $script:BtnSiteRestore, $script:BtnOdRestore)) { $b.IsEnabled = $script:SpOk }
        Set-Status "Verbonden als $($script:MeUpn)" 'success'
        Exit-Busy
        $script:OdLoaded = $false
        Import-Sites
    }
    catch {
        Write-Log "Inloggen mislukt: $($_.Exception.Message)" 'error'
        Set-Status 'Inloggen mislukt' 'error'
        Disconnect-All
        Exit-Busy
    }
}

function Invoke-Disconnect {
    if ($script:Elevated.Count) {
        $a = Show-Msg "Je bent nog tijdelijk sitebeheerder op $($script:Elevated.Count) site(s). Eerst intrekken?" 'Uitloggen' 'YesNoCancel' 'Question'
        if ($a -eq 'Cancel') { return }
        if ($a -eq 'Yes') { Invoke-ElevateRevoke -NoConfirm }
    }
    Disconnect-All
    $script:Sites = New-Object System.Collections.ArrayList
    $script:OneDrives = New-Object System.Collections.ArrayList
    $script:OdLoaded = $false
    $script:AccRows = @(); $script:GroupRows = @{}
    foreach ($P in 'S', 'O') {
        $x = $script:Ctx[$P]; $x.Site = $null; $x.PermRows = @(); $x.Libs = @()
        foreach ($g in @($x.PermGrid, $x.LibGrid, $x.LibPermGrid)) { $g.ItemsSource = $null }
    }
    foreach ($g in @($script:SiteGrid, $script:OdGrid, $script:AccGrid, $script:GrpGrid)) { $g.ItemsSource = $null }
    $script:LblConn.Text = 'Niet verbonden'
    $script:LblConn.Foreground = New-Brush '#FFC864'
    $script:BtnConnect.Content = 'Inloggen'
    $script:BtnDisconnect.IsEnabled = $false
    $script:MainArea.IsEnabled = $false
    Write-Log 'Uitgelogd' 'info'
    Set-Status 'Uitgelogd'
}

function Add-Handler {
    # Koppelt een event aan een opdracht als tekst, zodat er geen closure nodig is
    param($Control, [string]$Event, [string]$Command)
    $Control."Add_$Event"([scriptblock]::Create($Command))
}

function New-MainWindow {
    $xaml = $script:MainXamlTemplate.Replace('{{DETAIL_S}}', (Get-DetailXaml 'S')).Replace('{{DETAIL_O}}', (Get-DetailXaml 'O'))
    $script:MainXaml = $xaml
    $script:Window = ConvertFrom-Xaml $xaml.Replace('{{STYLES}}', $script:StylesXaml)
    $ctl = Get-XamlControls $script:Window $xaml
    $script:Ctx = @{ S = @{ Site = $null; PermRows = @(); Libs = @() }; O = @{ Site = $null; PermRows = @(); Libs = @() } }
    foreach ($n in $ctl.Keys) {
        if ($null -eq $ctl[$n]) { throw "XAML-element '$n' niet gevonden" }
        if ($n -match '^(S|O)_(\w+)$') { $script:Ctx[$Matches[1]][$Matches[2]] = $ctl[$n] }
        else { Set-Variable -Scope Script -Name $n -Value $ctl[$n] }
    }
    $script:Ctx.O.BtnM365Edit.Visibility = 'Collapsed'

    $s = Get-Settings
    $script:TxtAdmin.Text = [string]$s.LastAdmin
    $script:TxtTenant.Text = [string]$s.LastTenant
    $script:CmbSiteType.ItemsSource = @('Alle typen')
    $script:CmbSiteType.SelectedIndex = 0
    $script:CmbOdFilter.ItemsSource = @('Alle OneDrives', 'Account uit of verwijderd', 'Groter dan 10 GB', 'Groter dan 50 GB', 'Meer dan 80% vol', 'Inactief > 90 dagen')
    $script:CmbOdFilter.SelectedIndex = 0
    $script:CmbGrpFilter.ItemsSource = @('Alle groepen', 'Alleen statisch', 'Alleen dynamisch')
    $script:CmbGrpFilter.SelectedIndex = 0

    # Details laden kort na het selecteren van 1 regel
    $script:DetailTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:DetailTimer.Interval = [TimeSpan]::FromMilliseconds(450)
    $script:DetailTimerP = 'S'
    $script:DetailTimer.Add_Tick({
        $script:DetailTimer.Stop()
        if ($script:LongOp -or -not $script:SpOk) { return }
        $P = $script:DetailTimerP
        $sel = @(if ($P -eq 'O') { Get-SelectedOds } else { Get-SelectedSites })
        $cur = $script:Ctx[$P].Site
        if ($sel.Count -eq 1 -and (-not $cur -or $cur.Url -ne $sel[0].Url)) { Show-SiteDetail $P $sel[0] }
    })

    # --- Header ---
    $script:BtnConnect.Add_Click({ Invoke-Connect })
    $script:BtnDisconnect.Add_Click({ Invoke-Disconnect })
    $script:BtnSettings.Add_Click({ Show-SettingsDialog })

    # --- Sites ---
    $script:TxtSiteFilter.Add_TextChanged({ Update-SiteFilter })
    $script:CmbSiteType.Add_SelectionChanged({ Update-SiteFilter })
    $script:BtnSitesLoad.Add_Click({ Import-Sites })
    $script:BtnSiteScan.Add_Click({ Invoke-BulkDetails 'S' })
    $script:BtnSiteOpen.Add_Click({ foreach ($x in @(Get-SelectedSites | Select-Object -First 10)) { Start-Process $x.Url } })
    $script:BtnSiteCsv.Add_Click({ Export-GridCsv $script:SiteGrid 'SharePoint-sites' })
    $script:SiteGrid.Add_SelectionChanged({ $script:DetailTimerP = 'S'; $script:DetailTimer.Stop(); $script:DetailTimer.Start() })
    $script:SiteGrid.Add_MouseDoubleClick({ $sel = @(Get-SelectedSites); if ($sel.Count -eq 1) { Show-SiteDetail 'S' $sel[0] } })
    $script:BtnBulkAdd.Add_Click({ Invoke-BulkAccess -Mode Add })
    $script:BtnBulkRemove.Add_Click({ Invoke-BulkAccess -Mode Remove })
    $script:BtnBulkSharing.Add_Click({ Invoke-BulkSharing 'S' })
    $script:BtnBulkReport.Add_Click({ Invoke-PermReport })
    $script:BtnBulkGroups.Add_Click({ Invoke-GroupScan; $script:MainArea.SelectedIndex = 3 })
    $script:BtnSiteQuota.Add_Click({ Invoke-BulkQuota 'S' })
    $script:BtnSiteLock.Add_Click({ Invoke-BulkLock 'S' })
    $script:BtnSiteVersions.Add_Click({ Invoke-BulkVersions 'S' })
    $script:BtnSiteRecycle.Add_Click({ Invoke-BulkRecycle 'S' })
    $script:BtnSiteRestore.Add_Click({ Invoke-BulkRestore 'S' })
    $script:BtnElevRevoke.Add_Click({ Invoke-ElevateRevoke })

    # --- OneDrive ---
    $script:TxtOdFilter.Add_TextChanged({ Update-OdFilter })
    $script:CmbOdFilter.Add_SelectionChanged({ Update-OdFilter })
    $script:BtnOdLoad.Add_Click({ Import-OneDrives })
    $script:BtnOdScan.Add_Click({ Invoke-BulkDetails 'O' })
    $script:BtnOdOpen.Add_Click({ foreach ($x in @(Get-SelectedOds | Select-Object -First 10)) { Start-Process $x.Url } })
    $script:BtnOdCsv.Add_Click({ Export-GridCsv $script:OdGrid 'OneDrive-overzicht' })
    $script:OdGrid.Add_SelectionChanged({ $script:DetailTimerP = 'O'; $script:DetailTimer.Stop(); $script:DetailTimer.Start() })
    $script:OdGrid.Add_MouseDoubleClick({ $sel = @(Get-SelectedOds); if ($sel.Count -eq 1) { Show-SiteDetail 'O' $sel[0] } })
    $script:BtnOdAdminAdd.Add_Click({ Invoke-OdAdmin })
    $script:BtnOdAdminRemove.Add_Click({ Invoke-OdAdmin -Remove })
    $script:BtnOdSharing.Add_Click({ Invoke-BulkSharing 'O' })
    $script:BtnOdQuota.Add_Click({ Invoke-BulkQuota 'O' })
    $script:BtnOdLock.Add_Click({ Invoke-BulkLock 'O' })
    $script:BtnOdVersions.Add_Click({ Invoke-BulkVersions 'O' })
    $script:BtnOdRecycle.Add_Click({ Invoke-BulkRecycle 'O' })
    $script:BtnOdRestore.Add_Click({ Invoke-BulkRestore 'O' })
    $script:BtnOdEmpty.Add_Click({ Invoke-OdEmpty })
    $script:BtnOdElevRevoke.Add_Click({ Invoke-ElevateRevoke })
    $script:MainArea.Add_SelectionChanged({
        if ($_.OriginalSource -ne $script:MainArea) { return }
        if ($script:MainArea.SelectedIndex -eq 1 -and $script:Connected -and -not $script:OdLoaded -and -not $script:LongOp) { Import-OneDrives }
    })
    # Beide 'automatisch beheerder'-vinkjes delen dezelfde instelling
    $script:ChkAutoElevate.Add_Click({ $script:AutoElevate = [bool]$script:ChkAutoElevate.IsChecked; $script:ChkOdAutoElevate.IsChecked = $script:AutoElevate })
    $script:ChkOdAutoElevate.Add_Click({ $script:AutoElevate = [bool]$script:ChkOdAutoElevate.IsChecked; $script:ChkAutoElevate.IsChecked = $script:AutoElevate })

    # --- Detailpanelen (S en O) ---
    foreach ($P in 'S', 'O') {
        $x = $script:Ctx[$P]
        $csvName = if ($P -eq 'S') { 'SharePoint' } else { 'OneDrive' }
        Add-Handler $x.BtnPermReload 'Click' "Invoke-DetailReload '$P'"
        Add-Handler $x.BtnPermCsv 'Click' "Export-GridCsv `$script:Ctx['$P'].PermGrid '$csvName-machtigingen'"
        Add-Handler $x.BtnPermRemove 'Click' "Invoke-PermRows '$P' @(Get-SelectedRows `$script:Ctx['$P'].PermGrid | ForEach-Object { `$script:Ctx['$P'].PermRows[[int]`$_['Idx']] })"
        Add-Handler $x.BtnPermEditGroup 'Click' "Open-GroupFromGrid '$P' `$script:Ctx['$P'].PermGrid -Message"
        Add-Handler $x.PermGrid 'MouseDoubleClick' "Open-GroupFromGrid '$P' `$script:Ctx['$P'].PermGrid"
        Add-Handler $x.LibPermGrid 'MouseDoubleClick' "Open-GroupFromGrid '$P' `$script:Ctx['$P'].LibPermGrid"
        Add-Handler $x.BtnM365Edit 'Click' "`$s = `$script:Ctx['$P'].Site; if (`$s -and `$s.GroupId) { Open-GroupEditor `$s.GroupId }"
        Add-Handler $x.LibGrid 'SelectionChanged' "Update-LibPermGrid '$P'"
        Add-Handler $x.BtnLibBreak 'Click' "Invoke-LibBreak '$P'"
        Add-Handler $x.BtnLibReset 'Click' "Invoke-LibReset '$P'"
        Add-Handler $x.BtnLibGrant 'Click' "Invoke-LibGrant '$P'"
        Add-Handler $x.BtnLibPermRemove 'Click' "Invoke-LibPermRemove '$P'"
        Add-Handler $x.BtnLibVerLimit 'Click' "Invoke-LibVersionLimit '$P'"
        Add-Handler $x.BtnLibVerClean 'Click' "Invoke-LibVersionClean '$P'"
        Add-Handler $x.BtnLibOpen 'Click' "Open-Library '$P'"
        Add-Handler $x.BtnLibCsv 'Click' "Export-GridCsv `$script:Ctx['$P'].LibGrid '$csvName-bibliotheken'"
    }

    # --- Gebruiker-toegang ---
    $script:BtnAccPick.Add_Click({
        $p = @(Show-PrincipalPicker -Kinds user -Single -Title 'Gebruiker kiezen') | Select-Object -First 1
        if ($p) {
            $script:AccUser = $p
            $script:LblAccUser.Text = "$($p.Name)  ($($p.Upn))"
            $script:AccRows = @()
            Update-AccessGrid
        }
    })
    $script:BtnAccScan.Add_Click({ Invoke-AccessScan })
    $script:BtnAccRemove.Add_Click({ Invoke-AccessRemove })
    $script:BtnAccGroupEdit.Add_Click({
        $a = @(Get-SelectedRows $script:AccGrid | ForEach-Object { $script:AccRows[[int]$_['Idx']] } | Where-Object { $_.Row.ObjectId }) | Select-Object -First 1
        Open-GroupEditor $(if ($a) { $a.Row.ObjectId } else { '' })
    })
    $script:AccGrid.Add_MouseDoubleClick({
        $a = @(Get-SelectedRows $script:AccGrid | ForEach-Object { $script:AccRows[[int]$_['Idx']] }) | Select-Object -First 1
        if ($a -and $a.Row.ObjectId) { Open-GroupEditor $a.Row.ObjectId }
    })
    $script:BtnAccCsv.Add_Click({ Export-GridCsv $script:AccGrid 'SharePoint-gebruiker-toegang' })

    # --- Groepen ---
    $script:BtnGrpScan.Add_Click({ Invoke-GroupScan })
    $script:BtnGrpSearch.Add_Click({ Invoke-GroupSearch })
    $script:TxtGrpSearch.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-GroupSearch } })
    $script:CmbGrpFilter.Add_SelectionChanged({ Update-GroupFilter })
    $script:BtnGrpEdit.Add_Click({ $r = @(Get-SelectedRows $script:GrpGrid) | Select-Object -First 1; Open-GroupEditor $(if ($r) { [string]$r['Id'] } else { '' }) })
    $script:GrpGrid.Add_MouseDoubleClick({ $r = @(Get-SelectedRows $script:GrpGrid) | Select-Object -First 1; if ($r) { Open-GroupEditor ([string]$r['Id']) } })
    $script:BtnGrpCsv.Add_Click({ Export-GridCsv $script:GrpGrid 'Entra-groepen-in-sites' })

    # --- Algemeen ---
    $script:BtnStop.Add_Click({
        $script:StopRequested = $true
        $script:BtnStop.IsEnabled = $false
        Set-Status 'Stoppen na de huidige stap...'
    })
    $script:BtnClearLog.Add_Click({ $script:LogPara.Inlines.Clear() })
    $script:BtnOpenLogs.Add_Click({
        $d = if ($script:TenantId) { Get-LogDir } else { $script:AppDir }
        if (-not (Test-Path $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
        Start-Process explorer.exe "`"$d`""
    })
    $script:Window.Add_Closing({
        if ($script:LongOp) {
            $script:StopRequested = $true
            $_.Cancel = $true
            Set-Status 'Bezig - wordt gestopt, sluit daarna opnieuw'
            return
        }
        if ($script:Connected -and $script:Elevated.Count) {
            $a = Show-Msg "Je bent nog tijdelijk sitebeheerder op $($script:Elevated.Count) site(s). Nu intrekken?`n`n(Nee = later intrekken; de lijst wordt bewaard.)" 'Afsluiten' 'YesNoCancel' 'Question'
            if ($a -eq 'Cancel') { $_.Cancel = $true; return }
            if ($a -eq 'Yes') { Invoke-ElevateRevoke -NoConfirm }
        }
    })
    $script:Window.Add_Closed({ Disconnect-All })
}

function Open-GroupFromGrid {
    param([string]$P, $Grid, [switch]$Message)
    $x = $script:Ctx[$P]
    $r = @(Get-SelectedRows $Grid | ForEach-Object { $x.PermRows[[int]$_['Idx']] } | Where-Object { $_.ObjectId }) | Select-Object -First 1
    if ($r) { Open-GroupEditor $r.ObjectId }
    elseif ($Message) { Open-GroupEditor '' }
}

function Open-Library {
    param([string]$P)
    $x = $script:Ctx[$P]
    foreach ($l in @(Get-SelectedLibs $P | Select-Object -First 5)) {
        $u = ([uri]$x.Site.Url).GetLeftPart([UriPartial]::Authority) + $l.Path
        Start-Process $u
    }
}

function Show-MainWindow {
    New-MainWindow
    Write-Log 'Log in met je admin-account: eerst Graph (browser), daarna SharePoint (eigen venster). Kies beide keren je admin-account.' 'info'
    Write-Log "Wijzigingen worden gelogd in $($script:AppDir)\Logs" 'info'
    [void]$script:Window.ShowDialog()
}

#endregion

if (-not (Initialize-Msal)) {
    [System.Windows.MessageBox]::Show("MSAL (Microsoft.Identity.Client) niet gevonden.`n`nInstalleer bijvoorbeeld:`nInstall-Module Microsoft.Graph.Authentication -Scope CurrentUser`n(in Windows PowerShell 5.1)", 'Module ontbreekt', 'OK', 'Error') | Out-Null
    exit 1
}
Show-MainWindow
