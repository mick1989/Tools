#Requires -Version 5.1
<#
.SYNOPSIS
    Entra Admin Rollen Export (WPF) - alle accounts met een Entra-adminrol
    ophalen, bekijken en exporteren naar CSV.
.DESCRIPTION
    Haalt alle directory-roltoewijzingen op via Microsoft Graph:
      - Actieve toewijzingen (permanent en via PIM geactiveerd)
      - Geschikte toewijzingen (PIM eligible, vereist Entra ID P2)
      - Rollen die via een (role-assignable) groep zijn toegekend, optioneel
        uitgeklapt naar de leden van die groep
      - Serviceprincipals (apps) met een rol (optioneel)
    Per gebruiker optioneel: ingeschakeld, gast/lid, on-prem gesynchroniseerd,
    laatste aanmelding (vereist P1) en MFA-registratie.

    Weergave: 1 regel per roltoewijzing, of 1 regel per account (rollen
    samengevoegd). De zichtbare (gefilterde) regels worden geexporteerd naar
    CSV met puntkomma als scheidingsteken (opent direct goed in NL-Excel).

    Inloggen: altijd met account-kiezer, zodat je met je aparte admin-account
    inlogt. Optioneel een tenant (domein of tenant-ID) invullen voor GDAP.
    Token blijft alleen in dit proces.
.NOTES
    Vereist: Microsoft.Graph.Authentication (2.33.0 aanbevolen)
      Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser
    Rollen (minimaal): Global Reader (of Security Reader + Reports Reader).
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

$script:GraphRoot     = 'https://graph.microsoft.com/v1.0'
$script:GraphClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools
$script:GraphScopes   = @('RoleManagement.Read.Directory', 'Directory.Read.All', 'AuditLog.Read.All', 'UserAuthenticationMethod.Read.All')
$script:Msal          = $null
$script:MsalAuthority = $null
$script:MsalAvailable = $false
$script:Auth          = $null
$script:Connected     = $false
$script:TenantName    = ''
$script:TableRows     = $null   # per toewijzing
$script:TableAccounts = $null   # per account
$script:SettingsDir   = Join-Path $env:LOCALAPPDATA 'EntraAdminRollenExport'
$script:SettingsFile  = Join-Path $script:SettingsDir 'settings.json'
$script:LogFile       = Join-Path $script:SettingsDir 'log.txt'

#endregion

#region --- XAML ---

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Entra Admin Rollen Export" Height="760" Width="1280" MinHeight="500" MinWidth="900"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="12" Background="#F5F6F8">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="12,5"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="MinWidth" Value="90"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,0,18,0"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
  </Window.Resources>
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="120"/>
    </Grid.RowDefinitions>

    <!-- Verbinding -->
    <Border Grid.Row="0" Background="White" BorderBrush="#DDE1E6" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,8">
      <DockPanel LastChildFill="True">
        <TextBlock Text="Tenant (optioneel, GDAP):" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <TextBox x:Name="TxtTenant" Width="220" VerticalAlignment="Center" Padding="3" Margin="0,0,12,0"
                 ToolTip="Leeg = je eigen tenant. Vul een domein of tenant-ID in om via GDAP bij een klant in te loggen."/>
        <Button x:Name="BtnLogin" Content="Inloggen"/>
        <Button x:Name="BtnLogout" Content="Uitloggen" IsEnabled="False"/>
        <TextBlock x:Name="LblStatus" Text="Niet verbonden" VerticalAlignment="Center" Foreground="#B00020" Margin="8,0,0,0" TextTrimming="CharacterEllipsis"/>
      </DockPanel>
    </Border>

    <!-- Opties -->
    <Border Grid.Row="1" Background="White" BorderBrush="#DDE1E6" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,8">
      <StackPanel>
        <WrapPanel>
          <CheckBox x:Name="ChkEligible" Content="Geschikte toewijzingen (PIM) meenemen" IsChecked="True"/>
          <CheckBox x:Name="ChkExpand" Content="Groepen uitklappen naar leden" IsChecked="True"/>
          <CheckBox x:Name="ChkSp" Content="Serviceprincipals (apps) meenemen" IsChecked="True"/>
          <CheckBox x:Name="ChkSignIn" Content="Laatste aanmelding ophalen" IsChecked="True"/>
          <CheckBox x:Name="ChkMfa" Content="MFA-registratie ophalen" IsChecked="True"/>
        </WrapPanel>
        <WrapPanel Margin="0,10,0,0">
          <Button x:Name="BtnLoad" Content="Ophalen" IsEnabled="False" FontWeight="SemiBold"/>
          <TextBlock Text="Weergave:" VerticalAlignment="Center" Margin="10,0,6,0"/>
          <RadioButton x:Name="RbRows" Content="Per roltoewijzing" IsChecked="True" VerticalAlignment="Center" Margin="0,0,12,0"/>
          <RadioButton x:Name="RbAccounts" Content="Per account" VerticalAlignment="Center" Margin="0,0,12,0"/>
        </WrapPanel>
      </StackPanel>
    </Border>

    <!-- Filter -->
    <DockPanel Grid.Row="2" Margin="0,0,0,6">
      <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="TxtFilter" Width="300" Padding="3" VerticalAlignment="Center"
               ToolTip="Zoekt in rol, naam, UPN, toewijzing en via-groep"/>
      <TextBlock x:Name="LblCount" Text="" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="#555"/>
    </DockPanel>

    <DataGrid Grid.Row="3" x:Name="Grid" AutoGenerateColumns="True" IsReadOnly="True" CanUserSortColumns="True"
              HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#EEE"
              AlternatingRowBackground="#FAFBFC" SelectionMode="Extended" ClipboardCopyMode="IncludeHeader"
              Background="White" BorderBrush="#DDE1E6"/>

    <WrapPanel Grid.Row="4" Margin="0,8,0,8">
      <Button x:Name="BtnExport" Content="Exporteren naar CSV" IsEnabled="False" FontWeight="SemiBold"/>
      <TextBlock Text="Exporteert de zichtbare (gefilterde) regels van de huidige weergave, scheidingsteken ;" VerticalAlignment="Center" Foreground="#555"/>
    </WrapPanel>

    <TextBox Grid.Row="5" x:Name="TxtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"
             FontFamily="Consolas" FontSize="11" Background="#1E1E1E" Foreground="#DDD" BorderThickness="0" Padding="6"/>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$script:Window = [Windows.Markup.XamlReader]::Load($reader)
$script:ui = @{}
foreach ($n in 'TxtTenant','BtnLogin','BtnLogout','LblStatus','ChkEligible','ChkExpand','ChkSp','ChkSignIn','ChkMfa',
               'BtnLoad','RbRows','RbAccounts','TxtFilter','LblCount','Grid','BtnExport','TxtLog') {
    $script:ui[$n] = $script:Window.FindName($n)
}

#endregion

#region --- Hulpfuncties ---

function Update-UI {
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    $cb = [System.Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }
    [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $cb, $frame)
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Write-FileLog {
    param([string]$Line)
    try {
        if (-not (Test-Path $script:SettingsDir)) { New-Item -ItemType Directory -Path $script:SettingsDir -Force | Out-Null }
        Add-Content -Path $script:LogFile -Value ("{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Line) -Encoding UTF8
    } catch { }
}

function Write-Log {
    param([string]$Message, [string]$Level = 'info')
    $prefix = switch ($Level) { 'warning' { '[!] ' } 'error' { '[X] ' } 'success' { '[OK] ' } default { '' } }
    Write-FileLog "$prefix$Message"
    try {
        $script:ui.TxtLog.AppendText(("{0}  {1}{2}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $prefix, $Message))
        $script:ui.TxtLog.ScrollToEnd()
        Update-UI
    } catch { }
}

function Show-Crash {
    param($Err, [string]$Where)
    $msg = if ($Err -is [System.Management.Automation.ErrorRecord]) {
        "$($Err.Exception.GetBaseException().Message)`r`n$($Err.ScriptStackTrace)"
    } elseif ($Err -is [Exception]) { "$($Err.GetBaseException().Message)`r`n$($Err.StackTrace)" } else { "$Err" }
    Write-Log "Onverwachte fout ($Where): $msg" 'error'
    try { [System.Windows.MessageBox]::Show("Onverwachte fout ($Where):`n`n$msg`n`nLog: $($script:LogFile)", 'Fout', 'OK', 'Error') | Out-Null } catch { }
}

function Set-Busy {
    param([bool]$Busy)
    $script:ui.BtnLogin.IsEnabled  = -not $Busy
    $script:ui.BtnLogout.IsEnabled = (-not $Busy) -and $script:Connected
    $script:ui.BtnLoad.IsEnabled   = (-not $Busy) -and $script:Connected
    $script:ui.BtnExport.IsEnabled = (-not $Busy) -and ($null -ne $script:TableRows)
    $script:Window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
    Update-UI
}

function Get-Settings {
    try { if (Test-Path $script:SettingsFile) { return Get-Content $script:SettingsFile -Raw | ConvertFrom-Json } } catch { }
    return [pscustomobject]@{ LoginHint = ''; Tenant = '' }
}

function Save-Settings {
    param([string]$LoginHint, [string]$Tenant)
    try {
        if (-not (Test-Path $script:SettingsDir)) { New-Item -ItemType Directory -Path $script:SettingsDir -Force | Out-Null }
        [pscustomobject]@{ LoginHint = $LoginHint; Tenant = $Tenant } | ConvertTo-Json | Set-Content $script:SettingsFile -Encoding UTF8
    } catch { }
}

function Format-Date {
    param($Value)
    if (-not $Value) { return '' }
    try { return ([datetime]$Value).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { return "$Value" }
}

function Get-Prop {
    # Veilig een waarde uit een hashtable/object halen
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [System.Collections.IDictionary]) { return $Obj[$Name] }
    return $Obj.$Name
}

#endregion

#region --- Modules en verbinding ---

function Initialize-Modules {
    # Vanaf 2.34.0 dwingt Microsoft.Graph.Authentication WAM af (geen account-kiezer)
    try { Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -ErrorAction Stop }
    catch {
        try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
        catch { return $false }
    }
    $script:MsalAvailable = [bool]('Microsoft.Identity.Client.PublicClientApplicationBuilder' -as [type])
    if (-not $script:MsalAvailable) {
        $base = (Get-Module Microsoft.Graph.Authentication).ModuleBase
        $sub  = if ($PSVersionTable.PSEdition -eq 'Core') { 'Core' } else { 'Desktop' }
        $dll  = Join-Path $base "Dependencies\$sub\Microsoft.Identity.Client.dll"
        if (-not (Test-Path $dll)) {
            $dll = Get-ChildItem $base -Recurse -Filter 'Microsoft.Identity.Client.dll' -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -match "\\$sub\\" } | Select-Object -First 1 -ExpandProperty FullName
        }
        if ($dll) { try { Add-Type -Path $dll -ErrorAction Stop; $script:MsalAvailable = $true } catch { } }
    }
    return $true
}

function Wait-MsalTask {
    param($Task)
    while (-not $Task.IsCompleted) { Update-UI; Start-Sleep -Milliseconds 100 }
    if ($Task.IsCanceled) { throw 'Inloggen geannuleerd of verlopen' }
    if ($Task.IsFaulted) { throw $Task.Exception.GetBaseException().Message }
    return $Task.Result
}

function Get-MsalApp {
    param([string]$Tenant)
    $authority = "https://login.microsoftonline.com/$(if ($Tenant) { $Tenant } else { 'organizations' })"
    if (-not $script:Msal -or $script:MsalAuthority -ne $authority) {
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
    $script:Auth = @{ Account = $Result.Account; Expires = $Result.ExpiresOn.LocalDateTime }
}

function Connect-ToGraph {
    param([string]$Tenant, [string]$LoginHint)
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    $script:Auth = $null
    if ($script:MsalAvailable) {
        $app = Get-MsalApp $Tenant
        foreach ($a in @((Wait-MsalTask $app.GetAccountsAsync()))) { if ($a) { [void](Wait-MsalTask $app.RemoveAsync($a)) } }
        $b = $app.AcquireTokenInteractive([string[]]$script:GraphScopes).
            WithUseEmbeddedWebView($false).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)
        if ($LoginHint) { $b = $b.WithLoginHint($LoginHint) }
        $cts = New-Object System.Threading.CancellationTokenSource ([TimeSpan]::FromMinutes(5))
        $r = Wait-MsalTask $b.ExecuteAsync($cts.Token)
        Connect-GraphWithToken $r
        return $r.Account.Username
    }
    Write-Log 'MSAL niet beschikbaar - standaard Graph-login (browser kiest mogelijk automatisch een account)' 'warning'
    $p = @{ Scopes = $script:GraphScopes; NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
    if ($Tenant) { $p.TenantId = $Tenant }
    Connect-MgGraph @p | Out-Null
    return (Get-MgContext).Account
}

function Update-GraphToken {
    if (-not $script:Auth -or (Get-Date) -lt $script:Auth.Expires.AddMinutes(-5)) { return }
    try {
        $r = Wait-MsalTask $script:Msal.AcquireTokenSilent([string[]]$script:GraphScopes, $script:Auth.Account).ExecuteAsync()
        Connect-GraphWithToken $r
        Write-Log "Graph-token vernieuwd (geldig tot $($script:Auth.Expires.ToString('HH:mm')))"
    }
    catch {
        Write-Log "Token vernieuwen mislukt: $_" 'warning'
        $script:Auth.Expires = (Get-Date).AddHours(1)
    }
}

#endregion

#region --- Graph ---

function Invoke-Graph {
    param([string]$Uri, [string]$Method = 'GET', $Body)
    Update-GraphToken
    Update-UI
    $p = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $p.Body = $Body; $p.ContentType = 'application/json' }
    return Invoke-MgGraphRequest @p
}

function Get-GraphAll {
    param([string]$Uri)
    $out  = New-Object System.Collections.ArrayList
    $next = $Uri
    while ($next) {
        $r = Invoke-Graph $next
        foreach ($v in @($r['value'])) { if ($null -ne $v) { [void]$out.Add($v) } }
        $next = $r['@odata.nextLink']
    }
    return ,$out
}

function Invoke-GraphBatch {
    # $Requests: lijst van @{ id = '..'; url = '/relatief' }. Geeft hashtable id -> response.
    param([System.Collections.IList]$Requests, [string]$Label = 'items')
    $results = @{}
    $total = $Requests.Count
    for ($i = 0; $i -lt $total; $i += 20) {
        $pending = @($Requests[$i..([Math]::Min($i + 19, $total - 1))])
        $attempt = 0
        while ($pending.Count -gt 0 -and $attempt -lt 5) {
            $attempt++
            $body = @{ requests = @($pending | ForEach-Object { @{ id = $_.id; method = 'GET'; url = $_.url } }) } | ConvertTo-Json -Depth 5
            $r = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $body
            $retry = @(); $wait = 2
            foreach ($resp in @($r['responses'])) {
                if ([int]$resp['status'] -eq 429 -or [int]$resp['status'] -ge 500) {
                    $retry += $pending | Where-Object { $_.id -eq $resp['id'] }
                    $ra = Get-Prop $resp['headers'] 'Retry-After'
                    if ($ra -and [int]$ra -gt $wait) { $wait = [int]$ra }
                }
                else { $results[$resp['id']] = $resp }
            }
            $pending = $retry
            if ($pending.Count -gt 0) {
                $until = (Get-Date).AddSeconds([Math]::Min($wait, 30))
                while ((Get-Date) -lt $until) { Update-UI; Start-Sleep -Milliseconds 200 }
            }
        }
        if ($total -gt 40) { $script:ui.LblCount.Text = "$Label ophalen: $([Math]::Min($i + 20, $total)) / $total"; Update-UI }
    }
    return $results
}

function Get-PrincipalType {
    param($Principal)
    switch -Wildcard ([string](Get-Prop $Principal '@odata.type')) {
        '*user'             { return 'Gebruiker' }
        '*group'            { return 'Groep' }
        '*servicePrincipal' { return 'Serviceprincipal' }
        default             { return 'Onbekend' }
    }
}

#endregion

#region --- Ophalen ---

function Get-AdminRoleData {
    $optEligible = [bool]$script:ui.ChkEligible.IsChecked
    $optExpand   = [bool]$script:ui.ChkExpand.IsChecked
    $optSp       = [bool]$script:ui.ChkSp.IsChecked
    $optSignIn   = [bool]$script:ui.ChkSignIn.IsChecked
    $optMfa      = [bool]$script:ui.ChkMfa.IsChecked

    # 1. Roldefinities
    Write-Log 'Roldefinities ophalen...'
    $roleDefs = @{}
    foreach ($d in (Get-GraphAll "$($script:GraphRoot)/roleManagement/directory/roleDefinitions")) { $roleDefs[$d['id']] = $d }
    Write-Log "$($roleDefs.Count) roldefinities"

    # 2. Actieve toewijzingen
    Write-Log 'Actieve roltoewijzingen ophalen...'
    $active = Get-GraphAll "$($script:GraphRoot)/roleManagement/directory/roleAssignments?`$expand=principal"
    Write-Log "$($active.Count) actieve toewijzingen"

    # 3. PIM-info voor actieve toewijzingen (permanent / geactiveerd + einddatum)
    $instances = @{}
    try {
        foreach ($inst in (Get-GraphAll "$($script:GraphRoot)/roleManagement/directory/roleAssignmentScheduleInstances")) {
            $key = '{0}|{1}|{2}' -f $inst['principalId'], $inst['roleDefinitionId'], $inst['directoryScopeId']
            $instances[$key] = $inst
        }
    }
    catch { Write-Log 'PIM-info voor actieve toewijzingen niet beschikbaar (geen P2 of rechten) - einddatums ontbreken' 'warning' }

    # 4. Geschikte toewijzingen
    $eligible = @()
    if ($optEligible) {
        Write-Log 'Geschikte (PIM) toewijzingen ophalen...'
        try {
            $eligible = Get-GraphAll "$($script:GraphRoot)/roleManagement/directory/roleEligibilitySchedules?`$expand=principal"
            Write-Log "$($eligible.Count) geschikte toewijzingen"
        }
        catch { Write-Log "Geschikte toewijzingen niet op te halen (geen Entra ID P2 of rechten): $($_.Exception.Message)" 'warning' }
    }

    # 5. Ruwe toewijzingen samenvoegen
    $raw = New-Object System.Collections.ArrayList
    foreach ($a in $active) {
        $key  = '{0}|{1}|{2}' -f $a['principalId'], $a['roleDefinitionId'], $a['directoryScopeId']
        $inst = $instances[$key]
        $state = 'Actief (permanent)'; $end = $null; $start = $null
        if ($inst) {
            $start = $inst['startDateTime']; $end = $inst['endDateTime']
            if ($inst['assignmentType'] -eq 'Activated') { $state = 'Actief (PIM geactiveerd)' }
            elseif ($end) { $state = 'Actief (tijdelijk)' }
        }
        [void]$raw.Add(@{ RoleId = $a['roleDefinitionId']; Scope = $a['directoryScopeId']; State = $state
                          Start = $start; End = $end; Principal = $a['principal']; PrincipalId = $a['principalId'] })
    }
    foreach ($e in $eligible) {
        $si  = $e['scheduleInfo']
        $exp = Get-Prop $si 'expiration'
        $end = if ((Get-Prop $exp 'type') -eq 'noExpiration') { $null } else { Get-Prop $exp 'endDateTime' }
        [void]$raw.Add(@{ RoleId = $e['roleDefinitionId']; Scope = $e['directoryScopeId']; State = 'Geschikt (PIM)'
                          Start = (Get-Prop $si 'startDateTime'); End = $end; Principal = $e['principal']; PrincipalId = $e['principalId'] })
    }

    # 6. Regels opbouwen, groepen uitklappen
    $groupCache = @{}
    $lines = New-Object System.Collections.ArrayList
    foreach ($r in $raw) {
        $p = $r.Principal
        $type = if ($p) { Get-PrincipalType $p } else { 'Verwijderd object' }
        if ($type -eq 'Serviceprincipal' -and -not $optSp) { continue }
        [void]$lines.Add(@{ Raw = $r; Id = $r.PrincipalId; Type = $type; Principal = $p; ViaGroup = '' })

        if ($type -eq 'Groep' -and $optExpand) {
            $gid = $r.PrincipalId
            if (-not $groupCache.ContainsKey($gid)) {
                try { $groupCache[$gid] = Get-GraphAll "$($script:GraphRoot)/groups/$gid/transitiveMembers?`$select=id,displayName,userPrincipalName,appId" }
                catch { Write-Log "Leden van groep '$(Get-Prop $p 'displayName')' niet op te halen: $($_.Exception.Message)" 'warning'; $groupCache[$gid] = @() }
            }
            foreach ($m in $groupCache[$gid]) {
                $mt = Get-PrincipalType $m
                if ($mt -eq 'Groep') { continue }     # transitief: geneste groepen zelf overslaan
                if ($mt -eq 'Serviceprincipal' -and -not $optSp) { continue }
                [void]$lines.Add(@{ Raw = $r; Id = $m['id']; Type = $mt; Principal = $m; ViaGroup = [string](Get-Prop $p 'displayName') })
            }
        }
    }
    if ($groupCache.Count) { Write-Log "$($groupCache.Count) groep(en) met een rol uitgeklapt" }

    # 7. Gebruikersdetails
    $userIds = @($lines | Where-Object { $_.Type -eq 'Gebruiker' } | ForEach-Object { $_.Id } | Select-Object -Unique)
    $userInfo = @{}; $mfaInfo = @{}
    if ($userIds.Count) {
        $select = 'id,displayName,userPrincipalName,accountEnabled,userType,onPremisesSyncEnabled,createdDateTime'
        if ($optSignIn) {
            try { [void](Invoke-Graph "$($script:GraphRoot)/users/$($userIds[0])?`$select=id,signInActivity"); $select += ',signInActivity' }
            catch { Write-Log 'Laatste aanmelding niet beschikbaar (vereist Entra ID P1 en AuditLog.Read.All)' 'warning'; $optSignIn = $false }
        }
        if ($optMfa) {
            try { [void](Invoke-Graph "$($script:GraphRoot)/reports/authenticationMethods/userRegistrationDetails/$($userIds[0])") }
            catch {
                if ($_.Exception.Message -notmatch 'NotFound|404') { Write-Log 'MFA-registratie niet beschikbaar (vereist Entra ID P1 en Reports Reader)' 'warning'; $optMfa = $false }
            }
        }
        Write-Log "Details van $($userIds.Count) gebruiker(s) ophalen..."
        $reqs = New-Object System.Collections.ArrayList
        foreach ($id in $userIds) {
            [void]$reqs.Add(@{ id = "u_$id"; url = "/users/$id`?`$select=$select" })
            if ($optMfa) { [void]$reqs.Add(@{ id = "m_$id"; url = "/reports/authenticationMethods/userRegistrationDetails/$id" }) }
        }
        $res = Invoke-GraphBatch -Requests $reqs -Label 'Gebruikersdetails'
        foreach ($k in $res.Keys) {
            $resp = $res[$k]
            if ([int]$resp['status'] -ne 200) { continue }
            if ($k.StartsWith('u_')) { $userInfo[$k.Substring(2)] = $resp['body'] } else { $mfaInfo[$k.Substring(2)] = $resp['body'] }
        }
    }

    # 8. Bereiken (scopes) vertalen
    $scopeNames = @{ '/' = 'Tenant' }
    foreach ($s in @($raw | ForEach-Object { $_.Scope } | Select-Object -Unique)) {
        if (-not $s -or $scopeNames.ContainsKey($s)) { continue }
        $id = ($s -split '/')[-1]
        try {
            if ($s -like '/administrativeUnits/*') {
                $o = Invoke-Graph "$($script:GraphRoot)/directory/administrativeUnits/$id`?`$select=displayName"
                $scopeNames[$s] = "AU: $($o['displayName'])"
            }
            else {
                $o = Invoke-Graph "$($script:GraphRoot)/directoryObjects/$id"
                $scopeNames[$s] = "Object: $($o['displayName'])"
            }
        }
        catch { $scopeNames[$s] = $s }
    }

    # 9. Tabel per toewijzing
    $cols = 'Rol','Toewijzing','Startdatum','Einddatum','Bereik','Type','Naam','UPN / App-ID','Via groep',
            'Ingeschakeld','Gebruikerstype','On-prem gesync','Laatste aanmelding','MFA geregistreerd','MFA-methoden',
            'Aangemaakt','Rol ingebouwd','Rol bevoorrecht','Object-ID','Rol-ID'
    $dt = New-Object System.Data.DataTable
    foreach ($c in $cols) { [void]$dt.Columns.Add($c) }

    foreach ($l in $lines) {
        $r  = $l.Raw
        $rd = $roleDefs[$r.RoleId]
        $p  = $l.Principal
        $u  = $userInfo[$l.Id]
        $m  = $mfaInfo[$l.Id]
        $row = $dt.NewRow()
        $row['Rol']           = if ($rd) { $rd['displayName'] } else { $r.RoleId }
        $row['Toewijzing']    = $r.State
        $row['Startdatum']    = Format-Date $r.Start
        $row['Einddatum']     = if ($r.End) { Format-Date $r.End } else { 'Permanent' }
        $row['Bereik']        = if ($scopeNames.ContainsKey([string]$r.Scope)) { $scopeNames[[string]$r.Scope] } else { $r.Scope }
        $row['Type']          = $l.Type
        $row['Naam']          = if ($u) { $u['displayName'] } elseif ($p) { Get-Prop $p 'displayName' } else { '(verwijderd object)' }
        $row['UPN / App-ID']  = if ($u) { $u['userPrincipalName'] } elseif ($l.Type -eq 'Serviceprincipal') { Get-Prop $p 'appId' } else { Get-Prop $p 'userPrincipalName' }
        $row['Via groep']     = $l.ViaGroup
        $row['Object-ID']     = $l.Id
        $row['Rol-ID']        = $r.RoleId
        if ($rd) {
            $row['Rol ingebouwd']   = if ($rd['isBuiltIn']) { 'Ja' } else { 'Nee' }
            if ($rd.ContainsKey('isPrivileged')) { $row['Rol bevoorrecht'] = if ($rd['isPrivileged']) { 'Ja' } else { 'Nee' } }
        }
        if ($u) {
            $row['Ingeschakeld']   = if ($u['accountEnabled']) { 'Ja' } else { 'Nee' }
            $row['Gebruikerstype'] = $u['userType']
            $row['On-prem gesync'] = if ($u['onPremisesSyncEnabled']) { 'Ja' } else { 'Nee' }
            $row['Aangemaakt']     = Format-Date $u['createdDateTime']
            if ($optSignIn) {
                $sa = $u['signInActivity']
                $last = Get-Prop $sa 'lastSuccessfulSignInDateTime'
                if (-not $last) {
                    $cands = @((Get-Prop $sa 'lastSignInDateTime'), (Get-Prop $sa 'lastNonInteractiveSignInDateTime')) | Where-Object { $_ } |
                        ForEach-Object { [datetime]$_ } | Sort-Object -Descending
                    $last = $cands | Select-Object -First 1
                }
                $row['Laatste aanmelding'] = if ($last) { Format-Date $last } else { 'Nooit / onbekend' }
            }
        }
        if ($m) {
            $row['MFA geregistreerd'] = if ($m['isMfaRegistered']) { 'Ja' } else { 'Nee' }
            $row['MFA-methoden']      = (@($m['methodsRegistered']) -join ', ')
        }
        $dt.Rows.Add($row)
    }
    $dt.DefaultView.Sort = 'Rol ASC, Naam ASC'
    $script:TableRows = $dt.DefaultView.ToTable()

    # 10. Tabel per account
    $acc = New-Object System.Data.DataTable
    foreach ($c in 'Naam','UPN / App-ID','Type','Aantal rollen','Actieve rollen','Geschikte rollen (PIM)','Via groep',
                   'Ingeschakeld','Gebruikerstype','On-prem gesync','Laatste aanmelding','MFA geregistreerd','MFA-methoden','Aangemaakt','Object-ID') {
        [void]$acc.Columns.Add($c)
    }
    foreach ($g in ($script:TableRows.Rows | Group-Object { $_['Object-ID'] })) {
        $f = $g.Group[0]
        $label = { param($x) $n = $x['Rol']; if ($x['Bereik'] -ne 'Tenant') { $n += " ($($x['Bereik']))" }; $n }
        $act = @($g.Group | Where-Object { $_['Toewijzing'] -like 'Actief*' } | ForEach-Object { & $label $_ } | Select-Object -Unique)
        $elg = @($g.Group | Where-Object { $_['Toewijzing'] -like 'Geschikt*' } | ForEach-Object { & $label $_ } | Select-Object -Unique)
        $row = $acc.NewRow()
        foreach ($c in 'Naam','UPN / App-ID','Type','Ingeschakeld','Gebruikerstype','On-prem gesync','Laatste aanmelding','MFA geregistreerd','MFA-methoden','Aangemaakt','Object-ID') { $row[$c] = $f[$c] }
        $row['Aantal rollen']          = ($act + $elg | Select-Object -Unique).Count
        $row['Actieve rollen']         = $act -join '; '
        $row['Geschikte rollen (PIM)'] = $elg -join '; '
        $row['Via groep']              = (@($g.Group | ForEach-Object { $_['Via groep'] } | Where-Object { $_ } | Select-Object -Unique) -join '; ')
        $acc.Rows.Add($row)
    }
    $acc.DefaultView.Sort = 'Naam ASC'
    $script:TableAccounts = $acc.DefaultView.ToTable()

    $nAcc = @($script:TableAccounts.Rows | Where-Object { $_['Type'] -ne 'Groep' }).Count
    Write-Log "$($script:TableRows.Rows.Count) roltoewijzingen, $nAcc account(s) met een adminrol" 'success'
}

#endregion

#region --- Weergave en export ---

function Get-CurrentTable {
    # Komma voorkomt dat PowerShell de DataTable uitpakt tot losse DataRows
    if ($script:ui.RbAccounts.IsChecked) { return ,$script:TableAccounts } else { return ,$script:TableRows }
}

function Update-View {
    try { Update-ViewCore } catch { Show-Crash $_ 'weergave' }
}

function Update-ViewCore {
    $t = Get-CurrentTable
    if ($null -eq $t) { $script:ui.Grid.ItemsSource = $null; $script:ui.LblCount.Text = ''; return }
    $view = $t.DefaultView
    $q = $script:ui.TxtFilter.Text.Trim()
    if ($q) {
        $esc = ($q -replace "'", "''") -replace '([\[\]\*%])', '[$1]'
        $fields = @('Naam', 'UPN / App-ID', 'Via groep', 'Type')
        if ($t.Columns.Contains('Rol')) { $fields += 'Rol', 'Toewijzing', 'Bereik' } else { $fields += 'Actieve rollen', 'Geschikte rollen (PIM)' }
        $view.RowFilter = ($fields | ForEach-Object { "[$_] LIKE '%$esc%'" }) -join ' OR '
    }
    else { $view.RowFilter = '' }
    $script:ui.Grid.ItemsSource = $view
    $script:ui.LblCount.Text = "$($view.Count) van $($t.Rows.Count) regels"
}

function Export-CurrentView {
    $t = Get-CurrentTable
    if ($null -eq $t) { return }
    $view = $t.DefaultView
    $mode = if ($script:ui.RbAccounts.IsChecked) { 'PerAccount' } else { 'PerToewijzing' }
    $tenant = if ($script:TenantName) { $script:TenantName -replace '[^\w\.-]', '_' } else { 'tenant' }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV-bestand (*.csv)|*.csv'
    $dlg.FileName = "AdminRollen_${tenant}_${mode}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if (-not $dlg.ShowDialog($script:Window)) { return }
    $cols = @($t.Columns | ForEach-Object { $_.ColumnName })
    $data = foreach ($rv in $view) {
        $o = [ordered]@{}
        foreach ($c in $cols) { $o[$c] = [string]$rv[$c] }
        [pscustomobject]$o
    }
    try {
        @($data) | Export-Csv -Path $dlg.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log "$($view.Count) regels geexporteerd naar $($dlg.FileName)" 'success'
        $res = [System.Windows.MessageBox]::Show("Export gereed:`n$($dlg.FileName)`n`nBestand openen?", 'Export', 'YesNo', 'Information')
        if ($res -eq 'Yes') { Start-Process $dlg.FileName }
    }
    catch { Write-Log "Exporteren mislukt: $_" 'error' }
}

#endregion

#region --- Events ---

$script:ui.BtnLogin.Add_Click({
    Set-Busy $true
    try {
        $tenant = $script:ui.TxtTenant.Text.Trim()
        $settings = Get-Settings
        Write-Log "Inloggen$(if ($tenant) { " bij tenant $tenant" })... (kies je admin-account in de browser)"
        $upn = Connect-ToGraph -Tenant $tenant -LoginHint $settings.LoginHint
        $script:Connected = $true
        try {
            $org = Invoke-Graph "$($script:GraphRoot)/organization?`$select=displayName,verifiedDomains"
            $o = @($org['value'])[0]
            $script:TenantName = [string]$o['displayName']
            $def = @($o['verifiedDomains'] | Where-Object { $_['isDefault'] })[0]
            if ($def) { $script:TenantName = $def['name'] }
        } catch { $script:TenantName = $tenant }
        $script:ui.LblStatus.Text = "Verbonden als $upn  |  tenant: $($script:TenantName)"
        $script:ui.LblStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen
        Save-Settings -LoginHint $upn -Tenant $tenant
        Write-Log "Verbonden als $upn ($($script:TenantName))" 'success'
    }
    catch {
        $script:Connected = $false
        Write-Log "Inloggen mislukt: $_" 'error'
    }
    finally { Set-Busy $false }
})

$script:ui.BtnLogout.Add_Click({
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    $script:Connected = $false; $script:Auth = $null; $script:Msal = $null
    $script:TableRows = $null; $script:TableAccounts = $null; $script:TenantName = ''
    Update-View
    $script:ui.LblStatus.Text = 'Niet verbonden'
    $script:ui.LblStatus.Foreground = [System.Windows.Media.Brushes]::Firebrick
    Write-Log 'Uitgelogd'
    Set-Busy $false
})

$script:ui.BtnLoad.Add_Click({
    try {
        Set-Busy $true
        $script:ui.Grid.ItemsSource = $null
        Get-AdminRoleData
    }
    catch { Write-Log "Ophalen mislukt: $($_.Exception.GetBaseException().Message) | $($_.InvocationInfo.PositionMessage)" 'error' }
    try { Set-Busy $false } catch { }
    if ($null -ne $script:TableRows) { Update-View }
})

$script:ui.Grid.Add_AutoGeneratingColumn({
    param($s, $e)
    # Kolomnamen met '/', '(' of spaties breken de standaard binding: via indexer binden
    try {
        if ($e.Column -is [System.Windows.Controls.DataGridBoundColumn]) {
            $e.Column.Binding = New-Object System.Windows.Data.Binding ("[$($e.PropertyName)]")
        }
        $e.Column.Header = $e.PropertyName
    } catch { Write-FileLog "AutoGeneratingColumn: $_" }
})

$script:ui.BtnExport.Add_Click({ try { Export-CurrentView } catch { Show-Crash $_ 'export' } })
$script:ui.RbRows.Add_Checked({ Update-View })
$script:ui.RbAccounts.Add_Checked({ Update-View })
$script:ui.TxtFilter.Add_TextChanged({ Update-View })

$script:Window.Add_Closed({ try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { } })

#endregion

#region --- Start ---

$settings = Get-Settings
if ($settings.Tenant) { $script:ui.TxtTenant.Text = $settings.Tenant }

$script:Window.Add_ContentRendered({
    if (-not (Initialize-Modules)) {
        Write-Log 'Module Microsoft.Graph.Authentication niet gevonden. Installeer met:' 'error'
        Write-Log '  Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser' 'error'
        $script:ui.BtnLogin.IsEnabled = $false
        return
    }
    Write-Log "Graph-module geladen (v$((Get-Module Microsoft.Graph.Authentication).Version)). Klik op 'Inloggen'."
})

# Fouten die anders het venster (en PowerShell) laten afsluiten, afvangen en tonen
$script:Window.Dispatcher.Add_UnhandledException({
    param($s, $e)
    $e.Handled = $true
    Show-Crash $e.Exception 'UI'
})
[AppDomain]::CurrentDomain.Add_UnhandledException({
    param($s, $e)
    Write-FileLog "FATAAL: $($e.ExceptionObject)"
})

Write-FileLog '===== Tool gestart ====='
try { [void]$script:Window.ShowDialog() }
catch {
    Write-FileLog "ShowDialog afgebroken: $($_.Exception.GetBaseException().Message)`r`n$($_.ScriptStackTrace)"
    [System.Windows.MessageBox]::Show("De tool is onverwacht gestopt:`n`n$($_.Exception.GetBaseException().Message)`n`nLog: $($script:LogFile)", 'Fout', 'OK', 'Error') | Out-Null
}

#endregion
