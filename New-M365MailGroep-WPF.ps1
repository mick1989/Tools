#Requires -Version 5.1
<#
.SYNOPSIS
    Microsoft 365-groep aanmaken als mailgroep (vervanger van een
    distributielijst), met een interactieve WPF-interface.
.DESCRIPTION
    Maakt een Microsoft 365-groep (Unified Group) die zich gedraagt als een
    distributielijst:
      - leden ontvangen mail aan de groep in hun eigen postvak IN
        (AutoSubscribeNewMembers + bestaande leden als subscriber)
      - geen welkomstmail (UnifiedGroupWelcomeMessageEnabled uit, ingesteld
        VOORDAT er leden worden toegevoegd)
    Optioneel: externe afzenders toestaan, agenda-uitnodigingen doorsturen,
    verbergen in Outlook / adresboek, ledenlijst verbergen, Send As of Send on
    Behalf voor de leden, extra e-mailadressen en een MailTip.

    Leden en eigenaren zoeken in Exchange, plakken als lijst e-mailadressen of
    overnemen uit een bestaande distributielijst.

    De groep wordt aangemaakt via Microsoft Graph (New-UnifiedGroup geeft in
    sommige tenants een algemene serverfout). Via resourceBehaviorOptions
    staan welkomstmail uit en abonneren aan al vanaf het aanmaken, dus ook
    leden die direct worden toegevoegd krijgen geen welkomstmail. Er wordt geen
    Team aangemaakt; de SharePoint-site optioneel pas bij eerste gebruik.

    Stappen (elk met eigen status in het voortgangsvenster):
      Graph:    groep aanmaken, leden, eigenaren
      Exchange: wachten tot het groepspostvak bestaat, e-mailadres op eigen
                domein, leden abonneren, overige instellingen, controle

    Tabblad 'Bestaande groep beheren': M365-groepen zoeken (Graph) en van de
    gekozen groep naam, beschrijving, adressen, privacy, taal, MailTip en
    mailgedrag wijzigen (alleen gewijzigde velden, per stap). Leden toevoegen /
    verwijderen, lid/eigenaar wisselen, groepsmail, Send As en Send on Behalf
    per persoon aan/uit, leden exporteren en de groep verwijderen.

    Optioneel een tenant (domein of tenant-ID) invullen om via GDAP bij een
    klant-tenant in te loggen.
.NOTES
    Vereist: Microsoft.Graph.Authentication (2.33.0 aanbevolen) en
             ExchangeOnlineManagement 3.x
      Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
    Graph-scopes: Group.ReadWrite.All, User.Read.All
    Rol: Exchange Administrator + Groups Administrator (of Global Administrator).
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

#region --- Globale state ---

$script:Window         = $null
$script:ExoConnected   = $false
$script:GraphConnected = $false
$script:GraphRoot      = "https://graph.microsoft.com/v1.0"
$script:GraphScopes    = @('Group.ReadWrite.All', 'User.Read.All')
$script:AdminUpn       = $null
$script:AdminId        = $null
$script:Picked         = New-Object System.Collections.ArrayList   # @{ Naam; Adres; Type; Rol; Id (Entra object-id) }
$script:AliasTouched  = $false
$script:SettingAlias  = $false
$script:UserTypes     = @('UserMailbox', 'SharedMailbox', 'MailUser', 'GuestMailUser')

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

function Wait-UI {
    # Wachten zonder dat het venster bevriest
    param([int]$Seconds)
    $end = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $end) { Start-Sleep -Milliseconds 200; Update-UI }
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
        "api"     { "[EXO] " }
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
    Update-UI
}

function Get-PsErrorText {
    # Dezelfde fouttekst als PowerShell in de console toont (incl. CategoryInfo
    # en FullyQualifiedErrorId), aangevuld met eventuele onderliggende fouten.
    param($Err)
    $ErrorView = 'NormalView'
    $txt = ($Err | Out-String -Width 250).Trim()
    if ($Err.ErrorDetails -and $Err.ErrorDetails.Message -and $txt -notlike "*$($Err.ErrorDetails.Message)*") {
        $txt += "`n    ErrorDetails: $($Err.ErrorDetails.Message)"
    }
    $inner = $Err.Exception.InnerException
    while ($inner) {
        $txt += "`n    ---> $($inner.GetType().FullName): $($inner.Message)"
        $inner = $inner.InnerException
    }
    return $txt
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
    $script:MainArea.IsEnabled = $true
    $script:HeaderPanel.IsEnabled = $true
    $script:Window.Cursor = $null
}

function Set-GridRows {
    param($Grid, [string[]]$Columns, $Rows)
    $t = New-Object System.Data.DataTable
    foreach ($c in $Columns) { [void]$t.Columns.Add($c) }
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $dr = $t.NewRow()
        foreach ($c in $Columns) { $dr[$c] = [string]$r[$c] }
        [void]$t.Rows.Add($dr)
    }
    $Grid.ItemsSource = $t.DefaultView
}

function Get-SelectedRows {
    param($Grid)
    return @($Grid.SelectedItems | ForEach-Object { $_.Row })
}

function ConvertTo-OPathLiteral {
    param([string]$Text)
    return $Text.Replace("'", "''")
}

#endregion

#region --- Modules en verbinding ---

function Initialize-Modules {
    # ExchangeOnlineManagement en Microsoft.Graph.Authentication gebruiken allebei
    # MSAL; de eerst geladen versie wint. EXO levert de nieuwere mee, dus EXO
    # (en zijn MSAL) eerst laden - net als in de Entra Exchange Tool.
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $base = (Get-Module ExchangeOnlineManagement).ModuleBase
        $sub  = if ($PSVersionTable.PSEdition -eq 'Core') { 'netCore' } else { 'netFramework' }
        $msal = Join-Path $base "$sub\Microsoft.Identity.Client.dll"
        if (Test-Path $msal) { try { Add-Type -Path $msal -ErrorAction Stop } catch { } }
    }
    catch { return 'ExchangeOnlineManagement' }
    # Vanaf 2.34.0 dwingt Microsoft.Graph.Authentication WAM af; 2.33.0 gebruikt nog de browser-login
    try { Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -ErrorAction Stop }
    catch {
        try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
        catch { return 'Microsoft.Graph.Authentication' }
    }
    return $null
}

function Invoke-WithoutSyncContext {
    # MSAL wacht synchroon op een async login. Met de WPF SynchronizationContext
    # actief wil die login terug naar de UI-thread die juist staat te wachten:
    # deadlock, er verschijnt nooit een inlogvenster. Daarom tijdelijk zonder context.
    param([scriptblock]$Do)
    $ctx = [System.Threading.SynchronizationContext]::Current
    [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
    try { & $Do }
    finally { [System.Threading.SynchronizationContext]::SetSynchronizationContext($ctx) }
}

function Invoke-Graph {
    param([string]$Uri, [string]$Method = 'GET', $Body = $null, [hashtable]$Headers = $null)
    if ($Uri -notmatch '^https://') { $Uri = "$($script:GraphRoot)/$Uri" }
    Write-Log "Graph $Method $([uri]::UnescapeDataString(($Uri -replace '^https://graph\.microsoft\.com/v1\.0/', '')))" "api"
    $p = @{ Uri = $Uri; Method = $Method; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $p.Body = ($Body | ConvertTo-Json -Depth 10); $p.ContentType = 'application/json' }
    if ($Headers) { $p.Headers = $Headers }
    return Invoke-MgGraphRequest @p
}

function Connect-All {
    # Eerst Graph (aanmaken groep), dan Exchange met hetzelfde account (instellingen)
    $tenant = $script:TxtTenant.Text.Trim()
    $script:GraphConnected = $false
    $script:LblExo.Text = "Niet verbonden"
    $script:LblExo.Foreground = New-Brush "#FFC864"
    try {
        Set-Status "Verbinden met Microsoft Graph - inlogvenster opent zo..."
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        $gp = @{ Scopes = $script:GraphScopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($tenant) { $gp.TenantId = $tenant }
        Invoke-WithoutSyncContext { Connect-MgGraph @gp | Out-Null }
        $ctx = Get-MgContext
        $script:AdminUpn = $ctx.Account
        $missingScopes = @($script:GraphScopes | Where-Object { $ctx.Scopes -notcontains $_ })
        if ($missingScopes.Count) { Write-Log "Graph-toestemming ontbreekt voor: $($missingScopes -join ', ') - aanmaken kan mislukken" "warning" }
        try { $script:AdminId = (Invoke-Graph 'me?$select=id').id } catch { $script:AdminId = $null }
        $script:GraphConnected = $true
        Write-Log "Verbonden met Microsoft Graph als $($script:AdminUpn) (tenant $($ctx.TenantId))" "success"
    }
    catch {
        [void]$script:Window.Activate()
        Write-Log "Verbinden met Graph mislukt: $(Get-PsErrorText $_)" "error"
        Set-Status "Verbinden mislukt" "error"
        Show-Msg "Verbinding met Microsoft Graph mislukt:`n$(Get-PsErrorText $_)" "Fout" "OK" "Error" | Out-Null
        return
    }

    try {
        Set-Status "Verbinden met Exchange Online..."
        if ($script:ExoConnected) {
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
            $script:ExoConnected = $false
        }
        $p = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($tenant) { $p.DelegatedOrganization = $tenant }
        if ($script:AdminUpn) { $p.UserPrincipalName = $script:AdminUpn }
        # Liever het normale inlogvenster (account-kiezer) dan WAM
        if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) { $p.DisableWAM = $true }
        Invoke-WithoutSyncContext { Connect-ExchangeOnline @p | Out-Null }
        [void]$script:Window.Activate()
        Enter-Busy "Domeinen laden..."
        $script:ExoConnected = $true

        $domains = @(Get-AcceptedDomain -ErrorAction Stop | Sort-Object @{ Expression = { $_.Default }; Descending = $true }, DomainName)
        $script:CmbDomain.Items.Clear()
        $script:CmbMgDomain.Items.Clear()
        foreach ($d in $domains) {
            [void]$script:CmbDomain.Items.Add([string]$d.DomainName)
            [void]$script:CmbMgDomain.Items.Add([string]$d.DomainName)
        }
        if ($script:CmbDomain.Items.Count) { $script:CmbDomain.SelectedIndex = 0 }
        Clear-MgDetail

        $org = if ($tenant) { $tenant } elseif ($domains.Count) { [string]$domains[0].DomainName } else { '' }
        $script:LblExo.Text = "Graph + Exchange: $org ($($script:AdminUpn))"
        $script:LblExo.Foreground = New-Brush "#B4F0B4"
        Write-Log "Verbonden met Exchange Online - $($domains.Count) domein(en) geladen" "success"
        Test-RequiredCmdlets
        Set-Status "Verbonden" "success"
    }
    catch {
        $script:LblExo.Text = "Alleen Graph verbonden"
        Write-Log "Verbinden met Exchange mislukt: $(Get-PsErrorText $_)" "error"
        Set-Status "Verbinden met Exchange mislukt" "error"
        [void]$script:Window.Activate()
        Show-Msg "Verbinding met Exchange Online mislukt:`n$(Get-PsErrorText $_)" "Fout" "OK" "Error" | Out-Null
    }
    finally { Exit-Busy }
}

function Test-RequiredCmdlets {
    # Exchange Online laadt alleen de cmdlets uit de rollen van het account.
    # Ontbreekt er een, dan heeft het account (nog) niet de juiste Exchange-rol.
    $need = [ordered]@{
        'Set-UnifiedGroup'            = 'e-mailadres en instellingen'
        'Add-UnifiedGroupLinks'       = 'leden en abonnementen toevoegen'
        'Get-UnifiedGroupLinks'       = 'controle leden en abonnementen'
        'Remove-UnifiedGroupLinks'    = 'beheerder uit de groep halen'
        'Get-UnifiedGroup'            = 'controle instellingen achteraf'
        'Get-Recipient'               = 'adres controleren'
        'Add-RecipientPermission'     = 'Send As (alleen als gekozen)'
        'Get-DistributionGroup'       = 'leden uit distributielijst (optioneel)'
        'Get-DistributionGroupMember' = 'leden uit distributielijst (optioneel)'
    }
    $missing = @($need.Keys | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    try {
        $mod = (Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' } | Select-Object -Last 1).ModuleName
        if ($mod) { Write-Log "Exchange-sessie heeft $(@(Get-Command -Module $mod -ErrorAction Stop).Count) cmdlets geladen" "info" }
    }
    catch { }
    if (-not $missing.Count) { Write-Log "Alle benodigde Exchange-cmdlets zijn beschikbaar" "success"; return }

    foreach ($m in $missing) { Write-Log "Cmdlet ontbreekt: $m ($($need[$m]))" "error" }
    $core = @($missing | Where-Object { 'Set-UnifiedGroup', 'Add-UnifiedGroupLinks', 'Get-Recipient' -contains $_ })
    $msg = "Deze Exchange-cmdlets zijn niet beschikbaar voor $($script:AdminUpn):`n`n  $($missing -join "`n  ")`n`n" +
           "Exchange Online laadt alleen de cmdlets uit de rollen van het account. Controleer of het account de rol " +
           "Exchange Administrator (of Global Administrator) heeft. Een net toegekende rol kan tot een uur duren; " +
           "verbind daarna opnieuw."
    if ($core.Count) { $msg += "`n`nZonder deze cmdlets kan de groep wel via Graph worden aangemaakt, maar niet volledig worden ingesteld." }
    Show-Msg $msg "Rechten ontbreken" "OK" "Warning" | Out-Null
}

function Test-Connected {
    if ($script:ExoConnected -and $script:GraphConnected) { return $true }
    Show-Msg "Eerst verbinden (rechtsboven). De tool heeft zowel Microsoft Graph als Exchange Online nodig." "Niet verbonden" "OK" "Warning" | Out-Null
    return $false
}

#endregion

#region --- Groepsgegevens ---

function ConvertTo-Alias {
    # "Afdeling Financien & Co" -> "afdeling-financien-co"
    param([string]$Name)
    $n = $Name.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $n.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne 'NonSpacingMark') { [void]$sb.Append($ch) }
    }
    $a = $sb.ToString().ToLowerInvariant()
    $a = $a -replace '\s+', '-' -replace '[^a-z0-9._-]', '' -replace '-{2,}', '-' -replace '\.{2,}', '.'
    return $a.Trim('-', '.')
}

function Get-GroupAddress {
    $alias = $script:TxtAlias.Text.Trim()
    $dom = [string]$script:CmbDomain.SelectedItem
    if (-not $alias -or -not $dom) { return $null }
    return "$alias@$dom"
}

function Update-AddressPreview {
    $addr = Get-GroupAddress
    $script:LblAddress.Text = if ($addr) { $addr } else { '-' }
    $script:LblAvail.Text = ''
}

function Test-AddressFree {
    # Geeft $null terug als alias + adres vrij zijn, anders een melding
    param([string]$Alias, [string]$Address)
    $a = ConvertTo-OPathLiteral $Alias
    $m = ConvertTo-OPathLiteral $Address
    Write-Log "Get-Recipient (controle alias/adres $Address)" "api"
    $hit = @(Get-Recipient -Filter "EmailAddresses -eq 'smtp:$m'" -ResultSize 1 -ErrorAction Stop)
    if ($hit.Count) { return "Adres $Address is al in gebruik door '$($hit[0].DisplayName)' ($($hit[0].RecipientTypeDetails))." }
    $hit = @(Get-Recipient -Filter "Alias -eq '$a'" -ResultSize 1 -ErrorAction Stop)
    if ($hit.Count) { return "Alias '$Alias' is al in gebruik door '$($hit[0].DisplayName)' ($($hit[0].RecipientTypeDetails))." }
    return $null
}

function Invoke-CheckAddress {
    if (-not (Test-Connected)) { return }
    $addr = Get-GroupAddress
    if (-not $addr) { Show-Msg "Vul een alias in en kies een domein." "Controleren" "OK" "Warning" | Out-Null; return }
    Enter-Busy "Adres controleren..."
    try {
        $msg = Test-AddressFree $script:TxtAlias.Text.Trim() $addr
        if ($msg) {
            $script:LblAvail.Text = "Bezet"
            $script:LblAvail.Foreground = New-Brush "#C42B1C"
            Write-Log $msg "warning"
        }
        else {
            $script:LblAvail.Text = "Beschikbaar"
            $script:LblAvail.Foreground = New-Brush "#107C10"
            Write-Log "$addr is beschikbaar" "success"
        }
        Set-Status "Gereed"
    }
    catch { Write-Log "Controle mislukt: $(Get-PsErrorText $_)" "error"; Set-Status "Controle mislukt" "error" }
    finally { Exit-Busy }
}

#endregion

#region --- Leden en eigenaren ---

function Update-PickedGrid {
    $rows = @($script:Picked | Sort-Object @{ Expression = { switch ($_.Rol) { 'Eigenaar' { 0 } 'Alleen eigenaar' { 1 } default { 2 } } } }, Naam)
    Set-GridRows $script:GridPicked @('Rol', 'Naam', 'Adres', 'Type') $rows
    $own = @($script:Picked | Where-Object { $_.Rol -ne 'Lid' }).Count
    $only = @($script:Picked | Where-Object { $_.Rol -eq 'Alleen eigenaar' }).Count
    $script:LblPicked.Text = "$($script:Picked.Count) gekozen, waarvan $own eigenaar$(if ($only) { " ($only geen lid)" })"
    $script:LblPicked.Foreground = if ($own -eq 0 -and $script:Picked.Count) { New-Brush "#B05A00" } else { New-Brush "#0063B1" }
}

function Add-Picked {
    # Voegt 1 ontvanger toe; een bestaand lid wordt alleen opgewaardeerd naar eigenaar
    param([string]$Naam, [string]$Adres, [string]$Type, [string]$Rol = 'Lid', [string]$Id = '')
    $hit = $script:Picked | Where-Object { $_.Adres -eq $Adres } | Select-Object -First 1
    if ($hit) {
        if ($Rol -ne 'Lid' -and $hit.Rol -eq 'Lid') { $hit.Rol = $Rol }
        return $false
    }
    [void]$script:Picked.Add(@{ Naam = $Naam; Adres = $Adres; Type = $Type; Rol = $Rol; Id = $Id })
    return $true
}

function Invoke-Search {
    if (-not (Test-Connected)) { return }
    Enter-Busy "Zoeken..."
    try {
        $rows = @(Search-Recipients $script:TxtSearch.Text.Trim())
        Set-GridRows $script:GridSearch @('Naam', 'Adres', 'Type', 'Id') $rows
        $script:LblSearch.Text = "$($rows.Count) resultaten"
        Set-Status "Gereed"
    }
    catch { Write-Log "Zoeken mislukt: $(Get-PsErrorText $_)" "error"; Set-Status "Zoeken mislukt" "error" }
    finally { Exit-Busy }
}

function Search-Recipients {
    # Gebruikers die lid kunnen zijn van een M365-groep; geeft rij-hashtables terug
    param([string]$q)
    $p = @{ RecipientTypeDetails = $script:UserTypes; ResultSize = 1000; ErrorAction = 'Stop' }
    if ($q) { $p.Anr = $q }
    Write-Log "Get-EXORecipient $(if ($q) { "-Anr '$q'" } else { '(alle gebruikers)' })" "api"
    $res = @(Get-EXORecipient @p)
    if ($q) {
        # -Anr matcht ruim (o.a. op voor-/achternaam-delen) en wordt soms
        # genegeerd; daarom zelf nafilteren: alle woorden moeten voorkomen
        # in naam, e-mailadres of alias.
        $words = @($q -split '\s+' | Where-Object { $_ })
        $res = @($res | Where-Object {
            $hay = "$($_.DisplayName) $($_.PrimarySmtpAddress) $($_.Alias)"
            -not ($words | Where-Object { $hay.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -lt 0 })
        })
    }
    if ($q -match '@') {
        # UPN (inlognaam) wijkt soms af van het e-mailadres en matcht dan niet via -Anr
        try {
            $one = Get-EXORecipient -Identity $q -ErrorAction Stop
            if ($script:UserTypes -contains [string]$one.RecipientTypeDetails -and
                -not ($res | Where-Object { [string]$_.PrimarySmtpAddress -eq [string]$one.PrimarySmtpAddress })) { $res += $one }
        }
        catch { }
    }
    foreach ($r in ($res | Sort-Object DisplayName)) {
        @{ Naam = $r.DisplayName; Adres = [string]$r.PrimarySmtpAddress; Type = [string]$r.RecipientTypeDetails; Id = [string]$r.ExternalDirectoryObjectId }
    }
}

function Add-FromSearch {
    param([string]$Rol)
    # @() nodig: bij 1 selectie rolt PowerShell de array uit tot een losse DataRow
    $sel = @(Get-SelectedRows $script:GridSearch)
    if (-not $sel.Count) { Write-Log "Selecteer eerst 1 of meer personen in de zoekresultaten" "warning"; return }
    $n = 0
    foreach ($r in $sel) { if (Add-Picked $r['Naam'] $r['Adres'] $r['Type'] $Rol $r['Id']) { $n++ } }
    Update-PickedGrid
    Write-Log "$($sel.Count) toegevoegd/bijgewerkt als $($Rol.ToLower()) ($n nieuw)" "info"
}

function Add-Me {
    param([string]$Rol)
    if (-not (Test-Connected)) { return }
    if (-not $script:AdminUpn) { Write-Log "Ingelogd account onbekend" "warning"; return }
    Enter-Busy "Ingelogd account opzoeken..."
    try {
        $r = Resolve-Recipient $script:AdminUpn
        [void](Add-Picked $r.DisplayName ([string]$r.PrimarySmtpAddress) ([string]$r.RecipientTypeDetails) $Rol ([string]$r.ExternalDirectoryObjectId))
        Update-PickedGrid
        Write-Log "$($r.DisplayName) toegevoegd als $($Rol.ToLower())" "success"
        Set-Status "Gereed"
    }
    catch {
        $why = if ($script:TxtTenant.Text.Trim()) {
            "Je bent via GDAP ingelogd; je beheeraccount bestaat niet in deze klant-tenant en kan geen lid of eigenaar worden."
        }
        else {
            "Waarschijnlijk heeft $($script:AdminUpn) geen mailbox (geen Exchange-licentie). Een M365-groep kan alleen gebruikers met mailbox of mail-gebruikers als lid hebben."
        }
        Write-Log "$($script:AdminUpn) niet gevonden: $(Get-PsErrorText $_)" "warning"
        Show-Msg "$($script:AdminUpn) kan niet worden toegevoegd.`n`n$why" "Ingelogd account" "OK" "Warning" | Out-Null
        Set-Status "Gereed"
    }
    finally { Exit-Busy }
}

function Resolve-Recipient {
    param([string]$Id)
    $r = Get-EXORecipient -Identity $Id -ErrorAction Stop
    if ($script:UserTypes -notcontains [string]$r.RecipientTypeDetails) {
        throw "type $($r.RecipientTypeDetails) kan geen lid zijn van een Microsoft 365-groep"
    }
    return $r
}

function Invoke-PasteAdd {
    if (-not (Test-Connected)) { return }
    $ids = @($script:TxtPaste.Text -split '[\s;,]+' | Where-Object { $_ } | Select-Object -Unique)
    if (-not $ids.Count) { return }
    $rol = if ($script:ChkPasteOwner.IsChecked) { 'Eigenaar' } else { 'Lid' }
    Enter-Busy "Adressen opzoeken..."
    $ok = 0; $fail = @()
    try {
        foreach ($id in $ids) {
            try {
                $r = Resolve-Recipient $id
                [void](Add-Picked $r.DisplayName ([string]$r.PrimarySmtpAddress) ([string]$r.RecipientTypeDetails) $rol ([string]$r.ExternalDirectoryObjectId))
                $ok++
            }
            catch { $fail += $id; Write-Log "$id overgeslagen: $(Get-PsErrorText $_)" "warning" }
        }
        Update-PickedGrid
        Write-Log "Plakken: $ok gevonden, $($fail.Count) niet gevonden" $(if ($fail.Count) { 'warning' } else { 'success' })
        $script:TxtPaste.Text = ($fail -join "`r`n")
        Set-Status "Gereed"
    }
    finally { Exit-Busy }
}

function Invoke-ImportDl {
    if (-not (Test-Connected)) { return }
    $dl = $script:TxtDl.Text.Trim()
    if (-not $dl) { Show-Msg "Vul de naam of het adres van de distributielijst in." "Importeren" "OK" "Warning" | Out-Null; return }
    Enter-Busy "Distributielijst lezen..."
    try {
        Write-Log "Get-DistributionGroup $dl" "api"
        $g = Get-DistributionGroup -Identity $dl -ErrorAction Stop
        Write-Log "Get-DistributionGroupMember $($g.PrimarySmtpAddress)" "api"
        $mem = @(Get-DistributionGroupMember -Identity $g.Identity -ResultSize Unlimited -ErrorAction Stop)
        $ok = 0
        foreach ($m in $mem) {
            if ($script:UserTypes -contains [string]$m.RecipientTypeDetails) {
                if (Add-Picked $m.DisplayName ([string]$m.PrimarySmtpAddress) ([string]$m.RecipientTypeDetails) 'Lid' ([string]$m.ExternalDirectoryObjectId)) { $ok++ }
            }
            else {
                Write-Log "Overgeslagen: $($m.DisplayName) ($($m.RecipientTypeDetails)) - kan geen lid zijn van een M365-groep" "warning"
            }
        }

        $own = 0
        if (@($g.ManagedBy).Count) {
            $r = Show-Msg "De distributielijst heeft $(@($g.ManagedBy).Count) beheerder(s).`nOok overnemen als eigenaar van de nieuwe groep?" "Beheerders" "YesNo" "Question"
            if ($r -eq 'Yes') {
                foreach ($mb in @($g.ManagedBy)) {
                    try {
                        $rc = Resolve-Recipient ([string]$mb)
                        [void](Add-Picked $rc.DisplayName ([string]$rc.PrimarySmtpAddress) ([string]$rc.RecipientTypeDetails) 'Eigenaar' ([string]$rc.ExternalDirectoryObjectId))
                        $own++
                    }
                    catch { Write-Log "Beheerder '$mb' overgeslagen: $(Get-PsErrorText $_)" "warning" }
                }
            }
        }

        if (-not $script:TxtName.Text.Trim()) { $script:TxtName.Text = $g.DisplayName }
        Update-PickedGrid
        Write-Log "Uit '$($g.DisplayName)': $ok nieuwe leden, $own eigenaar(s)" "success"
        Write-Log "Let op: het adres $($g.PrimarySmtpAddress) blijft bij de distributielijst. Wil je dat adres op de nieuwe groep, verwijder of hernoem dan eerst de distributielijst." "info"
        Set-Status "Gereed"
    }
    catch { Write-Log "Importeren mislukt: $(Get-PsErrorText $_)" "error"; Set-Status "Importeren mislukt" "error" }
    finally { Exit-Busy }
}

#endregion

#region --- Aanmaken ---

function Invoke-WithRetry {
    # Nieuwe groepen zijn niet direct overal gerepliceerd; vervolgstappen
    # kunnen de eerste minuut 'object not found' geven. Na de laatste poging
    # wordt de fout doorgegooid.
    param([scriptblock]$Do, [string]$Label, [int]$Tries = 8, [int]$Delay = 10)
    for ($try = 1; $try -le $Tries; $try++) {
        try { return (& $Do) }
        catch {
            if ($try -eq $Tries) { throw }
            Write-Log "$Label nog niet gelukt (poging $try/$Tries): $($_.Exception.Message) - opnieuw over $Delay s" "warning"
            Wait-UI $Delay
        }
    }
}

function Invoke-SetGroup {
    # 1 Set-UnifiedGroup-aanroep per instelling, zodat per stap duidelijk is wat lukt
    param([string]$Label, [hashtable]$Params, [int]$Tries = 6)
    $txt = ($Params.GetEnumerator() | ForEach-Object { "-$($_.Key) $($_.Value)" }) -join ' '
    Write-Log "Set-UnifiedGroup $txt" "api"
    Invoke-WithRetry -Label $Label -Tries $Tries -Do { Set-UnifiedGroup -Identity $script:Gid @Params -ErrorAction Stop } | Out-Null
}

function Add-GroupLinks {
    # Eerst in 1 keer; lukt dat niet, dan per adres zodat 1 fout de rest niet
    # blokkeert. Geeft het aantal gelukte adressen terug.
    param([string]$GroupId, [string]$LinkType, [string[]]$Links)
    if (-not $Links.Count) { return 0 }
    Write-Log "Add-UnifiedGroupLinks -LinkType $LinkType ($($Links.Count))" "api"
    try {
        Invoke-WithRetry -Label "$LinkType toevoegen" -Tries 3 -Do {
            Add-UnifiedGroupLinks -Identity $GroupId -LinkType $LinkType -Links $Links -Confirm:$false -ErrorAction Stop
        } | Out-Null
        return $Links.Count
    }
    catch { Write-Log "$LinkType in 1 keer lukte niet, nu per adres: $($_.Exception.Message)" "warning" }
    $ok = 0
    foreach ($l in $Links) {
        try { Add-UnifiedGroupLinks -Identity $GroupId -LinkType $LinkType -Links $l -Confirm:$false -ErrorAction Stop | Out-Null; $ok++ }
        catch { Write-Log "$LinkType ${l}: $(Get-PsErrorText $_)" "error" }
    }
    return $ok
}

function Get-CountResult {
    param([int]$Ok, [int]$Total, [string]$What)
    if ($Ok -eq $Total) { return "$Ok $What" }
    if ($Ok -eq 0) { return @{ Status = 'Mislukt'; Detail = "0 van $Total $What - zie logboek" } }
    return @{ Status = 'Deels gelukt'; Detail = "$Ok van $Total $What - zie logboek" }
}

function Get-FormSettings {
    $access = [string]$script:CmbAccess.SelectedItem.Tag
    if ($script:ChkHideMembership.IsChecked) { $access = 'Private' }
    $sendMode = if ($script:RbSendAs.IsChecked) { 'SendAs' } elseif ($script:RbSendOnBehalf.IsChecked) { 'SendOnBehalf' } else { 'None' }
    return @{
        Name           = $script:TxtName.Text.Trim()
        Alias          = $script:TxtAlias.Text.Trim()
        Address        = Get-GroupAddress
        Notes          = $script:TxtNotes.Text.Trim()
        Access         = $access
        Language       = [string]$script:CmbLang.SelectedItem.Tag
        Extra          = @($script:TxtExtra.Text -split '[\s;,]+' | Where-Object { $_ })
        MailTip        = $script:TxtMailTip.Text.Trim()
        Subscribe      = [bool]$script:ChkSubscribe.IsChecked
        NoWelcome      = [bool]$script:ChkNoWelcome.IsChecked
        Calendar       = [bool]$script:ChkCalendar.IsChecked
        External       = [bool]$script:ChkExternal.IsChecked
        HideOutlook    = [bool]$script:ChkHideOutlook.IsChecked
        HideGal        = [bool]$script:ChkHideGal.IsChecked
        HideMembership = [bool]$script:ChkHideMembership.IsChecked
        NoSite         = [bool]$script:ChkNoSite.IsChecked
        SendMode       = $sendMode
    }
}

function Resolve-GraphUserId {
    # Entra object-id van een gekozen lid; komt normaal al mee uit Exchange
    # (ExternalDirectoryObjectId), anders opzoeken via UPN of mail.
    param([hashtable]$Item)
    if ($Item.Id) { return $Item.Id }
    try {
        $u = Invoke-Graph "users/$([uri]::EscapeDataString($Item.Adres))?`$select=id"
        $Item.Id = [string]$u.id
        return $Item.Id
    }
    catch { }
    $u = Invoke-Graph "users?`$filter=mail eq '$($Item.Adres.Replace("'", "''"))'&`$select=id"
    if (@($u.value).Count) { $Item.Id = [string]$u.value[0].id; return $Item.Id }
    throw "Geen Entra-gebruiker gevonden voor $($Item.Adres)"
}

function Add-GraphLinks {
    # Leden of eigenaren 1 voor 1 toevoegen; geeft het aantal gelukte terug
    param([string]$GroupId, [string]$Rel, $Items)
    $ok = 0
    foreach ($it in @($Items)) {
        try {
            $id = Resolve-GraphUserId $it
            Invoke-Graph "groups/$GroupId/$Rel/`$ref" 'POST' @{ '@odata.id' = "$($script:GraphRoot)/directoryObjects/$id" } | Out-Null
            $ok++
        }
        catch {
            if ($_.Exception.Message -match 'already exist') { $ok++; continue }
            Write-Log "$Rel $($it.Adres): $(Get-PsErrorText $_)" "error"
        }
    }
    return $ok
}

#region --- Voortgang per stap ---

$script:StepDefs = [ordered]@{
    create  = 'Groep aanmaken (Graph)'
    members = 'Leden toevoegen (Graph)'
    owners  = 'Eigenaren toevoegen (Graph)'
    ready   = 'Groepspostvak beschikbaar (Exchange)'
    address = 'E-mailadres instellen'
    subs    = 'Leden abonneren op groepsmail'
    cal     = 'Agenda-uitnodigingen naar leden'
    ext     = 'Externe afzenders'
    info    = 'Taal en MailTip'
    send    = 'Verzendrechten'
    extra   = 'Extra e-mailadressen'
    hide    = 'Verbergen in adresboek'
    admin   = 'Beheeraccount uit de groep halen'
    check   = 'Controle'
}

function Open-StepsWindow {
    # $Defs: geordende key -> omschrijving; standaard de stappen voor een nieuwe groep
    param([string]$Title, $Defs = $script:StepDefs)
    $script:Steps = New-Object System.Collections.ArrayList
    $nr = 0
    foreach ($k in $Defs.Keys) {
        $nr++
        [void]$script:Steps.Add(@{ Key = $k; Nr = "$nr"; Stap = $Defs[$k]; Status = 'Wachtend'; Detail = '' })
    }
    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$script:StepsXaml)))
    $w.Owner = $script:Window
    $script:StepsWin     = $w
    $script:StepsGrid    = $w.FindName('SW_Grid')
    $script:StepsSummary = $w.FindName('SW_Summary')
    $script:StepsClose   = $w.FindName('SW_Close')
    $w.FindName('SW_Title').Text = $Title
    $script:StepsBusy = $true
    $script:StepsClose.Add_Click({ $script:StepsWin.Close() })
    $w.Add_Closing({ if ($script:StepsBusy) { $_.Cancel = $true } })
    Update-StepsGrid
    $w.Show()
    Update-UI
}

function Update-StepsGrid {
    Set-GridRows $script:StepsGrid @('Nr', 'Stap', 'Status', 'Detail') $script:Steps
    Update-UI
}

function Set-Step {
    param([string]$Key, [string]$Status, [string]$Detail = '')
    $st = $script:Steps | Where-Object { $_.Key -eq $Key } | Select-Object -First 1
    if (-not $st) { return }
    $st.Status = $Status
    $st.Detail = $Detail
    if ($Status -eq 'Bezig') { Set-Status "$($st.Stap)..." }
    else {
        $type = switch ($Status) { 'Gelukt' { 'success' } 'Mislukt' { 'error' } 'Deels gelukt' { 'warning' } default { 'info' } }
        Write-Log "Stap $($st.Nr) '$($st.Stap)': $Status$(if ($Detail) { " - $Detail" })" $type
    }
    Update-StepsGrid
}

function Skip-Step {
    param([string]$Key, [string]$Why = 'niet gekozen')
    Set-Step $Key 'Overgeslagen' $Why
}

function Invoke-Step {
    # Voert 1 stap uit. Het scriptblock geeft een detailtekst terug (= Gelukt)
    # of een hashtable @{ Status; Detail } voor Deels gelukt / Mislukt / Overgeslagen.
    # Een exception = Mislukt. Geeft $true terug tenzij de stap mislukt is.
    param([string]$StepKey, [scriptblock]$StepDo)
    Set-Step $StepKey 'Bezig'
    try {
        $r = & $StepDo
        if ($r -is [array]) { $r = $r[-1] }
        if ($r -is [hashtable]) { Set-Step $StepKey $r.Status $r.Detail; return ($r.Status -ne 'Mislukt') }
        Set-Step $StepKey 'Gelukt' ([string]$r)
        return $true
    }
    catch {
        $label = ($script:Steps | Where-Object { $_.Key -eq $StepKey } | Select-Object -First 1).Stap
        Write-Log "Fout bij '$label': $(Get-PsErrorText $_)" "error"
        Set-Step $StepKey 'Mislukt' $_.Exception.Message
        return $false
    }
}

function Complete-Steps {
    # Resterende stappen op Overgeslagen, samenvatting tonen en venster sluitbaar maken
    param([string]$StopReason)
    foreach ($st in $script:Steps) {
        if ($st.Status -eq 'Wachtend') { $st.Status = 'Overgeslagen'; $st.Detail = $StopReason }
    }
    Update-StepsGrid
    $bad = @($script:Steps | Where-Object { 'Mislukt', 'Deels gelukt' -contains $_.Status })
    $good = @($script:Steps | Where-Object { $_.Status -eq 'Gelukt' }).Count
    if ($StopReason) {
        $txt = "Gestopt: $StopReason"; $col = '#C42B1C'; Set-Status "Gestopt - zie stappen en logboek" "error"
    }
    elseif ($bad.Count) {
        $txt = "Klaar met $($bad.Count) probleem/problemen: $(($bad | ForEach-Object { $_.Stap }) -join ', ')"; $col = '#B05A00'; Set-Status "Klaar met problemen - zie stappen en logboek" "error"
    }
    else {
        $txt = "Alles gelukt ($good stappen). Het kan tot een uur duren voordat de groep overal zichtbaar is."; $col = '#107C10'; Set-Status "Groep aangemaakt" "success"
    }
    $script:StepsSummary.Text = $txt
    $script:StepsSummary.Foreground = New-Brush $col
    $script:StepsBusy = $false
    $script:StepsClose.IsEnabled = $true
}

#endregion

function Invoke-CreateGroup {
    if (-not (Test-Connected)) { return }
    $s = Get-FormSettings
    # 'Alleen eigenaar' = eigenaar zonder lidmaatschap (ontvangt geen groepsmail)
    $ownerItems  = @($script:Picked | Where-Object { $_.Rol -ne 'Lid' } | Sort-Object @{ Expression = { $_.Rol -ne 'Eigenaar' } })
    $memberItems = @($script:Picked | Where-Object { $_.Rol -ne 'Alleen eigenaar' })
    $members     = @($memberItems | ForEach-Object { $_.Adres })

    $errors = @()
    if (-not $s.Name) { $errors += "Weergavenaam is verplicht." }
    if (-not $s.Alias) { $errors += "Alias is verplicht." }
    elseif ($s.Alias -notmatch '^[A-Za-z0-9!#$%&''*+\-/=?^_`{|}~.]+$' -or $s.Alias -match '^\.|\.$|\.\.') { $errors += "Alias bevat ongeldige tekens." }
    if (-not $s.Address) { $errors += "Kies een domein." }
    if (-not $ownerItems.Count) { $errors += "Kies minimaal 1 eigenaar." }
    foreach ($e in $s.Extra) { if ($e -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $errors += "Ongeldig extra adres: $e" } }
    if ($errors.Count) { Show-Msg ($errors -join "`n") "Controleer het formulier" "OK" "Warning" | Out-Null; return }

    Enter-Busy "Controleren..."
    try {
        $taken = Test-AddressFree $s.Alias $s.Address
        foreach ($e in $s.Extra) {
            if ($taken) { break }
            $hit = @(Get-Recipient -Filter "EmailAddresses -eq 'smtp:$(ConvertTo-OPathLiteral $e)'" -ResultSize 1 -ErrorAction Stop)
            if ($hit.Count) { $taken = "Extra adres $e is al in gebruik door '$($hit[0].DisplayName)'." }
        }
    }
    catch { $taken = "Controle mislukt: $(Get-PsErrorText $_)" }
    finally { Exit-Busy }
    if ($taken) { Write-Log $taken "warning"; Show-Msg $taken "Kan niet aanmaken" "OK" "Warning" | Out-Null; Set-Status "Gereed"; return }

    $sendTxt = @{ None = 'geen'; SendAs = 'Send As voor alle leden'; SendOnBehalf = 'Send on Behalf voor alle leden' }[$s.SendMode]
    $jn = { param($b) if ($b) { 'ja' } else { 'nee' } }
    $summary = @"
Naam:              $($s.Name)
Adres:             $($s.Address)
Privacy:           $($s.Access)
Eigenaren:         $($ownerItems.Count) (waarvan $(@($ownerItems | Where-Object { $_.Rol -eq 'Alleen eigenaar' }).Count) geen lid)
Leden (incl. eig.): $($memberItems.Count)

Mail naar postvak leden:  $(& $jn $s.Subscribe)
Geen welkomstmail:        $(& $jn $s.NoWelcome)
Agenda-uitnodigingen:     $(& $jn $s.Calendar)
Externe afzenders:        $(& $jn $s.External)
Verbergen in Outlook:     $(& $jn $s.HideOutlook)
Verbergen in adresboek:   $(& $jn $s.HideGal)
Ledenlijst verborgen:     $(& $jn $s.HideMembership)
SharePoint-site pas later: $(& $jn $s.NoSite)
Verzendrechten leden:     $sendTxt

De groep wordt via Graph aangemaakt (zonder Team), daarna
volgen de instellingen stap voor stap. Groep nu aanmaken?
"@
    if ((Show-Msg $summary "Groep aanmaken" "YesNo" "Question") -ne 'Yes') { return }

    Enter-Busy "Groep aanmaken..."
    Open-StepsWindow "$($s.Name)  <$($s.Address)>"
    $script:GroupObjId = $null
    $script:Gid = $null
    try {
        # --- 1. Groep via Graph. Welkomstmail uit en abonneren staan via
        #        resourceBehaviorOptions vanaf het begin goed (kan alleen bij aanmaken).
        $ok = Invoke-Step 'create' {
            $fid = Resolve-GraphUserId $ownerItems[0]
            $opts = @()
            if ($s.NoWelcome)    { $opts += 'WelcomeEmailDisabled' }
            if ($s.Subscribe)    { $opts += 'SubscribeNewGroupMembers' }
            if (-not $s.Calendar) { $opts += 'SubscribeMembersToCalendarEventsDisabled' }
            if ($s.HideOutlook)  { $opts += 'HideGroupInOutlook' }
            if ($s.NoSite)       { $opts += 'ProvisionSiteOnDemand' }
            $bind = "$($script:GraphRoot)/users/$fid"
            $body = [ordered]@{
                displayName             = $s.Name
                mailNickname            = $s.Alias
                mailEnabled             = $true
                securityEnabled         = $false
                groupTypes              = @('Unified')
                visibility              = $(if ($s.HideMembership) { 'HiddenMembership' } elseif ($s.Access -eq 'Public') { 'Public' } else { 'Private' })
                resourceBehaviorOptions = $opts
                'owners@odata.bind'     = @($bind)
            }
            if ($ownerItems[0].Rol -eq 'Eigenaar') { $body.'members@odata.bind' = @($bind) }
            if ($s.Notes) { $body.description = $s.Notes }
            $g = Invoke-Graph 'groups' 'POST' $body
            $script:GroupObjId = [string]$g.id
            "id $($script:GroupObjId) - $(if ($opts.Count) { $opts -join ', ' } else { 'geen extra opties' })"
        }
        if (-not $ok) {
            Complete-Steps "groep niet aangemaakt"
            Show-Msg "De groep kon niet worden aangemaakt via Graph. Zie het voortgangsvenster en het logboek." "Mislukt" "OK" "Error" | Out-Null
            return
        }

        # --- 2. Leden en eigenaren via Graph (welkomstmail staat al uit) ---
        $first = $ownerItems[0].Adres
        $addMembers = @($memberItems | Where-Object { $_.Adres -ne $first })
        $addOwners  = @($ownerItems | Where-Object { $_.Adres -ne $first })
        if ($addMembers.Count) { [void](Invoke-Step 'members' { Get-CountResult (Add-GraphLinks $script:GroupObjId 'members' $addMembers) $addMembers.Count 'leden' }) }
        else { Skip-Step 'members' $(if ($memberItems.Count) { 'alleen de eigenaar (bij aanmaken)' } else { 'geen leden gekozen' }) }
        if ($addOwners.Count) { [void](Invoke-Step 'owners' { Get-CountResult (Add-GraphLinks $script:GroupObjId 'owners' $addOwners) $addOwners.Count 'eigenaren' }) }
        else { Set-Step 'owners' 'Gelukt' "$first (bij aanmaken)" }

        # --- 3. Wachten tot Exchange het groepspostvak heeft aangemaakt ---
        $ok = Invoke-Step 'ready' {
            $found = Invoke-WithRetry -Label "Groepspostvak zoeken" -Tries 30 -Delay 10 -Do {
                $hit = @(Get-Recipient -Filter "Alias -eq '$(ConvertTo-OPathLiteral $s.Alias)'" -RecipientTypeDetails GroupMailbox -ResultSize 1 -ErrorAction Stop)
                if (-not $hit.Count) { throw "Groepspostvak nog niet aangemaakt" }
                $hit[0]
            }
            $script:Gid = if ($found.Guid) { [string]$found.Guid } else { [string]$found.PrimarySmtpAddress }
            $script:CurrentSmtp = [string]$found.PrimarySmtpAddress
            "gevonden ($($script:CurrentSmtp))"
        }
        if (-not $ok) {
            Complete-Steps "groepspostvak nog niet beschikbaar"
            Show-Msg "De groep en leden zijn aangemaakt (welkomstmail staat al uit), maar Exchange heeft het groepspostvak na 5 minuten nog niet klaar.`n`nDe Exchange-instellingen (o.a. e-mailadres $($s.Address)) zijn nog niet gezet. Probeer die later opnieuw of stel ze in via het Exchange-beheercentrum." "Gestopt" "OK" "Warning" | Out-Null
            return
        }

        # --- 4. Exchange-instellingen, elk apart ---
        if ($script:CurrentSmtp -ne $s.Address) {
            [void](Invoke-Step 'address' { Invoke-SetGroup 'E-mailadres' @{ PrimarySmtpAddress = $s.Address } 6; "$($s.Address) (was $($script:CurrentSmtp))" })
        }
        else { Set-Step 'address' 'Gelukt' "$($s.Address) (al juist)" }

        if ($s.Subscribe) {
            [void](Invoke-Step 'subs' {
                # SubscribeNewGroupMembers geldt voor nieuwe leden; leden die al via Graph
                # zijn toegevoegd voordat het groepspostvak bestond, hier expliciet abonneren.
                $subsNow = @()
                try { $subsNow = @(Get-UnifiedGroupLinks -Identity $script:Gid -LinkType Subscribers -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { [string]$_.PrimarySmtpAddress }) }
                catch { Write-Log "Huidige subscribers ophalen lukte niet - alle leden worden als subscriber toegevoegd" "info" }
                $todo = @($members | Where-Object { $subsNow -notcontains $_ })
                if (-not $todo.Count) { return "alle $($members.Count) leden al geabonneerd" }
                $n = Add-GroupLinks $script:Gid 'Subscribers' $todo
                Get-CountResult ($n + $members.Count - $todo.Count) $members.Count 'leden geabonneerd'
            })
        }
        else { Skip-Step 'subs' }

        if ($s.Calendar) { [void](Invoke-Step 'cal' { Invoke-SetGroup 'Agenda-uitnodigingen' @{ AlwaysSubscribeMembersToCalendarEvents = $true }; 'aan' }) }
        else { Skip-Step 'cal' 'uit (bij aanmaken)' }

        if ($s.External) { [void](Invoke-Step 'ext' { Invoke-SetGroup 'Externe afzenders' @{ RequireSenderAuthenticationEnabled = $false }; 'extern toegestaan' }) }
        else { Skip-Step 'ext' 'alleen interne afzenders (standaard)' }

        $info = @{}
        if ($s.Language) { $info.Language = $s.Language }
        if ($s.MailTip) { $info.MailTip = $s.MailTip }
        if ($info.Count) { [void](Invoke-Step 'info' { Invoke-SetGroup 'Taal/MailTip' $info; ($info.Keys -join ', ') }) }
        else { Skip-Step 'info' 'niets ingevuld' }

        if ($s.SendMode -eq 'SendAs') {
            [void](Invoke-Step 'send' {
                $n = 0
                foreach ($m in $members) {
                    Write-Log "Add-RecipientPermission $m SendAs" "api"
                    try { Add-RecipientPermission -Identity $script:Gid -Trustee $m -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null; $n++ }
                    catch { Write-Log "Send As ${m}: $(Get-PsErrorText $_)" "error" }
                }
                Get-CountResult $n $members.Count 'leden Send As'
            })
        }
        elseif ($s.SendMode -eq 'SendOnBehalf') {
            [void](Invoke-Step 'send' { Invoke-SetGroup 'Send on Behalf' @{ GrantSendOnBehalfTo = @{ Add = $members } } 3; "$($members.Count) leden Send on Behalf" })
        }
        else { Skip-Step 'send' }

        if ($s.Extra.Count) {
            [void](Invoke-Step 'extra' {
                Invoke-SetGroup 'Extra adressen' @{ EmailAddresses = @{ Add = @($s.Extra | ForEach-Object { "smtp:$_" }) } } 3
                ($s.Extra -join ', ')
            })
        }
        else { Skip-Step 'extra' 'geen' }

        if ($s.HideOutlook -or $s.HideGal) {
            # Verbergen in Outlook verbergt de groep ook in het adresboek; expliciet naar de keuze zetten
            [void](Invoke-Step 'hide' {
                Invoke-SetGroup 'Adresboek-zichtbaarheid' @{ HiddenFromAddressListsEnabled = $s.HideGal } 3
                $(if ($s.HideGal) { 'verborgen' } else { 'zichtbaar' })
            })
        }
        else { Skip-Step 'hide' }

        # --- 5. Beheeraccount eruit als dat automatisch is toegevoegd ---
        # Gekozen als 'Alleen eigenaar': alleen het automatische lidmaatschap weghalen
        $adminPick = @($script:Picked | Where-Object { ($script:AdminId -and $_.Id -eq $script:AdminId) -or $_.Adres -eq $script:AdminUpn }) | Select-Object -First 1
        $adminChosen = $adminPick -and $adminPick.Rol -ne 'Alleen eigenaar'
        if ($script:AdminId -and -not $adminChosen) {
            [void](Invoke-Step 'admin' {
                $done = @()
                $own = @((Invoke-Graph "groups/$($script:GroupObjId)/owners?`$select=id").value | ForEach-Object { [string]$_.id })
                if (-not $adminPick -and $own -contains $script:AdminId) {
                    Invoke-Graph "groups/$($script:GroupObjId)/owners/$($script:AdminId)/`$ref" 'DELETE' | Out-Null
                    $done += 'eigenaar'
                }
                $mem = @((Invoke-Graph "groups/$($script:GroupObjId)/members?`$select=id&`$top=999").value | ForEach-Object { [string]$_.id })
                if ($mem -contains $script:AdminId) {
                    Invoke-Graph "groups/$($script:GroupObjId)/members/$($script:AdminId)/`$ref" 'DELETE' | Out-Null
                    $done += 'lid'
                }
                if ($done.Count) { "verwijderd als $($done -join ' en ')" } else { 'stond er niet in' }
            })
        }
        else { Skip-Step 'admin' $(if ($adminChosen) { 'beheeraccount is gekozen als lid' } else { 'niet van toepassing' }) }

        # --- 6. Controle ---
        [void](Invoke-Step 'check' {
            $warn = Test-CreatedGroup $script:Gid $s $members
            if ($warn) { @{ Status = 'Deels gelukt'; Detail = "$warn afwijking(en) - zie logboek" } } else { 'alles in orde' }
        })

        Complete-Steps ''
    }
    catch {
        Write-Log "Onverwachte fout: $(Get-PsErrorText $_)" "error"
        Complete-Steps "onverwachte fout - zie logboek"
    }
    finally {
        $script:StepsBusy = $false
        if ($script:StepsClose) { $script:StepsClose.IsEnabled = $true }
        Exit-Busy
    }
}

function Test-CreatedGroup {
    # Geeft het aantal afwijkingen terug (0 = alles in orde)
    param([string]$GroupId, [hashtable]$S, [string[]]$Members)
    $warn = 0
    try {
        if (Get-Command Get-UnifiedGroup -ErrorAction SilentlyContinue) {
            Write-Log "Get-UnifiedGroup (controle)" "api"
            $v = Invoke-WithRetry -Label "Groep ophalen" -Tries 3 -Do { Get-UnifiedGroup -Identity $GroupId -ErrorAction Stop }
            $chk = @(
                @{ L = 'Welkomstmail uit';        Ok = (-not $S.NoWelcome) -or (-not $v.WelcomeMessageEnabled) }
                @{ L = 'AutoSubscribeNewMembers'; Ok = (-not $S.Subscribe) -or [bool]$v.AutoSubscribeNewMembers }
                @{ L = 'Privacy';                 Ok = [string]$v.AccessType -eq $S.Access }
                @{ L = 'Externe afzenders';       Ok = ([bool]$v.RequireSenderAuthenticationEnabled) -ne $S.External }
                @{ L = 'Verbergen in Outlook';    Ok = [bool]$v.HiddenFromExchangeClientsEnabled -eq $S.HideOutlook }
            )
            foreach ($c in $chk) {
                if (-not $c.Ok) { $warn++ }
                Write-Log "Controle $($c.L): $(if ($c.Ok) { 'OK' } else { 'AFWIJKEND' })" $(if ($c.Ok) { 'success' } else { 'warning' })
            }
        }
        else { Write-Log "Get-UnifiedGroup niet beschikbaar - controle van de instellingen overgeslagen" "warning" }

        $memNow = @(Get-UnifiedGroupLinks -Identity $GroupId -LinkType Members -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { [string]$_.PrimarySmtpAddress })
        $missing = @($Members | Where-Object { $memNow -notcontains $_ })
        if ($missing.Count) { $warn++; Write-Log "Geen lid (nog niet?): $($missing -join ', ')" "warning" }
        else { Write-Log "Leden: $($memNow.Count) - alle gekozen leden aanwezig" "success" }

        if ($S.Subscribe) {
            $subNow = @(Get-UnifiedGroupLinks -Identity $GroupId -LinkType Subscribers -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { [string]$_.PrimarySmtpAddress })
            $noSub = @($memNow | Where-Object { $subNow -notcontains $_ })
            if ($noSub.Count) { $warn++; Write-Log "Ontvangt (nog) geen groepsmail in postvak: $($noSub -join ', ')" "warning" }
            else { Write-Log "Alle $($memNow.Count) leden ontvangen groepsmail in hun postvak" "success" }
        }
    }
    catch { $warn++; Write-Log "Controle mislukt: $(Get-PsErrorText $_)" "warning" }
    return $warn
}

function Reset-Form {
    $script:SettingAlias = $true
    $script:TxtName.Text = ''; $script:TxtAlias.Text = ''; $script:TxtNotes.Text = ''
    $script:TxtExtra.Text = ''; $script:TxtMailTip.Text = ''
    $script:SettingAlias = $false
    $script:AliasTouched = $false
    $script:CmbAccess.SelectedIndex = 0
    $script:CmbLang.SelectedIndex = 0
    $script:ChkSubscribe.IsChecked = $true; $script:ChkNoWelcome.IsChecked = $true
    $script:ChkCalendar.IsChecked = $false; $script:ChkExternal.IsChecked = $false
    $script:ChkHideOutlook.IsChecked = $false; $script:ChkHideGal.IsChecked = $false
    $script:ChkHideMembership.IsChecked = $false
    $script:ChkNoSite.IsChecked = $true
    $script:RbSendNone.IsChecked = $true
    $script:Picked.Clear()
    Update-PickedGrid
    Update-AddressPreview
    Set-Status "Gereed"
}

#endregion

#region --- Bestaande groep beheren ---

# $script:Mg = @{ Id (Entra); Gid (Exchange-identity); Mail; Name; Orig (formwaarden zoals geladen); HiddenMembership }
$script:Mg        = $null
$script:MgRows    = @()
$script:MgSelIdx  = -1
$script:MgSuppress = $false

function ConvertFrom-MailTip {
    # Exchange bewaart een MailTip als HTML
    param($Html)
    if (-not $Html) { return '' }
    return [System.Net.WebUtility]::HtmlDecode(([string]$Html -replace '<[^>]+>', '')).Trim()
}

function Get-NormalizedList {
    param([string]$Text)
    return (@($Text -split '[\s;,]+' | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique) -join "`r`n")
}

function Invoke-GroupSearch {
    if (-not (Test-Connected)) { return }
    $q = $script:TxtMgSearch.Text.Trim()
    Enter-Busy "Groepen zoeken..."
    try {
        $sel = 'id,displayName,mail,visibility,resourceProvisioningOptions'
        $flt = [uri]::EscapeDataString("groupTypes/any(c:c eq 'Unified')")
        $headers = $null
        if ($q) {
            $srch = [uri]::EscapeDataString("`"displayName:$($q.Replace('"', ''))`" OR `"mail:$($q.Replace('"', ''))`"")
            $uri = "groups?`$filter=$flt&`$search=$srch&`$select=$sel&`$top=200&`$count=true"
            $headers = @{ ConsistencyLevel = 'eventual' }
        }
        else { $uri = "groups?`$filter=$flt&`$select=$sel&`$top=999" }
        $all = @()
        $page = 0
        while ($uri -and $page -lt 10) {
            $r = Invoke-Graph $uri -Headers $headers
            $all += @($r.value)
            $uri = $r.'@odata.nextLink'
            $page++
        }
        $rows = foreach ($g in ($all | Sort-Object { $_.displayName })) {
            @{ Naam = $g.displayName; Adres = $g.mail; Team = $(if (@($g.resourceProvisioningOptions) -contains 'Team') { 'ja' } else { '' }); Id = $g.id }
        }
        $script:MgSuppress = $true
        Set-GridRows $script:GridGroups @('Naam', 'Adres', 'Team', 'Id') $rows
        $script:MgSuppress = $false
        $script:MgSelIdx = -1
        $script:LblMgSearch.Text = "$(@($rows).Count) groep(en)"
        Set-Status "Gereed"
    }
    catch { Write-Log "Groepen zoeken mislukt: $(Get-PsErrorText $_)" "error"; Set-Status "Zoeken mislukt" "error" }
    finally { Exit-Busy }
}

function Open-ManagedGroup {
    param([string]$Id)
    Enter-Busy "Groep laden..."
    try {
        $gg = Invoke-Graph "groups/$Id`?`$select=id,displayName,description,mail,visibility,resourceProvisioningOptions,createdDateTime"
        Write-Log "Get-UnifiedGroup $($gg.mail)" "api"
        $g = Get-UnifiedGroup -Identity $gg.mail -IncludeAllProperties -ErrorAction Stop

        $o = @{
            Name        = [string]$gg.displayName
            Notes       = [string]$gg.description
            Primary     = ([string]$g.PrimarySmtpAddress).ToLowerInvariant()
            Extra       = Get-NormalizedList ((@($g.EmailAddresses) | Where-Object { $_ -cmatch '^smtp:' } | ForEach-Object { $_.Substring(5) }) -join ' ')
            Access      = [string]$g.AccessType
            Language    = [string]$g.Language
            MailTip     = ConvertFrom-MailTip $g.MailTip
            Subscribe   = [bool]$g.AutoSubscribeNewMembers
            NoWelcome   = -not [bool]$g.WelcomeMessageEnabled
            Calendar    = [bool]$g.AlwaysSubscribeMembersToCalendarEvents
            External    = -not [bool]$g.RequireSenderAuthenticationEnabled
            HideOutlook = [bool]$g.HiddenFromExchangeClientsEnabled
            HideGal     = [bool]$g.HiddenFromAddressListsEnabled
        }
        $script:Mg = @{
            Id               = [string]$gg.id
            Gid              = $(if ($g.Guid) { [string]$g.Guid } else { [string]$gg.mail })
            Mail             = [string]$gg.mail
            Name             = [string]$gg.displayName
            Orig             = $o
            HiddenMembership = [bool]$g.HiddenGroupMembershipEnabled
        }
        Set-MgForm $o

        $team = if (@($gg.resourceProvisioningOptions) -contains 'Team') { 'ja' } else { 'nee' }
        $created = try { ([datetime]$gg.createdDateTime).ToString('dd-MM-yyyy') } catch { '?' }
        $script:LblMgTitle.Text = $gg.displayName
        $script:LblMgInfo.Text = "Team: $team  |  Ledenlijst verborgen: $(if ($script:Mg.HiddenMembership) { 'ja' } else { 'nee' })  |  Aangemaakt: $created  |  Id: $($gg.id)"
        Update-MgMembers
        $script:MgDetail.IsEnabled = $true
        Write-Log "Groep '$($gg.displayName)' geladen" "success"
        Set-Status "Gereed"
    }
    catch {
        Write-Log "Groep laden mislukt: $(Get-PsErrorText $_)" "error"
        Set-Status "Groep laden mislukt" "error"
    }
    finally { Exit-Busy }
}

function Set-MgForm {
    param([hashtable]$o)
    $script:TxtMgName.Text = $o.Name
    $script:TxtMgNotes.Text = $o.Notes
    $local, $dom = $o.Primary -split '@', 2
    $script:TxtMgAlias.Text = $local
    if ($dom -and -not ($script:CmbMgDomain.Items -contains $dom)) { [void]$script:CmbMgDomain.Items.Add($dom) }
    $script:CmbMgDomain.SelectedItem = $dom
    $script:TxtMgExtra.Text = $o.Extra
    $script:CmbMgAccess.SelectedIndex = if ($o.Access -eq 'Public') { 1 } else { 0 }
    $script:CmbMgAccess.IsEnabled = -not $script:Mg.HiddenMembership
    $script:CmbMgLang.Text = $o.Language
    $script:TxtMgMailTip.Text = $o.MailTip
    $script:ChkMgSubscribe.IsChecked = $o.Subscribe
    $script:ChkMgNoWelcome.IsChecked = $o.NoWelcome
    $script:ChkMgCalendar.IsChecked = $o.Calendar
    $script:ChkMgExternal.IsChecked = $o.External
    $script:ChkMgHideOutlook.IsChecked = $o.HideOutlook
    $script:ChkMgHideGal.IsChecked = $o.HideGal
}

function Get-MgFormValues {
    $dom = [string]$script:CmbMgDomain.SelectedItem
    return @{
        Name        = $script:TxtMgName.Text.Trim()
        Notes       = $script:TxtMgNotes.Text.Trim()
        Primary     = "$($script:TxtMgAlias.Text.Trim())@$dom".ToLowerInvariant()
        Extra       = Get-NormalizedList $script:TxtMgExtra.Text
        Access      = [string]$script:CmbMgAccess.SelectedItem.Tag
        Language    = $script:CmbMgLang.Text.Trim()
        MailTip     = $script:TxtMgMailTip.Text.Trim()
        Subscribe   = [bool]$script:ChkMgSubscribe.IsChecked
        NoWelcome   = [bool]$script:ChkMgNoWelcome.IsChecked
        Calendar    = [bool]$script:ChkMgCalendar.IsChecked
        External    = [bool]$script:ChkMgExternal.IsChecked
        HideOutlook = [bool]$script:ChkMgHideOutlook.IsChecked
        HideGal     = [bool]$script:ChkMgHideGal.IsChecked
    }
}

function Get-MgChanges {
    # Namen van gewijzigde velden t.o.v. de geladen groep
    if (-not $script:Mg) { return @() }
    $o = $script:Mg.Orig
    $n = Get-MgFormValues
    $ch = @()
    foreach ($k in $o.Keys) {
        if ($k -eq 'Language' -and -not $n.Language) { continue }   # taal kan niet leeg gemaakt worden
        if ([string]$o[$k] -cne [string]$n[$k]) { $ch += $k }
    }
    return $ch
}

function Test-MgLeave {
    # $true als er geen onopgeslagen wijzigingen zijn of de gebruiker die wil weggooien
    $ch = @(Get-MgChanges)
    if (-not $ch.Count) { return $true }
    $r = Show-Msg "Er zijn $($ch.Count) niet-opgeslagen wijziging(en) voor '$($script:Mg.Name)'.`n`nDeze wijzigingen weggooien?" "Niet opgeslagen" "YesNo" "Warning"
    return ($r -eq 'Yes')
}

function Update-MgMembers {
    $gid = $script:Mg.Gid
    Write-Log "Get-UnifiedGroupLinks Members/Owners/Subscribers" "api"
    $mem = @(Get-UnifiedGroupLinks -Identity $gid -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
    $own = @(Get-UnifiedGroupLinks -Identity $gid -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
    $sub = @(Get-UnifiedGroupLinks -Identity $gid -LinkType Subscribers -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { ([string]$_.PrimarySmtpAddress).ToLowerInvariant() })
    $ownSet = @($own | ForEach-Object { ([string]$_.PrimarySmtpAddress).ToLowerInvariant() })
    $memSet = @($mem | ForEach-Object { ([string]$_.PrimarySmtpAddress).ToLowerInvariant() })

    $sendAs = @()
    if (Get-Command Get-RecipientPermission -ErrorAction SilentlyContinue) {
        try {
            $sendAs = @(Get-RecipientPermission -Identity $gid -AccessRights SendAs -ErrorAction Stop |
                Where-Object { [string]$_.Trustee -notlike 'NT AUTHORITY*' } | ForEach-Object { ([string]$_.Trustee).ToLowerInvariant() })
        }
        catch { Write-Log "Send As ophalen mislukt: $($_.Exception.Message)" "warning" }
    }
    $sob = @()
    try {
        $g = Get-UnifiedGroup -Identity $gid -ErrorAction Stop
        foreach ($x in @($g.GrantSendOnBehalfTo)) {
            if (-not $x) { continue }
            try { $sob += ([string](Get-Recipient -Identity ([string]$x) -ErrorAction Stop).PrimarySmtpAddress).ToLowerInvariant() }
            catch { $sob += ([string]$x).ToLowerInvariant() }
        }
    }
    catch { Write-Log "Send on Behalf ophalen mislukt: $($_.Exception.Message)" "warning" }

    $byAddr = [ordered]@{}
    foreach ($m in @($mem) + @($own)) {
        $a = ([string]$m.PrimarySmtpAddress).ToLowerInvariant()
        if ($a -and -not $byAddr.Contains($a)) { $byAddr[$a] = $m }
    }
    $rows = New-Object System.Collections.ArrayList
    $seen = @()
    foreach ($a in $byAddr.Keys) {
        $m = $byAddr[$a]
        $upn = ([string]$m.WindowsLiveID).ToLowerInvariant()
        $seen += $a; if ($upn) { $seen += $upn }
        [void]$rows.Add(@{
            Rol        = $(if ($ownSet -notcontains $a) { 'Lid' } elseif ($memSet -contains $a) { 'Eigenaar' } else { 'Alleen eigenaar' })
            Naam       = [string]$m.DisplayName
            Adres      = $a
            Groepsmail = $(if ($sub -contains $a) { 'ja' } else { '' })
            SendAs     = $(if ($sendAs -contains $a -or ($upn -and $sendAs -contains $upn)) { 'ja' } else { '' })
            Namens     = $(if ($sob -contains $a) { 'ja' } else { '' })
            Id         = [string]$m.ExternalDirectoryObjectId
            Upn        = $upn
        })
    }
    # Rechten voor accounts die geen lid zijn ook tonen, zodat ze weg te halen zijn
    foreach ($t in (@($sendAs) + @($sob) | Sort-Object -Unique)) {
        if ($seen -contains $t) { continue }
        [void]$rows.Add(@{ Rol = '(geen lid)'; Naam = $t; Adres = $t; Groepsmail = ''; SendAs = $(if ($sendAs -contains $t) { 'ja' } else { '' }); Namens = $(if ($sob -contains $t) { 'ja' } else { '' }); Id = ''; Upn = $t })
    }
    $script:MgRows = @($rows | Sort-Object @{ Expression = { switch ($_.Rol) { 'Eigenaar' { 0 } 'Alleen eigenaar' { 1 } 'Lid' { 2 } default { 3 } } } }, @{ Expression = { $_.Naam } })
    Set-GridRows $script:GridMgMembers @('Rol', 'Naam', 'Adres', 'Groepsmail', 'SendAs', 'Namens', 'Id', 'Upn') $script:MgRows

    $nm = @($script:MgRows | Where-Object { $_.Rol -eq 'Lid' -or $_.Rol -eq 'Eigenaar' }).Count
    $no = @($script:MgRows | Where-Object { $_.Rol -eq 'Eigenaar' -or $_.Rol -eq 'Alleen eigenaar' }).Count
    $ns = @($script:MgRows | Where-Object { $_.Groepsmail -eq 'ja' }).Count
    $script:LblMgMembers.Text = "$nm leden, $no eigenaar, $ns ontvangen groepsmail in hun postvak"
    $script:LblMgMembers.Foreground = if ($script:Mg.Orig.Subscribe -and $ns -lt $nm) { New-Brush "#B05A00" } else { New-Brush "#0063B1" }
}

function Get-MgSelected {
    $addrs = @(Get-SelectedRows $script:GridMgMembers | ForEach-Object { $_['Adres'] })
    return @($script:MgRows | Where-Object { $addrs -contains $_.Adres })
}

function Invoke-MgMemberAction {
    param([string]$Action)
    if (-not $script:Mg) { return }
    $sel = @(Get-MgSelected)
    if ($Action -eq 'SubAll') { $sel = @($script:MgRows | Where-Object { ($_.Rol -eq 'Lid' -or $_.Rol -eq 'Eigenaar') -and $_.Groepsmail -ne 'ja' }) }
    if (-not $sel.Count) {
        Write-Log $(if ($Action -eq 'SubAll') { "Alle leden ontvangen de groepsmail al" } else { "Selecteer eerst 1 of meer regels" }) "warning"
        return
    }
    if ($Action -ne 'SendAs' -and $Action -ne 'Sob') {
        # Verwijderen/rol wisselen kan ook bij 'Alleen eigenaar', groepsmail alleen bij leden
        $sel = @($sel | Where-Object { $_.Rol -ne '(geen lid)' -and ($_.Rol -ne 'Alleen eigenaar' -or $Action -eq 'Remove' -or $Action -eq 'Role') })
        if (-not $sel.Count) { Write-Log "Deze actie geldt alleen voor leden" "warning"; return }
    }
    if ($Action -eq 'Remove') {
        $names = ($sel | Select-Object -First 15 | ForEach-Object { "  $($_.Naam)" }) -join "`n"
        if ($sel.Count -gt 15) { $names += "`n  ... en $($sel.Count - 15) meer" }
        if ((Show-Msg "$($sel.Count) persoon/personen uit '$($script:Mg.Name)' verwijderen?`n`n$names" "Verwijderen" "YesNo" "Warning") -ne 'Yes') { return }
    }

    $gid = $script:Mg.Gid
    $gro = $script:Mg.Id
    Enter-Busy "Bezig..."
    $ok = 0; $fail = 0; $graphChanged = $false
    try {
        if ($Action -eq 'SubAll') {
            $n = Add-GroupLinks $gid 'Subscribers' @($sel | ForEach-Object { $_.Adres })
            $ok = $n; $fail = $sel.Count - $n
        }
        else {
            foreach ($r in $sel) {
                try {
                    switch ($Action) {
                        'Remove' {
                            $uid = Resolve-GraphUserId $r
                            if ($r.Rol -ne 'Lid') { Invoke-Graph "groups/$gro/owners/$uid/`$ref" 'DELETE' | Out-Null }
                            if ($r.Rol -ne 'Alleen eigenaar') { Invoke-Graph "groups/$gro/members/$uid/`$ref" 'DELETE' | Out-Null }
                            $graphChanged = $true
                        }
                        'Role' {
                            $uid = Resolve-GraphUserId $r
                            # Lid -> Eigenaar -> Alleen eigenaar -> Lid
                            $link = @{ '@odata.id' = "$($script:GraphRoot)/directoryObjects/$uid" }
                            switch ($r.Rol) {
                                'Lid'      { Invoke-Graph "groups/$gro/owners/`$ref" 'POST' $link | Out-Null }
                                'Eigenaar' { Invoke-Graph "groups/$gro/members/$uid/`$ref" 'DELETE' | Out-Null }
                                default {
                                    Invoke-Graph "groups/$gro/members/`$ref" 'POST' $link | Out-Null
                                    Invoke-Graph "groups/$gro/owners/$uid/`$ref" 'DELETE' | Out-Null
                                }
                            }
                            $graphChanged = $true
                        }
                        'Sub' {
                            $lt = 'Subscribers'
                            if ($r.Groepsmail -eq 'ja') {
                                Write-Log "Remove-UnifiedGroupLinks Subscribers $($r.Adres)" "api"
                                Remove-UnifiedGroupLinks -Identity $gid -LinkType $lt -Links $r.Adres -Confirm:$false -ErrorAction Stop
                            }
                            else {
                                Write-Log "Add-UnifiedGroupLinks Subscribers $($r.Adres)" "api"
                                Add-UnifiedGroupLinks -Identity $gid -LinkType $lt -Links $r.Adres -Confirm:$false -ErrorAction Stop
                            }
                        }
                        'SendAs' {
                            $trustee = if ($r.Upn) { $r.Upn } else { $r.Adres }
                            if ($r.SendAs -eq 'ja') {
                                Write-Log "Remove-RecipientPermission SendAs $trustee" "api"
                                Remove-RecipientPermission -Identity $gid -Trustee $trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            else {
                                Write-Log "Add-RecipientPermission SendAs $trustee" "api"
                                Add-RecipientPermission -Identity $gid -Trustee $trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                        }
                        'Sob' {
                            $change = if ($r.Namens -eq 'ja') { @{ Remove = $r.Adres } } else { @{ Add = $r.Adres } }
                            Write-Log "Set-UnifiedGroup -GrantSendOnBehalfTo $(@($change.Keys)[0]) $($r.Adres)" "api"
                            Set-UnifiedGroup -Identity $gid -GrantSendOnBehalfTo $change -ErrorAction Stop
                        }
                    }
                    $ok++
                }
                catch { $fail++; Write-Log "$($r.Naam): $(Get-PsErrorText $_)" "error" }
            }
        }
        Write-Log "$ok gelukt, $fail mislukt" $(if ($fail) { 'warning' } else { 'success' })
        # Wijzigingen via Graph zijn niet direct zichtbaar in Exchange
        if ($graphChanged) { Set-Status "Wachten op synchronisatie naar Exchange..."; Wait-UI 8 }
        Update-MgMembers
        Set-Status $(if ($fail) { "Klaar met fouten - zie logboek" } else { "Gereed" }) $(if ($fail) { 'error' } else { 'success' })
    }
    catch { Write-Log "Actie mislukt: $(Get-PsErrorText $_)" "error"; Set-Status "Mislukt" "error" }
    finally { Exit-Busy }
}

function Add-MgMembers {
    if (-not $script:Mg) { return }
    $res = Show-RecipientPicker "Leden toevoegen aan '$($script:Mg.Name)'"
    if (-not $res -or -not @($res.Items).Count) { return }
    $gro = $script:Mg.Id
    Enter-Busy "Leden toevoegen..."
    try {
        $items = @($res.Items)
        $n = Add-GraphLinks $gro 'members' $items
        $no = $items.Count
        if ($res.Rol -eq 'Eigenaar') { $no = Add-GraphLinks $gro 'owners' $items }
        Write-Log "Toegevoegd: $n van $($items.Count) als lid$(if ($res.Rol -eq 'Eigenaar') { ", $no als eigenaar" })" $(if ($n -eq $items.Count -and $no -eq $items.Count) { 'success' } else { 'warning' })

        if ($script:Mg.Orig.Subscribe) {
            # Net als bij aanmaken expliciet abonneren; Exchange ziet nieuwe leden pas na synchronisatie
            Set-Status "Leden abonneren op groepsmail (wacht op synchronisatie)..."
            $addrs = @($items | ForEach-Object { $_.Adres })
            try {
                Invoke-WithRetry -Label "Abonneren" -Tries 6 -Delay 10 -Do {
                    Add-UnifiedGroupLinks -Identity $script:Mg.Gid -LinkType Subscribers -Links $addrs -Confirm:$false -ErrorAction Stop
                } | Out-Null
                Write-Log "Nieuwe leden ontvangen de groepsmail" "success"
            }
            catch { Write-Log "Abonneren lukte (nog) niet: $($_.Exception.Message) - gebruik later 'Alle leden abonneren'" "warning" }
        }
        else { Wait-UI 8 }
        Update-MgMembers
        Set-Status "Gereed" "success"
    }
    catch { Write-Log "Leden toevoegen mislukt: $(Get-PsErrorText $_)" "error"; Set-Status "Mislukt" "error" }
    finally { Exit-Busy }
}

function Save-ManagedGroup {
    if (-not $script:Mg) { return }
    $o = $script:Mg.Orig
    $n = Get-MgFormValues
    $ch = @(Get-MgChanges)
    if (-not $ch.Count) { Show-Msg "Er zijn geen wijzigingen." "Opslaan" | Out-Null; return }

    $errors = @()
    if (-not $n.Name) { $errors += "Weergavenaam mag niet leeg zijn." }
    if ($n.Primary -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $errors += "Ongeldig e-mailadres: $($n.Primary)" }
    $extraNew = @($n.Extra -split "`r`n" | Where-Object { $_ })
    foreach ($e in $extraNew) { if ($e -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $errors += "Ongeldig extra adres: $e" } }
    if ($errors.Count) { Show-Msg ($errors -join "`n") "Controleer de gegevens" "OK" "Warning" | Out-Null; return }

    # Nieuwe adressen mogen niet bij een ander object in gebruik zijn
    $extraOld = @($o.Extra -split "`r`n" | Where-Object { $_ })
    $newAddrs = @(@($extraNew | Where-Object { $extraOld -notcontains $_ }) + $(if ($ch -contains 'Primary' -and $extraOld -notcontains $n.Primary) { $n.Primary }) | Where-Object { $_ })
    Enter-Busy "Adressen controleren..."
    $taken = $null
    try {
        foreach ($a in $newAddrs) {
            $hit = @(Get-Recipient -Filter "EmailAddresses -eq 'smtp:$(ConvertTo-OPathLiteral $a)'" -ResultSize 1 -ErrorAction Stop)
            if ($hit.Count -and [string]$hit[0].Guid -ne $script:Mg.Gid) { $taken = "$a is al in gebruik door '$($hit[0].DisplayName)'."; break }
        }
    }
    catch { $taken = "Controle mislukt: $(Get-PsErrorText $_)" }
    finally { Exit-Busy }
    if ($taken) { Show-Msg $taken "Kan niet opslaan" "OK" "Warning" | Out-Null; return }

    $labels = @{
        Name = 'Weergavenaam'; Notes = 'Beschrijving'; Primary = 'Primair e-mailadres'; Extra = 'Extra e-mailadressen'
        Access = 'Privacy'; Language = 'Taal'; MailTip = 'MailTip'; Subscribe = 'Groepsmail naar postvak nieuwe leden'
        NoWelcome = 'Welkomstmail uit'; Calendar = 'Agenda-uitnodigingen'; External = 'Externe afzenders'
        HideOutlook = 'Verbergen in Outlook'; HideGal = 'Verbergen in adresboek'
    }
    $fmt = { param($v) if ($v -is [bool]) { if ($v) { 'aan' } else { 'uit' } } elseif ("$v") { ("$v" -replace "`r`n", ', ') } else { '(leeg)' } }
    $lines = foreach ($k in $ch) { "  $($labels[$k]): $(& $fmt $o[$k])  ->  $(& $fmt $n[$k])" }
    $subAll = $ch -contains 'Subscribe' -and $n.Subscribe
    $msg = "Deze wijzigingen opslaan voor '$($script:Mg.Name)'?`n`n$($lines -join "`n")"
    if ($subAll) { $msg += "`n`nBestaande leden worden ook geabonneerd op de groepsmail." }
    if ((Show-Msg $msg "Opslaan" "YesNo" "Question") -ne 'Yes') { return }

    $defs = [ordered]@{}
    if ($ch -contains 'Name' -or $ch -contains 'Notes') { $defs.gen = 'Naam en beschrijving (Graph)' }
    foreach ($k in 'Primary', 'Extra', 'Access', 'Language', 'MailTip', 'Subscribe', 'NoWelcome', 'Calendar', 'External', 'HideOutlook', 'HideGal') {
        if ($ch -contains $k) { $defs[$k] = $labels[$k] }
    }
    if ($subAll) { $defs.SubAll = 'Bestaande leden abonneren' }

    Enter-Busy "Opslaan..."
    Open-StepsWindow "Wijzigen: $($script:Mg.Name)" $defs
    $script:Gid = $script:Mg.Gid
    try {
        if ($defs.Contains('gen')) {
            [void](Invoke-Step 'gen' {
                $body = @{}
                if ($ch -contains 'Name') { $body.displayName = $n.Name }
                if ($ch -contains 'Notes') { $body.description = $(if ($n.Notes) { $n.Notes } else { $null }) }
                Invoke-Graph "groups/$($script:Mg.Id)" 'PATCH' $body | Out-Null
                ($body.Keys -join ', ')
            })
        }
        if ($defs.Contains('Primary')) { [void](Invoke-Step 'Primary' { Invoke-SetGroup 'Primair adres' @{ PrimarySmtpAddress = $n.Primary } 3; $n.Primary }) }
        if ($defs.Contains('Extra')) {
            [void](Invoke-Step 'Extra' {
                # Het oude primaire adres blijft automatisch als extra adres staan; niet weghalen
                $add = @($extraNew | Where-Object { $extraOld -notcontains $_ } | ForEach-Object { "smtp:$_" })
                $rem = @($extraOld | Where-Object { $extraNew -notcontains $_ -and $_ -ne $n.Primary } | ForEach-Object { "smtp:$_" })
                $chg = @{}
                if ($add.Count) { $chg.Add = $add }
                if ($rem.Count) { $chg.Remove = $rem }
                if (-not $chg.Count) { return 'niets te wijzigen' }
                Invoke-SetGroup 'Extra adressen' @{ EmailAddresses = $chg } 3
                "$($add.Count) toegevoegd, $($rem.Count) verwijderd"
            })
        }
        if ($defs.Contains('Access'))      { [void](Invoke-Step 'Access'      { Invoke-SetGroup 'Privacy' @{ AccessType = $n.Access } 3; $n.Access }) }
        if ($defs.Contains('Language'))    { [void](Invoke-Step 'Language'    { Invoke-SetGroup 'Taal' @{ Language = $n.Language } 3; $n.Language }) }
        if ($defs.Contains('MailTip'))     { [void](Invoke-Step 'MailTip'     { Invoke-SetGroup 'MailTip' @{ MailTip = $(if ($n.MailTip) { $n.MailTip } else { $null }) } 3; $(if ($n.MailTip) { 'ingesteld' } else { 'verwijderd' }) }) }
        if ($defs.Contains('Subscribe'))   { [void](Invoke-Step 'Subscribe'   { Invoke-SetGroup 'Abonneren' @{ AutoSubscribeNewMembers = $n.Subscribe } 3; & $fmt $n.Subscribe }) }
        if ($defs.Contains('NoWelcome'))   { [void](Invoke-Step 'NoWelcome'   { Invoke-SetGroup 'Welkomstmail' @{ UnifiedGroupWelcomeMessageEnabled = (-not $n.NoWelcome) } 3; $(if ($n.NoWelcome) { 'welkomstmail uit' } else { 'welkomstmail aan' }) }) }
        if ($defs.Contains('Calendar'))    { [void](Invoke-Step 'Calendar'    { Invoke-SetGroup 'Agenda' @{ AlwaysSubscribeMembersToCalendarEvents = $n.Calendar } 3; & $fmt $n.Calendar }) }
        if ($defs.Contains('External'))    { [void](Invoke-Step 'External'    { Invoke-SetGroup 'Externe afzenders' @{ RequireSenderAuthenticationEnabled = (-not $n.External) } 3; $(if ($n.External) { 'extern toegestaan' } else { 'alleen intern' }) }) }
        if ($defs.Contains('HideOutlook')) { [void](Invoke-Step 'HideOutlook' { Invoke-SetGroup 'Verbergen in Outlook' @{ HiddenFromExchangeClientsEnabled = $n.HideOutlook } 3; & $fmt $n.HideOutlook }) }
        if ($defs.Contains('HideGal') -or $defs.Contains('HideOutlook')) {
            # Verbergen in Outlook verandert ook de adresboek-zichtbaarheid; altijd daarna naar de keuze zetten
            if (-not $defs.Contains('HideGal')) { $defs.HideGal = $labels.HideGal; [void]$script:Steps.Add(@{ Key = 'HideGal'; Nr = "$($script:Steps.Count + 1)"; Stap = $labels.HideGal; Status = 'Wachtend'; Detail = '' }) }
            [void](Invoke-Step 'HideGal' { Invoke-SetGroup 'Adresboek' @{ HiddenFromAddressListsEnabled = $n.HideGal } 3; $(if ($n.HideGal) { 'verborgen' } else { 'zichtbaar' }) })
        }
        if ($defs.Contains('SubAll')) {
            [void](Invoke-Step 'SubAll' {
                $todo = @($script:MgRows | Where-Object { ($_.Rol -eq 'Lid' -or $_.Rol -eq 'Eigenaar') -and $_.Groepsmail -ne 'ja' } | ForEach-Object { $_.Adres })
                if (-not $todo.Count) { return 'alle leden al geabonneerd' }
                Get-CountResult (Add-GroupLinks $script:Mg.Gid 'Subscribers' $todo) $todo.Count 'leden geabonneerd'
            })
        }
        Complete-Steps ''
    }
    catch {
        Write-Log "Onverwachte fout: $(Get-PsErrorText $_)" "error"
        Complete-Steps "onverwachte fout - zie logboek"
    }
    finally {
        $script:StepsBusy = $false
        if ($script:StepsClose) { $script:StepsClose.IsEnabled = $true }
        Exit-Busy
    }
    Open-ManagedGroup $script:Mg.Id
}

function Remove-ManagedGroup {
    if (-not $script:Mg) { return }
    $txt = "Groep '$($script:Mg.Name)' <$($script:Mg.Mail)> verwijderen?`n`nLeden ontvangen geen mail meer via dit adres. De groep (incl. postvak en eventuele SharePoint-site) kan 30 dagen worden hersteld via Entra > Groepen > Verwijderde groepen."
    if ((Show-Msg $txt "Groep verwijderen" "YesNo" "Warning") -ne 'Yes') { return }
    if ((Show-Msg "Weet je het zeker? '$($script:Mg.Name)' wordt nu verwijderd." "Groep verwijderen" "YesNo" "Warning") -ne 'Yes') { return }
    Enter-Busy "Groep verwijderen..."
    try {
        Invoke-Graph "groups/$($script:Mg.Id)" 'DELETE' | Out-Null
        Write-Log "Groep '$($script:Mg.Name)' verwijderd (30 dagen herstelbaar)" "success"
        Clear-MgDetail
        Set-Status "Groep verwijderd" "success"
    }
    catch { Write-Log "Verwijderen mislukt: $(Get-PsErrorText $_)" "error"; Set-Status "Verwijderen mislukt" "error" }
    finally { Exit-Busy }
    Invoke-GroupSearch
}

function Clear-MgDetail {
    $script:Mg = $null
    $script:MgRows = @()
    $script:MgDetail.IsEnabled = $false
    $script:LblMgTitle.Text = 'Kies links een groep'
    $script:LblMgInfo.Text = ''
    $script:LblMgMembers.Text = ''
    $script:GridMgMembers.ItemsSource = $null
    foreach ($t in $script:TxtMgName, $script:TxtMgNotes, $script:TxtMgAlias, $script:TxtMgExtra, $script:TxtMgMailTip) { $t.Text = '' }
}

function Export-MgMembers {
    if (-not $script:Mg -or -not $script:MgRows.Count) { return }
    $d = New-Object Microsoft.Win32.SaveFileDialog
    $d.Filter = 'CSV-bestand (*.csv)|*.csv'
    $d.FileName = "Leden - $($script:Mg.Name -replace '[\\/:*?"<>|]', '_').csv"
    if (-not $d.ShowDialog($script:Window)) { return }
    try {
        $script:MgRows | ForEach-Object {
            [pscustomobject]@{ Rol = $_.Rol; Naam = $_.Naam; Adres = $_.Adres; Groepsmail = $_.Groepsmail; SendAs = $_.SendAs; SendOnBehalf = $_.Namens }
        } | Export-Csv -Path $d.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Write-Log "Geexporteerd naar $($d.FileName)" "success"
    }
    catch { Write-Log "Exporteren mislukt: $(Get-PsErrorText $_)" "error" }
}

function Show-RecipientPicker {
    # Modale zoekdialoog. Handlers vuren tijdens ShowDialog, daarom alle state via $script:
    param([string]$Title)
    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$script:PickerXaml)))
    $w.Owner = $script:Window
    $w.Title = $Title
    $script:PkWin = $w
    $script:PkResult = $null
    foreach ($n in 'PK_Txt', 'PK_Search', 'PK_Grid', 'PK_Paste', 'PK_PasteBtn', 'PK_Member', 'PK_Owner', 'PK_Lbl') { Set-Variable -Scope Script -Name $n -Value $w.FindName($n) }

    $script:PkSearchAction = {
        $script:PkWin.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $rows = @(Search-Recipients $script:PK_Txt.Text.Trim())
            Set-GridRows $script:PK_Grid @('Naam', 'Adres', 'Type', 'Id') $rows
            $script:PK_Lbl.Text = "$($rows.Count) resultaten"
        }
        catch { $script:PK_Lbl.Text = "Zoeken mislukt: $($_.Exception.Message)" }
        finally { $script:PkWin.Cursor = $null }
    }
    $script:PkPick = {
        param([string]$Rol)
        $sel = @(Get-SelectedRows $script:PK_Grid)
        if (-not $sel.Count) { $script:PK_Lbl.Text = 'Selecteer eerst 1 of meer personen'; return }
        $script:PkResult = @{ Rol = $Rol; Items = @($sel | ForEach-Object { @{ Naam = $_['Naam']; Adres = $_['Adres']; Type = $_['Type']; Id = $_['Id'] } }) }
        $script:PkWin.Close()
    }
    $script:PK_Search.Add_Click({ & $script:PkSearchAction })
    $script:PK_Txt.Add_KeyDown({ if ($_.Key -eq 'Return') { & $script:PkSearchAction } })
    $script:PK_Member.Add_Click({ & $script:PkPick 'Lid' })
    $script:PK_Owner.Add_Click({ & $script:PkPick 'Eigenaar' })
    $script:PK_Grid.Add_MouseDoubleClick({ & $script:PkPick 'Lid' })
    $script:PK_PasteBtn.Add_Click({
        # Geplakte adressen opzoeken en als (geselecteerde) resultaten tonen
        $script:PkWin.Cursor = [System.Windows.Input.Cursors]::Wait
        $rows = @(); $fail = @()
        foreach ($id in @($script:PK_Paste.Text -split '[\s;,]+' | Where-Object { $_ } | Select-Object -Unique)) {
            try {
                $r = Resolve-Recipient $id
                $rows += @{ Naam = $r.DisplayName; Adres = [string]$r.PrimarySmtpAddress; Type = [string]$r.RecipientTypeDetails; Id = [string]$r.ExternalDirectoryObjectId }
            }
            catch { $fail += $id }
        }
        Set-GridRows $script:PK_Grid @('Naam', 'Adres', 'Type', 'Id') $rows
        $script:PK_Grid.SelectAll()
        $script:PK_Paste.Text = ($fail -join "`r`n")
        $script:PK_Lbl.Text = "$($rows.Count) gevonden en geselecteerd$(if ($fail.Count) { ", $($fail.Count) niet gevonden (staan nog in het plakvak)" })"
        $script:PkWin.Cursor = $null
    })
    [void]$w.ShowDialog()
    return $script:PkResult
}

#endregion

#region --- XAML ---

$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft 365 Mailgroep aanmaken" Width="1320" Height="920" MinWidth="1100" MinHeight="700"
        WindowStartupLocation="CenterScreen" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
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
    <Style x:Key="CellRole" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#555555"/>
      <Style.Triggers>
        <Trigger Property="Text" Value="Eigenaar"><Setter Property="Foreground" Value="#0063B1"/></Trigger>
        <Trigger Property="Text" Value="Alleen eigenaar"><Setter Property="Foreground" Value="#6B3FA0"/></Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="Background" Value="White"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,0,0,7"/>
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
    <Style x:Key="FieldCombo" TargetType="ComboBox">
      <Setter Property="Height" Value="26"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="Gray"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="*" MinHeight="420"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="170" MinHeight="70"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0078D4">
      <DockPanel Margin="16,0,12,0">
        <StackPanel x:Name="HeaderPanel" DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Tenant (optioneel, GDAP):" Foreground="White" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="TxtTenant" Width="190" Height="26" VerticalContentAlignment="Center"
                   ToolTip="Leeg = eigen tenant. Voor een klant via GDAP: domein (klant.onmicrosoft.com) of tenant-ID."/>
          <TextBlock x:Name="LblExo" Text="Niet verbonden" Foreground="#FFC864" VerticalAlignment="Center" Margin="14,0,0,0"/>
          <Button x:Name="BtnConnect" Content="Verbinden" Style="{StaticResource BtnDark}"
                  ToolTip="Logt in bij Microsoft Graph (groep aanmaken) en Exchange Online (instellingen) met hetzelfde account."/>
        </StackPanel>
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="Microsoft 365 Mailgroep" Foreground="White" FontSize="20" FontWeight="SemiBold"/>
          <TextBlock Text="M365-groep als vervanger van een distributielijst" Foreground="#CDE6FF"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <TabControl x:Name="MainArea" Grid.Row="1" Margin="6,6,6,0" Background="#F5F5F5">
      <TabItem Header="  Nieuwe mailgroep  " FontSize="13">
    <Grid Margin="0,6,0,0" TextElement.FontSize="12">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="5*" MinWidth="440"/>
        <ColumnDefinition Width="6"/>
        <ColumnDefinition Width="6*" MinWidth="520"/>
      </Grid.ColumnDefinitions>

      <!-- Links: groep en opties -->
      <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
        <StackPanel>
          <GroupBox Header="1. Groep">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="120"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Grid.Row="0" Text="Weergavenaam *" Style="{StaticResource FieldLabel}"/>
              <TextBox   Grid.Row="0" Grid.Column="1" x:Name="TxtName" Style="{StaticResource FieldBox}"/>
              <TextBlock Grid.Row="1" Text="Alias *" Style="{StaticResource FieldLabel}"/>
              <TextBox   Grid.Row="1" Grid.Column="1" x:Name="TxtAlias" Style="{StaticResource FieldBox}"
                         ToolTip="Wordt automatisch gevuld op basis van de naam; zelf aanpassen mag."/>
              <TextBlock Grid.Row="2" Text="Domein *" Style="{StaticResource FieldLabel}"/>
              <ComboBox  Grid.Row="2" Grid.Column="1" x:Name="CmbDomain" Style="{StaticResource FieldCombo}"/>
              <TextBlock Grid.Row="3" Text="E-mailadres" Style="{StaticResource FieldLabel}"/>
              <DockPanel Grid.Row="3" Grid.Column="1" Margin="0,0,0,6">
                <Button x:Name="BtnCheck" DockPanel.Dock="Right" Content="Controleren" Padding="10,3" Margin="8,0,0,0"/>
                <TextBlock x:Name="LblAvail" DockPanel.Dock="Right" FontWeight="SemiBold" VerticalAlignment="Center" Margin="8,0,0,0"/>
                <TextBlock x:Name="LblAddress" Text="-" FontWeight="SemiBold" Foreground="#0063B1" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
              </DockPanel>
              <TextBlock Grid.Row="4" Text="Beschrijving" Style="{StaticResource FieldLabel}" VerticalAlignment="Top" Margin="0,4,8,6"/>
              <TextBox   Grid.Row="4" Grid.Column="1" x:Name="TxtNotes" Height="48" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto" Margin="0,0,0,6" Padding="4,2"/>
              <TextBlock Grid.Row="5" Text="Privacy" Style="{StaticResource FieldLabel}"/>
              <ComboBox  Grid.Row="5" Grid.Column="1" x:Name="CmbAccess" Style="{StaticResource FieldCombo}">
                <ComboBoxItem Content="Prive - alleen leden zien inhoud (aanbevolen)" Tag="Private" IsSelected="True"/>
                <ComboBoxItem Content="Openbaar - iedereen in de organisatie kan lid worden" Tag="Public"/>
              </ComboBox>
              <TextBlock Grid.Row="6" Text="Taal" Style="{StaticResource FieldLabel}"/>
              <ComboBox  Grid.Row="6" Grid.Column="1" x:Name="CmbLang" Style="{StaticResource FieldCombo}"
                         ToolTip="Taal van systeemmeldingen van de groep.">
                <ComboBoxItem Content="Nederlands (nl-NL)" Tag="nl-NL" IsSelected="True"/>
                <ComboBoxItem Content="Engels (en-US)" Tag="en-US"/>
                <ComboBoxItem Content="Duits (de-DE)" Tag="de-DE"/>
                <ComboBoxItem Content="Frans (fr-FR)" Tag="fr-FR"/>
              </ComboBox>
              <TextBlock Grid.Row="7" Text="Extra adressen" Style="{StaticResource FieldLabel}"/>
              <TextBox   Grid.Row="7" Grid.Column="1" x:Name="TxtExtra" Style="{StaticResource FieldBox}"
                         ToolTip="Optioneel: extra e-mailadressen (aliassen) voor de groep, gescheiden door ; of spatie."/>
              <TextBlock Grid.Row="8" Text="MailTip" Style="{StaticResource FieldLabel}"/>
              <TextBox   Grid.Row="8" Grid.Column="1" x:Name="TxtMailTip" Style="{StaticResource FieldBox}"
                         ToolTip="Optioneel: melding die afzenders in Outlook zien als ze naar de groep mailen."/>
            </Grid>
          </GroupBox>

          <GroupBox Header="2. Mailgedrag">
            <StackPanel>
              <CheckBox x:Name="ChkSubscribe" IsChecked="True" FontWeight="SemiBold"
                        Content="Leden ontvangen mail aan de groep in hun eigen postvak IN"
                        ToolTip="AutoSubscribeNewMembers + alle leden als subscriber. Zonder deze optie komt mail alleen in het groepspostvak."/>
              <CheckBox x:Name="ChkNoWelcome" IsChecked="True" FontWeight="SemiBold"
                        Content="Geen welkomstmail naar leden"
                        ToolTip="UnifiedGroupWelcomeMessageEnabled uit. Wordt ingesteld voordat er leden worden toegevoegd."/>
              <CheckBox x:Name="ChkCalendar" Content="Leden ontvangen ook agenda-uitnodigingen aan de groep"
                        ToolTip="AlwaysSubscribeMembersToCalendarEvents"/>
              <CheckBox x:Name="ChkExternal" Content="Externe afzenders mogen naar de groep mailen"
                        ToolTip="RequireSenderAuthenticationEnabled uit. Nodig voor bijv. info@ of facturen@ adressen."/>
              <CheckBox x:Name="ChkHideOutlook" Content="Groep verbergen in Outlook (niet in de mappenlijst van leden)"
                        ToolTip="HiddenFromExchangeClientsEnabled. Leden krijgen de mail nog steeds in hun postvak IN."/>
              <CheckBox x:Name="ChkHideGal" Content="Verbergen in het adresboek (GAL)"
                        ToolTip="Collega's kunnen de groep dan niet opzoeken; het adres werkt wel."/>
              <CheckBox x:Name="ChkHideMembership" Content="Ledenlijst verbergen voor niet-leden (alleen Prive, later niet te wijzigen)"
                        ToolTip="Graph visibility HiddenMembership - kan alleen bij aanmaken worden ingesteld."/>
              <CheckBox x:Name="ChkNoSite" IsChecked="True" Content="Geen SharePoint-site aanmaken (pas als iemand die opent)"
                        ToolTip="ProvisionSiteOnDemand. Er wordt sowieso geen Team aangemaakt."/>
            </StackPanel>
          </GroupBox>

          <GroupBox Header="3. Verzenden als de groep">
            <StackPanel>
              <RadioButton x:Name="RbSendNone" GroupName="Send" IsChecked="True" Content="Leden mogen niet als de groep verzenden"/>
              <RadioButton x:Name="RbSendAs" GroupName="Send" Content="Send As - leden verzenden als het groepsadres"/>
              <RadioButton x:Name="RbSendOnBehalf" GroupName="Send" Content="Send on Behalf - 'Naam namens Groep'"/>
              <TextBlock Style="{StaticResource Hint}" Text="Geldt voor alle gekozen leden. Leden die later worden toegevoegd krijgen dit recht niet automatisch."/>
            </StackPanel>
          </GroupBox>
        </StackPanel>
      </ScrollViewer>

      <!-- Rechts: leden -->
      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="6*"/>
          <RowDefinition Height="5*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <GroupBox Grid.Row="0" Header="4. Leden en eigenaren zoeken">
          <DockPanel>
            <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
              <Button x:Name="BtnSearch" DockPanel.Dock="Right" Content="Zoeken" Margin="8,0,0,0" Padding="14,4"/>
              <TextBox x:Name="TxtSearch" Height="26" VerticalContentAlignment="Center" Padding="4,0"
                       ToolTip="Naam, e-mailadres of deel daarvan. Leeg = eerste 1000 gebruikers."/>
            </DockPanel>
            <StackPanel DockPanel.Dock="Bottom">
              <DockPanel Margin="0,6,0,0">
                <TextBlock x:Name="LblSearch" DockPanel.Dock="Right" Foreground="Gray" VerticalAlignment="Center"/>
                <StackPanel Orientation="Horizontal">
                  <Button x:Name="BtnAddMember" Content="+ Als lid"/>
                  <Button x:Name="BtnAddOwner" Content="+ Als eigenaar" Style="{StaticResource BtnDark}" Margin="0,0,8,6"/>
                  <Button x:Name="BtnAddMe" Content="+ Mijn account als eigenaar" Style="{StaticResource BtnGray}"
                          ToolTip="Voegt het ingelogde account toe als eigenaar (alleen als dat account een mailbox heeft in deze tenant)."/>
                </StackPanel>
              </DockPanel>
              <Expander Header="Plakken of overnemen uit distributielijst" Margin="0,2,0,0">
                <Grid Margin="0,6,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Style="{StaticResource Hint}" Text="E-mailadressen (1 per regel, of gescheiden door ; ,)"/>
                    <TextBox x:Name="TxtPaste" Height="70" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Margin="0,0,0,6"/>
                    <StackPanel Orientation="Horizontal">
                      <Button x:Name="BtnPaste" Content="Toevoegen"/>
                      <CheckBox x:Name="ChkPasteOwner" Content="als eigenaar" VerticalAlignment="Center" Margin="0,0,0,6"/>
                    </StackPanel>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Style="{StaticResource Hint}" Text="Bestaande distributielijst (naam of adres) - leden worden overgenomen"/>
                    <TextBox x:Name="TxtDl" Style="{StaticResource FieldBox}"/>
                    <Button x:Name="BtnDl" Content="Leden overnemen" HorizontalAlignment="Left"/>
                  </StackPanel>
                </Grid>
              </Expander>
            </StackPanel>
            <DataGrid x:Name="GridSearch">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="2*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="E-mailadres" Binding="{Binding Adres}" Width="2*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="*" ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
          </DockPanel>
        </GroupBox>

        <GroupBox Grid.Row="1" Header="5. Gekozen leden en eigenaren">
          <DockPanel>
            <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
              <TextBlock x:Name="LblPicked" DockPanel.Dock="Right" FontWeight="SemiBold" VerticalAlignment="Center"/>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="BtnToggleRole" Content="Rol wisselen" ToolTip="Lid &gt; Eigenaar &gt; Alleen eigenaar (geen lid) &gt; Lid"/>
                <Button x:Name="BtnRemove" Content="Verwijderen" Style="{StaticResource BtnRed}"/>
                <Button x:Name="BtnClearPicked" Content="Alles leegmaken" Style="{StaticResource BtnGray}"/>
              </StackPanel>
            </DockPanel>
            <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}" Text="Eigenaren zijn standaard ook lid en ontvangen de mail; 'Alleen eigenaar' beheert de groep zonder lid te zijn. Minimaal 1 eigenaar verplicht. Dubbelklik wisselt de rol (Lid / Eigenaar / Alleen eigenaar)."/>
            <DataGrid x:Name="GridPicked">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Rol" Binding="{Binding Rol}" Width="110" ElementStyle="{StaticResource CellRole}"/>
                <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="2*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="E-mailadres" Binding="{Binding Adres}" Width="2*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="*" ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
          </DockPanel>
        </GroupBox>

        <Border Grid.Row="2" Background="White" BorderBrush="#CCCCCC" BorderThickness="1" Padding="10,8,10,2" Margin="0,0,0,6">
          <DockPanel>
            <Button x:Name="BtnCreate" DockPanel.Dock="Right" Content="Groep aanmaken" Style="{StaticResource BtnGreen}"
                    FontSize="14" FontWeight="SemiBold" Padding="22,8" Margin="8,0,0,6"/>
            <Button x:Name="BtnReset" DockPanel.Dock="Right" Content="Nieuw formulier" Style="{StaticResource BtnGray}" Padding="14,8"/>
            <TextBlock Style="{StaticResource Hint}" VerticalAlignment="Center"
                       Text="De groep wordt via Graph aangemaakt met welkomstmail uit en abonneren aan; daarna volgen de instellingen stap voor stap."/>
          </DockPanel>
        </Border>
      </Grid>
    </Grid>
      </TabItem>

      <TabItem Header="  Bestaande groep beheren  " FontSize="13">
    <Grid Margin="0,6,0,0" TextElement.FontSize="12">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="320" MinWidth="250"/>
        <ColumnDefinition Width="6"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <GroupBox Grid.Column="0" Header="Groep zoeken">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
            <Button x:Name="BtnMgSearch" DockPanel.Dock="Right" Content="Zoeken" Margin="8,0,0,0" Padding="12,4"/>
            <TextBox x:Name="TxtMgSearch" Height="26" VerticalContentAlignment="Center" Padding="4,0"
                     ToolTip="Naam of e-mailadres (begin van een woord). Leeg = alle Microsoft 365-groepen."/>
          </DockPanel>
          <TextBlock x:Name="LblMgSearch" DockPanel.Dock="Bottom" Foreground="Gray" Margin="0,6,0,0"
                     Text="Zoek op naam of adres, of laat leeg voor alle groepen."/>
          <DataGrid x:Name="GridGroups" SelectionMode="Single">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="3*" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="E-mailadres" Binding="{Binding Adres}" Width="3*" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Team" Binding="{Binding Team}" Width="42" ElementStyle="{StaticResource CellText}"/>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </GroupBox>

      <Grid x:Name="MgDetail" Grid.Column="2" IsEnabled="False">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="5*" MinWidth="380"/>
          <ColumnDefinition Width="6"/>
          <ColumnDefinition Width="6*" MinWidth="430"/>
        </Grid.ColumnDefinitions>

        <DockPanel Grid.Column="0">
          <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#CCCCCC" BorderThickness="1" Padding="10,8,10,2" Margin="0,0,0,6">
            <DockPanel>
              <Button x:Name="BtnMgSave" DockPanel.Dock="Right" Content="Wijzigingen opslaan" Style="{StaticResource BtnGreen}"
                      FontWeight="SemiBold" Padding="16,7" Margin="8,0,0,6"/>
              <Button x:Name="BtnMgUndo" DockPanel.Dock="Right" Content="Ongedaan maken" Style="{StaticResource BtnGray}" Padding="12,7"/>
              <Button x:Name="BtnMgDelete" Content="Groep verwijderen" Style="{StaticResource BtnRed}" HorizontalAlignment="Left" Padding="12,7"/>
            </DockPanel>
          </Border>
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock x:Name="LblMgTitle" Text="Kies links een groep" FontSize="16" FontWeight="SemiBold" Margin="2,0,0,2" TextTrimming="CharacterEllipsis"/>
              <TextBlock x:Name="LblMgInfo" Style="{StaticResource Hint}" Margin="2,0,0,8"/>
              <GroupBox Header="Algemeen">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Text="Weergavenaam" Style="{StaticResource FieldLabel}"/>
                  <TextBox   Grid.Row="0" Grid.Column="1" x:Name="TxtMgName" Style="{StaticResource FieldBox}"/>
                  <TextBlock Grid.Row="1" Text="Beschrijving" Style="{StaticResource FieldLabel}" VerticalAlignment="Top" Margin="0,4,8,6"/>
                  <TextBox   Grid.Row="1" Grid.Column="1" x:Name="TxtMgNotes" Height="44" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" Margin="0,0,0,6" Padding="4,2"/>
                  <TextBlock Grid.Row="2" Text="E-mailadres" Style="{StaticResource FieldLabel}"/>
                  <DockPanel Grid.Row="2" Grid.Column="1" Margin="0,0,0,6">
                    <ComboBox x:Name="CmbMgDomain" DockPanel.Dock="Right" Width="170" Height="26" VerticalContentAlignment="Center"/>
                    <TextBlock DockPanel.Dock="Right" Text="@" VerticalAlignment="Center" Margin="4,0"/>
                    <TextBox x:Name="TxtMgAlias" Height="26" VerticalContentAlignment="Center" Padding="4,0"
                             ToolTip="Primair e-mailadres. Het oude adres blijft als extra adres werken."/>
                  </DockPanel>
                  <TextBlock Grid.Row="3" Text="Extra adressen" Style="{StaticResource FieldLabel}" VerticalAlignment="Top" Margin="0,4,8,6"/>
                  <TextBox   Grid.Row="3" Grid.Column="1" x:Name="TxtMgExtra" Height="62" TextWrapping="NoWrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" Margin="0,0,0,6" Padding="4,2"
                             ToolTip="1 adres per regel. Regel weghalen = adres verwijderen."/>
                  <TextBlock Grid.Row="4" Text="Privacy" Style="{StaticResource FieldLabel}"/>
                  <ComboBox  Grid.Row="4" Grid.Column="1" x:Name="CmbMgAccess" Style="{StaticResource FieldCombo}">
                    <ComboBoxItem Content="Prive - alleen leden zien inhoud" Tag="Private"/>
                    <ComboBoxItem Content="Openbaar - iedereen in de organisatie kan lid worden" Tag="Public"/>
                  </ComboBox>
                  <TextBlock Grid.Row="5" Text="Taal" Style="{StaticResource FieldLabel}"/>
                  <ComboBox  Grid.Row="5" Grid.Column="1" x:Name="CmbMgLang" IsEditable="True" Style="{StaticResource FieldCombo}">
                    <ComboBoxItem Content="nl-NL"/><ComboBoxItem Content="en-US"/><ComboBoxItem Content="de-DE"/><ComboBoxItem Content="fr-FR"/>
                  </ComboBox>
                  <TextBlock Grid.Row="6" Text="MailTip" Style="{StaticResource FieldLabel}"/>
                  <TextBox   Grid.Row="6" Grid.Column="1" x:Name="TxtMgMailTip" Style="{StaticResource FieldBox}"/>
                </Grid>
              </GroupBox>
              <GroupBox Header="Mailgedrag">
                <StackPanel>
                  <CheckBox x:Name="ChkMgSubscribe" FontWeight="SemiBold" Content="Nieuwe leden ontvangen groepsmail in hun postvak IN"
                            ToolTip="AutoSubscribeNewMembers. Bij aanzetten worden bestaande leden ook geabonneerd."/>
                  <CheckBox x:Name="ChkMgNoWelcome" FontWeight="SemiBold" Content="Geen welkomstmail naar nieuwe leden"/>
                  <CheckBox x:Name="ChkMgCalendar" Content="Leden ontvangen ook agenda-uitnodigingen aan de groep"/>
                  <CheckBox x:Name="ChkMgExternal" Content="Externe afzenders mogen naar de groep mailen"/>
                  <CheckBox x:Name="ChkMgHideOutlook" Content="Groep verbergen in Outlook (mappenlijst leden)"/>
                  <CheckBox x:Name="ChkMgHideGal" Content="Verbergen in het adresboek (GAL)"/>
                </StackPanel>
              </GroupBox>
            </StackPanel>
          </ScrollViewer>
        </DockPanel>

        <GroupBox Grid.Column="2" Header="Leden, eigenaren en rechten">
          <DockPanel>
            <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}"
                       Text="Selecteer 1 of meer regels en kies een actie. Deze acties worden direct doorgevoerd. Aan/uit-knoppen draaien de huidige waarde per regel om."/>
            <StackPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
              <WrapPanel>
                <Button x:Name="BtnMgAdd" Content="+ Leden toevoegen" Style="{StaticResource BtnGreen}"/>
                <Button x:Name="BtnMgRemove" Content="Uit groep verwijderen" Style="{StaticResource BtnRed}"/>
                <Button x:Name="BtnMgRole" Content="Rol wisselen" ToolTip="Lid &gt; Eigenaar &gt; Alleen eigenaar (geen lid) &gt; Lid"/>
              </WrapPanel>
              <WrapPanel>
                <Button x:Name="BtnMgSub" Content="Groepsmail aan/uit" Style="{StaticResource BtnGray}"/>
                <Button x:Name="BtnMgSubAll" Content="Alle leden abonneren" Style="{StaticResource BtnGray}"/>
                <Button x:Name="BtnMgSendAs" Content="Send As aan/uit" Style="{StaticResource BtnGray}"/>
                <Button x:Name="BtnMgSob" Content="Send on Behalf aan/uit" Style="{StaticResource BtnGray}"/>
                <Button x:Name="BtnMgExport" Content="Exporteren (CSV)" Style="{StaticResource BtnGray}"/>
                <Button x:Name="BtnMgReload" Content="Vernieuwen" Style="{StaticResource BtnGray}"/>
              </WrapPanel>
              <TextBlock x:Name="LblMgMembers" FontWeight="SemiBold" Foreground="#0063B1"/>
            </StackPanel>
            <DataGrid x:Name="GridMgMembers">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Rol" Binding="{Binding Rol}" Width="110" ElementStyle="{StaticResource CellRole}"/>
                <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="2*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="E-mailadres" Binding="{Binding Adres}" Width="2*" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Groepsmail" Binding="{Binding Groepsmail}" Width="72" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Send As" Binding="{Binding SendAs}" Width="58" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Namens" Binding="{Binding Namens}" Width="58" ElementStyle="{StaticResource CellText}"/>
              </DataGrid.Columns>
            </DataGrid>
          </DockPanel>
        </GroupBox>
      </Grid>
    </Grid>
      </TabItem>
    </TabControl>

    <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" Background="#D0D0D0" ResizeDirection="Rows"/>

    <DockPanel Grid.Row="3" Background="#1E1E1E">
      <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Background="#2D2D2D">
        <TextBlock Text="Logboek" Foreground="#B4B4B4" FontWeight="SemiBold" VerticalAlignment="Center" Margin="10,0,10,0"/>
        <Button x:Name="BtnClearLog" Content="Wissen" Style="{StaticResource BtnLog}"/>
        <Button x:Name="BtnCopyLog" Content="Kopieren" Style="{StaticResource BtnLog}" ToolTip="Hele logboek naar het klembord"/>
      </StackPanel>
      <RichTextBox x:Name="LogBox" IsReadOnly="True" Background="#1E1E1E" BorderThickness="0"
                   FontFamily="Consolas" FontSize="12" VerticalScrollBarVisibility="Auto" Padding="6,4">
        <FlowDocument><Paragraph x:Name="LogPara" Margin="0"/></FlowDocument>
      </RichTextBox>
    </DockPanel>

    <Border x:Name="StatusBar" Grid.Row="4" Background="#0078D4">
      <TextBlock x:Name="LblStatus" Text="Gereed" Foreground="White" VerticalAlignment="Center" Margin="12,0"/>
    </Border>
  </Grid>
</Window>
'@

$script:PickerXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Leden toevoegen" Width="760" Height="600" MinWidth="560" MinHeight="420"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <DockPanel Margin="12">
    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
      <Button x:Name="PK_Search" DockPanel.Dock="Right" Content="Zoeken" Width="90" Height="28" Margin="8,0,0,0"/>
      <TextBox x:Name="PK_Txt" Height="28" VerticalContentAlignment="Center" Padding="4,0"
               ToolTip="Naam, e-mailadres of deel daarvan. Leeg = eerste 1000 gebruikers."/>
    </DockPanel>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button x:Name="PK_Owner" DockPanel.Dock="Right" Content="Toevoegen als eigenaar" Height="30" Padding="12,0" Margin="8,0,0,0"/>
      <Button x:Name="PK_Member" DockPanel.Dock="Right" Content="Toevoegen als lid" Height="30" Padding="12,0" FontWeight="SemiBold"/>
      <TextBlock x:Name="PK_Lbl" Foreground="Gray" VerticalAlignment="Center" TextWrapping="Wrap"/>
    </DockPanel>
    <Expander DockPanel.Dock="Bottom" Header="E-mailadressen plakken" Margin="0,8,0,0">
      <DockPanel Margin="0,6,0,0">
        <Button x:Name="PK_PasteBtn" DockPanel.Dock="Right" Content="Opzoeken" Width="90" Height="28" VerticalAlignment="Top" Margin="8,0,0,0"/>
        <TextBox x:Name="PK_Paste" Height="70" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                 ToolTip="1 per regel, of gescheiden door ; ,"/>
      </DockPanel>
    </Expander>
    <DataGrid x:Name="PK_Grid" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
              HorizontalGridLinesBrush="#E6E6E6" Background="White" BorderBrush="#DDDDDD" CanUserAddRows="False"
              SelectionMode="Extended" SelectionUnit="FullRow" MinRowHeight="24">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Naam" Binding="{Binding Naam}" Width="2*"/>
        <DataGridTextColumn Header="E-mailadres" Binding="{Binding Adres}" Width="2*"/>
        <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="*"/>
      </DataGrid.Columns>
    </DataGrid>
  </DockPanel>
</Window>
'@

$script:StepsXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Voortgang groep aanmaken" Width="860" Height="600" MinWidth="600" MinHeight="400"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <DockPanel Margin="12">
    <TextBlock x:Name="SW_Title" DockPanel.Dock="Top" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10" TextTrimming="CharacterEllipsis"/>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
      <Button x:Name="SW_Close" DockPanel.Dock="Right" Content="Sluiten" IsEnabled="False" Width="110" Height="30" Margin="10,0,0,0"/>
      <TextBlock x:Name="SW_Summary" Text="Bezig..." FontWeight="SemiBold" Foreground="#0063B1" VerticalAlignment="Center" TextWrapping="Wrap"/>
    </DockPanel>
    <DataGrid x:Name="SW_Grid" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
              GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E6E6E6" Background="White"
              BorderBrush="#DDDDDD" CanUserAddRows="False" CanUserSortColumns="False" SelectionMode="Single" MinRowHeight="28">
      <DataGrid.Columns>
        <DataGridTextColumn Header="#" Binding="{Binding Nr}" Width="34"/>
        <DataGridTextColumn Header="Stap" Binding="{Binding Stap}" Width="250"/>
        <DataGridTemplateColumn Header="Status" Width="130">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock x:Name="t" Text="{Binding Status}" FontWeight="SemiBold" Margin="6,0" VerticalAlignment="Center" Foreground="#888888"/>
              <DataTemplate.Triggers>
                <DataTrigger Binding="{Binding Status}" Value="Bezig">
                  <Setter TargetName="t" Property="Foreground" Value="#0063B1"/><Setter TargetName="t" Property="Text" Value="&#x25B6; Bezig..."/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="Gelukt">
                  <Setter TargetName="t" Property="Foreground" Value="#107C10"/><Setter TargetName="t" Property="Text" Value="&#x2714; Gelukt"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="Deels gelukt">
                  <Setter TargetName="t" Property="Foreground" Value="#B05A00"/><Setter TargetName="t" Property="Text" Value="! Deels gelukt"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="Mislukt">
                  <Setter TargetName="t" Property="Foreground" Value="#C42B1C"/><Setter TargetName="t" Property="Text" Value="&#x2716; Mislukt"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="Overgeslagen">
                  <Setter TargetName="t" Property="Foreground" Value="#999999"/><Setter TargetName="t" Property="Text" Value="- Overgeslagen"/>
                </DataTrigger>
              </DataTemplate.Triggers>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
        <DataGridTemplateColumn Header="Details" Width="*">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock Text="{Binding Detail}" TextWrapping="Wrap" Margin="6,3" VerticalAlignment="Center" ToolTip="{Binding Detail}"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
      </DataGrid.Columns>
    </DataGrid>
  </DockPanel>
</Window>
'@

#endregion

#region --- Venster ---

function New-MainWindow {
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$script:MainXaml)
    $script:Window = [System.Windows.Markup.XamlReader]::Load($reader)

    $names = 'HeaderPanel', 'TxtTenant', 'LblExo', 'BtnConnect', 'MainArea',
             'TxtName', 'TxtAlias', 'CmbDomain', 'LblAddress', 'LblAvail', 'BtnCheck', 'TxtNotes', 'CmbAccess', 'CmbLang', 'TxtExtra', 'TxtMailTip',
             'ChkSubscribe', 'ChkNoWelcome', 'ChkCalendar', 'ChkExternal', 'ChkHideOutlook', 'ChkHideGal', 'ChkHideMembership', 'ChkNoSite',
             'RbSendNone', 'RbSendAs', 'RbSendOnBehalf',
             'TxtSearch', 'BtnSearch', 'GridSearch', 'LblSearch', 'BtnAddMember', 'BtnAddOwner', 'BtnAddMe',
             'TxtPaste', 'BtnPaste', 'ChkPasteOwner', 'TxtDl', 'BtnDl',
             'GridPicked', 'LblPicked', 'BtnToggleRole', 'BtnRemove', 'BtnClearPicked',
             'BtnCreate', 'BtnReset', 'BtnClearLog', 'BtnCopyLog', 'LogBox', 'LogPara', 'StatusBar', 'LblStatus',
             'TxtMgSearch', 'BtnMgSearch', 'LblMgSearch', 'GridGroups', 'MgDetail', 'LblMgTitle', 'LblMgInfo',
             'TxtMgName', 'TxtMgNotes', 'TxtMgAlias', 'CmbMgDomain', 'TxtMgExtra', 'CmbMgAccess', 'CmbMgLang', 'TxtMgMailTip',
             'ChkMgSubscribe', 'ChkMgNoWelcome', 'ChkMgCalendar', 'ChkMgExternal', 'ChkMgHideOutlook', 'ChkMgHideGal',
             'BtnMgSave', 'BtnMgUndo', 'BtnMgDelete', 'GridMgMembers', 'LblMgMembers',
             'BtnMgAdd', 'BtnMgRemove', 'BtnMgRole', 'BtnMgSub', 'BtnMgSubAll', 'BtnMgSendAs', 'BtnMgSob', 'BtnMgExport', 'BtnMgReload'
    foreach ($n in $names) { Set-Variable -Scope Script -Name $n -Value $script:Window.FindName($n) }

    # --- Verbinding ---
    $script:BtnConnect.Add_Click({ Connect-All })
    $script:TxtTenant.Add_KeyDown({ if ($_.Key -eq 'Return') { Connect-All } })

    # --- Groep ---
    $script:TxtName.Add_TextChanged({
        if ($script:AliasTouched) { return }
        $script:SettingAlias = $true
        $script:TxtAlias.Text = ConvertTo-Alias $script:TxtName.Text
        $script:SettingAlias = $false
        Update-AddressPreview
    })
    $script:TxtAlias.Add_TextChanged({
        if (-not $script:SettingAlias) { $script:AliasTouched = [bool]$script:TxtAlias.Text }
        Update-AddressPreview
    })
    $script:CmbDomain.Add_SelectionChanged({ Update-AddressPreview })
    $script:BtnCheck.Add_Click({ Invoke-CheckAddress })
    $script:ChkHideMembership.Add_Checked({ $script:CmbAccess.SelectedIndex = 0; $script:CmbAccess.IsEnabled = $false })
    $script:ChkHideMembership.Add_Unchecked({ $script:CmbAccess.IsEnabled = $true })
    $script:ChkHideOutlook.Add_Checked({
        Write-Log "Let op: 'Verbergen in Outlook' verbergt de groep ook in het adresboek, tenzij je dat hieronder uit laat; de tool zet het daarna terug naar jouw keuze." "info"
    })

    # --- Leden zoeken ---
    $script:BtnSearch.Add_Click({ Invoke-Search })
    $script:TxtSearch.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-Search } })
    $script:BtnAddMember.Add_Click({ Add-FromSearch 'Lid' })
    $script:BtnAddOwner.Add_Click({ Add-FromSearch 'Eigenaar' })
    $script:BtnAddMe.Add_Click({ Add-Me 'Eigenaar' })
    $script:GridSearch.Add_MouseDoubleClick({ Add-FromSearch 'Lid' })
    $script:BtnPaste.Add_Click({ Invoke-PasteAdd })
    $script:BtnDl.Add_Click({ Invoke-ImportDl })
    $script:TxtDl.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-ImportDl } })

    # --- Gekozen ---
    $toggle = {
        $sel = @(Get-SelectedRows $script:GridPicked | ForEach-Object { $_['Adres'] })
        foreach ($p in $script:Picked) {
            if ($sel -contains $p.Adres) { $p.Rol = switch ($p.Rol) { 'Lid' { 'Eigenaar' } 'Eigenaar' { 'Alleen eigenaar' } default { 'Lid' } } }
        }
        Update-PickedGrid
    }
    $script:BtnToggleRole.Add_Click($toggle)
    $script:GridPicked.Add_MouseDoubleClick($toggle)
    $script:BtnRemove.Add_Click({
        $sel = @(Get-SelectedRows $script:GridPicked | ForEach-Object { $_['Adres'] })
        foreach ($p in @($script:Picked)) { if ($sel -contains $p.Adres) { $script:Picked.Remove($p) } }
        Update-PickedGrid
    })
    $script:GridPicked.Add_KeyDown({ if ($_.Key -eq 'Delete') { $script:BtnRemove.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } })
    $script:BtnClearPicked.Add_Click({
        if (-not $script:Picked.Count) { return }
        if ((Show-Msg "Alle $($script:Picked.Count) gekozen leden verwijderen?" "Leegmaken" "YesNo" "Question") -eq 'Yes') {
            $script:Picked.Clear(); Update-PickedGrid
        }
    })

    # --- Aanmaken ---
    $script:BtnCreate.Add_Click({ Invoke-CreateGroup })
    $script:BtnReset.Add_Click({
        if ((Show-Msg "Formulier leegmaken voor een nieuwe groep?" "Nieuw formulier" "YesNo" "Question") -eq 'Yes') { Reset-Form }
    })

    # --- Bestaande groep beheren ---
    $script:BtnMgSearch.Add_Click({ Invoke-GroupSearch })
    $script:TxtMgSearch.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-GroupSearch } })
    $script:GridGroups.Add_SelectionChanged({
        if ($script:MgSuppress) { return }
        $idx = $script:GridGroups.SelectedIndex
        if ($idx -lt 0 -or $idx -eq $script:MgSelIdx) { return }
        if ($script:Mg -and -not (Test-MgLeave)) {
            # Terug naar de groep met openstaande wijzigingen
            $script:MgSuppress = $true
            $script:GridGroups.SelectedIndex = $script:MgSelIdx
            $script:MgSuppress = $false
            return
        }
        $script:MgSelIdx = $idx
        Open-ManagedGroup ([string]$script:GridGroups.SelectedItem.Row['Id'])
    })
    $script:BtnMgSave.Add_Click({ Save-ManagedGroup })
    $script:BtnMgUndo.Add_Click({ if ($script:Mg) { Set-MgForm $script:Mg.Orig; Set-Status "Formulier teruggezet" } })
    $script:BtnMgDelete.Add_Click({ Remove-ManagedGroup })
    $script:BtnMgReload.Add_Click({ if ($script:Mg -and (Test-MgLeave)) { Open-ManagedGroup $script:Mg.Id } })
    $script:BtnMgAdd.Add_Click({ Add-MgMembers })
    $script:BtnMgRemove.Add_Click({ Invoke-MgMemberAction 'Remove' })
    $script:BtnMgRole.Add_Click({ Invoke-MgMemberAction 'Role' })
    $script:BtnMgSub.Add_Click({ Invoke-MgMemberAction 'Sub' })
    $script:BtnMgSubAll.Add_Click({ Invoke-MgMemberAction 'SubAll' })
    $script:BtnMgSendAs.Add_Click({ Invoke-MgMemberAction 'SendAs' })
    $script:BtnMgSob.Add_Click({ Invoke-MgMemberAction 'Sob' })
    $script:BtnMgExport.Add_Click({ Export-MgMembers })

    # --- Algemeen ---
    $script:BtnClearLog.Add_Click({ $script:LogPara.Inlines.Clear() })
    $script:BtnCopyLog.Add_Click({
        $doc = $script:LogBox.Document
        $txt = (New-Object System.Windows.Documents.TextRange($doc.ContentStart, $doc.ContentEnd)).Text
        if ($txt.Trim()) { [System.Windows.Clipboard]::SetText($txt); Set-Status "Logboek gekopieerd naar het klembord" "success" }
    })
    $script:Window.Add_Closed({
        if ($script:ExoConnected) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { } }
        if ($script:GraphConnected) { try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { } }
    })

    Update-PickedGrid
}

function Show-MainWindow {
    New-MainWindow
    Write-Log "Klik 'Verbinden' (rechtsboven): eerst Microsoft Graph, daarna Exchange Online. Vul optioneel een tenant in voor GDAP." "info"
    Write-Log "Standaard: mail aan de groep komt in het postvak IN van alle leden, zonder welkomstmail en zonder Team." "info"
    [void]$script:Window.ShowDialog()
}

#endregion

$missingModule = Initialize-Modules
if ($missingModule) {
    $cmd = if ($missingModule -like 'Microsoft.Graph*') { "Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser" } else { "Install-Module ExchangeOnlineManagement -Scope CurrentUser" }
    [System.Windows.MessageBox]::Show("Module $missingModule niet gevonden.`n`nInstalleer met:`n$cmd", "Module ontbreekt", "OK", "Error") | Out-Null
    exit 1
}
Show-MainWindow
