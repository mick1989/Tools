#Requires -Version 5.1
<#
.SYNOPSIS
    Entra & Exchange Gebruikersbeheer (WPF) - per gebruiker en in bulk,
    met een wachtrij die alle wijzigingen in 1 keer uitvoert.
.DESCRIPTION
    WPF-app om Entra ID-eigenschappen en Exchange Online rechten aan te passen,
    voor 1 gebruiker of voor een hele selectie tegelijk.

    WACHTRIJ: geen enkele wijziging wordt direct doorgevoerd. Alles (per
    gebruiker, bulk en CSV) komt eerst in de wachtrij onderin. Met 'Alles
    uitvoeren' worden alle wijzigingen in 1 keer afgehandeld:
      1. Exchange-wijzigingen (in volgorde van de wachtrij)
      2. Entra-wijzigingen, per gebruiker samengevoegd, via Graph $batch
    Exchange gaat eerst, zodat die nog op de huidige UPN werkt als de UPN in
    dezelfde run wijzigt. Mislukte items blijven staan en worden bij de
    volgende run opnieuw geprobeerd.

    Links: gebruikers laden (Microsoft Graph), filteren op naam/UPN/afdeling/
    functie/bedrijf en 1 of meerdere gebruikers selecteren. Na verbinden met
    Exchange tonen de kolommen Mailbox (gebruiker/gedeeld/ruimte/apparatuur/
    geen) en Doorsturen de stand per account, met filter op mailboxtype.

    1 gebruiker:  Entra eigenschappen, mailbox rechten (FullAccess / SendAs /
                  SendOnBehalf), doorsturen (intern/extern, kopie bewaren,
                  inbox-regels die doorsturen), map rechten (agenda e.a.) en
                  'toegang tot andere mailboxen' (offboarding).
    Doorsturen:   overzicht van alle mailboxen in de tenant met doorsturen,
                  met export en 'uitzetten voor selectie'.
    Ontvangers:   alles met een mailadres (mailboxen, M365-groepen, distributie-
                  lijsten, beveiligingsgroepen, contacten, gasten) met aliassen,
                  beheerder, verborgen-status, leden per groep en export.
                  Leden toevoegen/verwijderen (ook adressen plakken uit Outlook/
                  Excel) en contactpersonen aanmaken voor externe adressen.
                  M365-groep: groepsinstellingen (groepsmail in de inbox van
                  nieuwe/huidige leden, welkomstmail, extern mailen, MailTip,
                  privacy, verbergen in GAL/Outlook, agenda alleen-lezen,
                  beschrijving, verzenden als/namens) en per lid de groepsmail
                  in de eigen inbox aan/uit (abonnees).
    Nieuwe groep: distributielijst, beveiligingsgroep met e-mail, Microsoft
                  365-groep, dynamische distributielijst of ruimtelijst aanmaken,
                  met eigenaren, leden (ook plakken) en instellingen (extern
                  mailen, verbergen, lid worden, privacy, filter). Via de wachtrij.
    Bulk:         Entra eigenschappen, mailbox rechten en map rechten voor de
                  hele selectie.
    Overzicht:    tabel met Entra-velden van alle zichtbare (of geselecteerde)
                  gebruikers; waarden per cel aanpassen, plakken uit Excel of
                  een waarde in 1 keer vervangen (bijv. afdeling opschonen).
    CSV:          exporteren, aanpassen in Excel, importeren met preview.

    Optioneel een tenant (domein of tenant-ID) invullen om via GDAP bij een
    klant-tenant in te loggen.
.NOTES
    Vereist: Microsoft.Graph.Authentication (2.33.0 aanbevolen) en
             ExchangeOnlineManagement 3.x
    Installeren:
      Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
    Rollen: User Administrator (Entra) en Exchange Administrator (Exchange).
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

$script:GraphRoot      = "https://graph.microsoft.com/v1.0"
$script:GraphScopes    = @("User.ReadWrite.All")
$script:Window         = $null
$script:ExoConnected   = $false
$script:GraphConnected = $false
$script:ClearToken     = '#WISSEN'

$script:CurrentUser     = $null   # hashtable van Graph
$script:CurrentUserId   = $null
$script:ActualManager   = $null   # manager zoals nu in Entra
$script:Orig            = @{}     # prop -> waarde zoals getoond (Entra + wachtrij)
$script:QueuedActual    = @{}     # prop -> huidige Entra-waarde, voor velden met wachtrij-wijziging
$script:QueuedManager   = $false
$script:OrigEnabled     = $true
$script:OrigManager     = $null
$script:PendingManager  = $null
$script:ManagerChanged  = $false
$script:FieldBoxes      = @{}
$script:BulkBoxes       = @{}
$script:Loading         = $false
$script:SuppressSel     = $false
$script:ExoState        = @{}
$script:UserTable       = $null
$script:SelectedUsers   = @()
$script:BulkManager     = $null
$script:BulkMbxTarget   = $null
$script:BulkFolTarget   = $null
$script:FolderPathCache = @{}
$script:LongOp          = $false
$script:StopRequested   = $false
$script:AutoFwdText     = $null   # uitgaand spambeleid (extern doorsturen), 1x per verbinding
$script:AcceptedDomains = $null   # eigen domeinen van de tenant, 1x per verbinding
$script:RcptCache       = $null   # alle ontvangers voor de ledenzoeker, 1x per verbinding (Vernieuwen in de zoeker)
$script:ExoScopes       = @("https://outlook.office365.com/.default")
$script:ExoMsal         = $null   # MSAL-app voor Exchange-login via de browser
$script:ExoMsalAuthority = $null
$script:ExoAuth         = $null   # account + verloopmoment van het Exchange-token

$script:UserColumns   = @('Naam', 'UPN', 'Afdeling', 'Functie', 'Bedrijf', 'Status', 'Type', 'Id', 'Mailbox', 'Doorsturen')
$script:UserSearchAll = $false  # lijst bevat alle gebruikers (zoekveld leeg)
$script:MbxSummary    = $null   # tellingen per mailboxtype, gevuld na Exchange-verbinding
$script:MbxTypeLabels = @{
    UserMailbox = 'Gebruiker'; SharedMailbox = 'Gedeeld'; RoomMailbox = 'Ruimte'
    EquipmentMailbox = 'Apparatuur'; SchedulingMailbox = 'Planning'
}

# Ontvangers-overzicht: omschrijving en filtercategorie per RecipientTypeDetails
$script:RcptTypes = @{
    UserMailbox                    = @{ Label = 'Gebruikersmailbox';          Cat = 'Mailbox' }
    SharedMailbox                  = @{ Label = 'Gedeelde mailbox';           Cat = 'Mailbox' }
    RoomMailbox                    = @{ Label = 'Ruimte';                     Cat = 'Mailbox' }
    EquipmentMailbox               = @{ Label = 'Apparatuur';                 Cat = 'Mailbox' }
    SchedulingMailbox              = @{ Label = 'Planningsmailbox';           Cat = 'Mailbox' }
    DiscoveryMailbox               = @{ Label = 'Discovery-mailbox';          Cat = 'Overig' }
    GroupMailbox                   = @{ Label = 'Microsoft 365-groep';        Cat = 'Groep' }
    MailUniversalDistributionGroup = @{ Label = 'Distributielijst';           Cat = 'Groep' }
    MailNonUniversalGroup          = @{ Label = 'Distributielijst (oud)';     Cat = 'Groep' }
    MailUniversalSecurityGroup     = @{ Label = 'Beveiligingsgroep (mail)';   Cat = 'Groep' }
    DynamicDistributionGroup       = @{ Label = 'Dynamische distributielijst'; Cat = 'Groep' }
    RoomList                       = @{ Label = 'Ruimtelijst';                Cat = 'Groep' }
    MailContact                    = @{ Label = 'Contactpersoon';             Cat = 'Contact' }
    MailUser                       = @{ Label = 'E-mailgebruiker';            Cat = 'Contact' }
    GuestMailUser                  = @{ Label = 'Gast';                       Cat = 'Contact' }
    PublicFolder                   = @{ Label = 'Openbare map';               Cat = 'Overig' }
}

# Nieuwe groep: omschrijving en RecipientTypeDetails per soort
$script:NewGroupLabels = @{
    Distribution = 'Distributielijst'; Security = 'Beveiligingsgroep (mail)'; Unified = 'Microsoft 365-groep'
    Dynamic = 'Dynamische distributielijst'; RoomList = 'Ruimtelijst'
}
$script:NewGroupRType = @{
    Distribution = 'MailUniversalDistributionGroup'; Security = 'MailUniversalSecurityGroup'; Unified = 'GroupMailbox'
    Dynamic = 'DynamicDistributionGroup'; RoomList = 'RoomList'
}

# Instellingen M365-groep (aan/uit): checkbox, eigenschap bij Get-UnifiedGroup,
# parameter bij Set-UnifiedGroup en of de checkbox het omgekeerde toont.
$script:GroupBoolSettings = @(
    [pscustomobject]@{ Box = 'GS_ChkAutoSub';     Get = 'AutoSubscribeNewMembers';                Set = 'AutoSubscribeNewMembers';                Invert = $false; Label = 'Groepsmail naar inbox nieuwe leden' }
    [pscustomobject]@{ Box = 'GS_ChkCalSub';      Get = 'AlwaysSubscribeMembersToCalendarEvents'; Set = 'AlwaysSubscribeMembersToCalendarEvents'; Invert = $false; Label = 'Groepsafspraken naar agenda nieuwe leden' }
    [pscustomobject]@{ Box = 'GS_ChkSubEnabled';  Get = 'SubscriptionEnabled';                    Set = 'SubscriptionEnabled';                    Invert = $false; Label = 'Abonneren toegestaan' }
    [pscustomobject]@{ Box = 'GS_ChkWelcome';     Get = 'WelcomeMessageEnabled';                  Set = 'UnifiedGroupWelcomeMessageEnabled';      Invert = $false; Label = 'Welkomstmail nieuwe leden' }
    [pscustomobject]@{ Box = 'GS_ChkExternal';    Get = 'RequireSenderAuthenticationEnabled';     Set = 'RequireSenderAuthenticationEnabled';     Invert = $true;  Label = 'Extern mailen naar de groep' }
    [pscustomobject]@{ Box = 'GS_ChkHideGal';     Get = 'HiddenFromAddressListsEnabled';          Set = 'HiddenFromAddressListsEnabled';          Invert = $false; Label = 'Verborgen in adresboek' }
    [pscustomobject]@{ Box = 'GS_ChkHideOutlook'; Get = 'HiddenFromExchangeClientsEnabled';       Set = 'HiddenFromExchangeClientsEnabled';       Invert = $false; Label = 'Verborgen in Outlook' }
    [pscustomobject]@{ Box = 'GS_ChkCalRO';       Get = 'CalendarMemberReadOnly';                 Set = 'CalendarMemberReadOnly';                 Invert = $false; Label = 'Groepsagenda alleen-lezen voor leden' }
)

$script:Queue      = New-Object System.Collections.ArrayList
$script:QueueSeq   = 0
$script:QueueTable = $null

# Bewerkbare Entra-velden. Group bepaalt het blok, Bulk of het veld in het
# bulkscherm staat (persoonlijke/unieke velden alleen per gebruiker of CSV).
$script:EntraFields = @(
    [pscustomobject]@{ Prop = 'givenName';         Label = 'Voornaam';          Group = 'Identiteit';  Bulk = $false }
    [pscustomobject]@{ Prop = 'surname';           Label = 'Achternaam';        Group = 'Identiteit';  Bulk = $false }
    [pscustomobject]@{ Prop = 'displayName';       Label = 'Weergavenaam';      Group = 'Identiteit';  Bulk = $false }
    [pscustomobject]@{ Prop = 'userPrincipalName'; Label = 'UPN (inlognaam)';   Group = 'Identiteit';  Bulk = $false }
    [pscustomobject]@{ Prop = 'employeeId';        Label = 'Personeelsnummer';  Group = 'Identiteit';  Bulk = $false }
    [pscustomobject]@{ Prop = 'employeeType';      Label = 'Medewerkertype';    Group = 'Identiteit';  Bulk = $true }
    [pscustomobject]@{ Prop = 'jobTitle';          Label = 'Functie';           Group = 'Organisatie'; Bulk = $true }
    [pscustomobject]@{ Prop = 'department';        Label = 'Afdeling';          Group = 'Organisatie'; Bulk = $true }
    [pscustomobject]@{ Prop = 'companyName';       Label = 'Bedrijf';           Group = 'Organisatie'; Bulk = $true }
    [pscustomobject]@{ Prop = 'officeLocation';    Label = 'Kantoor';           Group = 'Organisatie'; Bulk = $true }
    [pscustomobject]@{ Prop = 'businessPhones';    Label = 'Telefoon werk';     Group = 'Contact';     Bulk = $true }
    [pscustomobject]@{ Prop = 'mobilePhone';       Label = 'Mobiel';            Group = 'Contact';     Bulk = $false }
    [pscustomobject]@{ Prop = 'faxNumber';         Label = 'Fax';               Group = 'Contact';     Bulk = $true }
    [pscustomobject]@{ Prop = 'streetAddress';     Label = 'Straat + nr';       Group = 'Adres';       Bulk = $true }
    [pscustomobject]@{ Prop = 'postalCode';        Label = 'Postcode';          Group = 'Adres';       Bulk = $true }
    [pscustomobject]@{ Prop = 'city';              Label = 'Plaats';            Group = 'Adres';       Bulk = $true }
    [pscustomobject]@{ Prop = 'state';             Label = 'Provincie';         Group = 'Adres';       Bulk = $true }
    [pscustomobject]@{ Prop = 'country';           Label = 'Land';              Group = 'Adres';       Bulk = $true }
    [pscustomobject]@{ Prop = 'usageLocation';     Label = 'Gebruikslocatie';   Group = 'Adres';       Bulk = $true }
)

$script:FolderRights = @(
    [pscustomobject]@{ Value = 'Owner';            Label = 'Owner - eigenaar (alles, incl. rechten beheren)' }
    [pscustomobject]@{ Value = 'PublishingEditor'; Label = 'PublishingEditor - alles + submappen maken' }
    [pscustomobject]@{ Value = 'Editor';           Label = 'Editor - lezen, maken, wijzigen, verwijderen' }
    [pscustomobject]@{ Value = 'PublishingAuthor'; Label = 'PublishingAuthor - maken, lezen, eigen items + submappen' }
    [pscustomobject]@{ Value = 'Author';           Label = 'Author - maken, lezen, eigen items wijzigen' }
    [pscustomobject]@{ Value = 'NonEditingAuthor'; Label = 'NonEditingAuthor - maken, lezen, eigen items verwijderen' }
    [pscustomobject]@{ Value = 'Reviewer';         Label = 'Reviewer - alleen lezen' }
    [pscustomobject]@{ Value = 'Contributor';      Label = 'Contributor - alleen items maken' }
    [pscustomobject]@{ Value = 'AvailabilityOnly'; Label = 'AvailabilityOnly - alleen beschikbaarheid (agenda)' }
    [pscustomobject]@{ Value = 'LimitedDetails';   Label = 'LimitedDetails - beschikbaarheid + onderwerp/locatie (agenda)' }
    [pscustomobject]@{ Value = 'None';             Label = 'None - geen toegang' }
)

$script:FolderNames = @{ Calendar = 'Agenda'; Contacts = 'Contactpersonen'; Inbox = 'Postvak IN' }

# Overzicht bewerken: standaard zichtbare kolommen (onthouden binnen de sessie)
# en velden die ook bij AD-gesynchroniseerde gebruikers in de cloud beheerd worden.
$script:OvDefaultProps = @('department', 'jobTitle', 'companyName', 'officeLocation')
$script:CloudOnlyProps = @('usageLocation')
$script:Ov             = $null

#endregion

#region --- UI basis-helpers ---

function New-Brush {
    param([string]$Hex)
    $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
    $b.Freeze()
    return $b
}

function Update-UI {
    # Laat WPF de wachtrij verwerken zodat statusbalk/log (en de Stop-knop)
    # werken tijdens langlopende Graph/Exchange-calls op de UI-thread.
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
    $script:QueueButtons.IsEnabled = $false
    $script:Window.Cursor = [System.Windows.Input.Cursors]::Wait
    if ($Text) { Set-Status $Text }
}

function Exit-Busy {
    if ($script:LongOp) { return }
    $script:MainArea.IsEnabled = $true
    $script:HeaderPanel.IsEnabled = $true
    $script:QueueButtons.IsEnabled = $true
    $script:Window.Cursor = $null
}

function Start-LongOp {
    # Busy + een Stop-knop in de statusbalk (buiten het uitgeschakelde gebied)
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

#endregion

#region --- Modules en verbinding ---

function Initialize-Modules {
    # ExchangeOnlineManagement en Microsoft.Graph.Authentication gebruiken
    # allebei Microsoft.Identity.Client (MSAL); de eerst geladen versie wint.
    # EXO 3.9 levert een nieuwere MSAL mee dan Graph 2.33 en Graph kan met de
    # nieuwere overweg, andersom niet ('Method not found' bij
    # Connect-ExchangeOnline). Daarom laden we de MSAL van EXO als eerste.
    $script:ExoAvailable = $false
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $base = (Get-Module ExchangeOnlineManagement).ModuleBase
        $sub  = if ($PSVersionTable.PSEdition -eq 'Core') { 'netCore' } else { 'netFramework' }
        $msal = Join-Path $base "$sub\Microsoft.Identity.Client.dll"
        if (Test-Path $msal) { try { Add-Type -Path $msal -ErrorAction Stop } catch { } }
        $script:ExoAvailable = $true
    }
    catch { }

    # Vanaf module-versie 2.34.0 dwingt Microsoft.Graph.Authentication WAM
    # af (geen account-kiezer). 2.33.0 gebruikt nog de normale browser-login.
    try {
        Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -ErrorAction Stop
    }
    catch {
        try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
        catch { return $false }
    }
    return $true
}

function Get-TenantInput {
    $t = $script:TxtTenant.Text.Trim()
    if ($t) { return $t }
    return $null
}

function Connect-ToGraph {
    try {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        $p = @{ Scopes = $script:GraphScopes; NoWelcome = $true; ErrorAction = 'Stop' }
        $tenant = Get-TenantInput
        if ($tenant) { $p.TenantId = $tenant }
        Connect-MgGraph @p | Out-Null
        return $true
    }
    catch {
        Show-Msg "Verbinding met Graph mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" | Out-Null
        return $false
    }
}

function Start-LoginWindowWatcher {
    # Het Exchange-inlogvenster (zonder WAM) is een WinForms-venster in dit
    # proces dat achter het hoofdvenster opent; de UI-thread wacht dan op de
    # login en de tool lijkt vast te lopen. Een losse thread haalt elk nieuw
    # WinForms-venster van dit proces naar voren.
    if (-not ('LoginWindowWatcher' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
public static class LoginWindowWatcher {
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    static volatile bool running;
    public static void Start(int timeoutSeconds) {
        running = true;
        uint me = (uint)Process.GetCurrentProcess().Id;
        var t = new Thread(() => {
            var done = new HashSet<IntPtr>();
            var end = DateTime.Now.AddSeconds(timeoutSeconds);
            while (running && DateTime.Now < end) {
                EnumWindows((h, l) => {
                    uint pid; GetWindowThreadProcessId(h, out pid);
                    if (pid != me || done.Contains(h) || !IsWindowVisible(h)) return true;
                    var sb = new StringBuilder(256); GetClassName(h, sb, 256);
                    if (!sb.ToString().StartsWith("WindowsForms10.")) return true;
                    done.Add(h);
                    ShowWindow(h, 9);                                  // SW_RESTORE
                    SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, 0x43); // TOPMOST, geen move/size, tonen
                    SetWindowPos(h, new IntPtr(-2), 0, 0, 0, 0, 0x43); // weer NOTOPMOST
                    SetForegroundWindow(h);
                    return true;
                }, IntPtr.Zero);
                Thread.Sleep(300);
            }
        });
        t.IsBackground = true;
        t.Start();
    }
    public static void Stop() { running = false; }
}
'@
    }
    [LoginWindowWatcher]::Start(600)
}

function Wait-MsalTask {
    # Wacht op een MSAL-task terwijl de UI blijft reageren
    param($Task)
    while (-not $Task.IsCompleted) { Update-UI; Start-Sleep -Milliseconds 100 }
    if ($Task.IsCanceled) { throw "Inloggen geannuleerd of verlopen" }
    if ($Task.IsFaulted) { throw $Task.Exception.GetBaseException().Message }
    return $Task.Result
}

function Get-ExoBrowserToken {
    # Inloggen via de standaardbrowser, net als Graph 2.33, met de
    # app-registratie van de EXO-module. Het ingebouwde inlogvenster van
    # Connect-ExchangeOnline (zonder WAM) blijft achter de WPF-app hangen.
    param([string]$Tenant, [string]$LoginHint)
    $authority = "https://login.microsoftonline.com/$(if ($Tenant) { $Tenant } else { 'organizations' })"
    if (-not $script:ExoMsal -or $script:ExoMsalAuthority -ne $authority) {
        $script:ExoMsal = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create('fb78d390-0c51-40cd-8e17-fdbfab77341b').
            WithAuthority($authority).WithRedirectUri('http://localhost').Build()
        $script:ExoMsalAuthority = $authority
    }
    $b = $script:ExoMsal.AcquireTokenInteractive([string[]]$script:ExoScopes).
        WithUseEmbeddedWebView($false).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)
    if ($LoginHint) { $b = $b.WithLoginHint($LoginHint) }
    $cts = New-Object System.Threading.CancellationTokenSource ([TimeSpan]::FromMinutes(5))
    return (Wait-MsalTask $b.ExecuteAsync($cts.Token))
}

function Connect-ExoWithToken {
    param($Auth, [string]$Tenant)
    $p = @{ AccessToken = $Auth.AccessToken; UserPrincipalName = $Auth.Account.Username; ShowBanner = $false; ErrorAction = 'Stop' }
    if ($Tenant) { $p.DelegatedOrganization = $Tenant }
    Connect-ExchangeOnline @p | Out-Null
    $script:ExoAuth = @{ Account = $Auth.Account; Expires = $Auth.ExpiresOn.LocalDateTime; Tenant = $Tenant }
}

function Update-ExoToken {
    # Een sessie met -AccessToken ververst zichzelf niet: kort voor het
    # verlopen stil een nieuw token halen en opnieuw verbinden.
    if (-not $script:ExoAuth -or (Get-Date) -lt $script:ExoAuth.Expires.AddMinutes(-5)) { return }
    try {
        $r = Wait-MsalTask $script:ExoMsal.AcquireTokenSilent([string[]]$script:ExoScopes, $script:ExoAuth.Account).ExecuteAsync()
        Connect-ExoWithToken $r $script:ExoAuth.Tenant
        Write-Log "Exchange-token vernieuwd (geldig tot $($script:ExoAuth.Expires.ToString('HH:mm')))" "info"
    }
    catch {
        $script:ExoAuth = $null
        Write-Log "Exchange-token vernieuwen mislukt, klik zo nodig op 'Exchange opnieuw': $_" "warning"
    }
}

function Connect-ToExchange {
    $tenant = Get-TenantInput
    $upn = if ($script:GraphConnected) { (Get-MgContext).Account } else { $null }
    $script:ExoAuth = $null

    # 1. Via de browser (zoals Graph); hoofdvenster even minimaliseren
    $browserErr = $null
    $script:Window.WindowState = 'Minimized'
    try {
        $auth = Get-ExoBrowserToken -Tenant $tenant -LoginHint $upn
        Connect-ExoWithToken $auth $tenant
        return $true
    }
    catch { $browserErr = "$_" }
    finally {
        $script:Window.WindowState = 'Normal'
        [void]$script:Window.Activate()
    }
    if ($browserErr -match 'geannuleerd|cancel|access_denied') {
        Show-Msg "Inloggen bij Exchange Online geannuleerd." "Exchange" "OK" "Warning" | Out-Null
        return $false
    }
    Write-Log "Browser-login voor Exchange mislukt ($browserErr), terugval op standaard inlogvenster" "warning"

    # 2. Terugval: standaard inlogvenster van de module, naar voren gehaald
    Start-LoginWindowWatcher
    try {
        $p = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($tenant) { $p.DelegatedOrganization = $tenant }
        if ($upn) { $p.UserPrincipalName = $upn }
        if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) { $p.DisableWAM = $true }
        Connect-ExchangeOnline @p | Out-Null
        return $true
    }
    catch {
        Show-Msg "Verbinding met Exchange Online mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" | Out-Null
        return $false
    }
    finally { [LoginWindowWatcher]::Stop() }
}

function Invoke-Graph {
    param([string]$Uri, [string]$Method = "GET", $Body = $null, [hashtable]$Headers = $null, [switch]$AllowNotFound, [switch]$Quiet)
    $short = $Uri -replace 'https://graph\.microsoft\.com/(v1\.0|beta)/', ''
    if (-not $Quiet) { Write-Log "$Method $short" "api" }
    $p = @{ Uri = $Uri; Method = $Method }
    if ($null -ne $Body) { $p.Body = ($Body | ConvertTo-Json -Depth 10); $p.ContentType = "application/json" }
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
    param([string]$Uri, [hashtable]$Headers = $null)
    $all = New-Object System.Collections.ArrayList
    do {
        $r = Invoke-Graph -Uri $Uri -Headers $Headers
        foreach ($v in @($r.value)) { [void]$all.Add($v) }
        $Uri = $r['@odata.nextLink']
        if ($Uri) { Set-Status "$($all.Count) geladen..." }
    } while ($Uri -and -not $script:StopRequested)
    return ,$all
}

function Invoke-GraphBatch {
    # Voert losse Graph-requests uit via /$batch (max 20 per aanroep).
    # Requests: @{ id; method; url (relatief, bijv. /users/x); body }
    # Geeft een hashtable id -> @{ status; body } terug. 429 (throttling)
    # wordt maximaal 3x opnieuw geprobeerd.
    param([System.Collections.IList]$Requests, [string]$What = "Graph")
    $results = @{}
    $total = $Requests.Count
    for ($i = 0; $i -lt $total; $i += 20) {
        if ($script:StopRequested) { Write-Log "Gestopt door gebruiker na $i van $total requests" "warning"; break }
        $todo = @($Requests[$i..([Math]::Min($i + 19, $total - 1))])
        Set-Status "$What $([Math]::Min($i + 20, $total)) / $total..."
        for ($attempt = 1; $attempt -le 4 -and $todo.Count -gt 0; $attempt++) {
            $payload = @{
                requests = @(foreach ($r in $todo) {
                    $o = [ordered]@{ id = [string]$r.id; method = $r.method; url = $r.url }
                    if ($null -ne $r.body) { $o.body = $r.body; $o.headers = @{ 'Content-Type' = 'application/json' } }
                    $o
                })
            }
            $resp = Invoke-Graph "$($script:GraphRoot)/`$batch" -Method POST -Body $payload -Quiet
            $retry = @()
            $wait = 2
            foreach ($x in @($resp.responses)) {
                if ([int]$x.status -eq 429 -and $attempt -lt 4) {
                    $retry += @($todo | Where-Object { [string]$_.id -eq [string]$x.id })
                    if ($x.headers -and $x.headers['Retry-After']) { $wait = [Math]::Max($wait, [int]$x.headers['Retry-After']) }
                }
                else { $results[[string]$x.id] = $x }
            }
            $todo = $retry
            if ($todo.Count -gt 0) {
                Write-Log "Throttling: $($todo.Count) request(s) opnieuw over $wait s" "warning"
                Start-Sleep -Seconds ([Math]::Min($wait, 30))
            }
        }
    }
    return $results
}

function Get-BatchError {
    param($Response)
    if ($null -eq $Response) { return "Niet uitgevoerd" }
    $m = $null
    try { $m = $Response.body.error.message } catch { }
    if (-not $m) { $m = "HTTP $($Response.status)" }
    return $m
}

function Test-BatchOk {
    param($Response)
    return ($null -ne $Response -and [int]$Response.status -ge 200 -and [int]$Response.status -lt 300)
}

function Invoke-Exo {
    # Voert een Exchange-cmdlet uit met logging. De scriptblock moet zelf
    # -ErrorAction Stop meegeven: EXO-cmdlets zijn module-functies en zien
    # $ErrorActionPreference van de aanroeper niet.
    param([string]$Desc, [scriptblock]$Script)
    Update-ExoToken
    Write-Log $Desc "api"
    try {
        return & $Script
    }
    catch {
        $m = $_.Exception.Message
        Write-Log "FOUT: $m" "error"
        throw $m
    }
}

function Test-ExoReady {
    if ($script:ExoConnected) { return $true }
    Set-Status "Verbind eerst met Exchange Online (knop rechtsboven)" "error"
    Show-Msg "Verbind eerst met Exchange Online (knop rechtsboven)." "Niet verbonden" "OK" "Warning" | Out-Null
    return $false
}

#endregion

#region --- Wachtrij ---

function Get-StatusColor {
    param([string]$Status)
    if ($Status -eq 'Wachtend')        { return '#0063B1' }
    if ($Status -like 'Bezig*')        { return '#8A6D00' }
    if ($Status -like 'Gereed*')       { return '#107C10' }
    if ($Status -like 'Mislukt*')      { return '#C42B1C' }
    return '#555555'
}

function Test-QueueOpen {
    # Nog uit te voeren: wachtend of eerder mislukt (wordt opnieuw geprobeerd)
    param($Item)
    return ($Item.Status -eq 'Wachtend' -or $Item.Status -like 'Mislukt*')
}

function Add-QueueItem {
    # Zet 1 wijziging in de wachtrij. Kind: GraphPatch, GraphManager,
    # MbxRight, FolderPerm, Forward, MailContact, NewGroup, GroupMember. Identieke openstaande items worden niet dubbel
    # toegevoegd. Geeft $true terug als het item is toegevoegd.
    param([string]$Kind, [string]$Target, [string]$Soort, [string]$Text, [hashtable]$Data)
    $sig = "$Kind|" + ((@($Data.Keys) | Sort-Object | ForEach-Object { "$_=" + ($Data[$_] | ConvertTo-Json -Compress -Depth 6) }) -join '|')
    if (@($script:Queue | Where-Object { $_.Sig -eq $sig -and (Test-QueueOpen $_) }).Count -gt 0) { return $false }
    $script:QueueSeq++
    $item = [pscustomobject]@{
        Key = [string]$script:QueueSeq; Kind = $Kind; Target = $Target; Soort = $Soort
        Text = $Text; Data = $Data; Status = 'Wachtend'; Sig = $sig
    }
    [void]$script:Queue.Add($item)
    $dr = $script:QueueTable.NewRow()
    $dr['Nr'] = $item.Key; $dr['Doel'] = $Target; $dr['Soort'] = $Soort; $dr['Wijziging'] = $Text
    $dr['Status'] = 'Wachtend'; $dr['Kleur'] = Get-StatusColor 'Wachtend'; $dr['Key'] = $item.Key
    [void]$script:QueueTable.Rows.Add($dr)
    return $true
}

function Set-QueueStatus {
    param($Item, [string]$Status)
    $Item.Status = $Status
    foreach ($dr in $script:QueueTable.Select("Key = '$($Item.Key)'")) {
        $dr['Status'] = $Status
        $dr['Kleur'] = Get-StatusColor $Status
    }
}

function Remove-QueueItems {
    param([string[]]$Keys)
    $n = 0
    foreach ($k in $Keys) {
        $it = @($script:Queue | Where-Object { $_.Key -eq $k }) | Select-Object -First 1
        if (-not $it -or $it.Status -like 'Bezig*') { continue }
        $script:Queue.Remove($it)
        foreach ($dr in $script:QueueTable.Select("Key = '$k'")) { $dr.Delete() }
        $n++
    }
    $script:QueueTable.AcceptChanges()
    return $n
}

function Update-QueueSummary {
    $open = @($script:Queue | Where-Object { Test-QueueOpen $_ })
    $fail = @($open | Where-Object { $_.Status -like 'Mislukt*' }).Count
    $done = @($script:Queue | Where-Object { $_.Status -like 'Gereed*' }).Count
    $script:BtnRunQueue.Content = "Alles uitvoeren ($($open.Count))"
    $script:BtnRunQueue.IsEnabled = ($open.Count -gt 0)
    $parts = @("$($open.Count) te doen")
    if ($fail) { $parts += "waarvan $fail eerder mislukt" }
    if ($done) { $parts += "$done gereed" }
    $script:LblQueue.Text = "   " + ($parts -join ', ')
    Update-UserQueueNotice
}

function Get-UserQueueItems {
    # Openstaande items die over de huidige gebruiker gaan (Entra of zijn mailbox)
    if (-not $script:CurrentUser) { return @() }
    $ids = @([string]$script:CurrentUser.userPrincipalName, [string]$script:CurrentUser.mail) | Where-Object { $_ }
    if ($script:ExoState.Mailbox) { $ids += [string]$script:ExoState.Mailbox.PrimarySmtpAddress }
    return @($script:Queue | Where-Object {
        (Test-QueueOpen $_) -and (
            $_.Data.UserId -eq $script:CurrentUserId -or
            ($_.Data.Mailbox -and $ids -contains $_.Data.Mailbox) -or
            ($_.Data.Who -and $ids -contains $_.Data.Who))
    })
}

function Update-UserQueueNotice {
    if (-not $script:LblUserQueue) { return }
    $n = @(Get-UserQueueItems).Count
    $script:LblUserQueue.Text = if ($n) { "$n wijziging(en) voor deze gebruiker staan in de wachtrij (nog niet uitgevoerd)" } else { "" }
}

function Get-PendingUserOverlay {
    # Samengevoegde openstaande Entra-wijzigingen voor 1 gebruiker, in
    # wachtrijvolgorde (latere wint).
    param([string]$UserId)
    $body = [ordered]@{}
    $mgr = $null
    foreach ($it in $script:Queue) {
        if (-not (Test-QueueOpen $it) -or $it.Data.UserId -ne $UserId) { continue }
        if ($it.Kind -eq 'GraphPatch') { foreach ($k in @($it.Data.Body.Keys)) { $body[$k] = $it.Data.Body[$k] } }
        elseif ($it.Kind -eq 'GraphManager') { $mgr = $it.Data }
    }
    return @{ Body = $body; Manager = $mgr }
}

function Format-PatchValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) { return (@($Value) -join '; ') }
    return [string]$Value
}

function Format-BodyText {
    # Leesbare omschrijving van een PATCH-body voor de wachtrij
    param($Body, [hashtable]$Old = $null)
    $parts = foreach ($k in @($Body.Keys)) {
        if ($k -eq 'accountEnabled') { if ($Body[$k]) { 'Account INSCHAKELEN' } else { 'Account UITSCHAKELEN' }; continue }
        $f = $script:EntraFields | Where-Object { $_.Prop -eq $k } | Select-Object -First 1
        $label = if ($f) { $f.Label } else { $k }
        $new = Format-PatchValue $Body[$k]
        $from = if ($Old -and $Old.ContainsKey($k)) { "'$($Old[$k])' -> " } else { "" }
        if ($new) { "$label`: $from'$new'" } else { "$label`: $($from)LEEGMAKEN" }
    }
    return (@($parts) -join ';  ')
}

function Add-MbxRightToQueue {
    param([string]$Target, [string]$Kind, [string]$Mailbox, [string]$Who, [string]$WhoLabel, [bool]$Grant, [bool]$AutoMap = $true)
    if (-not $WhoLabel) { $WhoLabel = $Who }
    $verb = if ($Grant) { 'toekennen' } else { 'intrekken' }
    $am = if ($Grant -and $Kind -eq 'FullAccess') { " (automapping $(if ($AutoMap) { 'aan' } else { 'uit' }))" } else { "" }
    return (Add-QueueItem -Kind 'MbxRight' -Target $Target -Soort $Kind `
        -Text "$Kind $verb voor $WhoLabel op mailbox $Mailbox$am" `
        -Data @{ Kind = $Kind; Mailbox = $Mailbox; Who = $Who; WhoLabel = $WhoLabel; Grant = $Grant; AutoMap = $AutoMap })
}

function Add-FolderPermToQueue {
    # Op: Upsert (toevoegen of bijwerken), Set (bestaande wijzigen), Remove
    param([string]$Target, [string]$Op, [string]$Mailbox, [string]$FolderType, [string]$Path,
          [string]$Who, [string]$WhoLabel, [string]$Rights, [string]$Flags = 'None', [bool]$SendFlags = $false)
    if (-not $WhoLabel) { $WhoLabel = $Who }
    $fname = $script:FolderNames[$FolderType]
    $where = "$fname van $Mailbox"
    $text = switch ($Op) {
        'Upsert' { "Rechten $Rights$(if ($Flags -ne 'None') { " ($Flags)" }) instellen voor $WhoLabel op $where" }
        'Set'    { "Rechten van $WhoLabel op $where wijzigen naar $Rights$(if ($SendFlags -and $Flags -ne 'None') { " ($Flags)" })" }
        'Remove' { "Rechten van $WhoLabel op $where verwijderen" }
    }
    return (Add-QueueItem -Kind 'FolderPerm' -Target $Target -Soort $fname -Text $text `
        -Data @{ Op = $Op; Mailbox = $Mailbox; FolderType = $FolderType; Path = $Path; Who = $Who; WhoLabel = $WhoLabel
                 Rights = $Rights; Flags = $Flags; SendFlags = $SendFlags })
}

function Invoke-QueueItemExo {
    # Voert 1 Exchange-item uit. Geeft een korte resultaattekst terug of gooit een fout.
    param($Item)
    $d = $Item.Data
    switch ($Item.Kind) {
        'MbxRight' {
            Invoke-MbxRight -Kind $d.Kind -Mailbox $d.Mailbox -Who $d.Who -Grant $d.Grant -AutoMap $d.AutoMap
            return ''
        }
        'FolderPerm' {
            $path = if ($d.Path) { $d.Path } else { Get-FolderPathFor $d.Mailbox $d.FolderType }
            $who = $d.Who
            switch ($d.Op) {
                'Upsert' { return (Set-FolderPermission -Path $path -Who $who -Rights $d.Rights -Flags $d.Flags) }
                'Set' {
                    $p = @{ Identity = $path; User = $who; AccessRights = $d.Rights; Confirm = $false; ErrorAction = 'Stop' }
                    if ($d.SendFlags) { $p.SharingPermissionFlags = $d.Flags }
                    Invoke-Exo "Set-MailboxFolderPermission $path -User $who -AccessRights $($d.Rights)" { Set-MailboxFolderPermission @p | Out-Null }
                    return 'gewijzigd'
                }
                'Remove' {
                    Invoke-Exo "Remove-MailboxFolderPermission $path -User $who" {
                        Remove-MailboxFolderPermission -Identity $path -User $who -Confirm:$false -ErrorAction Stop
                    }
                    return 'verwijderd'
                }
            }
        }
        'Forward' {
            # Altijd beide adresvelden zetten: ForwardingAddress wint anders
            # stilletjes van een nieuw ForwardingSmtpAddress.
            $mb = $d.Mailbox
            $p = @{ Identity = $mb; ForwardingAddress = $null; ForwardingSmtpAddress = $null; DeliverToMailboxAndForward = $false; Confirm = $false; ErrorAction = 'Stop' }
            switch ($d.Mode) {
                'Internal' { $p.ForwardingAddress = $d.Address; $p.DeliverToMailboxAndForward = [bool]$d.KeepCopy }
                'External' { $p.ForwardingSmtpAddress = "smtp:$($d.Address)"; $p.DeliverToMailboxAndForward = [bool]$d.KeepCopy }
            }
            Invoke-Exo "Set-Mailbox $mb doorsturen: $($d.Mode)$(if ($d.Address) { " -> $($d.Address) (kopie $([bool]$d.KeepCopy))" })" { Set-Mailbox @p | Out-Null }
            if ($d.Mode -eq 'Off') { return 'uitgezet' }
            return 'ingesteld'
        }
        'MailContact' {
            $p = @{ Name = $d.Name; DisplayName = $d.Name; ExternalEmailAddress = $d.Address; ErrorAction = 'Stop' }
            Invoke-Exo "New-MailContact '$($d.Name)' <$($d.Address)>" { New-MailContact @p | Out-Null }
            if ($d.Hide) {
                # Net aangemaakt: kan even duren voordat het contact vindbaar is
                Invoke-WithRetry -Tries 4 {
                    Invoke-Exo "Set-MailContact $($d.Address) -HiddenFromAddressListsEnabled `$true" {
                        Set-MailContact -Identity $d.Address -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                    }
                }
                return 'aangemaakt, verborgen'
            }
            return 'aangemaakt'
        }
        'NewGroup' {
            return (Invoke-NewGroup $d)
        }
        'GroupMember' {
            # Extra pogingen als het lid of de groep in deze wachtrij net is aangemaakt
            $fresh = @($script:Queue | Where-Object { $_.Kind -eq 'NewGroup' -and $_.Data.Address -eq $d.Group }).Count -gt 0
            $tries = if ($d.Add -and ($d.NewContact -or $fresh)) { 4 } else { 1 }
            Invoke-WithRetry -Tries $tries { Invoke-GroupMemberChange $d }
            if ($d.Add) { return 'toegevoegd' }
            return 'verwijderd'
        }
        'GroupSettings' {
            return (Invoke-GroupSettings $d)
        }
        'GroupSubscriber' {
            $grp = [string]$d.Group
            $links = @($d.Members)
            if ($d.Add) {
                Invoke-Exo "Add-UnifiedGroupLinks $grp -LinkType Subscribers -Links $($links -join ',')" { Add-UnifiedGroupLinks -Identity $grp -LinkType Subscribers -Links $links -Confirm:$false -ErrorAction Stop }
                return 'geabonneerd'
            }
            Invoke-Exo "Remove-UnifiedGroupLinks $grp -LinkType Subscribers -Links $($links -join ',')" { Remove-UnifiedGroupLinks -Identity $grp -LinkType Subscribers -Links $links -Confirm:$false -ErrorAction Stop }
            return 'afgemeld'
        }
    }
}

function Invoke-WithRetry {
    # Opnieuw proberen zolang Exchange het object (nog) niet vindt, bijv.
    # direct na het aanmaken van een contactpersoon in dezelfde run.
    param([int]$Tries = 1, [scriptblock]$Action)
    for ($t = 1; ; $t++) {
        try { return (& $Action) }
        catch {
            if ($t -ge $Tries -or "$_" -notmatch "couldn't be found|could not be found|not found|niet gevonden|wasn't found") { throw }
            Write-Log "Nog niet gevonden (replicatie), poging $($t + 1) van $Tries over 5 seconden" "warning"
            for ($s = 0; $s -lt 10; $s++) { Start-Sleep -Milliseconds 500; Update-UI }
        }
    }
}

function Invoke-GroupMemberChange {
    # M365-groepen via UnifiedGroupLinks, de rest (distributielijsten,
    # mail-enabled beveiligingsgroepen, ruimtelijsten) via DistributionGroupMember.
    param($D)
    $grp = $D.Group; $mem = $D.Member
    if ($D.GroupType -eq 'GroupMailbox') {
        if ($D.Add) { Invoke-Exo "Add-UnifiedGroupLinks $grp -LinkType Members -Links $mem" { Add-UnifiedGroupLinks -Identity $grp -LinkType Members -Links $mem -Confirm:$false -ErrorAction Stop } }
        else { Invoke-Exo "Remove-UnifiedGroupLinks $grp -LinkType Members -Links $mem" { Remove-UnifiedGroupLinks -Identity $grp -LinkType Members -Links $mem -Confirm:$false -ErrorAction Stop } }
    }
    else {
        if ($D.Add) { Invoke-Exo "Add-DistributionGroupMember $grp -Member $mem" { Add-DistributionGroupMember -Identity $grp -Member $mem -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop } }
        else { Invoke-Exo "Remove-DistributionGroupMember $grp -Member $mem" { Remove-DistributionGroupMember -Identity $grp -Member $mem -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop } }
    }
}

function Add-GroupMemberToQueue {
    param([string]$GroupName, [string]$Group, [string]$GroupType, [string]$Member, [string]$MemberLabel, [bool]$Add, [bool]$NewContact = $false)
    if (-not $MemberLabel) { $MemberLabel = $Member }
    $text = if ($Add) { "$MemberLabel TOEVOEGEN aan groep $Group" } else { "$MemberLabel VERWIJDEREN uit groep $Group" }
    return (Add-QueueItem -Kind 'GroupMember' -Target $GroupName -Soort 'Groepslid' -Text $text `
        -Data @{ Group = $Group; GroupType = $GroupType; Member = $Member; MemberLabel = $MemberLabel; Add = $Add; NewContact = $NewContact })
}

function Add-MailContactToQueue {
    param([string]$Name, [string]$Address, [bool]$Hide = $false)
    if (-not $Name) { $Name = $Address }
    if ($Name.Length -gt 64) { $Name = $Name.Substring(0, 64) }
    $extra = if ($Hide) { ', verborgen in adresboek' } else { '' }
    return (Add-QueueItem -Kind 'MailContact' -Target $Name -Soort 'Contactpersoon' -Text "Contactpersoon AANMAKEN: $Name <$Address>$extra" `
        -Data @{ Name = $Name; Address = $Address; Hide = $Hide })
}

function Split-ListText {
    # "a; b, c" -> @('a','b','c')
    param([string]$Text)
    return @($Text -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Invoke-NewGroup {
    # Maakt 1 groep aan (Data uit Add-NewGroupToQueue) en zet daarna de
    # instellingen die niet bij het aanmaken kunnen. Lukt het aanmaken wel maar
    # een vervolgstap niet, dan wordt bij een nieuwe poging alleen de rest gedaan.
    param($D)
    $addr = [string]$D.Address
    $name = [string]$D.Name
    $short = if ($name.Length -gt 64) { $name.Substring(0, 64) } else { $name }
    $owners = @(@($D.Owners) | Where-Object { $_ })
    $res = @()
    if (-not $D.Created) {
        switch ($D.Type) {
            'Unified' {
                $p = @{ DisplayName = $name; Alias = $D.Alias; PrimarySmtpAddress = $addr; AccessType = $D.Privacy
                        RequireSenderAuthenticationEnabled = (-not $D.AllowExternal); ErrorAction = 'Stop' }
                if ($D.Description) { $p.Notes = $D.Description }
                if ($owners.Count) { $p.Owner = $owners[0] }
                Invoke-Exo "New-UnifiedGroup '$name' <$addr> ($($D.Privacy))" { New-UnifiedGroup @p | Out-Null }
            }
            'Dynamic' {
                $p = @{ Name = $short; DisplayName = $name; Alias = $D.Alias; PrimarySmtpAddress = $addr; ErrorAction = 'Stop' }
                if ($D.Filter) { $p.RecipientFilter = $D.Filter }
                else {
                    $p.IncludedRecipients = $D.Included
                    $dep = @(Split-ListText $D.Department); if ($dep.Count) { $p.ConditionalDepartment = $dep }
                    $com = @(Split-ListText $D.Company);    if ($com.Count) { $p.ConditionalCompany = $com }
                }
                if ($owners.Count) { $p.ManagedBy = $owners[0] }
                if ($D.Description) { $p.Notes = $D.Description }
                Invoke-Exo "New-DynamicDistributionGroup '$name' <$addr>" { New-DynamicDistributionGroup @p | Out-Null }
            }
            default {
                # Distribution, Security, RoomList
                $p = @{ Name = $short; DisplayName = $name; Alias = $D.Alias; PrimarySmtpAddress = $addr
                        Type = $(if ($D.Type -eq 'Security') { 'Security' } else { 'Distribution' })
                        MemberJoinRestriction = $D.Join; MemberDepartRestriction = $D.Depart
                        RequireSenderAuthenticationEnabled = (-not $D.AllowExternal); ErrorAction = 'Stop' }
                if ($D.Type -eq 'RoomList') { $p.RoomList = $true }
                if ($owners.Count) { $p.ManagedBy = $owners }
                if ($D.Description) { $p.Description = @($D.Description) }
                Invoke-Exo "New-DistributionGroup '$name' <$addr> -Type $($p.Type)$(if ($p.RoomList) { ' -RoomList' })" { New-DistributionGroup @p | Out-Null }
            }
        }
        $D.Created = $true
        $res += 'aangemaakt'
    }
    else { $res += 'was al aangemaakt' }

    # Vervolgstappen op het net aangemaakte object: opnieuw proberen bij replicatievertraging
    switch ($D.Type) {
        'Unified' {
            $extra = @($owners | Select-Object -Skip 1)
            if ($extra.Count -and -not $D.OwnersDone) {
                # Eigenaren moeten eerst lid zijn
                Invoke-WithRetry -Tries 4 {
                    Invoke-Exo "Add-UnifiedGroupLinks $addr -LinkType Members -Links $($extra -join ',')" { Add-UnifiedGroupLinks -Identity $addr -LinkType Members -Links $extra -Confirm:$false -ErrorAction Stop }
                }
                Invoke-Exo "Add-UnifiedGroupLinks $addr -LinkType Owners -Links $($extra -join ',')" { Add-UnifiedGroupLinks -Identity $addr -LinkType Owners -Links $extra -Confirm:$false -ErrorAction Stop }
                $D.OwnersDone = $true
                $res += "$($extra.Count) extra eigenaar(s)"
            }
            if ($D.Hide -or $D.AutoSubscribe) {
                $s = @{ Identity = $addr; ErrorAction = 'Stop' }
                if ($D.Hide) { $s.HiddenFromAddressListsEnabled = $true }
                if ($D.AutoSubscribe) { $s.AutoSubscribeNewMembers = $true }
                Invoke-WithRetry -Tries 4 { Invoke-Exo "Set-UnifiedGroup $addr$(if ($D.Hide) { ' -HiddenFromAddressListsEnabled' })$(if ($D.AutoSubscribe) { ' -AutoSubscribeNewMembers' })" { Set-UnifiedGroup @s } }
            }
        }
        'Dynamic' {
            # RequireSenderAuthenticationEnabled bestaat niet bij New-DynamicDistributionGroup
            $s = @{ Identity = $addr; RequireSenderAuthenticationEnabled = (-not $D.AllowExternal); ErrorAction = 'Stop' }
            if ($D.Hide) { $s.HiddenFromAddressListsEnabled = $true }
            Invoke-WithRetry -Tries 4 { Invoke-Exo "Set-DynamicDistributionGroup $addr (extern mailen $([bool]$D.AllowExternal)$(if ($D.Hide) { ', verborgen' }))" { Set-DynamicDistributionGroup @s } }
        }
        default {
            if ($D.Hide) {
                Invoke-WithRetry -Tries 4 {
                    Invoke-Exo "Set-DistributionGroup $addr -HiddenFromAddressListsEnabled `$true" { Set-DistributionGroup -Identity $addr -HiddenFromAddressListsEnabled $true -BypassSecurityGroupManagerCheck -ErrorAction Stop }
                }
            }
        }
    }
    if ($D.Hide) { $res += 'verborgen' }
    return ($res -join ', ')
}

function Add-NewGroupToQueue {
    param([hashtable]$Data)
    $parts = @("$($script:NewGroupLabels[$Data.Type]) AANMAKEN: $($Data.Name) <$($Data.Address)>")
    if (@($Data.Owners).Count) { $parts += "eigenaar: $(@($Data.Owners) -join ', ')" }
    if ($Data.Type -eq 'Unified') { $parts += $(if ($Data.Privacy -eq 'Private') { "priv$([char]0x00E9)" } else { 'openbaar' }) }
    if ($Data.Type -eq 'Dynamic') {
        $f = if ($Data.Filter) { $Data.Filter } else {
            $x = @("$($Data.Included)")
            if ($Data.Department) { $x += "afdeling $($Data.Department)" }
            if ($Data.Company) { $x += "bedrijf $($Data.Company)" }
            $x -join ' + '
        }
        $parts += "filter: $f"
    }
    if ($Data.Type -eq 'Distribution' -and $Data.Join -ne 'Closed') { $parts += "lid worden: $($Data.Join)" }
    if ($Data.AllowExternal) { $parts += 'extern mailen toegestaan' }
    if ($Data.Hide) { $parts += 'verborgen in adresboek' }
    return (Add-QueueItem -Kind 'NewGroup' -Target $Data.Name -Soort 'Nieuwe groep' -Text ($parts -join ', ') -Data $Data)
}

function Add-GroupSettingsToQueue {
    # Instellingen van een Microsoft 365-groep (Set-UnifiedGroup). Set bevat
    # parameternaam -> waarde, SobAdd/SobRemove de 'verzenden namens'-wijzigingen.
    param([string]$GroupName, [string]$Group, [hashtable]$Set = @{}, [string[]]$SobAdd = @(), [string[]]$SobRemove = @(), [string]$Text)
    return (Add-QueueItem -Kind 'GroupSettings' -Target $GroupName -Soort 'Groepsinstelling' -Text $Text `
        -Data @{ Group = $Group; Set = $Set; SobAdd = @($SobAdd | Where-Object { $_ }); SobRemove = @($SobRemove | Where-Object { $_ }) })
}

function Add-GroupSubscriberToQueue {
    # Groepsmail en -afspraken in de eigen inbox (abonnee van de M365-groep) aan/uit
    param([string]$GroupName, [string]$Group, [string[]]$Members, [string]$Label, [bool]$Add)
    $text = if ($Add) { "Groepsmail in inbox AANZETTEN voor $Label (groep $Group)" } else { "Groepsmail in inbox UITZETTEN voor $Label (groep $Group)" }
    return (Add-QueueItem -Kind 'GroupSubscriber' -Target $GroupName -Soort 'Groepsabonnement' -Text $text `
        -Data @{ Group = $Group; Members = @($Members); Label = $Label; Add = $Add })
}

function Invoke-GroupSettings {
    # Voert 1 GroupSettings-item uit. Volgorde is belangrijk:
    # - SubscriptionEnabled aan voordat AutoSubscribeNewMembers aan kan,
    #   uit pas nadat AutoSubscribeNewMembers uit staat;
    # - HiddenFromExchangeClientsEnabled verbergt zelf al in het adresboek,
    #   daarom HiddenFromAddressListsEnabled apart en daarna.
    param($D)
    $grp = [string]$D.Group
    $set = if ($D.Set) { $D.Set } else { @{} }
    if ($set.ContainsKey('SubscriptionEnabled') -and $set.SubscriptionEnabled) {
        Invoke-Exo "Set-UnifiedGroup $grp -SubscriptionEnabled:`$true" { Set-UnifiedGroup -Identity $grp -SubscriptionEnabled:$true -ErrorAction Stop }
    }
    $p = @{ Identity = $grp; ErrorAction = 'Stop' }
    $desc = @()
    foreach ($k in @($set.Keys)) {
        if ($k -in 'SubscriptionEnabled', 'HiddenFromAddressListsEnabled') { continue }
        $p[$k] = $set[$k]
        $desc += "-$k $(if ($null -eq $set[$k]) { '$null' } elseif ($set[$k] -is [bool]) { "`$$($set[$k])" } else { "'$($set[$k])'" })"
    }
    $sob = @{}
    if (@($D.SobAdd).Count) { $sob.Add = @($D.SobAdd) }
    if (@($D.SobRemove).Count) { $sob.Remove = @($D.SobRemove) }
    if ($sob.Count) {
        $p.GrantSendOnBehalfTo = $sob
        $sp = @()
        if ($sob.ContainsKey('Add')) { $sp += "Add=$($sob['Add'] -join ',')" }
        if ($sob.ContainsKey('Remove')) { $sp += "Remove=$($sob['Remove'] -join ',')" }
        $desc += "-GrantSendOnBehalfTo @{$($sp -join '; ')}"
    }
    if ($p.Count -gt 2) { Invoke-Exo "Set-UnifiedGroup $grp $($desc -join ' ')" { Set-UnifiedGroup @p } }
    $hideOutlook = ($set.ContainsKey('HiddenFromExchangeClientsEnabled') -and $set.HiddenFromExchangeClientsEnabled)
    if ($set.ContainsKey('HiddenFromAddressListsEnabled') -and -not $hideOutlook) {
        $v = [bool]$set.HiddenFromAddressListsEnabled
        Invoke-Exo "Set-UnifiedGroup $grp -HiddenFromAddressListsEnabled `$$v" { Set-UnifiedGroup -Identity $grp -HiddenFromAddressListsEnabled $v -ErrorAction Stop }
    }
    if ($set.ContainsKey('SubscriptionEnabled') -and -not $set.SubscriptionEnabled) {
        Invoke-Exo "Set-UnifiedGroup $grp -SubscriptionEnabled:`$false" { Set-UnifiedGroup -Identity $grp -SubscriptionEnabled:$false -ErrorAction Stop }
    }
    return 'ingesteld'
}

function Invoke-Queue {
    $todo = @($script:Queue | Where-Object { Test-QueueOpen $_ })
    if ($todo.Count -eq 0) { Show-Msg "De wachtrij is leeg." "Wachtrij" | Out-Null; return }
    $exo = @($todo | Where-Object { $_.Kind -in 'MbxRight', 'FolderPerm', 'Forward', 'MailContact', 'NewGroup', 'GroupMember', 'GroupSettings', 'GroupSubscriber' })
    $gr  = @($todo | Where-Object { $_.Kind -in 'GraphPatch', 'GraphManager' })
    if ($exo.Count -and -not $script:ExoConnected) { Show-Msg "De wachtrij bevat $($exo.Count) Exchange-wijziging(en). Verbind eerst met Exchange Online." "Niet verbonden" "OK" "Warning" | Out-Null; return }
    if ($gr.Count -and -not $script:GraphConnected) { Show-Msg "Verbind eerst met Microsoft Graph." "Niet verbonden" "OK" "Warning" | Out-Null; return }
    $r = Show-Msg "Wachtrij uitvoeren?`n`n$($gr.Count) Entra-wijziging(en)`n$($exo.Count) Exchange-wijziging(en)" "Alles uitvoeren" "YesNo" "Question"
    if ($r -ne 'Yes') { return }

    Write-Log "=== Wachtrij uitvoeren: $($gr.Count) Entra, $($exo.Count) Exchange ===" "info"
    $ok = 0; $fail = 0
    Start-LongOp "Wachtrij uitvoeren..."
    try {
        # 1. Exchange, in volgorde van de wachtrij
        $i = 0
        foreach ($it in $exo) {
            $i++
            if ($script:StopRequested) { Write-Log "Gestopt door gebruiker" "warning"; break }
            Set-Status "Exchange $i / $($exo.Count): $($it.Text)"
            Set-QueueStatus $it 'Bezig...'
            $grant = switch ($it.Kind) {
                'MbxRight'    { [bool]$it.Data.Grant }
                'GroupMember' { [bool]$it.Data.Add }
                'GroupSubscriber' { [bool]$it.Data.Add }
                'FolderPerm'  { $it.Data.Op -ne 'Remove' }
                default       { $true }
            }
            try {
                $res = Invoke-QueueItemExo $it
                Set-QueueStatus $it $(if ($res) { "Gereed ($res)" } else { 'Gereed' })
                $ok++
            }
            catch {
                # Bij een nieuwe groep betekent 'already' (adres al in gebruik) juist een echte fout
                if ($it.Kind -ne 'NewGroup' -and (Test-AlreadyDone "$_" $grant)) { Set-QueueStatus $it 'Gereed (was al zo)'; $ok++ }
                else { Set-QueueStatus $it "Mislukt: $_"; $fail++ }
            }
        }

        # 2. Entra: per gebruiker samenvoegen (latere wijziging wint), dan $batch
        if ($gr.Count -and -not $script:StopRequested) {
            $users = [ordered]@{}
            foreach ($it in $gr) {
                $uid = [string]$it.Data.UserId
                if (-not $users.Contains($uid)) { $users[$uid] = @{ Body = [ordered]@{}; PatchItems = @(); Mgr = $null; MgrItems = @() } }
                $u = $users[$uid]
                if ($it.Kind -eq 'GraphPatch') {
                    foreach ($k in @($it.Data.Body.Keys)) { $u.Body[$k] = $it.Data.Body[$k] }
                    $u.PatchItems += $it
                }
                else { $u.Mgr = $it.Data; $u.MgrItems += $it }
                Set-QueueStatus $it 'Bezig...'
            }
            $reqs = New-Object System.Collections.ArrayList
            $n = 0
            $keys = @{}
            foreach ($uid in @($users.Keys)) {
                $n++
                $u = $users[$uid]
                $keys[$uid] = $n
                if ($u.Body.Count -gt 0) { [void]$reqs.Add(@{ id = "p$n"; method = 'PATCH'; url = "/users/$uid"; body = $u.Body }) }
                if ($u.Mgr) {
                    if ($u.Mgr.Action -eq 'Delete') { [void]$reqs.Add(@{ id = "m$n"; method = 'DELETE'; url = "/users/$uid/manager/`$ref" }) }
                    else { [void]$reqs.Add(@{ id = "m$n"; method = 'PUT'; url = "/users/$uid/manager/`$ref"; body = @{ '@odata.id' = "$($script:GraphRoot)/users/$($u.Mgr.ManagerId)" } }) }
                }
            }
            Write-Log "Entra: $($reqs.Count) request(s) voor $($users.Count) gebruiker(s) via `$batch" "api"
            $res = Invoke-GraphBatch -Requests $reqs -What "Entra"
            foreach ($uid in @($users.Keys)) {
                $u = $users[$uid]
                $k = $keys[$uid]
                if ($u.PatchItems.Count) {
                    $x = $res["p$k"]
                    if (Test-BatchOk $x) {
                        foreach ($it in $u.PatchItems) { Set-QueueStatus $it 'Gereed'; $ok++ }
                        Update-UserGridValues $uid $u.Body
                    }
                    elseif ($null -ne $x) {
                        $e = Get-BatchError $x
                        foreach ($it in $u.PatchItems) { Set-QueueStatus $it "Mislukt: $e"; $fail++ }
                        Write-Log "$($u.PatchItems[0].Target): $e" "error"
                    }
                }
                if ($u.MgrItems.Count) {
                    $x = $res["m$k"]
                    if ((Test-BatchOk $x) -or ($u.Mgr.Action -eq 'Delete' -and $x -and [int]$x.status -eq 404)) {
                        foreach ($it in $u.MgrItems) { Set-QueueStatus $it 'Gereed'; $ok++ }
                    }
                    elseif ($null -ne $x) {
                        $e = Get-BatchError $x
                        foreach ($it in $u.MgrItems) { Set-QueueStatus $it "Mislukt: $e"; $fail++ }
                        Write-Log "$($u.MgrItems[0].Target) (manager): $e" "error"
                    }
                }
            }
        }
    }
    catch {
        Show-Msg "Uitvoeren afgebroken:`n$_" "Fout" "OK" "Error" | Out-Null
    }
    finally {
        # Niet uitgevoerd (Stop of afgebroken): terug naar wachtend
        foreach ($it in $todo) { if ($it.Status -like 'Bezig*') { Set-QueueStatus $it 'Wachtend' } }
        Stop-LongOp
    }

    # Nieuwe groepen/contactpersonen: ledenzoeker moet ze opnieuw ophalen
    if (@($exo | Where-Object { $_.Kind -in 'NewGroup', 'MailContact' -and $_.Status -like 'Gereed*' }).Count) { $script:RcptCache = $null }

    $msg = "Wachtrij: $ok gereed, $fail mislukt"
    Write-Log "=== $msg ===" $(if ($fail) { 'warning' } else { 'success' })
    Set-Status $msg $(if ($fail) { 'error' } else { 'success' })
    Update-QueueSummary
    if ($fail) { Show-Msg "$msg.`nMislukte items blijven in de wachtrij staan (rood) en worden bij 'Alles uitvoeren' opnieuw geprobeerd." "Wachtrij" "OK" "Warning" | Out-Null }

    # Huidige weergave verversen met de werkelijke stand
    $script:FolderPathCache = @{}
    if (@($exo | Where-Object { $_.Kind -eq 'Forward' -and $_.Status -like 'Gereed*' }).Count) { Update-UserExoColumns }
    if ($script:CurrentUserId) {
        $id = $script:CurrentUserId
        $script:CurrentUserId = $null
        Open-User $id
    }
}

#endregion

#region --- Pickers ---

function Search-GraphUsers {
    param([string]$Text, [switch]$All)
    $sel = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,department,jobTitle,companyName'
    $top = if ($All) { 999 } else { 100 }
    $h = @{ ConsistencyLevel = 'eventual' }
    if ([string]::IsNullOrWhiteSpace($Text)) {
        $uri = "$($script:GraphRoot)/users?`$select=$sel&`$top=$top&`$orderby=displayName&`$count=true"
    }
    else {
        $t = $Text.Trim().Replace('"', '')
        $q = [uri]::EscapeDataString("`"displayName:$t`" OR `"userPrincipalName:$t`" OR `"mail:$t`"")
        $uri = "$($script:GraphRoot)/users?`$search=$q&`$select=$sel&`$top=$top&`$orderby=displayName&`$count=true"
    }
    if ($All) { return (Get-GraphPaged -Uri $uri -Headers $h) }
    return @((Invoke-Graph $uri -Headers $h).value)
}

$script:GraphUserSearch = {
    param($Text)
    foreach ($u in (Search-GraphUsers $Text)) {
        [pscustomobject]@{ Label = "$($u.displayName)  <$($u.userPrincipalName)>"; Value = [string]$u.userPrincipalName; Id = [string]$u.id }
    }
}

$script:ExoRecipientSearch = {
    param($Text)
    if ($Text.Trim().Length -lt 2) {
        Show-Msg "Typ minimaal 2 tekens." "Zoeken" | Out-Null
        return
    }
    $r = Invoke-Exo "Get-EXORecipient -Anr '$Text'" { Get-EXORecipient -Anr $Text -ResultSize 50 -ErrorAction Stop }
    foreach ($x in @($r)) {
        [pscustomobject]@{ Label = "$($x.DisplayName)  <$($x.PrimarySmtpAddress)>  [$($x.RecipientTypeDetails)]"; Value = [string]$x.PrimarySmtpAddress; Id = [string]$x.ExternalDirectoryObjectId; Name = [string]$x.DisplayName; RType = [string]$x.RecipientTypeDetails }
    }
}

function Show-Picker {
    # Modale zoekdialoog. Alle handlers vuren tijdens ShowDialog(), dus lokale
    # variabelen zijn hier gewoon bereikbaar.
    param([string]$Title, [string]$Prompt, [scriptblock]$Search)
    $w = ConvertFrom-Xaml ($script:PickerXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Title = $Title
    $w.Owner = $script:Window
    $lbl    = $w.FindName('PK_Prompt')
    $txt    = $w.FindName('PK_Txt')
    $list   = $w.FindName('PK_List')
    $btnS   = $w.FindName('PK_BtnSearch')
    $btnOk  = $w.FindName('PK_BtnOk')
    $lbl.Text = $Prompt

    $doSearch = {
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $list.Items.Clear()
            foreach ($i in @(& $Search $txt.Text)) {
                if ($null -eq $i) { continue }
                $li = New-Object System.Windows.Controls.ListBoxItem
                $li.Content = $i.Label
                $li.Tag = $i
                [void]$list.Items.Add($li)
            }
            if ($list.Items.Count -eq 0) { $list.Items.Add("(geen resultaten)") | Out-Null }
        }
        catch { Show-Msg "Zoeken mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" | Out-Null }
        finally { $w.Cursor = $null }
    }
    $accept = {
        if ($list.SelectedItem -is [System.Windows.Controls.ListBoxItem]) { $w.DialogResult = $true }
    }
    $btnS.Add_Click($doSearch)
    $txt.Add_KeyDown({ if ($_.Key -eq 'Return') { & $doSearch } })
    $list.Add_MouseDoubleClick($accept)
    $btnOk.Add_Click($accept)
    $w.Add_ContentRendered({ [void]$txt.Focus() })

    if ($w.ShowDialog()) { return $list.SelectedItem.Tag }
    return $null
}

function Get-RecipientCache {
    # Alle ontvangers (1x per verbinding) voor de ledenzoeker, zodat lokaal en
    # live gezocht kan worden op elk deel van naam, adres of alias.
    param([switch]$Refresh)
    if ($Refresh -or $null -eq $script:RcptCache) {
        $list = New-Object System.Collections.ArrayList
        foreach ($x in (Get-AllRecipients)) {
            $prim = [string]$x.PrimarySmtpAddress
            if (-not $prim) { continue }
            $type = [string]$x.RecipientTypeDetails
            $info = Get-RecipientTypeInfo $type
            $alias = @(@($x.EmailAddresses) | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^smtp:' } |
                ForEach-Object { $_ -replace '^smtp:', '' } | Where-Object { $_ -ne $prim })
            $name = [string]$x.DisplayName
            [void]$list.Add(@{
                Naam = $name; Adres = $prim; Type = $info.Label; Categorie = $info.Cat; RType = $type
                Id = [string]$x.ExternalDirectoryObjectId; Aliassen = ($alias -join '; ')
                Zoek = "$name $prim $($alias -join ' ') $([string]$x.Alias)".ToLower()
            })
        }
        $script:RcptCache = $list
        Write-Log "Ledenzoeker: $($list.Count) ontvanger(s) opgehaald" "info"
    }
    return ,$script:RcptCache
}

function Show-RecipientPicker {
    # Zoekvenster voor ontvangers: live zoeken op elk deel van naam, adres of
    # alias (meerdere woorden = allemaal), filter op soort en meerdere tegelijk
    # kiezen (Ctrl/Shift-klik). -Allow beperkt tot bepaalde RecipientTypeDetails,
    # -Exclude (adressen, kleine letters) verbergt wat al gekozen of lid is.
    # Geeft objecten (Label, Value, RType, Name, Id) terug; leeg bij annuleren.
    param([string]$Title, [string]$Prompt, $Owner = $script:Window, [string[]]$Allow = @(), [hashtable]$Exclude = @{}, [switch]$Single)
    $w = ConvertFrom-Xaml ($script:RcptPickerXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Title = $Title
    $w.Owner = $Owner
    $c = @{}
    foreach ($n in 'RP_Prompt', 'RP_Txt', 'RP_CmbCat', 'RP_BtnRefresh', 'RP_Grid', 'RP_LblCount', 'RP_BtnOk') { $c[$n] = $w.FindName($n) }
    $c.RP_Prompt.Text = $Prompt
    if ($Single) { $c.RP_Grid.SelectionMode = 'Single' }
    $cols = @('Naam', 'Adres', 'Type', 'Aliassen', 'Categorie', 'RType', 'Id', 'Zoek')
    $state = @{ Table = $null; Result = @() }

    $updateCount = {
        if (-not $state.Table) { return }
        $n = @($c.RP_Grid.SelectedItems).Count
        $c.RP_LblCount.Text = "$($state.Table.DefaultView.Count) van $($state.Table.Rows.Count) zichtbaar$(if ($n) { "   |   $n geselecteerd" })"
        $c.RP_BtnOk.Content = if ($n -gt 1) { "Toevoegen ($n)" } else { "Toevoegen" }
        $c.RP_BtnOk.IsEnabled = ($n -gt 0)
    }
    $filter = {
        if (-not $state.Table) { return }
        $parts = @()
        foreach ($t in @($c.RP_Txt.Text.Trim().ToLower() -split '\s+' | Where-Object { $_ })) { $parts += "Zoek LIKE '*$(ConvertTo-LikeLiteral $t)*'" }
        $cat = [string]$c.RP_CmbCat.SelectedItem.Tag
        if ($cat) { $parts += "($cat)" }
        $state.Table.DefaultView.RowFilter = $parts -join ' AND '
        # Precies 1 treffer: meteen selecteren, dan is Enter genoeg
        if ($state.Table.DefaultView.Count -eq 1) { $c.RP_Grid.SelectedIndex = 0 }
        & $updateCount
    }
    $fill = {
        param([bool]$Refresh)
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $c.RP_LblCount.Text = "Ontvangers ophalen... (eerste keer kan even duren)"; Update-UI
            $rows = foreach ($r in (Get-RecipientCache -Refresh:$Refresh)) {
                if ($Allow.Count -and $Allow -notcontains $r.RType) { continue }
                if ($Exclude.ContainsKey(([string]$r.Adres).ToLower())) { continue }
                $r
            }
            $state.Table = Set-GridRows $c.RP_Grid $cols $rows -PassThru
            $state.Table.DefaultView.Sort = 'Naam'
            & $filter
        }
        catch {
            $c.RP_LblCount.Text = "Ophalen mislukt"
            Show-Msg "Ontvangers ophalen mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
        }
        finally { $w.Cursor = $null }
    }
    $accept = {
        $sel = @(Get-SelectedRows $c.RP_Grid)
        if (-not $sel.Count) { return }
        $state.Result = @($sel | ForEach-Object {
            [pscustomobject]@{ Label = "$($_['Naam'])  <$($_['Adres'])>  [$($_['Type'])]"; Value = [string]$_['Adres']
                               RType = [string]$_['RType']; Name = [string]$_['Naam']; Id = [string]$_['Id'] }
        })
        $w.DialogResult = $true
    }

    $c.RP_Txt.Add_TextChanged($filter)
    $c.RP_Txt.Add_PreviewKeyDown({
        if ($_.Key -eq 'Return') { $_.Handled = $true; & $accept }
        elseif ($_.Key -eq 'Down' -and $state.Table -and $state.Table.DefaultView.Count) {
            $_.Handled = $true
            if ($c.RP_Grid.SelectedIndex -lt 0) { $c.RP_Grid.SelectedIndex = 0 }
            [void]$c.RP_Grid.Focus()
        }
    })
    $c.RP_CmbCat.Add_SelectionChanged($filter)
    $c.RP_Grid.Add_SelectionChanged($updateCount)
    $c.RP_Grid.Add_MouseDoubleClick($accept)
    $c.RP_Grid.Add_PreviewKeyDown({ if ($_.Key -eq 'Return') { $_.Handled = $true; & $accept } })
    $c.RP_BtnOk.Add_Click($accept)
    $c.RP_BtnRefresh.Add_Click({ & $fill $true })
    $w.Add_ContentRendered({ [void]$c.RP_Txt.Focus(); & $fill $false })
    [void]$w.ShowDialog()
    return $state.Result
}

#endregion

#region --- Gebruikerslijst en selectie ---

function Invoke-UserSearch {
    if (-not $script:GraphConnected) { Set-Status "Verbind eerst met Microsoft Graph" "error"; return }
    if (-not (Confirm-LeaveUser)) { return }
    Start-LongOp "Gebruikers laden..."
    try {
        $users = Search-GraphUsers $script:TxtSearch.Text -All
        $rows = foreach ($u in $users) {
            @{
                Naam     = $u.displayName
                UPN      = $u.userPrincipalName
                Afdeling = $u.department
                Functie  = $u.jobTitle
                Bedrijf  = $u.companyName
                Status   = if ($u.accountEnabled) { 'Actief' } else { 'Uitgeschakeld' }
                Type     = $u.userType
                Id       = $u.id
            }
        }
        $script:UserSearchAll = [string]::IsNullOrWhiteSpace($script:TxtSearch.Text)
        $script:MbxSummary = $null
        $script:SuppressSel = $true
        $script:UserTable = Set-GridRows $script:UserGrid $script:UserColumns $rows -PassThru
        $script:SuppressSel = $false
        $script:CurrentUser = $null; $script:CurrentUserId = $null
        Update-UserFilter
        Set-ViewMode 'None'
        Set-Status "$(@($users).Count) gebruiker(s) geladen" "success"
        Update-UserExoColumns
    }
    catch {
        Set-Status "Laden mislukt" "error"
        Show-Msg "Gebruikers laden mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" | Out-Null
    }
    finally { $script:SuppressSel = $false; Stop-LongOp }
}

function Update-UserFilter {
    if (-not $script:UserTable) { return }
    $parts = @()
    $t = $script:TxtFilter.Text.Trim()
    if ($t) {
        $l = ConvertTo-LikeLiteral $t
        $parts += "(Naam LIKE '*$l*' OR UPN LIKE '*$l*' OR Afdeling LIKE '*$l*' OR Functie LIKE '*$l*' OR Bedrijf LIKE '*$l*' OR Doorsturen LIKE '*$l*' OR Mailbox LIKE '*$l*')"
    }
    if ($script:ChkHideGuests.IsChecked)  { $parts += "Type <> 'Guest'" }
    # Gedeelde/resource-mailboxen hebben bewust een uitgeschakeld account: die blijven zichtbaar
    if ($script:ChkOnlyEnabled.IsChecked) { $parts += "(Status = 'Actief' OR Mailbox IN ('Gedeeld', 'Ruimte', 'Apparatuur'))" }
    if ($script:ChkOnlyFwd.IsChecked)     { $parts += "Doorsturen <> ''" }
    switch ([string]$script:CmbMbxType.SelectedItem.Tag) {
        'User'     { $parts += "Mailbox = 'Gebruiker'" }
        'Shared'   { $parts += "Mailbox = 'Gedeeld'" }
        'Resource' { $parts += "Mailbox IN ('Ruimte', 'Apparatuur')" }
        'None'     { $parts += "Mailbox = 'Geen'" }
    }
    $script:UserTable.DefaultView.RowFilter = $parts -join ' AND '
    Update-UserCount
}

function Update-UserCount {
    if (-not $script:UserTable) { $script:LblUserCount.Text = ""; return }
    $sel = $script:UserGrid.SelectedItems.Count
    $txt = "$($script:UserTable.DefaultView.Count) van $($script:UserTable.Rows.Count) gebruikers zichtbaar   |   $sel geselecteerd"
    if ($script:MbxSummary) { $txt += "`n$($script:MbxSummary)" }
    $script:LblUserCount.Text = $txt
}

function Update-UserGridValues {
    # Werkt na een wijziging de zichtbare kolommen bij zonder alles te herladen
    param([string]$Id, $Values)
    if (-not $script:UserTable) { return }
    $map = @{ displayName = 'Naam'; userPrincipalName = 'UPN'; department = 'Afdeling'; jobTitle = 'Functie'; companyName = 'Bedrijf' }
    foreach ($row in $script:UserTable.Select("Id = '$Id'")) {
        foreach ($k in @($Values.Keys)) {
            if ($map.ContainsKey($k)) { $row[$map[$k]] = Format-PatchValue $Values[$k] }
            elseif ($k -eq 'accountEnabled') { $row['Status'] = if ($Values[$k]) { 'Actief' } else { 'Uitgeschakeld' } }
        }
    }
}

function Select-UserRow {
    param([string]$Id)
    $script:SuppressSel = $true
    try {
        $script:UserGrid.SelectedItem = $null
        foreach ($item in $script:UserGrid.Items) {
            if ($item.Row['Id'] -eq $Id) { $script:UserGrid.SelectedItem = $item; break }
        }
    }
    finally { $script:SuppressSel = $false }
}

function Set-ViewMode {
    param([ValidateSet('None', 'Single', 'Bulk')][string]$Mode)
    $script:DetailTabs.Visibility = if ($Mode -eq 'Bulk') { 'Collapsed' } else { 'Visible' }
    $script:BulkTabs.Visibility   = if ($Mode -eq 'Bulk') { 'Visible' } else { 'Collapsed' }
    if ($Mode -ne 'Single') { $script:SyncWarn.Visibility = 'Collapsed'; $script:LblUserQueue.Text = "" }
    if ($Mode -eq 'None') {
        $script:DetailTabs.IsEnabled = $false
        $script:LblUser.Text = "Selecteer links een of meer gebruikers"
        $script:LblUserSub.Text = "Tip: Ctrl/Shift-klik of 'Alles selecteren' voor bulk-aanpassingen"
    }
}

function Confirm-LeaveUser {
    # Bij niet-in-wachtrij-gezette formulierwijzigingen: toevoegen, weggooien
    # of blijven. Geeft $true terug als er verder gegaan mag worden.
    if (-not (Test-EntraDirty)) { return $true }
    $r = Show-Msg "De wijzigingen voor '$($script:CurrentUser.displayName)' staan nog niet in de wachtrij.`n`nJa = toevoegen aan wachtrij`nNee = wijzigingen weggooien`nAnnuleren = terug naar het formulier" "Niet in wachtrij" "YesNoCancel" "Warning"
    if ($r -eq 'Cancel') { return $false }
    if ($r -eq 'Yes') { return (Add-EntraToQueue) }
    return $true
}

function On-UserSelectionChanged {
    Update-UserCount
    if ($script:SuppressSel) { return }
    $items = @($script:UserGrid.SelectedItems)
    if ($items.Count -eq 0) { return }
    $isSameSingle = ($items.Count -eq 1 -and [string]$items[0].Row['Id'] -eq $script:CurrentUserId)
    if (-not $isSameSingle -and -not (Confirm-LeaveUser)) { Select-UserRow $script:CurrentUserId; return }

    if ($items.Count -gt 1) {
        $script:CurrentUser = $null
        $script:CurrentUserId = $null
        $script:SelectedUsers = @($items | ForEach-Object {
            [pscustomobject]@{ Id = [string]$_.Row['Id']; Name = [string]$_.Row['Naam']; Upn = [string]$_.Row['UPN'] }
        })
        Set-ViewMode 'Bulk'
        $n = $script:SelectedUsers.Count
        $script:LblUser.Text = "$n gebruikers geselecteerd (bulk)"
        $names = @($script:SelectedUsers | Select-Object -First 6 | ForEach-Object { $_.Name })
        $script:LblUserSub.Text = ($names -join ', ') + $(if ($n -gt 6) { " en nog $($n - 6)" } else { "" })
        return
    }

    Set-ViewMode 'Single'
    if ($isSameSingle) { Set-EntraForm; return }
    Open-User ([string]$items[0].Row['Id'])
}

function Open-User {
    param([string]$Id)
    $props = @($script:EntraFields | ForEach-Object { $_.Prop }) + @('id', 'mail', 'userType', 'createdDateTime', 'onPremisesSyncEnabled', 'proxyAddresses', 'accountEnabled')
    $sel = ($props | Select-Object -Unique) -join ','
    Enter-Busy "Gebruiker laden..."
    $ok = $false
    try {
        $u = Invoke-Graph "$($script:GraphRoot)/users/$($Id)?`$select=$sel"
        $m = Invoke-Graph "$($script:GraphRoot)/users/$($Id)/manager?`$select=id,displayName,userPrincipalName" -AllowNotFound
        $script:CurrentUser   = $u
        $script:CurrentUserId = $Id
        $script:ActualManager = if ($m -and $m.id) { [pscustomobject]@{ Id = $m.id; Label = "$($m.displayName)  <$($m.userPrincipalName)>" } } else { $null }
        Reset-ExoState
        Set-EntraForm
        $script:DetailTabs.IsEnabled = $true
        Write-Log "Gebruiker geladen: $($u.displayName) ($($u.userPrincipalName))" "info"
        Set-Status "Geladen: $($u.displayName)" "success"
        $ok = $true
    }
    catch {
        Set-Status "Laden mislukt" "error"
        Show-Msg "Gebruiker laden mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" | Out-Null
    }
    finally { Exit-Busy }
    if ($ok) { Invoke-TabActivated }
}

#endregion

#region --- Entra eigenschappen (1 gebruiker) ---

function Get-EntraValue {
    param($User, [string]$Prop)
    $v = $User[$Prop]
    if ($null -eq $v) { return '' }
    if ($Prop -eq 'businessPhones') { return (@($v) -join '; ') }
    return [string]$v
}

function ConvertTo-PatchValue {
    # Zet tekst uit het formulier/CSV om naar de waarde voor de PATCH-body.
    # Geeft een foutmelding terug via -ErrorMessage als de waarde ongeldig is.
    param([string]$Prop, [string]$Text, [ref]$ErrorMessage)
    $ErrorMessage.Value = $null
    switch ($Prop) {
        'businessPhones' { return ,([string[]]@($Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) }
        { $_ -in 'displayName', 'userPrincipalName', 'mailNickname' } {
            if (-not $Text) { $ErrorMessage.Value = "mag niet leeg zijn" }
            return $Text
        }
        'usageLocation' {
            if ($Text -and $Text -notmatch '^[A-Za-z]{2}$') { $ErrorMessage.Value = "moet een landcode van 2 letters zijn (bijv. NL)" }
            if ($Text) { return $Text.ToUpper() }
            return $null
        }
        default { if ($Text) { return $Text }; return $null }
    }
}

function Set-EntraForm {
    # Vult het formulier met de huidige Entra-waarden, met daaroverheen wat
    # al in de wachtrij staat (blauw). Gele velden = nog niet in de wachtrij.
    $u = $script:CurrentUser
    $ov = Get-PendingUserOverlay $u.id
    $script:Loading = $true
    try {
        $script:Orig = @{}
        $script:QueuedActual = @{}
        foreach ($f in $script:EntraFields) {
            $actual = Get-EntraValue $u $f.Prop
            $val = $actual
            if ($ov.Body.Contains($f.Prop)) {
                $val = Format-PatchValue $ov.Body[$f.Prop]
                $script:QueuedActual[$f.Prop] = $actual
            }
            $script:Orig[$f.Prop] = $val
            $script:FieldBoxes[$f.Prop].Text = $val
        }
        $script:OrigEnabled = [bool]$u.accountEnabled
        if ($ov.Body.Contains('accountEnabled')) {
            $script:QueuedActual['accountEnabled'] = "$([bool]$u.accountEnabled)"
            $script:OrigEnabled = [bool]$ov.Body['accountEnabled']
        }
        $script:ChkAccountEnabled.IsChecked = $script:OrigEnabled

        $script:OrigManager = $script:ActualManager
        $script:QueuedManager = $false
        if ($ov.Manager) {
            $script:QueuedManager = $true
            $script:OrigManager = if ($ov.Manager.Action -eq 'Set') { [pscustomobject]@{ Id = $ov.Manager.ManagerId; Label = $ov.Manager.ManagerLabel } } else { $null }
        }
        $script:PendingManager = $script:OrigManager
        $script:ManagerChanged = $false
        $script:TxtManager.Text = if ($script:OrigManager) { $script:OrigManager.Label } else { '(geen)' }

        $script:LblUser.Text = $u.displayName
        $state = if ($u.accountEnabled) { 'actief' } else { 'UITGESCHAKELD' }
        $script:LblUserSub.Text = "$($u.userPrincipalName)   |   $state   |   $($u.userType)"
        $script:SyncWarn.Visibility = if ($u.onPremisesSyncEnabled) { 'Visible' } else { 'Collapsed' }

        $src = if ($u.onPremisesSyncEnabled) { 'On-premises AD (gesynchroniseerd)' } else { 'Cloud' }
        $script:TxtInfo.Text = @(
            "Object ID:      $($u.id)"
            "E-mail:         $($u.mail)"
            "Type:           $($u.userType)"
            "Aangemaakt:     $($u.createdDateTime)"
            "Bron:           $src"
            "Proxyadressen:  $(@($u.proxyAddresses) -join ', ')"
        ) -join "`r`n"
    }
    finally { $script:Loading = $false }
    Update-EntraDirty
    Update-UserQueueNotice
}

function Get-EntraChanges {
    # Lijst met velden die afwijken van wat getoond werd: Prop, Label, Old, New
    $changes = @()
    foreach ($f in $script:EntraFields) {
        $new = $script:FieldBoxes[$f.Prop].Text.Trim()
        $old = $script:Orig[$f.Prop]
        if ($new -cne $old) { $changes += [pscustomobject]@{ Prop = $f.Prop; Label = $f.Label; Old = $old; New = $new } }
    }
    $en = [bool]$script:ChkAccountEnabled.IsChecked
    if ($null -ne $script:CurrentUser -and $en -ne $script:OrigEnabled) {
        $changes += [pscustomobject]@{ Prop = 'accountEnabled'; Label = 'Account ingeschakeld'; Old = "$($script:OrigEnabled)"; New = "$en" }
    }
    return ,$changes
}

function Test-EntraDirty {
    if ($null -eq $script:CurrentUser) { return $false }
    return ((Get-EntraChanges).Count -gt 0 -or $script:ManagerChanged)
}

function Update-EntraDirty {
    if ($script:Loading -or $null -eq $script:CurrentUser) { return }
    $changed = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($c in (Get-EntraChanges)) { [void]$changed.Add($c.Prop) }
    $yellow = New-Brush "#FFF4CE"
    $blue   = New-Brush "#DCEBFF"
    foreach ($f in $script:EntraFields) {
        $box = $script:FieldBoxes[$f.Prop]
        if ($changed.Contains($f.Prop)) { $box.Background = $yellow; $box.ToolTip = "Gewijzigd - nog niet in de wachtrij" }
        elseif ($script:QueuedActual.ContainsKey($f.Prop)) { $box.Background = $blue; $box.ToolTip = "In de wachtrij. Nu in Entra: '$($script:QueuedActual[$f.Prop])'" }
        else { $box.Background = [System.Windows.Media.Brushes]::White; $box.ToolTip = $null }
    }
    $script:ChkAccountEnabled.Foreground = if ($changed.Contains('accountEnabled')) { New-Brush "#B05A00" } elseif ($script:QueuedActual.ContainsKey('accountEnabled')) { New-Brush "#0063B1" } else { [System.Windows.Media.Brushes]::Black }
    $script:TxtManager.Background = if ($script:ManagerChanged) { $yellow } elseif ($script:QueuedManager) { $blue } else { New-Brush "#F7F7F7" }
    $n = $changed.Count + [int]$script:ManagerChanged
    $q = $script:QueuedActual.Count + [int]$script:QueuedManager
    $parts = @()
    if ($n) { $parts += "$n wijziging(en) nog niet in de wachtrij (geel)" }
    if ($q) { $parts += "$q veld(en) al in de wachtrij (blauw)" }
    $script:LblEntraDirty.Text = $parts -join '   |   '
}

function Invoke-PickManager {
    if (-not $script:CurrentUser) { return }
    $sel = Show-Picker -Title "Manager kiezen" -Prompt "Zoek op naam, UPN of e-mail:" -Search $script:GraphUserSearch
    if (-not $sel) { return }
    if ($sel.Id -eq $script:CurrentUserId) {
        Show-Msg "Een gebruiker kan niet zijn eigen manager zijn." "Manager" "OK" "Warning" | Out-Null
        return
    }
    $script:PendingManager = $sel
    $origId = if ($script:OrigManager) { $script:OrigManager.Id } else { $null }
    $script:ManagerChanged = ($sel.Id -ne $origId)
    $script:TxtManager.Text = $sel.Label
    Update-EntraDirty
}

function Invoke-ClearManager {
    if (-not $script:CurrentUser) { return }
    $script:PendingManager = $null
    $script:ManagerChanged = ($null -ne $script:OrigManager)
    $script:TxtManager.Text = '(geen)'
    Update-EntraDirty
}

function Add-EntraToQueue {
    # Zet de formulierwijzigingen van de huidige gebruiker in de wachtrij.
    if (-not $script:CurrentUser) { return $false }
    $changes = Get-EntraChanges
    if ($changes.Count -eq 0 -and -not $script:ManagerChanged) {
        Show-Msg "Geen wijzigingen om toe te voegen." "Niets te doen" | Out-Null
        return $false
    }
    $body = [ordered]@{}
    $old = @{}
    foreach ($c in $changes) {
        if ($c.Prop -eq 'accountEnabled') { $body.accountEnabled = [bool]$script:ChkAccountEnabled.IsChecked; continue }
        $err = $null
        $val = ConvertTo-PatchValue -Prop $c.Prop -Text $c.New -ErrorMessage ([ref]$err)
        if ($err) { Show-Msg "'$($c.Label)' $err." "Ongeldige invoer" "OK" "Warning" | Out-Null; return $false }
        $body[$c.Prop] = $val
        $old[$c.Prop] = $c.Old
    }
    if ($body.Contains('userPrincipalName')) {
        $r = Show-Msg "LET OP: de UPN wijzigt naar '$($body.userPrincipalName)'. De gebruiker moet daarna met de nieuwe naam inloggen.`n`nToevoegen aan de wachtrij?" "UPN wijzigen" "YesNo" "Warning"
        if ($r -ne 'Yes') { return $false }
    }
    $u = $script:CurrentUser
    $name = "$($u.displayName)"
    if ($body.Count -gt 0) {
        [void](Add-QueueItem -Kind 'GraphPatch' -Target $name -Soort 'Entra' -Text (Format-BodyText $body $old) -Data @{ UserId = $u.id; Body = $body })
    }
    if ($script:ManagerChanged) {
        if ($script:PendingManager) {
            [void](Add-QueueItem -Kind 'GraphManager' -Target $name -Soort 'Manager' -Text "Manager instellen op $($script:PendingManager.Label)" `
                -Data @{ UserId = $u.id; Action = 'Set'; ManagerId = $script:PendingManager.Id; ManagerLabel = $script:PendingManager.Label })
        }
        else {
            [void](Add-QueueItem -Kind 'GraphManager' -Target $name -Soort 'Manager' -Text "Manager verwijderen" -Data @{ UserId = $u.id; Action = 'Delete' })
        }
    }
    Write-Log "Entra-wijzigingen voor $name toegevoegd aan wachtrij" "info"
    Update-QueueSummary
    Set-EntraForm
    Set-Status "Toegevoegd aan wachtrij: $name" "success"
    return $true
}

#endregion

#region --- Exchange: gedeelde helpers ---

function Reset-ExoState {
    $script:ExoState = @{ Mailbox = $null; Looked = $false; MbxLoaded = $false; FolderLoaded = $false; FolderPath = $null; FolderType = $null
                          Fa = @(); Sa = @(); Sob = @(); FolderRows = @(); FwdLoaded = $false; Fwd = $null; FwdPick = $null }
    foreach ($g in $script:MbxFullGrid, $script:MbxSendAsGrid, $script:MbxSobGrid, $script:FolderGrid, $script:ScanGrid, $script:FwdRuleGrid) { $g.ItemsSource = $null }
    foreach ($l in $script:LblMbxInfo, $script:LblFolderPath, $script:LblScan, $script:LblFwdInfo, $script:LblFwdState, $script:LblFwdDetail, $script:LblFwdQueue, $script:LblFwdRules) { $l.Text = "" }
    $script:FwdStateBorder.Background = New-Brush "#F0F0F0"
    $script:MbxContent.IsEnabled = $false
    $script:FolderContent.IsEnabled = $false
    $script:FwdContent.IsEnabled = $false
}

function Get-CurrentMailbox {
    # Mailbox van de huidige gebruiker, 1x per gebruiker opgehaald.
    if ($script:ExoState.Looked) { return $script:ExoState.Mailbox }
    $id = $script:CurrentUserId
    $mbx = $null
    try {
        $mbx = Invoke-Exo "Get-EXOMailbox (ExternalDirectoryObjectId $id)" {
            Get-EXOMailbox -ExternalDirectoryObjectId $id -Properties GrantSendOnBehalfTo, RecipientTypeDetails, Alias, Name, DisplayName -ErrorAction Stop
        }
    }
    catch {
        Write-Log "Geen mailbox gevonden voor deze gebruiker" "warning"
    }
    $script:ExoState.Looked = $true
    $script:ExoState.Mailbox = $mbx
    return $mbx
}

function Get-CurrentMailboxIds {
    # Schrijfwijzen waarmee wachtrij-items naar deze mailbox kunnen verwijzen
    $ids = @([string]$script:CurrentUser.userPrincipalName, [string]$script:CurrentUser.mail)
    if ($script:ExoState.Mailbox) { $ids += [string]$script:ExoState.Mailbox.PrimarySmtpAddress }
    return @($ids | Where-Object { $_ })
}

function Invoke-TabActivated {
    if (-not $script:CurrentUser) { return }
    $tab = $script:DetailTabs.SelectedItem
    if ($tab -eq $script:TabMbx)        { if (-not $script:ExoState.MbxLoaded)    { Load-MailboxTab } else { Show-MailboxGrids } }
    elseif ($tab -eq $script:TabFwd)    { if (-not $script:ExoState.FwdLoaded)    { Load-ForwardTab } else { Show-ForwardState } }
    elseif ($tab -eq $script:TabFolder) { if (-not $script:ExoState.FolderLoaded) { Load-FolderTab } else { Show-FolderGrid } }
}

function Invoke-MbxRight {
    # Kent 1 mailboxrecht toe of trekt het in. Gooit een fout bij mislukken.
    param(
        [ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')][string]$Kind,
        [string]$Mailbox, [string]$Who, [bool]$Grant, [bool]$AutoMap = $true
    )
    switch ("$Kind|$Grant") {
        'FullAccess|True' {
            Invoke-Exo "Add-MailboxPermission $Mailbox -User $Who -AccessRights FullAccess -AutoMapping $AutoMap" {
                Add-MailboxPermission -Identity $Mailbox -User $Who -AccessRights FullAccess -InheritanceType All -AutoMapping $AutoMap -Confirm:$false -ErrorAction Stop | Out-Null
            }
        }
        'FullAccess|False' {
            Invoke-Exo "Remove-MailboxPermission $Mailbox -User $Who -AccessRights FullAccess" {
                Remove-MailboxPermission -Identity $Mailbox -User $Who -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
            }
        }
        'SendAs|True' {
            Invoke-Exo "Add-RecipientPermission $Mailbox -Trustee $Who -AccessRights SendAs" {
                Add-RecipientPermission -Identity $Mailbox -Trustee $Who -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            }
        }
        'SendAs|False' {
            Invoke-Exo "Remove-RecipientPermission $Mailbox -Trustee $Who -AccessRights SendAs" {
                Remove-RecipientPermission -Identity $Mailbox -Trustee $Who -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            }
        }
        'SendOnBehalf|True' {
            Invoke-Exo "Set-Mailbox $Mailbox -GrantSendOnBehalfTo @{Add=$Who}" {
                Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{ Add = $Who } -Confirm:$false -ErrorAction Stop
            }
        }
        'SendOnBehalf|False' {
            Invoke-Exo "Set-Mailbox $Mailbox -GrantSendOnBehalfTo @{Remove=$Who}" {
                Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{ Remove = $Who } -Confirm:$false -ErrorAction Stop
            }
        }
    }
}

function Get-FolderPathFor {
    # Volledig mappad (bijv. jan@x.nl:\Agenda). De mapnaam is taalafhankelijk,
    # dus opzoeken via FolderType. Resultaat wordt per mailbox gecachet.
    param([string]$Mailbox, [string]$Type)
    $key = "$Mailbox|$Type".ToLower()
    if ($script:FolderPathCache.ContainsKey($key)) { return $script:FolderPathCache[$key] }
    $stats = Invoke-Exo "Get-EXOMailboxFolderStatistics $Mailbox -FolderScope $Type" {
        Get-EXOMailboxFolderStatistics -Identity $Mailbox -FolderScope $Type -ErrorAction Stop
    }
    $folder = @($stats | Where-Object { $_.FolderType -eq $Type }) | Select-Object -First 1
    if (-not $folder) { $folder = @($stats) | Select-Object -First 1 }
    if (-not $folder) { throw "Map van type $Type niet gevonden in $Mailbox." }
    $path = "$($Mailbox):\" + (([string]$folder.FolderPath).TrimStart('/') -replace '/', '\')
    $script:FolderPathCache[$key] = $path
    return $path
}

function Set-FolderPermission {
    # Toevoegen, of bijwerken als de gebruiker al rechten heeft op de map.
    # Geeft 'toegevoegd' of 'bijgewerkt' terug.
    param([string]$Path, [string]$Who, [string]$Rights, [string]$Flags = 'None')
    $p = @{ Identity = $Path; User = $Who; AccessRights = $Rights; Confirm = $false; ErrorAction = 'Stop' }
    if ($Flags -and $Flags -ne 'None') { $p.SharingPermissionFlags = $Flags }
    if ($Who -match '^(Default|Anonymous)$') {
        Invoke-Exo "Set-MailboxFolderPermission $Path -User $Who -AccessRights $Rights" { Set-MailboxFolderPermission @p | Out-Null }
        return 'bijgewerkt'
    }
    try {
        Invoke-Exo "Add-MailboxFolderPermission $Path -User $Who -AccessRights $Rights $(if ($p.SharingPermissionFlags) { "-SharingPermissionFlags $Flags" })" { Add-MailboxFolderPermission @p | Out-Null }
        return 'toegevoegd'
    }
    catch {
        if ("$_" -notmatch 'existing permission entry|already') { throw }
        Write-Log "Bestaat al, rechten worden bijgewerkt" "info"
        Invoke-Exo "Set-MailboxFolderPermission $Path -User $Who -AccessRights $Rights" { Set-MailboxFolderPermission @p | Out-Null }
        return 'bijgewerkt'
    }
}

function Test-AlreadyDone {
    # Fouten die betekenen dat de gewenste situatie al bestaat
    param([string]$Message, [bool]$Grant)
    if ($Grant) { return ($Message -match 'already|bestaat al|has already') }
    return ($Message -match "isn't present|not present|doesn't exist|does not exist|no existing permission|Could not find|niet gevonden|There is no existing|isn't a member|is not a member|isn't a subscriber|is not a subscriber")
}

#endregion

#region --- Exchange: mailbox rechten (1 gebruiker) ---

function Resolve-RecipientLabel {
    param([string]$Identity)
    try {
        $r = Get-EXORecipient -Identity $Identity -ErrorAction Stop
        return [pscustomobject]@{ Label = "$($r.DisplayName)  <$($r.PrimarySmtpAddress)>"; Key = [string]$r.PrimarySmtpAddress }
    }
    catch { return [pscustomobject]@{ Label = $Identity; Key = $Identity } }
}

function Load-MailboxTab {
    if (-not $script:ExoConnected) { $script:LblMbxInfo.Text = "Niet verbonden met Exchange Online (knop rechtsboven)"; return }
    Enter-Busy "Mailboxrechten laden..."
    try {
        $mbx = Get-CurrentMailbox
        if (-not $mbx) {
            $script:LblMbxInfo.Text = "Deze gebruiker heeft geen Exchange Online mailbox."
            $script:MbxContent.IsEnabled = $false
            Set-Status "Geen mailbox" "error"
            return
        }
        $smtp = [string]$mbx.PrimarySmtpAddress
        $script:LblMbxInfo.Text = "$($mbx.RecipientTypeDetails)   |   $smtp"

        $fa = Invoke-Exo "Get-EXOMailboxPermission $smtp" { Get-EXOMailboxPermission -Identity $smtp -ErrorAction Stop }
        $script:ExoState.Fa = @(foreach ($p in @($fa)) {
            if ($p.IsInherited -or $p.Deny) { continue }
            if ([string]$p.User -like 'NT AUTHORITY\*' -or [string]$p.User -like 'S-1-5-*') { continue }
            if ((@($p.AccessRights) -join ',') -notmatch 'FullAccess') { continue }
            @{ Gebruiker = [string]$p.User; Rechten = (@($p.AccessRights) -join ', '); Key = [string]$p.User }
        })

        $sa = Invoke-Exo "Get-EXORecipientPermission $smtp" { Get-EXORecipientPermission -Identity $smtp -AccessRights SendAs -ErrorAction Stop }
        $script:ExoState.Sa = @(foreach ($p in @($sa)) {
            if ($p.IsInherited -or [string]$p.Trustee -like 'NT AUTHORITY\*') { continue }
            @{ Gebruiker = [string]$p.Trustee; Key = [string]$p.Trustee }
        })

        $script:ExoState.Sob = @(foreach ($v in @($mbx.GrantSendOnBehalfTo)) {
            if (-not $v) { continue }
            $r = Resolve-RecipientLabel ([string]$v)
            @{ Gebruiker = $r.Label; Key = $r.Key }
        })

        $script:MbxContent.IsEnabled = $true
        $script:ExoState.MbxLoaded = $true
        Show-MailboxGrids
        Update-UserQueueNotice
        Set-Status "Mailboxrechten geladen" "success"
    }
    catch {
        Set-Status "Laden mislukt" "error"
        Show-Msg "Mailboxrechten laden mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
    }
    finally { Exit-Busy }
}

function Show-MailboxGrids {
    # Huidige rechten + wat er in de wachtrij staat voor deze mailbox
    if (-not $script:ExoState.MbxLoaded) { return }
    $ids = Get-CurrentMailboxIds
    $cols = @('Gebruiker', 'Rechten', 'Key', 'Wachtrij', 'QKey')
    foreach ($kind in 'FullAccess', 'SendAs', 'SendOnBehalf') {
        $src = switch ($kind) { 'FullAccess' { $script:ExoState.Fa } 'SendAs' { $script:ExoState.Sa } 'SendOnBehalf' { $script:ExoState.Sob } }
        $rows = New-Object System.Collections.ArrayList
        foreach ($r in @($src)) { [void]$rows.Add(@{ Gebruiker = $r.Gebruiker; Rechten = $r.Rechten; Key = $r.Key; Wachtrij = ''; QKey = '' }) }
        $pend = @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'MbxRight' -and $_.Data.Kind -eq $kind -and $ids -contains $_.Data.Mailbox })
        foreach ($p in $pend) {
            $match = @($rows | Where-Object { $_.Key -eq $p.Data.Who -or $_.Gebruiker -eq $p.Data.Who }) | Select-Object -First 1
            if ($p.Data.Grant) {
                if ($match) { $match.Wachtrij = '+ toevoegen (bestaat al)'; $match.QKey = $p.Key }
                else { [void]$rows.Add(@{ Gebruiker = $p.Data.WhoLabel; Rechten = ''; Key = $p.Data.Who; Wachtrij = '+ toevoegen'; QKey = $p.Key }) }
            }
            else {
                if ($match) { $match.Wachtrij = '- verwijderen'; $match.QKey = $p.Key }
                else { [void]$rows.Add(@{ Gebruiker = $p.Data.WhoLabel; Rechten = ''; Key = $p.Data.Who; Wachtrij = '- verwijderen'; QKey = $p.Key }) }
            }
        }
        $grid = switch ($kind) { 'FullAccess' { $script:MbxFullGrid } 'SendAs' { $script:MbxSendAsGrid } 'SendOnBehalf' { $script:MbxSobGrid } }
        Set-GridRows $grid $cols $rows
    }
}

function Invoke-MbxAdd {
    param([string]$Kind)
    $mbx = $script:ExoState.Mailbox
    if (-not $mbx) { return }
    $sel = Show-Picker -Title "$Kind toekennen" -Prompt "Zoek gebruiker of groep (naam, e-mail):" -Search $script:ExoRecipientSearch
    if (-not $sel) { return }
    $smtp = [string]$mbx.PrimarySmtpAddress
    $added = Add-MbxRightToQueue -Target $script:CurrentUser.displayName -Kind $Kind -Mailbox $smtp -Who $sel.Value -WhoLabel $sel.Label -Grant $true -AutoMap ([bool]$script:ChkAutoMap.IsChecked)
    if (-not $added) { Set-Status "Staat al in de wachtrij" }
    Update-QueueSummary
    Show-MailboxGrids
}

function Invoke-MbxRemove {
    # Bestaand recht -> intrekken in de wachtrij. Een wachtrij-regel ('+ toevoegen')
    # selecteren en op Verwijderen klikken haalt die uit de wachtrij.
    param([string]$Kind)
    $mbx = $script:ExoState.Mailbox
    if (-not $mbx) { return }
    $grid = switch ($Kind) { 'FullAccess' { $script:MbxFullGrid } 'SendAs' { $script:MbxSendAsGrid } 'SendOnBehalf' { $script:MbxSobGrid } }
    $rows = Get-SelectedRows $grid
    if ($rows.Count -eq 0) { Show-Msg "Selecteer eerst een of meer regels." "$Kind verwijderen" | Out-Null; return }
    $smtp = [string]$mbx.PrimarySmtpAddress
    foreach ($row in $rows) {
        $q = [string]$row['QKey']
        if ($q -and [string]$row['Wachtrij'] -like '+*') { [void](Remove-QueueItems @($q)); continue }
        if ($q) { continue }
        [void](Add-MbxRightToQueue -Target $script:CurrentUser.displayName -Kind $Kind -Mailbox $smtp -Who ([string]$row['Key']) -WhoLabel ([string]$row['Gebruiker']) -Grant $false)
    }
    Update-QueueSummary
    Show-MailboxGrids
}

function Invoke-MbxReload {
    # Mailbox-object opnieuw ophalen (GrantSendOnBehalfTo kan gewijzigd zijn)
    $script:ExoState.Looked = $false
    $script:ExoState.Mailbox = $null
    $script:ExoState.MbxLoaded = $false
    Load-MailboxTab
}

#endregion

#region --- Exchange: doorsturen ---

function Get-ForwardInfo {
    # Doorstuurinstelling van een mailbox-object. Mode: Off, Internal, External.
    # ForwardingAddress (intern) heeft voorrang op ForwardingSmtpAddress.
    param($Mailbox, [hashtable]$Cache = $null)
    $int  = [string]$Mailbox.ForwardingAddress
    $ext  = ([string]$Mailbox.ForwardingSmtpAddress) -replace '^smtp:', ''
    $keep = [bool]$Mailbox.DeliverToMailboxAndForward
    if ($int) {
        if ($Cache -and $Cache.ContainsKey($int)) { $r = $Cache[$int] }
        else { $r = Resolve-RecipientLabel $int; if ($Cache) { $Cache[$int] = $r } }
        return @{ Mode = 'Internal'; Address = $r.Key; Label = $r.Label; KeepCopy = $keep; Ext = $ext }
    }
    if ($ext) { return @{ Mode = 'External'; Address = $ext; Label = $ext; KeepCopy = $keep; Ext = '' } }
    return @{ Mode = 'Off'; Address = ''; Label = ''; KeepCopy = $keep; Ext = '' }
}

function Format-ForwardText {
    param([string]$Mode, [string]$Label, [bool]$KeepCopy)
    $copy = if ($KeepCopy) { 'kopie bewaren' } else { 'geen kopie in mailbox' }
    switch ($Mode) {
        'Internal' { return "doorsturen naar $Label (intern, $copy)" }
        'External' { return "doorsturen naar $Label (EXTERN, $copy)" }
        default    { return 'niet doorsturen' }
    }
}

function Format-ForwardShort {
    # Korte tekst voor de gebruikerslijst / het overzicht
    param($Info)
    if ($Info.Mode -eq 'Off') { return '' }
    $soort = if ($Info.Mode -eq 'External') { 'Extern' } else { 'Intern' }
    $copy = if ($Info.KeepCopy) { '' } else { ' (geen kopie)' }
    return "$soort`: $($Info.Address)$copy"
}

function Get-PendingForward {
    # Laatste openstaande doorstuur-item voor deze mailbox (latere wint)
    param([string[]]$Ids)
    return (@($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'Forward' -and $Ids -contains $_.Data.Mailbox }) | Select-Object -Last 1)
}

function Add-ForwardQueueItem {
    # Een mailbox heeft maar 1 doorstuurinstelling: eerdere openstaande items
    # voor dezelfde mailbox vervallen.
    param([string]$Target, [string]$Mailbox, [string[]]$Ids, [string]$Mode, [string]$Address = '', [string]$Label = '', [bool]$KeepCopy = $false)
    if (-not $Ids) { $Ids = @($Mailbox) }
    $old = @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'Forward' -and $Ids -contains $_.Data.Mailbox } | ForEach-Object { $_.Key })
    if ($old.Count) { [void](Remove-QueueItems $old) }
    $text = if ($Mode -eq 'Off') { "Doorsturen UITZETTEN op mailbox $Mailbox" } else { "Mailbox $Mailbox`: $(Format-ForwardText $Mode $Label $KeepCopy)" }
    return (Add-QueueItem -Kind 'Forward' -Target $Target -Soort 'Doorsturen' -Text $text `
        -Data @{ Mailbox = $Mailbox; Mode = $Mode; Address = $Address; Label = $Label; KeepCopy = $KeepCopy })
}

function Get-AutoForwardPolicyText {
    # Extern doorsturen werkt alleen als het uitgaande spambeleid het toestaat
    if ($null -eq $script:AutoFwdText) {
        try {
            $pol = Invoke-Exo "Get-HostedOutboundSpamFilterPolicy" { Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop }
            $parts = foreach ($x in @($pol)) {
                $m = switch ([string]$x.AutoForwardingMode) { 'On' { 'toegestaan' } 'Off' { 'geblokkeerd' } default { 'geblokkeerd (Automatic)' } }
                "$($x.Name): $m"
            }
            $script:AutoFwdText = "Extern doorsturen volgens uitgaand spambeleid - " + (@($parts) -join ';  ')
        }
        catch { $script:AutoFwdText = "Uitgaand spambeleid kon niet gelezen worden (extern doorsturen is standaard geblokkeerd)." }
    }
    return $script:AutoFwdText
}

function Get-ForwardingMailboxes {
    # Alle mailboxen met een doorstuuradres (ForwardingAddress of ForwardingSmtpAddress)
    $props = 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward'
    try {
        return @(Invoke-Exo "Get-EXOMailbox -Filter (doorsturen ingesteld)" {
            Get-EXOMailbox -ResultSize Unlimited -Filter "ForwardingAddress -ne `$null -or ForwardingSmtpAddress -ne `$null" -Properties $props -ErrorAction Stop
        })
    }
    catch {
        Write-Log "Filteren in Exchange niet gelukt, alle mailboxen worden opgehaald" "warning"
        return @(Invoke-Exo "Get-EXOMailbox -ResultSize Unlimited" { Get-EXOMailbox -ResultSize Unlimited -Properties $props -ErrorAction Stop } |
            Where-Object { $_.ForwardingAddress -or $_.ForwardingSmtpAddress })
    }
}

function Get-MailboxTypeLabel {
    param([string]$Type)
    if ($script:MbxTypeLabels.ContainsKey($Type)) { return $script:MbxTypeLabels[$Type] }
    return $Type
}

function Update-UserExoColumns {
    # Vult de kolommen 'Mailbox' en 'Doorsturen' in de gebruikerslijst met 1
    # Exchange-aanroep. Bij 'alle gebruikers' worden mailboxen die Graph niet
    # teruggaf (bijv. gedeelde mailboxen) alsnog aan de lijst toegevoegd.
    if (-not $script:UserTable -or -not $script:ExoConnected) { return }
    Set-Status "Mailboxen ophalen (type + doorsturen)..."
    try {
        $props = 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward'
        $all = @(Invoke-Exo "Get-EXOMailbox -ResultSize Unlimited (type + doorsturen)" {
            Get-EXOMailbox -ResultSize Unlimited -Properties $props -ErrorAction Stop
        })
        $cache = @{}
        $byId = @{}
        foreach ($m in $all) {
            $id = [string]$m.ExternalDirectoryObjectId
            if ($id) { $byId[$id] = $m }
        }
        $seen = @{}
        $nFwd = 0
        foreach ($row in $script:UserTable.Rows) {
            $id = [string]$row['Id']
            $seen[$id] = $true
            if ($byId.ContainsKey($id)) {
                $m = $byId[$id]
                $row['Mailbox'] = Get-MailboxTypeLabel ([string]$m.RecipientTypeDetails)
                $row['Doorsturen'] = Format-ForwardShort (Get-ForwardInfo $m -Cache $cache)
            }
            else { $row['Mailbox'] = 'Geen'; $row['Doorsturen'] = '' }
            if ($row['Doorsturen']) { $nFwd++ }
        }
        $added = 0
        if ($script:UserSearchAll) {
            foreach ($id in @($byId.Keys)) {
                if ($seen.ContainsKey($id)) { continue }
                $m = $byId[$id]
                $dr = $script:UserTable.NewRow()
                $dr['Naam'] = [string]$m.DisplayName
                $dr['UPN'] = [string]$m.UserPrincipalName
                $dr['Status'] = 'Onbekend'
                $dr['Type'] = 'Member'
                $dr['Id'] = $id
                $dr['Mailbox'] = Get-MailboxTypeLabel ([string]$m.RecipientTypeDetails)
                $dr['Doorsturen'] = Format-ForwardShort (Get-ForwardInfo $m -Cache $cache)
                if ($dr['Doorsturen']) { $nFwd++ }
                [void]$script:UserTable.Rows.Add($dr)
                $added++
            }
            if ($added) { Write-Log "Mailboxen: $added mailbox(en) uit Exchange toegevoegd die Graph niet teruggaf" "warning" }
        }
        $n = @{}
        foreach ($row in $script:UserTable.Rows) { $n[[string]$row['Mailbox']]++ }
        $script:MbxSummary = "Mailboxen: $([int]$n['Gebruiker']) gebruiker, $([int]$n['Gedeeld']) gedeeld, $([int]$n['Ruimte'] + [int]$n['Apparatuur']) ruimte/apparatuur, $([int]$n['Geen']) zonder mailbox   |   $nFwd met doorsturen"
        Update-UserFilter
        Write-Log "Mailboxen: $($all.Count) in Exchange, $([int]$n['Gedeeld']) gedeeld, $nFwd met doorsturen" "info"
        Set-Status "$($all.Count) mailbox(en) geladen, waarvan $([int]$n['Gedeeld']) gedeeld" "success"
    }
    catch { Set-Status "Mailboxen ophalen mislukt" "error"; Write-Log "Mailboxen ophalen mislukt: $_" "error" }
}

function Load-ForwardTab {
    if (-not $script:ExoConnected) { $script:LblFwdInfo.Text = "Niet verbonden met Exchange Online (knop rechtsboven)"; return }
    Enter-Busy "Doorstuurinstelling laden..."
    try {
        $mbx = Get-CurrentMailbox
        if (-not $mbx) {
            $script:LblFwdInfo.Text = "Deze gebruiker heeft geen Exchange Online mailbox."
            $script:FwdContent.IsEnabled = $false
            Set-Status "Geen mailbox" "error"
            return
        }
        $smtp = [string]$mbx.PrimarySmtpAddress
        $m = Invoke-Exo "Get-EXOMailbox $smtp (doorsturen)" {
            Get-EXOMailbox -Identity $smtp -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward -ErrorAction Stop
        }
        $script:ExoState.Fwd = Get-ForwardInfo $m
        $script:ExoState.FwdLoaded = $true
        $script:LblFwdInfo.Text = "$($mbx.RecipientTypeDetails)   |   $smtp"
        $script:LblFwdPolicy.Text = Get-AutoForwardPolicyText
        $script:FwdContent.IsEnabled = $true
        Show-ForwardState
        Set-ForwardForm
        Set-Status "Doorstuurinstelling geladen" "success"
    }
    catch {
        Set-Status "Laden mislukt" "error"
        Show-Msg "Doorstuurinstelling laden mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
    }
    finally { Exit-Busy }
}

function Show-ForwardState {
    # Huidige instelling (groot) + wat er in de wachtrij staat
    if (-not $script:ExoState.FwdLoaded) { return }
    $f = $script:ExoState.Fwd
    switch ($f.Mode) {
        'Internal' { $script:LblFwdState.Text = "AAN - doorsturen naar $($f.Label)";              $bg = '#FFF4CE' }
        'External' { $script:LblFwdState.Text = "AAN - doorsturen naar EXTERN adres $($f.Address)"; $bg = '#FDE7E9' }
        default    { $script:LblFwdState.Text = "UIT - mail wordt niet doorgestuurd";              $bg = '#DFF6DD' }
    }
    $script:FwdStateBorder.Background = New-Brush $bg
    $lines = @()
    if ($f.Mode -ne 'Off') { $lines += "Kopie bewaren in deze mailbox: $(if ($f.KeepCopy) { 'ja' } else { 'NEE - mail staat alleen bij het doorstuuradres' })" }
    if ($f.Ext) { $lines += "Ook ingesteld: extern adres $($f.Ext) (genegeerd zolang het interne adres is ingesteld)" }
    $script:LblFwdDetail.Text = $lines -join "`r`n"
    $p = Get-PendingForward (Get-CurrentMailboxIds)
    $script:LblFwdQueue.Text = if ($p) { "In de wachtrij: $($p.Text)" } else { "" }
}

function Set-ForwardForm {
    # Formulier vullen met de wachtrij-instelling, of anders de huidige
    $f = $script:ExoState.Fwd
    if (-not $f) { return }
    $mode = $f.Mode; $addr = $f.Address; $label = $f.Label; $keep = $f.KeepCopy
    $p = Get-PendingForward (Get-CurrentMailboxIds)
    if ($p) { $mode = $p.Data.Mode; $addr = $p.Data.Address; $label = $p.Data.Label; $keep = [bool]$p.Data.KeepCopy }
    $script:ExoState.FwdPick = if ($mode -eq 'Internal') { [pscustomobject]@{ Value = $addr; Label = $label } } else { $null }
    $script:TxtFwdInt.Text = if ($mode -eq 'Internal') { $label } else { '' }
    $script:TxtFwdExt.Text = if ($mode -eq 'External') { $addr } else { '' }
    $script:RbFwdOff.IsChecked = ($mode -eq 'Off')
    $script:RbFwdInt.IsChecked = ($mode -eq 'Internal')
    $script:RbFwdExt.IsChecked = ($mode -eq 'External')
    $script:ChkFwdKeep.IsChecked = if ($mode -eq 'Off') { $true } else { $keep }
    Update-ForwardOptions
}

function Update-ForwardOptions {
    # Kopie bewaren is alleen zinvol als er doorgestuurd wordt
    $script:ChkFwdKeep.IsEnabled = -not $script:RbFwdOff.IsChecked
}

function Invoke-PickForward {
    if (-not $script:ExoState.Mailbox) { return }
    $sel = Show-Picker -Title "Doorsturen naar" -Prompt "Zoek gebruiker, gedeelde mailbox, groep of contactpersoon (naam, e-mail):" -Search $script:ExoRecipientSearch
    if (-not $sel) { return }
    if ((Get-CurrentMailboxIds) -contains $sel.Value) {
        Show-Msg "Een mailbox kan niet naar zichzelf doorsturen." "Doorsturen" "OK" "Warning" | Out-Null
        return
    }
    $script:ExoState.FwdPick = $sel
    $script:TxtFwdInt.Text = $sel.Label
    $script:RbFwdInt.IsChecked = $true
    Update-ForwardOptions
}

function Add-ForwardToQueue {
    $mbx = $script:ExoState.Mailbox
    if (-not $mbx -or -not $script:ExoState.FwdLoaded) { return }
    $keep = [bool]$script:ChkFwdKeep.IsChecked
    if ($script:RbFwdInt.IsChecked) {
        $pick = $script:ExoState.FwdPick
        if (-not $pick) { Show-Msg "Kies eerst een interne ontvanger (knop Kiezen...)." "Doorsturen" "OK" "Warning" | Out-Null; return }
        $mode = 'Internal'; $addr = [string]$pick.Value; $label = [string]$pick.Label
    }
    elseif ($script:RbFwdExt.IsChecked) {
        $addr = $script:TxtFwdExt.Text.Trim()
        if ($addr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { Show-Msg "Vul een geldig extern e-mailadres in." "Doorsturen" "OK" "Warning" | Out-Null; return }
        $mode = 'External'; $label = $addr
    }
    else { $mode = 'Off'; $addr = ''; $label = ''; $keep = $false }

    $ids = Get-CurrentMailboxIds
    $f = $script:ExoState.Fwd
    $same = ($mode -eq $f.Mode -and $addr -eq $f.Address -and ($mode -eq 'Off' -or $keep -eq $f.KeepCopy) -and -not $f.Ext)
    if ($same) {
        $p = Get-PendingForward $ids
        if ($p) {
            [void](Remove-QueueItems @($p.Key))
            Update-QueueSummary
            Show-ForwardState
            Set-Status "Gelijk aan huidige instelling - wachtrij-item verwijderd"
        }
        else { Show-Msg "Dit is al de huidige instelling." "Geen wijziging" | Out-Null }
        return
    }
    $smtp = [string]$mbx.PrimarySmtpAddress
    [void](Add-ForwardQueueItem -Target $script:CurrentUser.displayName -Mailbox $smtp -Ids $ids -Mode $mode -Address $addr -Label $label -KeepCopy $keep)
    Write-Log "Doorsturen voor $smtp toegevoegd aan wachtrij: $(Format-ForwardText $mode $label $keep)" "info"
    Update-QueueSummary
    Show-ForwardState
    Set-Status "Toegevoegd aan wachtrij: doorsturen $smtp" "success"
}

function Invoke-ForwardRules {
    # Inbox-regels van de gebruiker die doorsturen of omleiden (alleen lezen)
    $mbx = $script:ExoState.Mailbox
    if (-not $mbx) { return }
    $smtp = [string]$mbx.PrimarySmtpAddress
    Enter-Busy "Inbox-regels laden..."
    try {
        $rules = @(Invoke-Exo "Get-InboxRule -Mailbox $smtp" { Get-InboxRule -Mailbox $smtp -WarningAction SilentlyContinue -ErrorAction Stop })
        $acts = @(@{ P = 'ForwardTo'; L = 'Doorsturen' }, @{ P = 'ForwardAsAttachmentTo'; L = 'Doorsturen als bijlage' }, @{ P = 'RedirectTo'; L = 'Omleiden' })
        $rows = New-Object System.Collections.ArrayList
        foreach ($r in $rules) {
            foreach ($a in $acts) {
                $to = @(@($r.($a.P)) | Where-Object { $_ } | ForEach-Object { ([string]$_) -replace '^"?([^"]*)"?\s*\[(?:SMTP|smtp|EX):([^\]]+)\]$', '$1 <$2>' })
                if ($to.Count) { [void]$rows.Add(@{ Regel = $r.Name; Actief = $(if ($r.Enabled) { 'Ja' } else { 'Nee' }); Actie = $a.L; Naar = ($to -join '; ') }) }
            }
        }
        Set-GridRows $script:FwdRuleGrid @('Regel', 'Actief', 'Actie', 'Naar') $rows
        $script:LblFwdRules.Text = "$($rows.Count) doorstuuractie(s) gevonden in $($rules.Count) inbox-regel(s)"
        Set-Status "Inbox-regels geladen" "success"
    }
    catch {
        Set-Status "Inbox-regels laden mislukt" "error"
        Show-Msg "Inbox-regels laden mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
    }
    finally { Exit-Busy }
}

function Show-ForwardOverview {
    # Alle mailboxen in de tenant met doorsturen. Handlers vuren tijdens
    # ShowDialog(), dus lokale variabelen zijn bereikbaar (zoals bij CSV).
    if (-not (Test-ExoReady)) { return }
    $w = ConvertFrom-Xaml ($script:FwdOverviewXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Owner = $script:Window
    $c = @{}
    foreach ($n in 'FO_BtnLoad', 'FO_ChkExtern', 'FO_LblStatus', 'FO_LblPolicy', 'FO_Grid', 'FO_BtnOff', 'FO_BtnExport') { $c[$n] = $w.FindName($n) }
    $state = @{ Table = $null }
    $c.FO_BtnOff.IsEnabled = $false
    $c.FO_BtnExport.IsEnabled = $false

    $applyFilter = {
        if ($state.Table) { $state.Table.DefaultView.RowFilter = if ($c.FO_ChkExtern.IsChecked) { "Soort = 'Extern'" } else { "" } }
    }
    $load = {
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $c.FO_LblStatus.Text = "Mailboxen ophalen..."; Update-UI
            $c.FO_LblPolicy.Text = Get-AutoForwardPolicyText
            $cache = @{}
            $rows = foreach ($m in (Get-ForwardingMailboxes)) {
                $info = Get-ForwardInfo $m -Cache $cache
                $smtp = [string]$m.PrimarySmtpAddress
                $p = Get-PendingForward @($smtp)
                @{
                    Mailbox = $smtp; Naam = $m.DisplayName; Type = [string]$m.RecipientTypeDetails; Naar = $info.Label
                    Soort = $(if ($info.Mode -eq 'External') { 'Extern' } else { 'Intern' })
                    Kopie = $(if ($info.KeepCopy) { 'Ja' } else { 'Nee' })
                    Wachtrij = $(if ($p) { $(if ($p.Data.Mode -eq 'Off') { '- uitzetten' } else { '-> wijzigen' }) } else { '' })
                }
            }
            $state.Table = Set-GridRows $c.FO_Grid @('Mailbox', 'Naam', 'Type', 'Naar', 'Soort', 'Kopie', 'Wachtrij') $rows -PassThru
            & $applyFilter
            $nExt = @($state.Table.Select("Soort = 'Extern'")).Count
            $c.FO_LblStatus.Text = "$($state.Table.Rows.Count) mailbox(en) met doorsturen, waarvan $nExt extern"
            $c.FO_BtnExport.IsEnabled = ($state.Table.Rows.Count -gt 0)
        }
        catch {
            $c.FO_LblStatus.Text = "Ophalen mislukt"
            Show-Msg "Ophalen mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
        }
        finally { $w.Cursor = $null }
    }

    $c.FO_BtnLoad.Add_Click($load)
    $c.FO_ChkExtern.Add_Click($applyFilter)
    $c.FO_Grid.Add_SelectionChanged({ $c.FO_BtnOff.IsEnabled = ($c.FO_Grid.SelectedItems.Count -gt 0) })
    $c.FO_BtnOff.Add_Click({
        $rows = Get-SelectedRows $c.FO_Grid
        if ($rows.Count -eq 0) { return }
        $n = 0
        foreach ($row in $rows) {
            if (Add-ForwardQueueItem -Target ([string]$row['Naam']) -Mailbox ([string]$row['Mailbox']) -Mode 'Off') { $n++ }
            $row['Wachtrij'] = '- uitzetten'
        }
        Update-QueueSummary
        if ($script:CurrentUser) { Show-ForwardState; Set-ForwardForm }
        Write-Log "Doorsturen-overzicht: $n mailbox(en) in de wachtrij om doorsturen uit te zetten" "info"
        $c.FO_LblStatus.Text = "$n mailbox(en) in de wachtrij gezet. Sluit dit venster en klik onderin op 'Alles uitvoeren'."
    })
    $c.FO_BtnExport.Add_Click({
        if (-not $state.Table) { return }
        $d = New-Object Microsoft.Win32.SaveFileDialog
        $d.Filter = "CSV (*.csv)|*.csv"
        $d.FileName = "Doorsturen-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
        if (-not $d.ShowDialog($w)) { return }
        try {
            $state.Table.DefaultView | ForEach-Object {
                $r = $_.Row
                [pscustomobject]@{ Mailbox = $r['Mailbox']; Naam = $r['Naam']; Type = $r['Type']; DoorsturenNaar = $r['Naar']; Soort = $r['Soort']; KopieBewaren = $r['Kopie'] }
            } | Export-Csv -Path $d.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
            $c.FO_LblStatus.Text = "Geexporteerd naar $($d.FileName)"
        }
        catch { Show-Msg "Exporteren mislukt:`n$_" "Fout" "OK" "Error" | Out-Null }
    })
    $w.Add_ContentRendered({ & $load })
    [void]$w.ShowDialog()
}

#endregion

#region --- Ontvangers-overzicht (alles met een mailadres) ---

function Get-RecipientTypeInfo {
    # Nederlandse omschrijving + categorie (voor het filter) per RecipientTypeDetails
    param([string]$Type)
    if ($script:RcptTypes.ContainsKey($Type)) { return $script:RcptTypes[$Type] }
    return @{ Label = $Type; Cat = 'Overig' }
}

function Get-AllRecipients {
    # Alle ontvangers in Exchange Online. Lukt het niet met extra eigenschappen
    # (oudere module), dan alleen de standaardset.
    try {
        return @(Invoke-Exo "Get-EXORecipient -ResultSize Unlimited" {
            Get-EXORecipient -ResultSize Unlimited -Properties HiddenFromAddressListsEnabled, ExternalEmailAddress, ManagedBy, IsDirSynced -ErrorAction Stop
        })
    }
    catch {
        Write-Log "Ontvangers ophalen met extra eigenschappen mislukt, opnieuw met standaardset" "warning"
        return @(Invoke-Exo "Get-EXORecipient -ResultSize Unlimited (standaard)" { Get-EXORecipient -ResultSize Unlimited -ErrorAction Stop })
    }
}

function Get-RecipientMembers {
    # Leden van een groep, afhankelijk van het groepstype
    param([string]$Identity, [string]$Type)
    switch ($Type) {
        'GroupMailbox' {
            return @(Invoke-Exo "Get-UnifiedGroupLinks $Identity -LinkType Members" { Get-UnifiedGroupLinks -Identity $Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop })
        }
        'DynamicDistributionGroup' {
            return @(Invoke-Exo "Get-DynamicDistributionGroupMember $Identity" { Get-DynamicDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop })
        }
        default {
            return @(Invoke-Exo "Get-DistributionGroupMember $Identity" { Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop })
        }
    }
}

function Get-AcceptedDomainList {
    # Eigen domeinen van de tenant (1x per verbinding): een onbekend adres op
    # een eigen domein is een tikfout, geen extern contact.
    if ($null -eq $script:AcceptedDomains) {
        try { $script:AcceptedDomains = @(Invoke-Exo "Get-AcceptedDomain" { Get-AcceptedDomain -ErrorAction Stop } | ForEach-Object { ([string]$_.DomainName).ToLower() }) }
        catch { $script:AcceptedDomains = @() }
    }
    return $script:AcceptedDomains
}

function ConvertFrom-AddressText {
    # Herkent adressen in geplakte tekst: 1 per regel of gescheiden door ;
    # Ook "Naam <adres>" (Outlook) en "Naam<tab>adres" (Excel) met naam.
    param([string]$Text)
    $rx = '[^\s<>;,"''()\[\]]+@[^\s<>;,"''()\[\]]+\.[A-Za-z]{2,}'
    $out = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($chunk in ($Text -split '[\r\n;]+')) {
        $ms = [regex]::Matches($chunk, $rx)
        if ($ms.Count -eq 0) { continue }
        $name = ''
        if ($ms.Count -eq 1) { $name = ($chunk.Replace($ms[0].Value, '') -replace '[<>"\t]', ' ' -replace '\s+', ' ').Trim([char[]]' ,') }
        foreach ($m in $ms) {
            $key = $m.Value.ToLower()
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            [void]$out.Add([pscustomobject]@{ Address = $m.Value; Name = $name })
        }
    }
    return ,$out
}

function Show-AddressDialog {
    # Plakvenster voor e-mailadressen. Geeft @{ Items; Hide } terug of $null.
    param([string]$Title, [string]$Prompt, $Owner)
    $w = ConvertFrom-Xaml ($script:AddressXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Title = $Title
    $w.Owner = $Owner
    $txt = $w.FindName('AD_Txt')
    $chk = $w.FindName('AD_ChkHide')
    $cnt = $w.FindName('AD_LblCount')
    $w.FindName('AD_Prompt').Text = $Prompt
    $txt.Add_TextChanged({ $cnt.Text = "$((ConvertFrom-AddressText $txt.Text).Count) adres(sen) herkend" })
    $w.FindName('AD_BtnOk').Add_Click({ if ((ConvertFrom-AddressText $txt.Text).Count) { $w.DialogResult = $true } })
    $w.Add_ContentRendered({ [void]$txt.Focus() })
    if (-not $w.ShowDialog()) { return $null }
    return @{ Items = (ConvertFrom-AddressText $txt.Text); Hide = [bool]$chk.IsChecked }
}

function Add-AddressesToQueue {
    # Zet adressen in de wachtrij. Met $Group: als lid toevoegen (bestaande
    # ontvanger, of eerst een nieuwe contactpersoon voor een extern adres).
    # Zonder $Group: alleen contactpersonen aanmaken. Geeft @{ Added; Lines }.
    param([object[]]$Items, [bool]$Hide, $Group = $null, [hashtable]$Existing = @{})
    Update-ExoToken
    $acc = Get-AcceptedDomainList
    $lines = New-Object System.Collections.ArrayList
    $n = 0
    foreach ($it in $Items) {
        $addr = [string]$it.Address
        if ($Existing.ContainsKey($addr.ToLower())) { [void]$lines.Add("$addr - is al lid"); continue }
        $r = $null
        try { $r = Get-EXORecipient -Identity $addr -ErrorAction Stop } catch { }
        if ($r) {
            $smtp = [string]$r.PrimarySmtpAddress
            if (-not $Group) { [void]$lines.Add("$addr - bestaat al als $((Get-RecipientTypeInfo ([string]$r.RecipientTypeDetails)).Label) ($($r.DisplayName))"); continue }
            if ($Existing.ContainsKey($smtp.ToLower())) { [void]$lines.Add("$addr - is al lid (als $smtp)"); continue }
            if (Add-GroupMemberToQueue -GroupName $Group.Name -Group $Group.Address -GroupType $Group.Type -Member $smtp -MemberLabel "$($r.DisplayName) <$smtp>" -Add $true) {
                [void]$lines.Add("$addr - toevoegen ($($r.DisplayName))"); $n++
            }
            else { [void]$lines.Add("$addr - staat al in de wachtrij") }
            continue
        }
        $domain = ($addr -split '@')[-1].ToLower()
        if ($acc -contains $domain) { [void]$lines.Add("$addr - NIET GEVONDEN (eigen domein, dus geen contactpersoon aangemaakt)"); continue }
        if ($Group -and $Group.Type -eq 'GroupMailbox') { [void]$lines.Add("$addr - extern adres kan geen lid worden van een Microsoft 365-groep (alleen als gast)"); continue }
        $name = if ($it.Name) { [string]$it.Name } else { $addr }
        $pending = @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'MailContact' -and $_.Data.Address -eq $addr })
        $newC = if ($pending.Count) { $false } else { Add-MailContactToQueue -Name $name -Address $addr -Hide $Hide }
        if ($Group) {
            [void](Add-GroupMemberToQueue -GroupName $Group.Name -Group $Group.Address -GroupType $Group.Type -Member $addr -MemberLabel "$name <$addr> (nieuw contact)" -Add $true -NewContact $true)
            [void]$lines.Add("$addr - $(if ($newC) { 'nieuwe contactpersoon + ' })toevoegen")
            $n++
        }
        elseif ($newC) { [void]$lines.Add("$addr - nieuwe contactpersoon '$name'"); $n++ }
        else { [void]$lines.Add("$addr - contactpersoon staat al in de wachtrij") }
    }
    Update-QueueSummary
    return @{ Added = $n; Lines = $lines }
}

function Show-AddResult {
    param($Result, [string]$Title, $Owner)
    $l = @($Result.Lines)
    $txt = ($l | Select-Object -First 30) -join "`n"
    if ($l.Count -gt 30) { $txt += "`n... en nog $($l.Count - 30)" }
    foreach ($x in $l) { Write-Log "$Title`: $x" "info" }
    [void][System.Windows.MessageBox]::Show($Owner, "$($Result.Added) wijziging(en) in de wachtrij gezet:`n`n$txt", $Title, 'OK', 'Information')
}

function ConvertTo-GroupAlias {
    # "Afdeling Financien & Control" -> "afdeling.financien.control" (accenten worden weggehaald)
    param([string]$Text)
    $n = $Text.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''
    $n = (($n -replace '\s+', '.') -replace '[^A-Za-z0-9._-]', '') -replace '\.{2,}', '.'
    $n = $n.Trim('.').ToLower()
    if ($n.Length -gt 64) { $n = $n.Substring(0, 64).TrimEnd('.') }
    return $n
}

function Show-NewGroupDialog {
    # Nieuwe distributielijst, beveiligingsgroep (mail), Microsoft 365-groep,
    # dynamische distributielijst of ruimtelijst. Zet de groep en daarna de
    # leden in de wachtrij. Geeft het aantal wachtrij-items terug (0 = geannuleerd).
    # Handlers vuren tijdens ShowDialog(), dus lokale variabelen zijn bereikbaar.
    param($Owner = $script:Window)
    if (-not (Test-ExoReady)) { return 0 }
    $doms = @()
    try {
        $doms = @(Invoke-Exo "Get-AcceptedDomain" { Get-AcceptedDomain -ErrorAction Stop } |
            Sort-Object @{ Expression = { -not $_.Default } }, DomainName | ForEach-Object { ([string]$_.DomainName).ToLower() })
    }
    catch { }
    if (-not $doms.Count) { Show-Msg "Kan de domeinen van de tenant niet ophalen (Get-AcceptedDomain)." "Nieuwe groep" "OK" "Error" | Out-Null; return 0 }

    $w = ConvertFrom-Xaml ($script:NewGroupXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Owner = $Owner
    $c = @{}
    foreach ($n in 'NG_RbDL', 'NG_RbSec', 'NG_RbM365', 'NG_RbDyn', 'NG_RbRoom', 'NG_LblTypeInfo', 'NG_TxtName', 'NG_TxtAlias', 'NG_CmbDomain', 'NG_TxtDesc',
                   'NG_ChkExternal', 'NG_ChkHide', 'NG_PnlJoin', 'NG_CmbJoin', 'NG_CmbDepart', 'NG_PnlM365', 'NG_CmbPrivacy', 'NG_ChkSubscribe',
                   'NG_PnlDyn', 'NG_CmbIncluded', 'NG_TxtDept', 'NG_TxtCompany', 'NG_TxtFilter',
                   'NG_GrpOwners', 'NG_LblOwners', 'NG_LstOwners', 'NG_BtnOwnerAdd', 'NG_BtnOwnerRemove',
                   'NG_GrpMembers', 'NG_LblMembers', 'NG_LstMembers', 'NG_BtnMemberAdd', 'NG_BtnMemberPaste', 'NG_BtnMemberRemove',
                   'NG_LblAddr', 'NG_BtnOk') { $c[$n] = $w.FindName($n) }
    foreach ($d in $doms) { [void]$c.NG_CmbDomain.Items.Add($d) }
    $c.NG_CmbDomain.SelectedIndex = 0
    $state = @{
        AliasTouched = $false; Setting = $false; Hide = $false; Result = 0
        Owners = New-Object System.Collections.ArrayList; Members = New-Object System.Collections.ArrayList; Pasted = New-Object System.Collections.ArrayList
    }
    $ownerTypes = @('UserMailbox', 'MailUser')

    $getType = {
        if ($c.NG_RbSec.IsChecked)  { return 'Security' }
        if ($c.NG_RbM365.IsChecked) { return 'Unified' }
        if ($c.NG_RbDyn.IsChecked)  { return 'Dynamic' }
        if ($c.NG_RbRoom.IsChecked) { return 'RoomList' }
        return 'Distribution'
    }
    $memberOk = {
        # Mag dit type ontvanger lid worden van het gekozen groepstype?
        param([string]$GroupType, [string]$RType)
        switch ($GroupType) {
            'Unified'  { return ($RType -in 'UserMailbox', 'MailUser', 'GuestMailUser') }
            'RoomList' { return ($RType -eq 'RoomMailbox') }
            default    { return $true }
        }
    }
    $render = {
        $t = & $getType
        $c.NG_LstOwners.Items.Clear()
        foreach ($o in $state.Owners) {
            $li = New-Object System.Windows.Controls.ListBoxItem; $li.Content = $o.Label; $li.Tag = $o
            [void]$c.NG_LstOwners.Items.Add($li)
        }
        $c.NG_LstMembers.Items.Clear()
        foreach ($m in $state.Members) {
            $li = New-Object System.Windows.Controls.ListBoxItem; $li.Content = $m.Label; $li.Tag = $m
            if (-not (& $memberOk $t $m.RType)) { $li.Foreground = New-Brush '#C42B1C'; $li.ToolTip = 'Kan geen lid worden van dit soort groep' }
            [void]$c.NG_LstMembers.Items.Add($li)
        }
        foreach ($p in $state.Pasted) {
            $li = New-Object System.Windows.Controls.ListBoxItem
            $li.Content = "$(if ($p.Name) { "$($p.Name) <$($p.Address)>" } else { $p.Address })   (geplakt, wordt gecontroleerd bij 'Naar wachtrij')"
            $li.Tag = $p; $li.Foreground = New-Brush '#555555'
            [void]$c.NG_LstMembers.Items.Add($li)
        }
        $c.NG_GrpOwners.Header = "Eigenaren ($($state.Owners.Count))"
        $c.NG_GrpMembers.Header = "Leden ($($state.Members.Count + $state.Pasted.Count))"
    }
    $updateAddr = {
        $a = $c.NG_TxtAlias.Text.Trim()
        $c.NG_LblAddr.Text = if ($a) { "Adres: $a@$($c.NG_CmbDomain.SelectedItem)" } else { "" }
    }
    $updateType = {
        $t = & $getType
        $vis = { param($El, [bool]$On) $El.Visibility = $(if ($On) { 'Visible' } else { 'Collapsed' }) }
        & $vis $c.NG_PnlJoin ($t -in 'Distribution', 'Security', 'RoomList')
        & $vis $c.NG_PnlM365 ($t -eq 'Unified')
        & $vis $c.NG_PnlDyn ($t -eq 'Dynamic')
        & $vis $c.NG_GrpMembers ($t -ne 'Dynamic')
        # Lid worden/opzeggen alleen vrij te kiezen bij een gewone distributielijst
        $free = ($t -eq 'Distribution')
        $c.NG_CmbJoin.IsEnabled = $free
        $c.NG_CmbDepart.IsEnabled = $free
        if (-not $free) { $c.NG_CmbJoin.SelectedIndex = 0; $c.NG_CmbDepart.SelectedIndex = 1 }
        $c.NG_BtnMemberPaste.IsEnabled = ($t -in 'Distribution', 'Security', 'Unified')
        $c.NG_LblTypeInfo.Text = switch ($t) {
            'Distribution' { "Leden: gebruikers, gedeelde mailboxen, contactpersonen en andere groepen. Externe adressen (via 'Adressen plakken') krijgen eerst automatisch een contactpersoon." }
            'Security'     { "Lid worden en opzeggen staan vast op 'Gesloten' (vereist bij beveiligingsgroepen). Te gebruiken voor mail en voor rechten (mailbox, agenda, SharePoint)." }
            'Unified'      { "Minimaal 1 eigenaar verplicht. Leden alleen gebruikers en gasten, geen contactpersonen. Teams koppelen kan daarna vanuit Teams." }
            'Dynamic'      { "Geen vaste leden: Exchange bepaalt bij elke mail wie er onder het filter valt. Maximaal 1 eigenaar." }
            'RoomList'     { "Leden zijn alleen ruimte-mailboxen. Verschijnt in Outlook bij de ruimtezoeker / Ruimtes toevoegen." }
        }
        $c.NG_LblOwners.Text = if ($t -eq 'Unified') { 'Verplicht. De eerste is de hoofd-eigenaar.' }
                               elseif ($t -eq 'Dynamic') { 'Optioneel, maximaal 1.' }
                               else { 'Optioneel. Zonder eigenaar wordt het account dat de wachtrij uitvoert eigenaar.' }
        $c.NG_LblMembers.Text = if ($t -eq 'RoomList') { 'Alleen ruimte-mailboxen.' } else { 'Rood = past niet bij dit soort groep.' }
        & $render
    }

    $addOwner = {
        $t = & $getType
        if ($t -eq 'Dynamic' -and $state.Owners.Count -ge 1) { Show-Msg "Een dynamische distributielijst kan maar 1 eigenaar hebben. Verwijder eerst de huidige eigenaar." "Eigenaar" "OK" "Warning" | Out-Null; return }
        $ex = @{}
        foreach ($o in $state.Owners) { $ex[([string]$o.Value).ToLower()] = $true }
        $sel = @(Show-RecipientPicker -Title "Eigenaar toevoegen" -Owner $w -Allow $ownerTypes -Exclude $ex -Single:($t -eq 'Dynamic') `
            -Prompt "Kies de gebruiker(s) die de groep gaan beheren. Alleen gebruikers (mailbox of e-mailgebruiker) worden getoond.")
        foreach ($s in $sel) { if ($s) { [void]$state.Owners.Add($s) } }
        & $render
    }
    $addMember = {
        $t = & $getType
        $allow = switch ($t) {
            'Unified'  { @('UserMailbox', 'MailUser', 'GuestMailUser') }
            'RoomList' { @('RoomMailbox') }
            default    { @() }
        }
        $ex = @{}
        foreach ($m in $state.Members) { $ex[([string]$m.Value).ToLower()] = $true }
        foreach ($p in $state.Pasted) { $ex[([string]$p.Address).ToLower()] = $true }
        $what = switch ($t) {
            'Unified'  { 'Alleen gebruikers en gasten worden getoond (een Microsoft 365-groep kan geen contactpersonen of groepen bevatten).' }
            'RoomList' { 'Alleen ruimte-mailboxen worden getoond.' }
            default    { 'Gebruikers, gedeelde mailboxen, groepen en contactpersonen. Een nieuw extern adres? Gebruik "Adressen plakken...".' }
        }
        $sel = @(Show-RecipientPicker -Title "Leden toevoegen" -Owner $w -Allow $allow -Exclude $ex -Prompt "Kies een of meer leden. $what")
        foreach ($s in $sel) { if ($s) { [void]$state.Members.Add($s) } }
        & $render
    }
    $pasteMembers = {
        $t = & $getType
        $extra = if ($t -eq 'Unified') { "Bij een Microsoft 365-groep worden alleen bestaande gebruikers en gasten toegevoegd." }
                 else { "Externe adressen die nog niet bestaan krijgen eerst automatisch een contactpersoon (Naam <adres> of Naam[tab]adres geeft meteen de weergavenaam)." }
        $res = Show-AddressDialog -Title "Leden plakken" -Owner $w -Prompt "Plak of typ e-mailadressen (1 per regel of gescheiden door ;), ook direct uit Outlook of Excel. Ze worden gecontroleerd bij 'Naar wachtrij'. $extra"
        if (-not $res) { return }
        $state.Hide = $res.Hide
        $have = @{}
        foreach ($x in $state.Members) { $have[([string]$x.Value).ToLower()] = $true }
        foreach ($x in $state.Pasted) { $have[([string]$x.Address).ToLower()] = $true }
        foreach ($i in $res.Items) { if (-not $have.ContainsKey(([string]$i.Address).ToLower())) { [void]$state.Pasted.Add($i); $have[([string]$i.Address).ToLower()] = $true } }
        & $render
    }
    $removeSel = {
        param($List)
        foreach ($li in @($List.SelectedItems)) {
            [void]$state.Owners.Remove($li.Tag); [void]$state.Members.Remove($li.Tag); [void]$state.Pasted.Remove($li.Tag)
        }
        & $render
    }

    $submit = {
        $t = & $getType
        $name = $c.NG_TxtName.Text.Trim()
        $alias = $c.NG_TxtAlias.Text.Trim()
        $dom = [string]$c.NG_CmbDomain.SelectedItem
        $errs = @()
        if (-not $name) { $errs += "Vul een weergavenaam in." }
        elseif ($name.Length -gt 256) { $errs += "De weergavenaam is te lang (max. 256 tekens)." }
        if (-not $alias) { $errs += "Vul het e-mailadres (alias) in." }
        elseif ($alias -notmatch '^[A-Za-z0-9_-]+(\.[A-Za-z0-9_-]+)*$') { $errs += "Alias mag alleen letters, cijfers, punt, - en _ bevatten (geen spaties, niet beginnen of eindigen met een punt)." }
        elseif ($alias.Length -gt 64) { $errs += "Alias is te lang (max. 64 tekens)." }
        if ($t -eq 'Unified' -and $state.Owners.Count -eq 0) { $errs += "Een Microsoft 365-groep heeft minimaal 1 eigenaar nodig." }
        if ($t -eq 'Dynamic' -and $c.NG_TxtFilter.Text.Trim().Length -gt 2048) { $errs += "Het eigen filter is te lang (max. 2048 tekens)." }
        if ($t -ne 'Dynamic') {
            $bad = @($state.Members | Where-Object { -not (& $memberOk $t $_.RType) })
            if ($bad.Count) { $errs += "Deze leden passen niet bij dit soort groep (rood), verwijder ze eerst:`n  " + (($bad | ForEach-Object { $_.Label }) -join "`n  ") }
        }
        if ($errs.Count) { Show-Msg ($errs -join "`n`n") "Nieuwe groep" "OK" "Warning" | Out-Null; return }

        $addr = "$alias@$dom"
        if (@($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'NewGroup' -and $_.Data.Address -eq $addr }).Count) {
            Show-Msg "Er staat al een nieuwe groep met adres $addr in de wachtrij." "Nieuwe groep" "OK" "Warning" | Out-Null; return
        }
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        $taken = $null
        try {
            Update-ExoToken
            foreach ($id in @($addr, $alias)) {
                if ($taken) { break }
                try { $taken = Get-EXORecipient -Identity $id -ErrorAction Stop | Select-Object -First 1 } catch { }
            }
        }
        finally { $w.Cursor = $null }
        if ($taken) { Show-Msg "Adres of alias is al in gebruik door:`n$($taken.DisplayName) <$($taken.PrimarySmtpAddress)> [$($taken.RecipientTypeDetails)]`n`nKies een ander adres." "Nieuwe groep" "OK" "Warning" | Out-Null; return }
        if ($t -notin 'Unified', 'Dynamic' -and $state.Owners.Count -eq 0) {
            $r = Show-Msg "Er is geen eigenaar gekozen. Het account dat de wachtrij uitvoert wordt dan eigenaar (bij GDAP: een account van de partner-tenant).`n`nToch doorgaan?" "Geen eigenaar" "YesNo" "Question"
            if ($r -ne 'Yes') { return }
        }

        $data = @{
            Type = $t; Name = $name; Alias = $alias; Address = $addr; Description = $c.NG_TxtDesc.Text.Trim()
            Owners = @($state.Owners | ForEach-Object { [string]$_.Value })
            AllowExternal = [bool]$c.NG_ChkExternal.IsChecked; Hide = [bool]$c.NG_ChkHide.IsChecked
            Join = [string]$c.NG_CmbJoin.SelectedItem.Tag; Depart = [string]$c.NG_CmbDepart.SelectedItem.Tag
            Privacy = [string]$c.NG_CmbPrivacy.SelectedItem.Tag; AutoSubscribe = [bool]$c.NG_ChkSubscribe.IsChecked
            Included = [string]$c.NG_CmbIncluded.SelectedItem.Tag; Department = $c.NG_TxtDept.Text.Trim()
            Company = $c.NG_TxtCompany.Text.Trim(); Filter = $c.NG_TxtFilter.Text.Trim()
            Created = $false; OwnersDone = $false
        }
        if (-not (Add-NewGroupToQueue $data)) { return }
        $n = 1
        if ($t -ne 'Dynamic') {
            # Leden na de groep in de wachtrij: ze worden pas toegevoegd als de groep bestaat
            $gt = $script:NewGroupRType[$t]
            $existing = @{}
            # Eigenaar van een M365-groep is automatisch lid
            if ($t -eq 'Unified' -and $data.Owners.Count) { $existing[$data.Owners[0].ToLower()] = $true }
            foreach ($m in $state.Members) {
                $existing[([string]$m.Value).ToLower()] = $true
                if ($t -eq 'Unified' -and $m.Value -eq $data.Owners[0]) { continue }
                if (Add-GroupMemberToQueue -GroupName $name -Group $addr -GroupType $gt -Member $m.Value -MemberLabel $m.Label -Add $true) { $n++ }
            }
            if ($state.Pasted.Count) {
                $w.Cursor = [System.Windows.Input.Cursors]::Wait
                try { $out = Add-AddressesToQueue -Items @($state.Pasted) -Hide $state.Hide -Group @{ Name = $name; Address = $addr; Type = $gt } -Existing $existing }
                finally { $w.Cursor = $null }
                $n += $out.Added
                Show-AddResult $out "Leden voor $name" $w
            }
        }
        Update-QueueSummary
        Write-Log "Nieuwe groep in wachtrij: $($script:NewGroupLabels[$t]) $name <$addr> ($n item(s))" "info"
        $state.Result = $n
        $w.DialogResult = $true
    }

    foreach ($rb in 'NG_RbDL', 'NG_RbSec', 'NG_RbM365', 'NG_RbDyn', 'NG_RbRoom') { $c[$rb].Add_Checked($updateType) }
    $c.NG_TxtName.Add_TextChanged({
        if (-not $state.AliasTouched) { $state.Setting = $true; $c.NG_TxtAlias.Text = ConvertTo-GroupAlias $c.NG_TxtName.Text; $state.Setting = $false }
        & $updateAddr
    })
    $c.NG_TxtAlias.Add_TextChanged({
        # Zelf aangepaste alias niet meer overschrijven; leegmaken = weer automatisch
        if (-not $state.Setting) { $state.AliasTouched = [bool]$c.NG_TxtAlias.Text.Trim() }
        & $updateAddr
    })
    $c.NG_CmbDomain.Add_SelectionChanged($updateAddr)
    $c.NG_BtnOwnerAdd.Add_Click($addOwner)
    $c.NG_BtnOwnerRemove.Add_Click({ & $removeSel $c.NG_LstOwners })
    $c.NG_BtnMemberAdd.Add_Click($addMember)
    $c.NG_BtnMemberPaste.Add_Click($pasteMembers)
    $c.NG_BtnMemberRemove.Add_Click({ & $removeSel $c.NG_LstMembers })
    $c.NG_BtnOk.Add_Click($submit)
    $w.Add_ContentRendered({ [void]$c.NG_TxtName.Focus() })
    & $updateType
    [void]$w.ShowDialog()
    return $state.Result
}

function Get-GroupSettingsOverlay {
    # Openstaande instellingen-wijzigingen voor 1 M365-groep (latere wint)
    param([string]$Group)
    $set = @{}
    foreach ($it in $script:Queue) {
        if ((Test-QueueOpen $it) -and $it.Kind -eq 'GroupSettings' -and $it.Data.Group -eq $Group) {
            foreach ($k in @($it.Data.Set.Keys)) { $set[$k] = $it.Data.Set[$k] }
        }
    }
    return $set
}

function Show-GroupSettingsDialog {
    # Instellingen van 1 Microsoft 365-groep: groepsmail in de inbox van leden
    # (abonneren), extern mailen, privacy, verbergen, beschrijving/MailTip en
    # verzenden als/namens. Alles via de wachtrij. Geeft het aantal nieuwe
    # wachtrij-items terug. Handlers vuren tijdens ShowDialog(), dus lokale
    # variabelen zijn bereikbaar.
    param([hashtable]$Group, $Owner = $script:Window)
    $ident = if ($Group.Id) { $Group.Id } else { $Group.Address }
    $addr = [string]$Group.Address
    $gname = [string]$Group.Name
    $Owner.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $g = @(Invoke-Exo "Get-UnifiedGroup $addr -IncludeAllProperties" { Get-UnifiedGroup -Identity $ident -IncludeAllProperties -ErrorAction Stop }) | Select-Object -First 1
        $members = @(Invoke-Exo "Get-UnifiedGroupLinks $addr -LinkType Members" { Get-UnifiedGroupLinks -Identity $ident -LinkType Members -ResultSize Unlimited -ErrorAction Stop })
        $subs = @(Invoke-Exo "Get-UnifiedGroupLinks $addr -LinkType Subscribers" { Get-UnifiedGroupLinks -Identity $ident -LinkType Subscribers -ResultSize Unlimited -ErrorAction Stop })
        $sendAs = @(try { Invoke-Exo "Get-RecipientPermission $addr -AccessRights SendAs" { Get-RecipientPermission -Identity $addr -AccessRights SendAs -ResultSize Unlimited -ErrorAction Stop } }
                    catch { Write-Log "Verzenden als ophalen mislukt: $_" "warning" })
    }
    catch {
        Show-Msg "Groep ophalen mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
        return 0
    }
    finally { $Owner.Cursor = $null }
    if (-not $g) { Show-Msg "Groep $addr niet gevonden." "Fout" "OK" "Error" | Out-Null; return 0 }

    # Leden die nu (of via de wachtrij) geen groepsmail in hun inbox krijgen
    $subKeys = @{}
    foreach ($s in $subs) { foreach ($k in @([string]$s.PrimarySmtpAddress, [string]$s.ExternalDirectoryObjectId)) { if ($k) { $subKeys[$k.ToLower()] = $true } } }
    $subNow = 0
    foreach ($m in $members) {
        $sm = ([string]$m.PrimarySmtpAddress).ToLower(); $oid = ([string]$m.ExternalDirectoryObjectId).ToLower()
        if (($sm -and $subKeys[$sm]) -or ($oid -and $subKeys[$oid])) { $subNow++ }
    }
    foreach ($q in @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'GroupSubscriber' -and $_.Data.Group -eq $addr -and $_.Data.Add })) {
        foreach ($m in @($q.Data.Members)) { $subKeys[([string]$m).ToLower()] = $true }
    }
    $notSub = @(foreach ($m in $members) {
        $sm = ([string]$m.PrimarySmtpAddress).ToLower(); $oid = ([string]$m.ExternalDirectoryObjectId).ToLower()
        if (($sm -and $subKeys[$sm]) -or ($oid -and $subKeys[$oid])) { continue }
        if ($m.PrimarySmtpAddress) { [string]$m.PrimarySmtpAddress } else { [string]$m.ExternalDirectoryObjectId }
    })

    $w = ConvertFrom-Xaml ($script:GroupSettingsXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Owner = $Owner
    $c = @{}
    foreach ($n in 'GS_LblTitle', 'GS_LblInfo', 'GS_LblSubs', 'GS_ChkSubExisting', 'GS_CmbPrivacy', 'GS_TxtDesc', 'GS_TxtMailTip',
                   'GS_DelGrid', 'GS_BtnAddSendAs', 'GS_BtnAddSob', 'GS_BtnDelRemove', 'GS_LblChanges', 'GS_BtnOk') { $c[$n] = $w.FindName($n) }
    foreach ($s in $script:GroupBoolSettings) { $c[$s.Box] = $w.FindName($s.Box) }
    $state = @{ Orig = @{}; Avail = @{}; Rows = (New-Object System.Collections.ArrayList); Seq = 0; Undo = @(); Result = 0; Setting = $true }
    $pend = Get-GroupSettingsOverlay $addr
    $readVal = {
        param([string]$Name)
        $pp = $g.PSObject.Properties[$Name]
        if ($pp) { return $pp.Value }
    }

    # Aan/uit-instellingen: Exchange-waarde, met wat al in de wachtrij staat erover
    foreach ($s in $script:GroupBoolSettings) {
        $box = $c[$s.Box]
        $v = & $readVal $s.Get
        if ($v -is [string]) { $v = ($v -eq 'True') }
        if ($pend.ContainsKey($s.Set)) { $v = $pend[$s.Set] }
        if ($null -eq $v) { $box.IsEnabled = $false; $box.ToolTip = "Niet beschikbaar (eigenschap $($s.Get) niet gevonden)"; continue }
        $v = [bool]$v
        $state.Orig[$s.Set] = $v
        $state.Avail[$s.Box] = $true
        $box.IsChecked = $(if ($s.Invert) { -not $v } else { $v })
    }
    $acc = [string](& $readVal 'AccessType')
    if ($pend.ContainsKey('AccessType')) { $acc = [string]$pend.AccessType }
    $state.Orig.AccessType = $acc
    $item = @($c.GS_CmbPrivacy.Items | Where-Object { [string]$_.Tag -eq $acc }) | Select-Object -First 1
    if (-not $item) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = "$acc (hier niet te wijzigen)"; $item.Tag = $acc
        [void]$c.GS_CmbPrivacy.Items.Add($item)
        $c.GS_CmbPrivacy.IsEnabled = $false
    }
    $c.GS_CmbPrivacy.SelectedItem = $item
    foreach ($t in @(@{ Prop = 'Notes'; Box = 'GS_TxtDesc' }, @{ Prop = 'MailTip'; Box = 'GS_TxtMailTip' })) {
        # MailTip komt als HTML terug
        $v = [System.Net.WebUtility]::HtmlDecode(([string](& $readVal $t.Prop)) -replace '<[^>]+>', '').Trim()
        if ($pend.ContainsKey($t.Prop)) { $v = [string]$pend[$t.Prop] }
        $state.Orig[$t.Prop] = $v
        $c[$t.Box].Text = $v
    }

    # Verzenden als / namens: huidige stand + wat al in de wachtrij staat
    $addRow = {
        param([string]$Naam, [string]$Kind, [string]$Key, [string]$Wachtrij = '', [string]$QKey = '', [bool]$New = $false)
        $state.Seq++
        [void]$state.Rows.Add(@{ Nr = [string]$state.Seq; Naam = $Naam; Kind = $Kind; Key = $Key; Wachtrij = $Wachtrij; QKey = $QKey; New = $New; Remove = $false
                                 Recht = $(if ($Kind -eq 'SendAs') { 'Verzenden als' } else { 'Verzenden namens' }) })
    }
    $findRow = {
        param([string]$Kind, [string]$Key)
        @($state.Rows | Where-Object { $_.Kind -eq $Kind -and ([string]$_.Key).ToLower() -eq $Key.ToLower() }) | Select-Object -First 1
    }
    foreach ($p in $sendAs) {
        if (-not $p -or $p.Deny -or [string]$p.Trustee -like 'NT AUTHORITY*') { continue }
        & $addRow ([string]$p.Trustee) 'SendAs' ([string]$p.Trustee)
    }
    foreach ($s in @(& $readVal 'GrantSendOnBehalfTo')) { if ($s) { & $addRow ([string]$s) 'SendOnBehalf' ([string]$s) } }
    foreach ($q in @($script:Queue | Where-Object { Test-QueueOpen $_ })) {
        if ($q.Kind -eq 'MbxRight' -and $q.Data.Kind -eq 'SendAs' -and $q.Data.Mailbox -eq $addr) {
            $r = & $findRow 'SendAs' ([string]$q.Data.Who)
            if ($q.Data.Grant) { if (-not $r) { & $addRow ([string]$q.Data.WhoLabel) 'SendAs' ([string]$q.Data.Who) '+ toevoegen' $q.Key } }
            elseif ($r) { $r.Wachtrij = '- verwijderen'; $r.QKey = $q.Key }
        }
        elseif ($q.Kind -eq 'GroupSettings' -and $q.Data.Group -eq $addr) {
            foreach ($x in @($q.Data.SobAdd)) { if (-not (& $findRow 'SendOnBehalf' $x)) { & $addRow $x 'SendOnBehalf' $x '+ toevoegen' $q.Key } }
            foreach ($x in @($q.Data.SobRemove)) { $r = & $findRow 'SendOnBehalf' $x; if ($r) { $r.Wachtrij = '- verwijderen'; $r.QKey = $q.Key } }
        }
    }

    $c.GS_LblTitle.Text = "$gname  <$addr>"
    $own = (@($g.ManagedBy) | Where-Object { $_ } | ForEach-Object { [string]$_ }) -join ', '
    $c.GS_LblInfo.Text = "Microsoft 365-groep   |   $($members.Count) lid/leden   |   eigenaar(s): $(if ($own) { $own } else { '-' })" +
        $(if ($pend.Count) { "`nLet op: er staan al instellingen voor deze groep in de wachtrij; die worden hier al getoond." })
    $c.GS_LblSubs.Text = "Deze twee instellingen gelden alleen voor leden die hierna worden toegevoegd (gasten worden altijd geabonneerd). " +
        "Nu ontvangen $subNow van $($members.Count) leden de groepsmail in hun eigen inbox; leden kunnen dit zelf aan/uit zetten in Outlook."
    $c.GS_ChkSubExisting.Content = if ($notSub.Count) { "Ook de $($notSub.Count) huidige lid/leden die nu geen groepsmail in hun inbox krijgen abonneren" }
                                   else { "Alle huidige leden krijgen de groepsmail al in hun inbox" }

    $collect = {
        # Gewijzigde instellingen ten opzichte van wat getoond werd
        $set = @{}; $txt = @()
        $hideOutlook = [bool]$c.GS_ChkHideOutlook.IsChecked
        foreach ($s in $script:GroupBoolSettings) {
            if (-not $state.Orig.ContainsKey($s.Set)) { continue }
            # Verbergen in Outlook verbergt zelf al in het adresboek
            if ($s.Set -eq 'HiddenFromAddressListsEnabled' -and $hideOutlook) { continue }
            $chk = [bool]$c[$s.Box].IsChecked
            $v = if ($s.Invert) { -not $chk } else { $chk }
            if ($v -ne [bool]$state.Orig[$s.Set]) { $set[$s.Set] = $v; $txt += "$($s.Label) $(if ($chk) { 'AAN' } else { 'UIT' })" }
        }
        $acc = [string]$c.GS_CmbPrivacy.SelectedItem.Tag
        if ($c.GS_CmbPrivacy.IsEnabled -and $acc -and $acc -ne $state.Orig.AccessType) {
            $set.AccessType = $acc
            $txt += "Privacy: $(if ($acc -eq 'Private') { "priv$([char]0x00E9)" } else { 'openbaar' })"
        }
        foreach ($t in @(@{ Prop = 'Notes'; Box = 'GS_TxtDesc'; Label = 'Beschrijving' }, @{ Prop = 'MailTip'; Box = 'GS_TxtMailTip'; Label = 'MailTip' })) {
            $v = $c[$t.Box].Text.Trim()
            if ($v -cne [string]$state.Orig[$t.Prop]) {
                $set[$t.Prop] = $(if ($v) { $v } else { $null })
                $txt += $(if ($v) { "$($t.Label): '$v'" } else { "$($t.Label) LEEGMAKEN" })
            }
        }
        return @{ Set = $set; Text = @($txt) }
    }
    $updateChanges = {
        if ($state.Setting) { return }
        $x = & $collect
        $n = $x.Text.Count + @($state.Rows | Where-Object { $_.New -or $_.Remove }).Count + @($state.Undo).Count
        if ($c.GS_ChkSubExisting.IsChecked) { $n++ }
        $c.GS_LblChanges.Text = if ($n) { "$n wijziging(en) - gaan pas door na 'Alles uitvoeren'" } else { 'Geen wijzigingen' }
        $c.GS_BtnOk.IsEnabled = ($n -gt 0)
    }
    $updateDeps = {
        # Afhankelijkheden tussen de opties
        $state.Setting = $true
        $subOk = (-not $state.Avail['GS_ChkSubEnabled']) -or [bool]$c.GS_ChkSubEnabled.IsChecked
        if ($state.Avail['GS_ChkAutoSub']) {
            if (-not $subOk) { $c.GS_ChkAutoSub.IsChecked = $false }
            $c.GS_ChkAutoSub.IsEnabled = $subOk
        }
        if ($state.Avail['GS_ChkCalSub']) { $c.GS_ChkCalSub.IsEnabled = ($subOk -and -not $c.GS_ChkAutoSub.IsChecked) }
        if (-not $subOk) { $c.GS_ChkSubExisting.IsChecked = $false }
        $c.GS_ChkSubExisting.IsEnabled = ($subOk -and $notSub.Count -gt 0)
        if ($state.Avail['GS_ChkHideGal']) {
            $ho = [bool]$c.GS_ChkHideOutlook.IsChecked
            if ($ho) { $c.GS_ChkHideGal.IsChecked = $true }
            $c.GS_ChkHideGal.IsEnabled = (-not $ho)
        }
        $state.Setting = $false
        & $updateChanges
    }
    $renderDel = {
        Set-GridRows $c.GS_DelGrid @('Nr', 'Naam', 'Recht', 'Wachtrij') @($state.Rows)
        & $updateChanges
    }
    $addDel = {
        param([string]$Kind)
        $ex = @{}
        $ex[$addr.ToLower()] = $true
        foreach ($r in $state.Rows) { if ($r.Kind -eq $Kind) { $ex[([string]$r.Key).ToLower()] = $true } }
        $what = if ($Kind -eq 'SendAs') { 'Verzenden als' } else { 'Verzenden namens' }
        $sel = @(Show-RecipientPicker -Title "$what $gname" -Owner $w -Allow @('UserMailbox', 'SharedMailbox', 'MailUser', 'MailUniversalSecurityGroup') -Exclude $ex `
            -Prompt "Wie mag '$what' de groep $gname mailen? Alleen mailboxen, mailgebruikers en beveiligingsgroepen met e-mail kunnen dit recht krijgen.")
        foreach ($s in $sel) { if ($s) { & $addRow ([string]$s.Label) $Kind ([string]$s.Value) '+ toevoegen' '' $true } }
        & $renderDel
    }
    $removeDel = {
        $sel = @(Get-SelectedRows $c.GS_DelGrid)
        if (-not $sel.Count) { Show-Msg "Selecteer eerst een of meer regels." "Verwijderen" | Out-Null; return }
        foreach ($dr in $sel) {
            $r = @($state.Rows | Where-Object { $_.Nr -eq [string]$dr['Nr'] }) | Select-Object -First 1
            if (-not $r) { continue }
            if ($r.New) { $state.Rows.Remove($r) }                                                   # net toegevoegd: weer weg
            elseif ($r.Wachtrij -eq '+ toevoegen') { $state.Undo += $r.QKey; $state.Rows.Remove($r) } # stond al in de wachtrij: item schrappen
            elseif ($r.Remove) { $r.Remove = $false; $r.Wachtrij = '' }                              # verwijderen ongedaan maken
            elseif (-not $r.Wachtrij) { $r.Remove = $true; $r.Wachtrij = '- verwijderen' }
        }
        & $renderDel
    }
    $submit = {
        $x = & $collect
        $n = 0
        if ($x.Text.Count -and (Add-GroupSettingsToQueue -GroupName $gname -Group $addr -Set $x.Set -Text "Instellingen $addr`: $($x.Text -join '; ')")) { $n++ }
        if ($c.GS_ChkSubExisting.IsChecked -and $notSub.Count) {
            if (Add-GroupSubscriberToQueue -GroupName $gname -Group $addr -Members $notSub -Label "$($notSub.Count) huidige lid/leden" -Add $true) { $n++ }
        }
        foreach ($r in @($state.Rows)) {
            if (-not ($r.New -or $r.Remove)) { continue }
            $grant = [bool]$r.New
            $ok = if ($r.Kind -eq 'SendAs') { Add-MbxRightToQueue -Target $gname -Kind 'SendAs' -Mailbox $addr -Who $r.Key -WhoLabel $r.Naam -Grant $grant }
                  elseif ($grant) { Add-GroupSettingsToQueue -GroupName $gname -Group $addr -SobAdd @($r.Key) -Text "Verzenden namens groep $addr TOEKENNEN aan $($r.Naam)" }
                  else { Add-GroupSettingsToQueue -GroupName $gname -Group $addr -SobRemove @($r.Key) -Text "Verzenden namens groep $addr INTREKKEN voor $($r.Naam)" }
            if ($ok) { $n++ }
        }
        if (@($state.Undo).Count) { [void](Remove-QueueItems $state.Undo) }
        Update-QueueSummary
        $state.Result = $n
        $w.DialogResult = $true
    }

    foreach ($s in $script:GroupBoolSettings) { $c[$s.Box].Add_Click($updateDeps) }
    $c.GS_ChkSubExisting.Add_Click($updateChanges)
    $c.GS_CmbPrivacy.Add_SelectionChanged($updateChanges)
    $c.GS_TxtDesc.Add_TextChanged($updateChanges)
    $c.GS_TxtMailTip.Add_TextChanged($updateChanges)
    $c.GS_BtnAddSendAs.Add_Click({ & $addDel 'SendAs' })
    $c.GS_BtnAddSob.Add_Click({ & $addDel 'SendOnBehalf' })
    $c.GS_BtnDelRemove.Add_Click($removeDel)
    $c.GS_BtnOk.Add_Click($submit)
    & $renderDel
    & $updateDeps
    [void]$w.ShowDialog()
    return [int]$state.Result
}

function Show-RecipientOverview {
    # Alle ontvangers met een mailadres: mailboxen, groepen, distributielijsten,
    # contactpersonen en mailgebruikers, met groepsleden beheren en
    # contactpersonen aanmaken (via de wachtrij). Handlers vuren tijdens
    # ShowDialog(), dus lokale variabelen zijn bereikbaar.
    if (-not (Test-ExoReady)) { return }
    $w = ConvertFrom-Xaml ($script:RcptOverviewXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Owner = $script:Window
    $c = @{}
    foreach ($n in 'RO_BtnLoad', 'RO_TxtFilter', 'RO_CmbCat', 'RO_ChkHidden', 'RO_LblStatus', 'RO_Grid', 'RO_LblMembers', 'RO_MemberGrid',
                   'RO_BtnMembers', 'RO_BtnOpen', 'RO_BtnExport', 'RO_BtnExportMembers', 'RO_BtnNewContact', 'RO_BtnNewGroup',
                   'RO_BtnAddPick', 'RO_BtnAddPaste', 'RO_BtnRemove', 'RO_BtnSettings', 'RO_BtnSubOn', 'RO_BtnSubOff') { $c[$n] = $w.FindName($n) }
    $state = @{ Table = $null; Members = $null; MemberData = $null; Group = $null; Added = 0 }
    $cols = @('Naam', 'Adres', 'Type', 'Categorie', 'Bron', 'Aliassen', 'Extern', 'Beheerder', 'Verborgen', 'Doorsturen', 'Wachtrij', 'Id', 'RType')
    $mcols = @('Naam', 'Adres', 'Type', 'Inbox', 'Wachtrij', 'Key', 'QKey')
    $inboxCol = @($c.RO_MemberGrid.Columns | Where-Object { [string]$_.Header -eq 'Inbox' }) | Select-Object -First 1
    foreach ($b in 'RO_BtnMembers', 'RO_BtnSettings', 'RO_BtnOpen', 'RO_BtnExport', 'RO_BtnExportMembers', 'RO_BtnAddPick', 'RO_BtnAddPaste', 'RO_BtnRemove') { $c[$b].IsEnabled = $false }

    $applyFilter = {
        if (-not $state.Table) { return }
        $parts = @()
        $t = $c.RO_TxtFilter.Text.Trim()
        if ($t) {
            $l = ConvertTo-LikeLiteral $t
            $parts += "(Naam LIKE '*$l*' OR Adres LIKE '*$l*' OR Aliassen LIKE '*$l*' OR Extern LIKE '*$l*' OR Type LIKE '*$l*' OR Beheerder LIKE '*$l*')"
        }
        $cat = [string]$c.RO_CmbCat.SelectedItem.Tag
        if ($cat) { $parts += "Categorie = '$cat'" }
        if ($c.RO_ChkHidden.IsChecked) { $parts += "Verborgen = 'Ja'" }
        $state.Table.DefaultView.RowFilter = $parts -join ' AND '
        $n = @{}
        foreach ($r in $state.Table.Rows) { $n[[string]$r['Categorie']]++ }
        $c.RO_LblStatus.Text = "$($state.Table.DefaultView.Count) van $($state.Table.Rows.Count) zichtbaar   |   " +
            "$([int]$n['Mailbox']) mailboxen, $([int]$n['Groep']) groepen/lijsten, $([int]$n['Contact']) contacten/mailgebruikers, $([int]$n['Overig']) overig"
    }
    $addPendingContacts = {
        # Contactpersonen die in de wachtrij staan om aangemaakt te worden
        $have = @{}
        foreach ($r in $state.Table.Rows) { $have[([string]$r['Adres']).ToLower()] = $true }
        foreach ($q in @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'MailContact' })) {
            if ($have.ContainsKey(([string]$q.Data.Address).ToLower())) { continue }
            $dr = $state.Table.NewRow()
            foreach ($col in $cols) { $dr[$col] = '' }
            $dr['Naam'] = $q.Data.Name; $dr['Adres'] = $q.Data.Address; $dr['Extern'] = $q.Data.Address
            $dr['Type'] = 'Contactpersoon'; $dr['Categorie'] = 'Contact'; $dr['Bron'] = 'Cloud'; $dr['RType'] = 'MailContact'
            $dr['Verborgen'] = if ($q.Data.Hide) { 'Ja' } else { '' }
            $dr['Wachtrij'] = '+ aanmaken'
            [void]$state.Table.Rows.Add($dr)
            $have[([string]$q.Data.Address).ToLower()] = $true
        }
        # Groepen die in de wachtrij staan om aangemaakt te worden
        foreach ($q in @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'NewGroup' })) {
            if ($have.ContainsKey(([string]$q.Data.Address).ToLower())) { continue }
            $rt = $script:NewGroupRType[$q.Data.Type]
            $dr = $state.Table.NewRow()
            foreach ($col in $cols) { $dr[$col] = '' }
            $dr['Naam'] = $q.Data.Name; $dr['Adres'] = $q.Data.Address
            $dr['Type'] = (Get-RecipientTypeInfo $rt).Label; $dr['Categorie'] = 'Groep'; $dr['Bron'] = 'Cloud'; $dr['RType'] = $rt
            $dr['Beheerder'] = @($q.Data.Owners) -join '; '
            $dr['Verborgen'] = if ($q.Data.Hide) { 'Ja' } else { '' }
            $dr['Wachtrij'] = '+ aanmaken'
            [void]$state.Table.Rows.Add($dr)
            $have[([string]$q.Data.Address).ToLower()] = $true
        }
    }
    $load = {
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $c.RO_LblStatus.Text = "Ontvangers ophalen..."; Update-UI
            # Doorsturen uit de gebruikerslijst hergebruiken (al opgehaald)
            $fwd = @{}
            if ($script:UserTable) { foreach ($r in $script:UserTable.Rows) { if ($r['Doorsturen']) { $fwd[[string]$r['Id']] = [string]$r['Doorsturen'] } } }
            $rows = foreach ($x in (Get-AllRecipients)) {
                $type = [string]$x.RecipientTypeDetails
                $info = Get-RecipientTypeInfo $type
                $prim = [string]$x.PrimarySmtpAddress
                $alias = @(@($x.EmailAddresses) | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^smtp:' } |
                    ForEach-Object { $_ -replace '^smtp:', '' } | Where-Object { $_ -ne $prim })
                $id = [string]$x.ExternalDirectoryObjectId
                @{
                    Naam = [string]$x.DisplayName; Adres = $prim; Type = $info.Label; Categorie = $info.Cat
                    Bron = $(if ($x.IsDirSynced) { 'AD' } else { 'Cloud' })
                    Aliassen = ($alias -join '; ')
                    Extern = (([string]$x.ExternalEmailAddress) -replace '^smtp:', '')
                    Beheerder = (@($x.ManagedBy) | Where-Object { $_ } | ForEach-Object { [string]$_ }) -join '; '
                    Verborgen = $(if ($x.HiddenFromAddressListsEnabled) { 'Ja' } else { '' })
                    Doorsturen = $(if ($id -and $fwd.ContainsKey($id)) { $fwd[$id] } else { '' })
                    Wachtrij = ''; Id = $id; RType = $type
                }
            }
            $state.Table = Set-GridRows $c.RO_Grid $cols $rows -PassThru
            & $addPendingContacts
            $state.Table.DefaultView.Sort = 'Naam'
            $state.Group = $null; $state.MemberData = $null
            $c.RO_MemberGrid.ItemsSource = $null
            $c.RO_LblMembers.Text = "Selecteer een groep en klik op 'Leden tonen'"
            & $applyFilter
            & $onSel
            $c.RO_BtnExport.IsEnabled = ($state.Table.Rows.Count -gt 0)
            Write-Log "Ontvangers-overzicht: $($state.Table.Rows.Count) ontvanger(s) geladen" "info"
        }
        catch {
            $c.RO_LblStatus.Text = "Ophalen mislukt"
            Show-Msg "Ophalen mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
        }
        finally { $w.Cursor = $null }
    }
    $selected = {
        $rows = @(Get-SelectedRows $c.RO_Grid)
        if ($rows.Count -eq 1) { return $rows[0] }
        return $null
    }
    $groupInfo = {
        # Groep waarvan de leden nu getoond worden + of die hier te wijzigen is
        param($r)
        $why = ''
        if ($r['RType'] -eq 'DynamicDistributionGroup') { $why = 'dynamische lijst: leden volgen uit een filter' }
        elseif ($r['Bron'] -eq 'AD') { $why = 'gesynchroniseerd vanuit AD: leden wijzigen in de lokale AD' }
        elseif ($r['Wachtrij']) { $why = 'staat nog in de wachtrij' }
        return @{ Name = [string]$r['Naam']; Address = [string]$r['Adres']; Id = [string]$r['Id']; Type = [string]$r['RType']; Why = $why }
    }
    $onSel = {
        $r = & $selected
        $c.RO_BtnMembers.IsEnabled = [bool]($r -and $r['Categorie'] -eq 'Groep' -and -not $r['Wachtrij'])
        $c.RO_BtnSettings.IsEnabled = [bool]($r -and $r['RType'] -eq 'GroupMailbox' -and -not $r['Wachtrij'])
        $c.RO_BtnOpen.IsEnabled = [bool]($r -and $r['Id'] -and $script:UserTable -and $script:UserTable.Select("Id = '$($r['Id'])'").Count)
    }
    $renderMembers = {
        # Huidige leden + wat er in de wachtrij staat voor deze groep
        $g = $state.Group
        $rows = New-Object System.Collections.ArrayList
        $byKey = @{}
        $m365 = ($g.Type -eq 'GroupMailbox')
        foreach ($m in $state.MemberData) {
            $row = @{ Naam = $m.Naam; Adres = $m.Adres; Type = $m.Type; Key = $m.Key; Inbox = $m.Inbox; Wachtrij = ''; QKey = '' }
            [void]$rows.Add($row)
            foreach ($k in @($m.Key, $m.Adres, $m.Id)) { if ($k) { $byKey[$k.ToLower()] = $row } }
        }
        # Groepsmail in de inbox (abonnees) die in de wachtrij staat
        foreach ($q in @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'GroupSubscriber' -and $_.Data.Group -eq $g.Address })) {
            foreach ($mk in @($q.Data.Members)) {
                $k = ([string]$mk).ToLower()
                if ($byKey.ContainsKey($k)) { $byKey[$k].Inbox = $(if ($q.Data.Add) { '+ aan' } else { '- uit' }) }
            }
        }
        foreach ($q in @($script:Queue | Where-Object { (Test-QueueOpen $_) -and $_.Kind -eq 'GroupMember' -and $_.Data.Group -eq $g.Address })) {
            $k = ([string]$q.Data.Member).ToLower()
            if ($q.Data.Add) {
                if ($byKey.ContainsKey($k)) { continue }
                $row = @{ Naam = $q.Data.MemberLabel; Adres = $q.Data.Member; Type = ''; Key = $q.Data.Member; Inbox = ''; Wachtrij = '+ toevoegen'; QKey = $q.Key }
                [void]$rows.Add($row); $byKey[$k] = $row
            }
            elseif ($byKey.ContainsKey($k)) { $byKey[$k].Wachtrij = '- verwijderen'; $byKey[$k].QKey = $q.Key }
        }
        $state.Members = Set-GridRows $c.RO_MemberGrid $mcols $rows -PassThru
        $state.Members.DefaultView.Sort = 'Naam'
        $nq = @($rows | Where-Object { $_.Wachtrij -or $_.Inbox -like '[+-]*' }).Count
        $subTxt = if ($m365) { "`n$(@($state.MemberData | Where-Object { $_.Inbox -eq 'Ja' }).Count) van $($state.MemberData.Count) krijgen de groepsmail in hun eigen inbox (kolom Inbox)" } else { '' }
        $c.RO_LblMembers.Text = "$($state.MemberData.Count) lid/leden van $($g.Name)$(if ($nq) { "  ($nq wijziging(en) in wachtrij)" })$subTxt$(if ($g.Why) { "`nAlleen-lezen: $($g.Why)" })"
        $edit = -not $g.Why
        $c.RO_BtnAddPick.IsEnabled = $edit
        $c.RO_BtnAddPaste.IsEnabled = $edit
        $c.RO_BtnRemove.IsEnabled = $edit
        $vis = if ($m365) { 'Visible' } else { 'Collapsed' }
        if ($inboxCol) { $inboxCol.Visibility = $vis }
        $c.RO_BtnSubOn.Visibility = $vis
        $c.RO_BtnSubOff.Visibility = $vis
        $c.RO_BtnExportMembers.IsEnabled = ($rows.Count -gt 0)
    }
    $existingMembers = {
        $h = @{}
        if ($state.Members) { foreach ($r in $state.Members.Rows) { if ($r['Wachtrij'] -ne '- verwijderen') { foreach ($k in @($r['Key'], $r['Adres'])) { if ($k) { $h[([string]$k).ToLower()] = $true } } } } }
        return $h
    }
    $loadMembers = {
        $r = & $selected
        if (-not $r -or $r['Categorie'] -ne 'Groep' -or $r['Wachtrij']) { return }
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $g = & $groupInfo $r
            $c.RO_LblMembers.Text = "Leden van $($g.Name) ophalen..."; Update-UI
            $ident = if ($g.Id) { $g.Id } else { $g.Address }
            # M365-groep: wie krijgt de groepsmail in de eigen inbox (abonnees)
            $subKeys = @{}
            if ($g.Type -eq 'GroupMailbox') {
                foreach ($s in @(Invoke-Exo "Get-UnifiedGroupLinks $($g.Address) -LinkType Subscribers" { Get-UnifiedGroupLinks -Identity $ident -LinkType Subscribers -ResultSize Unlimited -ErrorAction Stop })) {
                    foreach ($k in @([string]$s.PrimarySmtpAddress, [string]$s.ExternalDirectoryObjectId)) { if ($k) { $subKeys[$k.ToLower()] = $true } }
                }
            }
            $data = foreach ($m in (Get-RecipientMembers $ident $g.Type)) {
                $smtp = [string]$m.PrimarySmtpAddress
                $oid = [string]$m.ExternalDirectoryObjectId
                $key = if ($smtp) { $smtp } elseif ($oid) { $oid } else { [string]$m.Name }
                $inbox = if ($g.Type -ne 'GroupMailbox') { '' } elseif (($smtp -and $subKeys[$smtp.ToLower()]) -or ($oid -and $subKeys[$oid.ToLower()])) { 'Ja' } else { 'Nee' }
                [pscustomobject]@{ Naam = [string]$m.DisplayName; Adres = $smtp; Type = (Get-RecipientTypeInfo ([string]$m.RecipientTypeDetails)).Label; Key = $key; Id = $oid; Inbox = $inbox }
            }
            $state.Group = $g
            $state.MemberData = @($data)
            & $renderMembers
        }
        catch {
            $c.RO_LblMembers.Text = "Leden ophalen mislukt"
            Show-Msg "Leden ophalen mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
        }
        finally { $w.Cursor = $null }
    }
    $addPick = {
        $g = $state.Group
        if (-not $g -or $g.Why) { return }
        # Bestaande leden en de groep zelf niet tonen
        $ex = & $existingMembers
        $ex[([string]$g.Address).ToLower()] = $true
        $m365 = ($g.Type -eq 'GroupMailbox')
        $allow = if ($m365) { @('UserMailbox', 'MailUser', 'GuestMailUser') } elseif ($g.Type -eq 'RoomList') { @('RoomMailbox') } else { @() }
        $what = if ($m365) { 'Alleen gebruikers en gasten worden getoond (Microsoft 365-groep).' } elseif ($g.Type -eq 'RoomList') { 'Alleen ruimte-mailboxen worden getoond.' } else { "Een nieuw extern adres? Gebruik 'Adressen plakken...'." }
        $sel = @(Show-RecipientPicker -Title "Leden toevoegen aan $($g.Name)" -Owner $w -Allow $allow -Exclude $ex -Prompt "Kies een of meer leden voor $($g.Name). Bestaande leden worden niet getoond. $what")
        foreach ($s in $sel) {
            if ($s -and (Add-GroupMemberToQueue -GroupName $g.Name -Group $g.Address -GroupType $g.Type -Member $s.Value -MemberLabel $s.Label -Add $true)) { $state.Added++ }
        }
        Update-QueueSummary
        & $renderMembers
    }
    $addPaste = {
        $g = $state.Group
        if (-not $g -or $g.Why) { return }
        $extra = if ($g.Type -eq 'GroupMailbox') { "Let op: in een Microsoft 365-groep kunnen alleen bestaande gebruikers/gasten, geen externe contactpersonen." }
                 else { "Externe adressen die nog niet bestaan krijgen eerst automatisch een contactpersoon (Naam <adres> of Naam[tab]adres geeft meteen de weergavenaam)." }
        $res = Show-AddressDialog -Title "Leden toevoegen aan $($g.Name)" -Owner $w -Prompt "Plak of typ e-mailadressen (1 per regel of gescheiden door ;). Ook direct uit Outlook of Excel. $extra"
        if (-not $res) { return }
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try { $out = Add-AddressesToQueue -Items $res.Items -Hide $res.Hide -Group $g -Existing (& $existingMembers) }
        finally { $w.Cursor = $null }
        $state.Added += $out.Added
        & $addPendingContacts
        & $applyFilter
        & $renderMembers
        Show-AddResult $out "Leden toevoegen" $w
    }
    $removeMembers = {
        $g = $state.Group
        if (-not $g -or $g.Why) { return }
        $rows = @(Get-SelectedRows $c.RO_MemberGrid)
        if ($rows.Count -eq 0) { Show-Msg "Selecteer eerst een of meer leden (Ctrl/Shift-klik)." "Verwijderen" | Out-Null; return }
        $undo = @(); $n = 0
        foreach ($r in $rows) {
            if ($r['Wachtrij'] -eq '+ toevoegen') { $undo += [string]$r['QKey']; continue }   # nog niet uitgevoerd: uit de wachtrij halen
            if ($r['Wachtrij'] -eq '- verwijderen') { continue }
            if (Add-GroupMemberToQueue -GroupName $g.Name -Group $g.Address -GroupType $g.Type -Member ([string]$r['Key']) -MemberLabel "$($r['Naam']) <$($r['Adres'])>" -Add $false) { $n++ }
        }
        if ($undo.Count) { [void](Remove-QueueItems $undo) }
        $state.Added += $n
        Update-QueueSummary
        & $renderMembers
        $c.RO_LblStatus.Text = "$n lid/leden in de wachtrij om te verwijderen$(if ($undo.Count) { ", $($undo.Count) toevoeging(en) uit de wachtrij gehaald" })"
    }
    $newContacts = {
        $res = Show-AddressDialog -Title "Nieuwe contactpersonen" -Owner $w -Prompt "Externe e-mailadressen waarvoor een contactpersoon (MailContact) aangemaakt moet worden. 1 per regel of gescheiden door ;. Gebruik 'Naam <adres>' of 'Naam[tab]adres' (Excel) voor de weergavenaam, anders wordt het adres de naam."
        if (-not $res) { return }
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try { $out = Add-AddressesToQueue -Items $res.Items -Hide $res.Hide }
        finally { $w.Cursor = $null }
        $state.Added += $out.Added
        if ($state.Table) { & $addPendingContacts; & $applyFilter }
        Show-AddResult $out "Nieuwe contactpersonen" $w
    }
    $openUser = {
        # Selecteert de ontvanger in de gebruikerslijst (opent hem daar)
        $r = & $selected
        if (-not $r -or -not $r['Id']) { return }
        $item = @($script:UserGrid.Items | Where-Object { [string]$_.Row['Id'] -eq [string]$r['Id'] }) | Select-Object -First 1
        if (-not $item) { Show-Msg "Deze ontvanger staat niet zichtbaar in de gebruikerslijst (filter actief?)." "Openen" "OK" "Warning" | Out-Null; return }
        $w.Close()
        $script:UserGrid.SelectedItem = $item
        $script:UserGrid.ScrollIntoView($item)
    }
    $export = {
        param($Table, [string[]]$Columns, [string]$Name)
        $d = New-Object Microsoft.Win32.SaveFileDialog
        $d.Filter = "CSV (*.csv)|*.csv"
        $d.FileName = "$Name-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
        if (-not $d.ShowDialog($w)) { return }
        try {
            $Table.DefaultView | ForEach-Object {
                $o = [ordered]@{}
                foreach ($col in $Columns) { $o[$col] = $_.Row[$col] }
                [pscustomobject]$o
            } | Export-Csv -Path $d.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
            $c.RO_LblStatus.Text = "Geexporteerd naar $($d.FileName)"
        }
        catch { Show-Msg "Exporteren mislukt:`n$_" "Fout" "OK" "Error" | Out-Null }
    }

    $setSubscription = {
        # Groepsmail in de eigen inbox aan/uit voor de geselecteerde leden (M365-groep)
        param([bool]$On)
        $g = $state.Group
        if (-not $g -or $g.Type -ne 'GroupMailbox' -or $g.Why) { return }
        $rows = @(Get-SelectedRows $c.RO_MemberGrid)
        if ($rows.Count -eq 0) { Show-Msg "Selecteer eerst een of meer leden (Ctrl/Shift-klik)." "Groepsmail in inbox" | Out-Null; return }
        $want = if ($On) { 'Nee' } else { 'Ja' }
        $pick = @($rows | Where-Object { [string]$_['Inbox'] -eq $want -and -not $_['Wachtrij'] })
        if ($pick.Count -eq 0) {
            Show-Msg "Geen van de geselecteerde leden $(if ($On) { 'kan worden aangezet (staat al aan, in de wachtrij of nog niet uitgevoerd)' } else { 'krijgt de groepsmail nu in de inbox' })." "Groepsmail in inbox" | Out-Null
            return
        }
        $n = 0
        foreach ($r in $pick) {
            if (Add-GroupSubscriberToQueue -GroupName $g.Name -Group $g.Address -Members @([string]$r['Key']) -Label "$($r['Naam']) <$($r['Adres'])>" -Add $On) { $n++ }
        }
        $state.Added += $n
        Update-QueueSummary
        & $renderMembers
        $c.RO_LblStatus.Text = "Groepsmail in inbox $(if ($On) { 'aanzetten' } else { 'uitzetten' }) voor $n lid/leden in de wachtrij gezet"
    }
    $openSettings = {
        $r = & $selected
        if (-not $r -or $r['RType'] -ne 'GroupMailbox' -or $r['Wachtrij']) { return }
        $g = & $groupInfo $r
        $n = [int](@(Show-GroupSettingsDialog -Group $g -Owner $w) | Select-Object -Last 1)
        if (-not $n) { return }
        $state.Added += $n
        $c.RO_LblStatus.Text = "$n groepsinstelling(en) voor $($g.Name) in de wachtrij gezet. Wordt uitgevoerd bij 'Alles uitvoeren' of bij het sluiten van dit venster."
        if ($state.Group -and $state.Group.Address -eq $g.Address) { & $renderMembers }
    }

    $c.RO_BtnSettings.Add_Click($openSettings)
    $c.RO_BtnSubOn.Add_Click({ & $setSubscription $true })
    $c.RO_BtnSubOff.Add_Click({ & $setSubscription $false })
    $c.RO_BtnLoad.Add_Click($load)
    $c.RO_TxtFilter.Add_TextChanged($applyFilter)
    $c.RO_CmbCat.Add_SelectionChanged($applyFilter)
    $c.RO_ChkHidden.Add_Click($applyFilter)
    $c.RO_Grid.Add_SelectionChanged($onSel)
    $c.RO_Grid.Add_MouseDoubleClick({
        $r = & $selected
        if (-not $r) { return }
        if ($r['Categorie'] -eq 'Groep') { & $loadMembers } elseif ($c.RO_BtnOpen.IsEnabled) { & $openUser }
    })
    $c.RO_BtnMembers.Add_Click($loadMembers)
    $c.RO_BtnAddPick.Add_Click($addPick)
    $c.RO_BtnAddPaste.Add_Click($addPaste)
    $c.RO_BtnRemove.Add_Click($removeMembers)
    $c.RO_BtnNewContact.Add_Click($newContacts)
    $c.RO_BtnNewGroup.Add_Click({
        $n = Show-NewGroupDialog -Owner $w
        if (-not $n) { return }
        $state.Added += $n
        if ($state.Table) { & $addPendingContacts; & $applyFilter }
        $c.RO_LblStatus.Text = "Nieuwe groep in de wachtrij gezet ($n item(s)). Wordt aangemaakt bij 'Alles uitvoeren' of bij het sluiten van dit venster."
    })
    $c.RO_BtnOpen.Add_Click($openUser)
    $c.RO_BtnExport.Add_Click({ if ($state.Table) { & $export $state.Table @($cols | Where-Object { $_ -notin 'RType', 'Wachtrij' }) 'Ontvangers' } })
    $c.RO_BtnExportMembers.Add_Click({
        if ($state.Members) { & $export $state.Members @('Naam', 'Adres', 'Type', 'Wachtrij') "Leden-$($state.Group.Address -replace '[^\w\.-]', '_')" }
    })
    $w.Add_ContentRendered({ & $load })
    [void]$w.ShowDialog()

    # Invoke-Queue vraagt zelf om bevestiging (Nee = later via 'Alles uitvoeren')
    if ($state.Added) { Invoke-Queue }
}

#endregion

#region --- Exchange: map rechten (1 gebruiker) ---

function Get-FolderUserKey {
    # Identiteit waarmee een folder-permissie weer aan te spreken is.
    param($Perm)
    $u = $Perm.User
    $type = ''
    try { $type = [string]$u.UserType } catch { }
    if ($type -match '^(Default|Anonymous)$') { return $type }
    foreach ($path in 'RecipientPrincipal.PrimarySmtpAddress', 'ADRecipient.PrimarySmtpAddress') {
        $v = $u
        foreach ($seg in $path.Split('.')) { if ($null -ne $v) { $v = $v.$seg } }
        if ($v) { return [string]$v }
    }
    if ($u.DisplayName) { return [string]$u.DisplayName }
    return [string]$u
}

function Load-FolderTab {
    if (-not $script:ExoConnected) { $script:LblFolderPath.Text = "Niet verbonden met Exchange Online (knop rechtsboven)"; return }
    Enter-Busy "Maprechten laden..."
    try {
        $mbx = Get-CurrentMailbox
        if (-not $mbx) {
            $script:LblFolderPath.Text = "Deze gebruiker heeft geen Exchange Online mailbox."
            $script:FolderContent.IsEnabled = $false
            return
        }
        $ftype = [string]$script:CmbFolder.SelectedItem.Tag
        $path  = Get-FolderPathFor ([string]$mbx.PrimarySmtpAddress) $ftype
        $script:ExoState.FolderPath = $path
        $script:ExoState.FolderType = $ftype
        $script:LblFolderPath.Text = $path

        $perms = Invoke-Exo "Get-EXOMailboxFolderPermission $path" { Get-EXOMailboxFolderPermission -Identity $path -ErrorAction Stop }
        $script:ExoState.FolderRows = @(foreach ($p in @($perms)) {
            $flags = [string]$p.SharingPermissionFlags
            @{
                Gebruiker   = if ($p.User.DisplayName) { [string]$p.User.DisplayName } else { [string]$p.User }
                Rechten     = (@($p.AccessRights) -join ', ')
                Gemachtigde = if ($flags -match 'Delegate') { if ($flags -match 'CanViewPrivateItems') { 'Ja (+ prive)' } else { 'Ja' } } else { '' }
                Key         = Get-FolderUserKey $p
                Flags       = $flags
            }
        })
        $script:FolderContent.IsEnabled = $true
        $script:ExoState.FolderLoaded = $true
        Show-FolderGrid
        Update-DelegateOptions
        Set-Status "Maprechten geladen" "success"
    }
    catch {
        Set-Status "Laden mislukt" "error"
        Show-Msg "Maprechten laden mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
    }
    finally { Exit-Busy }
}

function Show-FolderGrid {
    if (-not $script:ExoState.FolderLoaded) { return }
    $ids = Get-CurrentMailboxIds
    $path = $script:ExoState.FolderPath
    $ftype = $script:ExoState.FolderType
    $rows = New-Object System.Collections.ArrayList
    foreach ($r in @($script:ExoState.FolderRows)) {
        $h = @{}; foreach ($k in $r.Keys) { $h[$k] = $r[$k] }
        $h.Wachtrij = ''; $h.QKey = ''
        [void]$rows.Add($h)
    }
    $pend = @($script:Queue | Where-Object {
        (Test-QueueOpen $_) -and $_.Kind -eq 'FolderPerm' -and
        ($_.Data.Path -eq $path -or ($ids -contains $_.Data.Mailbox -and $_.Data.FolderType -eq $ftype))
    })
    foreach ($p in $pend) {
        $match = @($rows | Where-Object { $_.Key -eq $p.Data.Who -or $_.Gebruiker -eq $p.Data.Who }) | Select-Object -First 1
        $txt = switch ($p.Data.Op) { 'Remove' { '- verwijderen' } default { "-> $($p.Data.Rights)" } }
        if ($match) { $match.Wachtrij = $txt; $match.QKey = $p.Key }
        else { [void]$rows.Add(@{ Gebruiker = $p.Data.WhoLabel; Rechten = ''; Gemachtigde = ''; Key = $p.Data.Who; Flags = ''; Wachtrij = "+ $($p.Data.Rights)"; QKey = $p.Key }) }
    }
    Set-GridRows $script:FolderGrid @('Gebruiker', 'Rechten', 'Gemachtigde', 'Key', 'Flags', 'Wachtrij', 'QKey') $rows
}

function Update-DelegateOptions {
    # Gemachtigde (delegate) kan alleen op de agenda in combinatie met Editor
    $ok = ($script:ExoState.FolderType -eq 'Calendar' -and $script:CmbRights.SelectedValue -eq 'Editor')
    $script:ChkDelegate.IsEnabled = $ok
    if (-not $ok) { $script:ChkDelegate.IsChecked = $false }
    $script:ChkPrivate.IsEnabled = $ok -and [bool]$script:ChkDelegate.IsChecked
    if (-not $script:ChkPrivate.IsEnabled) { $script:ChkPrivate.IsChecked = $false }
}

function Get-SharingFlags {
    param($ChkDelegate, $ChkPrivate)
    if (-not $ChkDelegate.IsChecked) { return 'None' }
    if ($ChkPrivate.IsChecked) { return 'Delegate,CanViewPrivateItems' }
    return 'Delegate'
}

function On-FolderRowSelected {
    $row = @(Get-SelectedRows $script:FolderGrid) | Select-Object -First 1
    if (-not $row) { return }
    $first = ([string]$row['Rechten']).Split(',')[0].Trim()
    if ($script:FolderRights.Value -contains $first) { $script:CmbRights.SelectedValue = $first }
    Update-DelegateOptions
    if ($script:ChkDelegate.IsEnabled) {
        $script:ChkDelegate.IsChecked = ([string]$row['Flags'] -match 'Delegate')
        Update-DelegateOptions
        $script:ChkPrivate.IsChecked = ([string]$row['Flags'] -match 'CanViewPrivateItems')
    }
}

function Invoke-FolderAdd {
    $path = $script:ExoState.FolderPath
    if (-not $path) { return }
    $sel = Show-Picker -Title "Maprechten toekennen" -Prompt "Zoek gebruiker of groep (naam, e-mail):" -Search $script:ExoRecipientSearch
    if (-not $sel) { return }
    [void](Add-FolderPermToQueue -Target $script:CurrentUser.displayName -Op 'Upsert' -Mailbox ([string]$script:ExoState.Mailbox.PrimarySmtpAddress) `
        -FolderType $script:ExoState.FolderType -Path $path -Who $sel.Value -WhoLabel $sel.Label `
        -Rights ([string]$script:CmbRights.SelectedValue) -Flags (Get-SharingFlags $script:ChkDelegate $script:ChkPrivate))
    Update-QueueSummary
    Show-FolderGrid
}

function Invoke-FolderSet {
    $path = $script:ExoState.FolderPath
    $rows = Get-SelectedRows $script:FolderGrid
    if (-not $path -or $rows.Count -eq 0) { Show-Msg "Selecteer eerst een of meer regels." "Wijzigen" | Out-Null; return }
    $rights = [string]$script:CmbRights.SelectedValue
    $flags  = Get-SharingFlags $script:ChkDelegate $script:ChkPrivate
    $isCal  = ($script:ExoState.FolderType -eq 'Calendar')
    foreach ($row in $rows) {
        $who = [string]$row['Key']
        $q = [string]$row['QKey']
        if ($q -and -not [string]$row['Rechten']) {
            # Nieuwe regel uit de wachtrij: dat item vervangen door de nieuwe rechten
            [void](Remove-QueueItems @($q))
            [void](Add-FolderPermToQueue -Target $script:CurrentUser.displayName -Op 'Upsert' -Mailbox ([string]$script:ExoState.Mailbox.PrimarySmtpAddress) `
                -FolderType $script:ExoState.FolderType -Path $path -Who $who -WhoLabel ([string]$row['Gebruiker']) -Rights $rights -Flags $flags)
            continue
        }
        if ($q) { [void](Remove-QueueItems @($q)) }
        # Flags alleen meesturen als het nodig is (ook om delegate weg te
        # halen): bij Default/Anonymous en niet-agenda's weigert Exchange ze.
        $send = $isCal -and $who -notmatch '^(Default|Anonymous)$' -and ($flags -ne 'None' -or [string]$row['Flags'] -match 'Delegate')
        [void](Add-FolderPermToQueue -Target $script:CurrentUser.displayName -Op 'Set' -Mailbox ([string]$script:ExoState.Mailbox.PrimarySmtpAddress) `
            -FolderType $script:ExoState.FolderType -Path $path -Who $who -WhoLabel ([string]$row['Gebruiker']) -Rights $rights -Flags $flags -SendFlags $send)
    }
    Update-QueueSummary
    Show-FolderGrid
}

function Invoke-FolderRemove {
    $path = $script:ExoState.FolderPath
    $rows = Get-SelectedRows $script:FolderGrid
    if (-not $path -or $rows.Count -eq 0) { Show-Msg "Selecteer eerst een of meer regels." "Verwijderen" | Out-Null; return }
    foreach ($row in $rows) {
        $who = [string]$row['Key']
        $q = [string]$row['QKey']
        if ($q -and -not [string]$row['Rechten']) { [void](Remove-QueueItems @($q)); continue }   # nieuwe regel uit wachtrij
        if ($who -match '^(Default|Anonymous)$') {
            Show-Msg "Default en Anonymous kunnen niet verwijderd worden.`nGebruik 'Rechten van selectie wijzigen' (bijv. None of AvailabilityOnly)." "Niet mogelijk" "OK" "Warning" | Out-Null
            continue
        }
        if ($q) { [void](Remove-QueueItems @($q)) }
        [void](Add-FolderPermToQueue -Target $script:CurrentUser.displayName -Op 'Remove' -Mailbox ([string]$script:ExoState.Mailbox.PrimarySmtpAddress) `
            -FolderType $script:ExoState.FolderType -Path $path -Who $who -WhoLabel ([string]$row['Gebruiker']))
    }
    Update-QueueSummary
    Show-FolderGrid
}

#endregion

#region --- Exchange: toegang tot andere mailboxen (1 gebruiker) ---

function Invoke-AccessScan {
    if (-not (Test-ExoReady)) { return }
    $u   = $script:CurrentUser
    $upn = [string]$u.userPrincipalName
    $types = if ($script:ChkScanShared.IsChecked) { @('SharedMailbox') } else { @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox') }
    $checkSob = [bool]$script:ChkScanSob.IsChecked

    # Mogelijke schrijfwijzen van deze gebruiker in GrantSendOnBehalfTo
    $ids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($v in @($upn, $u.mail, $u.displayName, $u.id)) { if ($v) { [void]$ids.Add([string]$v) } }
    $own = Get-CurrentMailbox
    if ($own) { foreach ($v in @($own.Name, $own.Alias, $own.DisplayName, [string]$own.PrimarySmtpAddress, $own.DistinguishedName, $own.Identity)) { if ($v) { [void]$ids.Add([string]$v) } } }

    Start-LongOp "Mailboxen ophalen..."
    $rows = New-Object System.Collections.ArrayList
    $cols = @('Mailbox', 'Naam', 'Type', 'Recht', 'Key', 'Wachtrij')
    try {
        $all = @(Invoke-Exo "Get-EXOMailbox -RecipientTypeDetails $($types -join ',')" {
            if ($checkSob) { Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $types -Properties GrantSendOnBehalfTo -ErrorAction Stop }
            else { Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $types -ErrorAction Stop }
        })
        $i = 0
        foreach ($m in $all) {
            $i++
            if ($script:StopRequested) { Write-Log "Zoeken gestopt na $($i - 1) van $($all.Count) mailboxen" "warning"; break }
            if ($own -and $m.ExternalDirectoryObjectId -eq $own.ExternalDirectoryObjectId) { continue }
            Update-ExoToken
            $smtp = [string]$m.PrimarySmtpAddress
            $script:LblScan.Text = "Controleren $i / $($all.Count): $smtp"
            Set-Status "Controleren $i / $($all.Count)..."

            $base = @{ Mailbox = $smtp; Naam = $m.DisplayName; Type = [string]$m.RecipientTypeDetails; Wachtrij = '' }
            $fa = Get-EXOMailboxPermission -Identity $smtp -User $upn -ErrorAction SilentlyContinue
            if (@($fa | Where-Object { -not $_.IsInherited -and -not $_.Deny -and (@($_.AccessRights) -join ',') -match 'FullAccess' }).Count -gt 0) {
                [void]$rows.Add($base + @{ Recht = 'FullAccess'; Key = $upn })
            }
            $sa = Get-EXORecipientPermission -Identity $smtp -Trustee $upn -AccessRights SendAs -ErrorAction SilentlyContinue
            if (@($sa | Where-Object { -not $_.IsInherited }).Count -gt 0) {
                [void]$rows.Add($base + @{ Recht = 'SendAs'; Key = $upn })
            }
            if ($checkSob) {
                $hit = @($m.GrantSendOnBehalfTo | Where-Object { $_ -and $ids.Contains([string]$_) }) | Select-Object -First 1
                if ($hit) { [void]$rows.Add($base + @{ Recht = 'SendOnBehalf'; Key = [string]$hit }) }
            }
            if ($i % 10 -eq 0) { Set-GridRows $script:ScanGrid $cols $rows }
        }
        Set-GridRows $script:ScanGrid $cols $rows
        $script:LblScan.Text = "$($rows.Count) recht(en) gevonden op $($all.Count) gecontroleerde mailbox(en)"
        Write-Log "Toegang van ${upn}: $($rows.Count) recht(en) gevonden" "success"
        Set-Status "Zoeken klaar" "success"
    }
    catch {
        Set-Status "Zoeken mislukt" "error"
        Show-Msg "Zoeken mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
    }
    finally { Stop-LongOp }
}

function Invoke-ScanRemove {
    $rows = Get-SelectedRows $script:ScanGrid
    if ($rows.Count -eq 0) { Show-Msg "Selecteer eerst een of meer regels." "Intrekken" | Out-Null; return }
    $n = 0
    foreach ($row in $rows) {
        if (Add-MbxRightToQueue -Target $script:CurrentUser.displayName -Kind ([string]$row['Recht']) -Mailbox ([string]$row['Mailbox']) `
                -Who ([string]$row['Key']) -WhoLabel $script:CurrentUser.displayName -Grant $false) { $n++ }
        $row['Wachtrij'] = '- intrekken'
    }
    Update-QueueSummary
    Set-Status "$n intrekking(en) toegevoegd aan de wachtrij" "success"
}

#endregion

#region --- Bulk: Entra ---

function Invoke-BulkPickManager {
    $sel = Show-Picker -Title "Manager kiezen (bulk)" -Prompt "Zoek op naam, UPN of e-mail:" -Search $script:GraphUserSearch
    if (-not $sel) { return }
    $script:BulkManager = $sel
    $script:TxtBulkMgr.Text = $sel.Label
    $script:ChkBulkMgr.IsChecked = $true
}

function Invoke-BulkEntra {
    $users = @($script:SelectedUsers)
    if ($users.Count -eq 0) { return }

    $body = [ordered]@{}
    foreach ($f in @($script:EntraFields | Where-Object { $_.Bulk })) {
        $b = $script:BulkBoxes[$f.Prop]
        if (-not $b.Chk.IsChecked) { continue }
        $err = $null
        $val = ConvertTo-PatchValue -Prop $f.Prop -Text $b.Box.Text.Trim() -ErrorMessage ([ref]$err)
        if ($err) { Show-Msg "'$($f.Label)' $err." "Ongeldige invoer" "OK" "Warning" | Out-Null; return }
        $body[$f.Prop] = $val
    }
    switch ($script:CmbBulkAccount.SelectedIndex) {
        1 { $body.accountEnabled = $true }
        2 { $body.accountEnabled = $false }
    }
    $doMgr = [bool]$script:ChkBulkMgr.IsChecked
    if ($body.Count -eq 0 -and -not $doMgr) {
        Show-Msg "Vink eerst aan welke velden je wilt aanpassen." "Niets gekozen" | Out-Null
        return
    }
    $text = Format-BodyText $body
    $n = 0
    foreach ($u in $users) {
        if ($body.Count -gt 0) {
            # Per gebruiker een eigen kopie van de body
            $b = [ordered]@{}; foreach ($k in @($body.Keys)) { $b[$k] = $body[$k] }
            if (Add-QueueItem -Kind 'GraphPatch' -Target $u.Name -Soort 'Entra (bulk)' -Text $text -Data @{ UserId = $u.Id; Body = $b }) { $n++ }
        }
        if ($doMgr) {
            if ($script:BulkManager) {
                if ($script:BulkManager.Id -eq $u.Id) { Write-Log "$($u.Name): overgeslagen als eigen manager" "warning"; continue }
                if (Add-QueueItem -Kind 'GraphManager' -Target $u.Name -Soort 'Manager (bulk)' -Text "Manager instellen op $($script:BulkManager.Label)" `
                    -Data @{ UserId = $u.Id; Action = 'Set'; ManagerId = $script:BulkManager.Id; ManagerLabel = $script:BulkManager.Label }) { $n++ }
            }
            else {
                if (Add-QueueItem -Kind 'GraphManager' -Target $u.Name -Soort 'Manager (bulk)' -Text "Manager verwijderen" -Data @{ UserId = $u.Id; Action = 'Delete' }) { $n++ }
            }
        }
    }
    Update-QueueSummary
    Write-Log "Bulk Entra: $n item(s) voor $($users.Count) gebruikers toegevoegd aan wachtrij" "info"
    Set-Status "$n item(s) toegevoegd aan de wachtrij" "success"
}

#endregion

#region --- Bulk: Exchange ---

function Update-BulkLabels {
    $script:LblBulkMbxTarget.Text = if ($script:RbMbxA.IsChecked) { "Mailbox (bijv. gedeelde mailbox):" } else { "Gebruiker/groep die toegang krijgt:" }
    $folA = [bool]$script:RbFolA.IsChecked
    $folC = [bool]$script:RbFolC.IsChecked
    $script:LblBulkFolTarget.Text = if ($folA) { "Mailbox van wie de map is:" } else { "Gebruiker/groep die toegang krijgt:" }
    $script:BulkFolTargetPanel.IsEnabled = -not $folC
    $script:BtnBulkFolRevoke.IsEnabled = -not $folC
    Update-BulkDelegateOptions
}

function Update-BulkDelegateOptions {
    $ok = ([string]$script:CmbBulkFolder.SelectedItem.Tag -eq 'Calendar' -and $script:CmbBulkRights.SelectedValue -eq 'Editor' -and -not $script:RbFolC.IsChecked)
    $script:ChkBulkDelegate.IsEnabled = $ok
    if (-not $ok) { $script:ChkBulkDelegate.IsChecked = $false }
    $script:ChkBulkPrivate.IsEnabled = $ok -and [bool]$script:ChkBulkDelegate.IsChecked
    if (-not $script:ChkBulkPrivate.IsEnabled) { $script:ChkBulkPrivate.IsChecked = $false }
}

function Invoke-BulkPickTarget {
    param([ValidateSet('Mbx', 'Fol')][string]$Which)
    if (-not (Test-ExoReady)) { return }
    $sel = Show-Picker -Title "Kiezen" -Prompt "Zoek mailbox, gebruiker of groep (naam, e-mail):" -Search $script:ExoRecipientSearch
    if (-not $sel) { return }
    if ($Which -eq 'Mbx') { $script:BulkMbxTarget = $sel; $script:TxtBulkMbxTarget.Text = $sel.Label }
    else { $script:BulkFolTarget = $sel; $script:TxtBulkFolTarget.Text = $sel.Label }
}

function Invoke-BulkMbx {
    param([bool]$Grant)
    $users = @($script:SelectedUsers)
    $t = $script:BulkMbxTarget
    if (-not $t) { Show-Msg "Kies eerst de mailbox / gebruiker." "Bulk mailbox rechten" | Out-Null; return }
    $kinds = @()
    if ($script:ChkBFull.IsChecked)   { $kinds += 'FullAccess' }
    if ($script:ChkBSendAs.IsChecked) { $kinds += 'SendAs' }
    if ($script:ChkBSob.IsChecked)    { $kinds += 'SendOnBehalf' }
    if ($kinds.Count -eq 0) { Show-Msg "Vink minimaal 1 recht aan." "Bulk mailbox rechten" | Out-Null; return }

    $dirA = [bool]$script:RbMbxA.IsChecked
    $auto = [bool]$script:ChkBAuto.IsChecked
    $n = 0
    foreach ($u in $users) {
        $mbx = if ($dirA) { $t.Value } else { $u.Upn }
        $who = if ($dirA) { $u.Upn } else { $t.Value }
        $whoLabel = if ($dirA) { $u.Name } else { $t.Label }
        if ($mbx -eq $who) { continue }
        foreach ($k in $kinds) {
            if (Add-MbxRightToQueue -Target $u.Name -Kind $k -Mailbox $mbx -Who $who -WhoLabel $whoLabel -Grant $Grant -AutoMap $auto) { $n++ }
        }
    }
    Update-QueueSummary
    Write-Log "Bulk mailbox rechten: $n item(s) toegevoegd aan wachtrij" "info"
    Set-Status "$n item(s) toegevoegd aan de wachtrij" "success"
}

function Invoke-BulkFolder {
    param([bool]$Grant)
    $users = @($script:SelectedUsers)
    $mode = if ($script:RbFolA.IsChecked) { 'A' } elseif ($script:RbFolB.IsChecked) { 'B' } else { 'C' }
    $t = $script:BulkFolTarget
    if ($mode -ne 'C' -and -not $t) { Show-Msg "Kies eerst de mailbox / gebruiker." "Bulk map rechten" | Out-Null; return }
    $ftype  = [string]$script:CmbBulkFolder.SelectedItem.Tag
    $rights = [string]$script:CmbBulkRights.SelectedValue
    $flags  = Get-SharingFlags $script:ChkBulkDelegate $script:ChkBulkPrivate
    $n = 0
    foreach ($u in $users) {
        $mbx = if ($mode -eq 'A') { $t.Value } else { $u.Upn }
        $who = switch ($mode) { 'A' { $u.Upn } 'B' { $t.Value } 'C' { 'Default' } }
        $whoLabel = switch ($mode) { 'A' { $u.Name } 'B' { $t.Label } 'C' { 'Default (iedereen)' } }
        if ($mbx -eq $who) { continue }
        $op = if ($Grant) { 'Upsert' } else { 'Remove' }
        if (Add-FolderPermToQueue -Target $u.Name -Op $op -Mailbox $mbx -FolderType $ftype -Who $who -WhoLabel $whoLabel -Rights $rights -Flags $flags) { $n++ }
    }
    Update-QueueSummary
    Write-Log "Bulk map rechten: $n item(s) toegevoegd aan wachtrij" "info"
    Set-Status "$n item(s) toegevoegd aan de wachtrij" "success"
}

#endregion

#region --- CSV import / export ---

function Get-CsvColumns {
    # Volgorde van kolommen in export en sjabloon
    $cols = @('UserPrincipalName')
    $cols += @($script:EntraFields | Where-Object { $_.Prop -ne 'userPrincipalName' } | ForEach-Object { $_.Prop })
    $cols += @('accountEnabled', 'managerUPN')
    return $cols
}

function Get-UserSelectProps {
    return (@($script:EntraFields | ForEach-Object { $_.Prop }) + @('id', 'accountEnabled') | Select-Object -Unique) -join ','
}

function Export-UsersCsv {
    param([string[]]$Ids, [string]$Path)
    $sel = Get-UserSelectProps
    $reqs = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $Ids.Count; $i++) {
        [void]$reqs.Add(@{ id = "u$i"; method = 'GET'; url = "/users/$($Ids[$i])?`$select=$sel" })
        [void]$reqs.Add(@{ id = "m$i"; method = 'GET'; url = "/users/$($Ids[$i])/manager?`$select=userPrincipalName" })
    }
    $res = Invoke-GraphBatch -Requests $reqs -What "Exporteren"
    $out = for ($i = 0; $i -lt $Ids.Count; $i++) {
        $x = $res["u$i"]
        if (-not (Test-BatchOk $x)) { Write-Log "Export: gebruiker $($Ids[$i]) overgeslagen: $(Get-BatchError $x)" "warning"; continue }
        $u = $x.body
        $o = [ordered]@{ UserPrincipalName = $u.userPrincipalName }
        foreach ($f in $script:EntraFields) { if ($f.Prop -ne 'userPrincipalName') { $o[$f.Prop] = Get-EntraValue $u $f.Prop } }
        $o.accountEnabled = if ($u.accountEnabled) { 'TRUE' } else { 'FALSE' }
        $m = $res["m$i"]
        $o.managerUPN = if (Test-BatchOk $m) { [string]$m.body.userPrincipalName } else { '' }
        [pscustomobject]$o
    }
    @($out) | Export-Csv -Path $Path -Delimiter ';' -NoTypeInformation -Encoding UTF8
    return @($out).Count
}

function Read-CsvFile {
    # Leest een CSV uit Excel: UTF-8 (met/zonder BOM) of ANSI, ; of , als scheidingsteken
    param([string]$Path)
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $text = [Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    else {
        try { $text = ([Text.UTF8Encoding]::new($false, $true)).GetString($bytes) }
        catch { $text = [Text.Encoding]::Default.GetString($bytes) }
    }
    $lines = @($text -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
    if ($lines.Count -lt 2) { throw "CSV bevat geen gegevensregels." }
    $delim = if ($lines[0].Split(';').Count -ge $lines[0].Split(',').Count) { ';' } else { ',' }
    return @($lines | ConvertFrom-Csv -Delimiter $delim)
}

function Build-CsvPlan {
    # Vergelijkt de CSV met de huidige waarden in Entra. Geeft @{ Plan; Table }
    # terug: Plan = per gebruiker de PATCH-body en manager-actie, Table = preview.
    param($Rows)
    $headers = @($Rows[0].PSObject.Properties.Name)
    $keyCol = $headers | Where-Object { $_ -in 'UserPrincipalName', 'UPN' } | Select-Object -First 1
    if (-not $keyCol) { throw "Kolom 'UserPrincipalName' (of 'UPN') ontbreekt in de CSV." }

    $map = [ordered]@{}
    foreach ($h in $headers) {
        if ($h -eq $keyCol) { continue }
        $f = $script:EntraFields | Where-Object { $_.Prop -eq $h -or $_.Label -eq $h } | Select-Object -First 1
        if ($f -and $f.Prop -ne 'userPrincipalName') { $map[$h] = $f }
        elseif ($h -eq 'accountEnabled') { $map[$h] = [pscustomobject]@{ Prop = 'accountEnabled'; Label = 'Account ingeschakeld' } }
        elseif ($h -in 'managerUPN', 'Manager') { $map[$h] = [pscustomobject]@{ Prop = 'manager'; Label = 'Manager' } }
        else { Write-Log "CSV-kolom '$h' wordt genegeerd (onbekend veld)" "warning" }
    }

    $sel = Get-UserSelectProps
    $reqs = New-Object System.Collections.ArrayList
    $mgrUpns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $hasMgr = @($map.Values | Where-Object { $_.Prop -eq 'manager' }).Count -gt 0
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $upn = ([string]$Rows[$i].$keyCol).Trim()
        if (-not $upn) { continue }
        $e = [uri]::EscapeDataString($upn)
        [void]$reqs.Add(@{ id = "u$i"; method = 'GET'; url = "/users/$($e)?`$select=$sel" })
        if ($hasMgr) {
            [void]$reqs.Add(@{ id = "m$i"; method = 'GET'; url = "/users/$e/manager?`$select=id,userPrincipalName" })
            foreach ($h in $map.Keys) {
                if ($map[$h].Prop -ne 'manager') { continue }
                $v = ([string]$Rows[$i].$h).Trim()
                if ($v -and $v -ne $script:ClearToken) { [void]$mgrUpns.Add($v) }
            }
        }
    }
    $j = 0
    $mgrKey = @{}
    foreach ($m in $mgrUpns) { $j++; $mgrKey[$m.ToLower()] = "x$j"; [void]$reqs.Add(@{ id = "x$j"; method = 'GET'; url = "/users/$([uri]::EscapeDataString($m))?`$select=id,displayName,userPrincipalName" }) }

    $res = Invoke-GraphBatch -Requests $reqs -What "Huidige waarden ophalen"

    $table = New-Table @('UPN', 'Veld', 'Huidig', 'Nieuw', 'Status', 'Key')
    $plan  = New-Object System.Collections.ArrayList
    $addRow = { param($upn, $veld, $old, $new, $status, $key)
        $dr = $table.NewRow(); $dr['UPN'] = $upn; $dr['Veld'] = $veld; $dr['Huidig'] = $old; $dr['Nieuw'] = $new; $dr['Status'] = $status; $dr['Key'] = $key
        [void]$table.Rows.Add($dr)
    }

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $row = $Rows[$i]
        $upn = ([string]$row.$keyCol).Trim()
        if (-not $upn) { continue }
        $x = $res["u$i"]
        if (-not (Test-BatchOk $x)) { & $addRow $upn '-' '' '' "FOUT: gebruiker niet gevonden ($(Get-BatchError $x))" ''; continue }
        $u = $x.body
        $key = [string]$plan.Count
        $body = [ordered]@{}
        $old = @{}
        $mgr = $null
        $bad = $false
        foreach ($h in $map.Keys) {
            $f = $map[$h]
            $cell = ([string]$row.$h).Trim()
            if (-not $cell) { continue }
            $clear = ($cell -eq $script:ClearToken)
            switch ($f.Prop) {
                'accountEnabled' {
                    $val = switch -Regex ($cell) { '^(true|waar|ja|yes|1)$' { $true } '^(false|onwaar|nee|no|0)$' { $false } default { $null } }
                    if ($null -eq $val) { & $addRow $upn $f.Label '' $cell "FOUT: gebruik TRUE of FALSE" ''; $bad = $true; break }
                    if ($val -ne [bool]$u.accountEnabled) { $body.accountEnabled = $val; & $addRow $upn $f.Label "$([bool]$u.accountEnabled)" "$val" 'Te wijzigen' $key }
                }
                'manager' {
                    $m = $res["m$i"]
                    $cur = if (Test-BatchOk $m) { [string]$m.body.userPrincipalName } else { '' }
                    if ($clear) {
                        if ($cur) { $mgr = @{ Action = 'Delete'; Label = '(verwijderen)' }; & $addRow $upn 'Manager' $cur '(verwijderen)' 'Te wijzigen' $key }
                    }
                    elseif ($cell -ne $cur) {
                        $mx = $res[$mgrKey[$cell.ToLower()]]
                        if (-not (Test-BatchOk $mx)) { & $addRow $upn 'Manager' $cur $cell "FOUT: manager niet gevonden" ''; $bad = $true }
                        elseif ($mx.body.id -eq $u.id) { & $addRow $upn 'Manager' $cur $cell "FOUT: kan niet eigen manager zijn" ''; $bad = $true }
                        else { $mgr = @{ Action = 'Set'; Id = $mx.body.id; Label = "$($mx.body.displayName)  <$cell>" }; & $addRow $upn 'Manager' $cur $cell 'Te wijzigen' $key }
                    }
                }
                default {
                    $new = if ($clear) { '' } else { $cell }
                    $cur = Get-EntraValue $u $f.Prop
                    if ($new -ceq $cur) { continue }
                    $err = $null
                    $val = ConvertTo-PatchValue -Prop $f.Prop -Text $new -ErrorMessage ([ref]$err)
                    if ($err) { & $addRow $upn $f.Label $cur $new "FOUT: $err" ''; $bad = $true; continue }
                    $body[$f.Prop] = $val
                    $old[$f.Prop] = $cur
                    & $addRow $upn $f.Label $cur $(if ($clear) { '(leegmaken)' } else { $new }) 'Te wijzigen' $key
                }
            }
        }
        if ($bad) {
            # Regel met een fout: niets van deze gebruiker toepassen
            foreach ($dr in $table.Select("Key = '$key'")) { $dr['Status'] = 'Overgeslagen: fout in deze regel'; $dr['Key'] = '' }
        }
        elseif ($body.Count -gt 0 -or $mgr) {
            [void]$plan.Add(@{ Id = $u.id; Upn = $upn; Name = [string]$u.displayName; Body = $body; Old = $old; Manager = $mgr })
        }
        else {
            & $addRow $upn '-' '' '' 'Geen wijzigingen' ''
        }
    }
    return @{ Plan = $plan; Table = $table }
}

function Add-CsvToQueue {
    param($Plan, $Table)
    $n = 0
    for ($k = 0; $k -lt $Plan.Count; $k++) {
        $p = $Plan[$k]
        $name = if ($p.Name) { $p.Name } else { $p.Upn }
        if ($p.Body.Count -gt 0) {
            if (Add-QueueItem -Kind 'GraphPatch' -Target $name -Soort 'Entra (CSV)' -Text (Format-BodyText $p.Body $p.Old) -Data @{ UserId = $p.Id; Body = $p.Body }) { $n++ }
        }
        if ($p.Manager) {
            if ($p.Manager.Action -eq 'Delete') {
                if (Add-QueueItem -Kind 'GraphManager' -Target $name -Soort 'Manager (CSV)' -Text "Manager verwijderen" -Data @{ UserId = $p.Id; Action = 'Delete' }) { $n++ }
            }
            else {
                if (Add-QueueItem -Kind 'GraphManager' -Target $name -Soort 'Manager (CSV)' -Text "Manager instellen op $($p.Manager.Label)" `
                    -Data @{ UserId = $p.Id; Action = 'Set'; ManagerId = $p.Manager.Id; ManagerLabel = $p.Manager.Label }) { $n++ }
            }
        }
        foreach ($dr in $Table.Select("Key = '$k'")) { $dr['Status'] = 'In wachtrij' }
    }
    Update-QueueSummary
    Write-Log "CSV import: $n item(s) voor $($Plan.Count) gebruiker(s) toegevoegd aan wachtrij" "info"
    return $n
}

function Show-CsvDialog {
    if (-not $script:GraphConnected) { Show-Msg "Verbind eerst met Microsoft Graph." "CSV" "OK" "Warning" | Out-Null; return }
    $w = ConvertFrom-Xaml ($script:CsvXaml.Replace('{{STYLES}}', $script:StylesXaml).Replace('{{CLEAR}}', $script:ClearToken).Replace('{{COLS}}', ((Get-CsvColumns) -join '; ')))
    $w.Owner = $script:Window
    $c = @{}
    foreach ($n in 'CV_BtnExportSel', 'CV_BtnExportAll', 'CV_BtnTemplate', 'CV_BtnLoad', 'CV_LblFile', 'CV_Grid', 'CV_BtnApply', 'CV_LblStatus', 'CV_ChkOnlyChanges') { $c[$n] = $w.FindName($n) }
    $state = @{ Plan = $null; Table = $null }
    $selCount = $script:UserGrid.SelectedItems.Count
    $c.CV_BtnExportSel.Content = "Exporteer selectie ($selCount)"
    $c.CV_BtnExportSel.IsEnabled = ($selCount -gt 0)
    $c.CV_BtnApply.IsEnabled = $false

    $saveDlg = {
        param($name)
        $d = New-Object Microsoft.Win32.SaveFileDialog
        $d.Filter = "CSV (*.csv)|*.csv"
        $d.FileName = $name
        if ($d.ShowDialog($w)) { return $d.FileName }
        return $null
    }
    $export = {
        param([string[]]$ids)
        $file = & $saveDlg "Gebruikers-export-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
        if (-not $file) { return }
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $c.CV_LblStatus.Text = "Exporteren..."; Update-UI
            $n = Export-UsersCsv -Ids $ids -Path $file
            $c.CV_LblStatus.Text = "$n gebruikers geexporteerd naar $file"
            Write-Log "$n gebruikers geexporteerd naar $file" "success"
        }
        catch { Show-Msg "Exporteren mislukt:`n$_" "Fout" "OK" "Error" | Out-Null }
        finally { $w.Cursor = $null }
    }
    $applyFilter = {
        if ($state.Table) {
            $state.Table.DefaultView.RowFilter = if ($c.CV_ChkOnlyChanges.IsChecked) { "Status <> 'Geen wijzigingen'" } else { "" }
        }
    }

    $c.CV_BtnExportSel.Add_Click({ & $export @($script:UserGrid.SelectedItems | ForEach-Object { [string]$_.Row['Id'] }) })
    $c.CV_BtnExportAll.Add_Click({
        if (-not $script:UserTable) { return }
        & $export @($script:UserTable.DefaultView | ForEach-Object { [string]$_.Row['Id'] })
    })
    $c.CV_BtnTemplate.Add_Click({
        $file = & $saveDlg "Gebruikers-sjabloon.csv"
        if (-not $file) { return }
        [IO.File]::WriteAllText($file, ((Get-CsvColumns) -join ';') + "`r`n", (New-Object System.Text.UTF8Encoding($true)))
        $c.CV_LblStatus.Text = "Leeg sjabloon opgeslagen: $file"
    })
    $c.CV_BtnLoad.Add_Click({
        $d = New-Object Microsoft.Win32.OpenFileDialog
        $d.Filter = "CSV (*.csv)|*.csv|Alle bestanden (*.*)|*.*"
        if (-not $d.ShowDialog($w)) { return }
        $w.Cursor = [System.Windows.Input.Cursors]::Wait
        $c.CV_BtnApply.IsEnabled = $false
        try {
            $c.CV_LblFile.Text = $d.FileName
            $c.CV_LblStatus.Text = "CSV lezen en vergelijken met Entra..."; Update-UI
            $rows = Read-CsvFile $d.FileName
            $r = Build-CsvPlan $rows
            $state.Plan = $r.Plan
            $state.Table = $r.Table
            $c.CV_Grid.ItemsSource = $r.Table.DefaultView
            & $applyFilter
            $nChg = @($r.Table.Select("Status = 'Te wijzigen'")).Count
            $nErr = @($r.Table.Select("Status LIKE 'FOUT*'")).Count
            $c.CV_LblStatus.Text = "$($rows.Count) regels gelezen: $nChg wijziging(en) bij $($r.Plan.Count) gebruiker(s), $nErr fout(en). Controleer en zet in de wachtrij."
            $c.CV_BtnApply.IsEnabled = ($r.Plan.Count -gt 0)
        }
        catch {
            $c.CV_LblStatus.Text = "Inlezen mislukt"
            Show-Msg "CSV inlezen mislukt:`n$_" "Fout" "OK" "Error" | Out-Null
        }
        finally { $w.Cursor = $null }
    })
    $c.CV_ChkOnlyChanges.Add_Click($applyFilter)
    $c.CV_BtnApply.Add_Click({
        if (-not $state.Plan -or $state.Plan.Count -eq 0) { return }
        $n = Add-CsvToQueue -Plan $state.Plan -Table $state.Table
        $c.CV_LblStatus.Text = "$n item(s) in de wachtrij gezet. Sluit dit venster en klik onderin op 'Alles uitvoeren'."
        $c.CV_BtnApply.IsEnabled = $false
        Set-Status "CSV: $n item(s) in de wachtrij" "success"
    })
    [void]$w.ShowDialog()
}

#endregion

#region --- Overzicht bewerken (tabel, meerdere gebruikers) ---
# Alle state staat in $script:Ov (handlers vuren tijdens ShowDialog, maar via
# $script: blijft het ook in Dispatcher-callbacks betrouwbaar bereikbaar).

function Get-OvFields {
    # UPN is de sleutel en wordt hier niet bewerkt (dat kan per gebruiker)
    return @($script:EntraFields | Where-Object { $_.Prop -ne 'userPrincipalName' })
}

function Show-OvMsg {
    param([string]$Text, [string]$Title = "Overzicht", [string]$Button = "OK", [string]$Icon = "Information")
    return [System.Windows.MessageBox]::Show($script:Ov.Win, $Text, $Title, $Button, $Icon)
}

function Get-OvUsers {
    # Haalt de bewerkbare velden op, in de volgorde van $Ids. Bij veel
    # gebruikers is 1 gepagineerde lijst sneller dan losse requests.
    param([string[]]$Ids)
    $sel = (Get-UserSelectProps) + ',onPremisesSyncEnabled'
    $byId = @{}
    if ($Ids.Count -gt 200) {
        foreach ($u in (Get-GraphPaged "$($script:GraphRoot)/users?`$select=$sel&`$top=999")) { $byId[[string]$u.id] = $u }
    }
    else {
        $reqs = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $Ids.Count; $i++) { [void]$reqs.Add(@{ id = "u$i"; method = 'GET'; url = "/users/$($Ids[$i])?`$select=$sel" }) }
        $res = Invoke-GraphBatch -Requests $reqs -What "Gebruikers ophalen"
        for ($i = 0; $i -lt $Ids.Count; $i++) {
            $x = $res["u$i"]
            if (Test-BatchOk $x) { $byId[$Ids[$i]] = $x.body }
        }
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($id in $Ids) {
        if ($byId.ContainsKey($id)) { [void]$out.Add($byId[$id]) }
        elseif (-not $script:StopRequested) { Write-Log "Overzicht: gebruiker $id niet opgehaald, overgeslagen" "warning" }
    }
    return ,$out
}

function New-OvTable {
    # Per veld 3 kolommen: <prop> (bewerkbaar), o_<prop> (waarde bij openen of
    # na in de wachtrij zetten) en s_<prop> (status voor de celkleur:
    # C = gewijzigd, E = ongeldig, Q = staat in de wachtrij).
    param($Users)
    $t = New-Object System.Data.DataTable
    foreach ($c in 'Id', 'Naam', 'UPN', 'Bron', 'Mailbox') { [void]$t.Columns.Add($c) }
    foreach ($c in 'Wijz', 'Hit') { [void]$t.Columns.Add($c, [int]) }
    foreach ($f in $script:Ov.Fields) { foreach ($c in $f.Prop, "o_$($f.Prop)", "s_$($f.Prop)") { [void]$t.Columns.Add($c) } }
    # Mailboxtype komt uit de gebruikerslijst (gevuld via Exchange)
    $mbx = @{}
    if ($script:UserTable) { foreach ($r in $script:UserTable.Rows) { $mbx[[string]$r['Id']] = [string]$r['Mailbox'] } }
    foreach ($u in $Users) {
        $id = [string]$u.id
        $pending = (Get-PendingUserOverlay $id).Body
        $dr = $t.NewRow()
        $dr['Id'] = $id; $dr['Naam'] = [string]$u.displayName; $dr['UPN'] = [string]$u.userPrincipalName
        $dr['Bron'] = if ($u.onPremisesSyncEnabled) { 'AD' } else { 'Cloud' }
        $dr['Mailbox'] = $mbx[$id]
        $dr['Wijz'] = 0; $dr['Hit'] = 1
        foreach ($f in $script:Ov.Fields) {
            $p = $f.Prop
            $val = Get-EntraValue $u $p
            $st = ''
            if ($pending.Contains($p)) { $val = Format-PatchValue $pending[$p]; $script:Ov.Queued["$id|$p"] = $true; $st = 'Q' }
            $dr[$p] = $val; $dr["o_$p"] = $val; $dr["s_$p"] = $st
        }
        [void]$t.Rows.Add($dr)
    }
    return ,$t
}

function Test-OvEditable {
    param([System.Data.DataRow]$Row, [string]$Prop)
    if (-not $script:Ov.Props.Contains($Prop)) { return $false }
    return ($Row['Bron'] -ne 'AD' -or $script:CloudOnlyProps -contains $Prop)
}

function Update-OvRowState {
    # Bepaalt per cel de status (kleur) en het aantal wijzigingen van de rij
    param([System.Data.DataRow]$Row)
    $n = 0
    foreach ($f in $script:Ov.Fields) {
        $p = $f.Prop
        $cur = ([string]$Row[$p]).Trim()
        $st = ''
        if ($cur -cne [string]$Row["o_$p"]) {
            $err = $null
            [void](ConvertTo-PatchValue -Prop $p -Text $cur -ErrorMessage ([ref]$err))
            $st = if ($err) { 'E' } else { 'C' }
            $n++
        }
        elseif ($script:Ov.Queued.ContainsKey("$($Row['Id'])|$p")) { $st = 'Q' }
        if ([string]$Row["s_$p"] -ne $st) { $Row["s_$p"] = $st }
    }
    if ([int]$Row['Wijz'] -ne $n) { $Row['Wijz'] = $n }
}

function Complete-OvEdit {
    # Rondt een lopende celbewerking af en werkt de status van die rij bij
    $g = $script:Ov.Grid
    [void]$g.CommitEdit('Row', $true)
    if ($g.CurrentItem -is [System.Data.DataRowView]) { Update-OvRowState $g.CurrentItem.Row }
}

function Set-OvCells {
    # Zet waarden in de tabel. $Cells: objecten met Row, Prop en Text.
    # Alleen-lezen cellen (AD-gesynchroniseerd) worden overgeslagen.
    param($Cells)
    Complete-OvEdit
    $touched = [System.Collections.Generic.HashSet[object]]::new()
    $set = 0; $locked = 0
    foreach ($c in $Cells) {
        if (-not (Test-OvEditable $c.Row $c.Prop)) { $locked++; continue }
        if ([string]$c.Row[$c.Prop] -cne $c.Text) { $c.Row[$c.Prop] = $c.Text; [void]$touched.Add($c.Row) }
        $set++
    }
    foreach ($r in $touched) { Update-OvRowState $r }
    return @{ Set = $set; Locked = $locked }
}

function Get-OvSelectedCells {
    # Geselecteerde cellen in de veldkolommen (niet Naam/UPN/Bron)
    $out = New-Object System.Collections.ArrayList
    foreach ($ci in @($script:Ov.Grid.SelectedCells)) {
        $p = [string]$ci.Column.SortMemberPath
        if ($script:Ov.Props.Contains($p) -and $ci.Item -is [System.Data.DataRowView]) {
            [void]$out.Add([pscustomobject]@{ Row = $ci.Item.Row; Prop = $p; Text = '' })
        }
    }
    return ,$out
}

function Get-OvLockedText {
    param([int]$Locked)
    if ($Locked) { return "; $Locked overgeslagen (AD-gesynchroniseerd, alleen-lezen)" }
    return ""
}

function Invoke-OvSetSelected {
    param([string]$Text, [switch]$Revert)
    $cells = Get-OvSelectedCells
    if ($cells.Count -eq 0) { Show-OvMsg "Selecteer eerst een of meer cellen in de tabel (Ctrl/Shift-klik, of Ctrl+A voor alles)." | Out-Null; return }
    foreach ($c in $cells) { $c.Text = if ($Revert) { [string]$c.Row["o_$($c.Prop)"] } else { $Text } }
    $r = Set-OvCells $cells
    $what = if ($Revert) { "hersteld" } elseif ($Text) { "ingesteld op '$Text'" } else { "leeggemaakt" }
    Update-OvSummary "$($r.Set) cel(len) $what$(Get-OvLockedText $r.Locked)"
}

function Invoke-OvPaste {
    # Plakken uit Excel/klembord: 1 waarde = in alle geselecteerde cellen,
    # een blok (tabs/regels) = vanaf de aangeklikte cel naar rechts/onder.
    param([string]$Text)
    if (-not $PSBoundParameters.ContainsKey('Text')) {
        try { $Text = [System.Windows.Clipboard]::GetText() }
        catch { Show-OvMsg "Het klembord is bezet door een ander programma, probeer het opnieuw." | Out-Null; return }
    }
    $text = $Text
    if (-not $text) { return }
    $lines = @($text -split "`r?`n")
    if ($lines.Count -gt 1 -and $lines[-1] -eq '') { $lines = @($lines[0..($lines.Count - 2)]) }
    if ($lines.Count -eq 1 -and $lines[0] -notmatch "`t") { Invoke-OvSetSelected $lines[0].Trim(); return }

    $grid = $script:Ov.Grid
    $cc = $grid.CurrentCell
    if ($null -eq $cc.Column -or $cc.Item -isnot [System.Data.DataRowView]) { Show-OvMsg "Klik eerst de cel aan waar het plakken moet beginnen." | Out-Null; return }
    $start = $cc.Column.DisplayIndex
    $cols = @($grid.Columns | Where-Object { $_.Visibility -eq 'Visible' -and $_.DisplayIndex -ge $start } | Sort-Object DisplayIndex)
    $r0 = $grid.Items.IndexOf($cc.Item)
    # Rijen vooraf vastleggen: door filter/sortering kan de volgorde verschuiven
    $rows = @()
    for ($i = 0; $i -lt $lines.Count -and ($r0 + $i) -lt $grid.Items.Count; $i++) { $rows += $grid.Items[$r0 + $i].Row }

    $cells = New-Object System.Collections.ArrayList
    $outside = 0
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $vals = $lines[$i] -split "`t"
        for ($j = 0; $j -lt $vals.Count; $j++) {
            $p = if ($j -lt $cols.Count) { [string]$cols[$j].SortMemberPath } else { '' }
            if (-not $script:Ov.Props.Contains($p)) { $outside++; continue }
            [void]$cells.Add([pscustomobject]@{ Row = $rows[$i]; Prop = $p; Text = $vals[$j].Trim() })
        }
    }
    $r = Set-OvCells $cells
    $extra = if ($outside) { "; $outside waarde(n) buiten de bewerkbare kolommen genegeerd" } else { "" }
    Update-OvSummary "Geplakt: $($r.Set) cel(len)$(Get-OvLockedText $r.Locked)$extra"
}

function Update-OvFilter {
    $parts = @()
    $t = $script:Ov.TxtFilter.Text.Trim()
    if ($t) {
        $l = ConvertTo-LikeLiteral $t
        $cols = @('Naam', 'UPN', 'Mailbox') + @($script:Ov.Fields | ForEach-Object { $_.Prop })
        $parts += '(' + (@($cols | ForEach-Object { "[$_] LIKE '*$l*'" }) -join ' OR ') + ')'
    }
    if ($script:Ov.ValueFilter) { $parts += 'Hit = 1' }
    if ($script:Ov.ChkOnlyChanged.IsChecked) { $parts += 'Wijz > 0' }
    $script:Ov.Table.DefaultView.RowFilter = $parts -join ' AND '
    Update-OvSummary
}

function Update-OvValueList {
    # Unieke waarden (hoofdlettergevoelig) van de gekozen kolom met aantallen,
    # zodat schrijfvarianten als 'Sales', 'sales' en 'Sales ' opvallen.
    $prop = [string]$script:Ov.CmbCol.SelectedValue
    $lst = $script:Ov.LstValues
    $keep = if ($lst.SelectedItem) { [string]$lst.SelectedItem.Tag } else { $null }
    $script:Ov.Refreshing = $true
    try {
        $lst.Items.Clear()
        if (-not $prop) { return }
        $counts = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::Ordinal)
        foreach ($r in $script:Ov.Table.Rows) {
            $v = [string]$r[$prop]
            if ($counts.ContainsKey($v)) { $counts[$v] = $counts[$v] + 1 } else { $counts[$v] = 1 }
        }
        foreach ($v in @($counts.Keys | Sort-Object)) {
            $label = if ($v -eq '') { '(leeg)' } elseif ($v -cne $v.Trim()) { "'$v'" } else { $v }
            $li = New-Object System.Windows.Controls.ListBoxItem
            $li.Content = "$label   ($($counts[$v]))"
            $li.Tag = $v
            [void]$lst.Items.Add($li)
            if ($null -ne $keep -and $v -ceq $keep) { $lst.SelectedItem = $li }
        }
        $script:Ov.LblValues.Text = "$($counts.Count) verschillende waarde(n). Klik een waarde om alleen die gebruikers te tonen."
    }
    finally { $script:Ov.Refreshing = $false }
}

function Update-OvSummary {
    param([string]$Msg)
    $chg = 0; $users = 0; $bad = 0; $queued = 0
    foreach ($r in $script:Ov.Table.Rows) {
        $n = [int]$r['Wijz']
        if ($n) { $users++; $chg += $n }
        foreach ($f in $script:Ov.Fields) {
            $s = [string]$r["s_$($f.Prop)"]
            if ($s -eq 'E') { $bad++ } elseif ($s -eq 'Q') { $queued++ }
        }
    }
    $parts = @("$($script:Ov.Table.DefaultView.Count) van $($script:Ov.Table.Rows.Count) gebruikers zichtbaar")
    $parts += if ($chg) { "$chg wijziging(en) bij $users gebruiker(s) nog niet in de wachtrij (geel)" } else { "geen openstaande wijzigingen" }
    if ($bad) { $parts += "$bad ongeldig (rood)" }
    if ($queued) { $parts += "$queued in de wachtrij (blauw)" }
    $script:Ov.LblStatus.Text = $parts -join '   |   '
    if ($PSBoundParameters.ContainsKey('Msg')) { $script:Ov.LblMsg.Text = $Msg }
    $script:Ov.BtnQueue.IsEnabled = ($chg -gt 0)
    Update-OvValueList
}

function Set-OvColumnVisible {
    param([string]$Prop, [bool]$Visible)
    $script:Ov.Cols[$Prop].Visibility = if ($Visible) { 'Visible' } else { 'Collapsed' }
    $script:Ov.Checks[$Prop].IsChecked = $Visible
    $script:OvDefaultProps = @($script:Ov.Fields | Where-Object { $script:Ov.Checks[$_.Prop].IsChecked } | ForEach-Object { $_.Prop })
}

function Invoke-OvValueSelected {
    # Klik op een waarde in de lijst: alleen gebruikers met die waarde tonen.
    # Ze blijven zichtbaar na aanpassen, zodat je ziet wat je wijzigde.
    if ($script:Ov.Refreshing) { return }
    $li = $script:Ov.LstValues.SelectedItem
    if (-not $li) { return }
    $prop = [string]$script:Ov.CmbCol.SelectedValue
    $v = [string]$li.Tag
    foreach ($r in $script:Ov.Table.Rows) {
        $hit = [int]([string]$r[$prop] -ceq $v)
        if ([int]$r['Hit'] -ne $hit) { $r['Hit'] = $hit }
    }
    $script:Ov.ValueFilter = $true
    $script:Ov.TxtReplace.Text = $v
    Update-OvFilter
}

function Clear-OvValueFilter {
    $script:Ov.ValueFilter = $false
    $script:Ov.Refreshing = $true
    try { $script:Ov.LstValues.SelectedItem = $null } finally { $script:Ov.Refreshing = $false }
    Update-OvFilter
}

function Invoke-OvReplace {
    $li = $script:Ov.LstValues.SelectedItem
    if (-not $li) { Show-OvMsg "Klik eerst in de lijst links de waarde aan die je wilt vervangen." | Out-Null; return }
    $prop = [string]$script:Ov.CmbCol.SelectedValue
    $old = [string]$li.Tag
    $new = $script:Ov.TxtReplace.Text.Trim()
    if ($new -ceq $old) { Show-OvMsg "De nieuwe waarde is gelijk aan de huidige." | Out-Null; return }
    $cells = New-Object System.Collections.ArrayList
    foreach ($row in $script:Ov.Table.Rows) {
        if ([string]$row[$prop] -ceq $old) { [void]$cells.Add([pscustomobject]@{ Row = $row; Prop = $prop; Text = $new }) }
    }
    $r = Set-OvCells $cells
    $label = @($script:Ov.Fields | Where-Object { $_.Prop -eq $prop })[0].Label
    $from = if ($old) { "'$old'" } else { '(leeg)' }
    $to = if ($new) { "'$new'" } else { '(leeg)' }
    Update-OvSummary "$label`: $from vervangen door $to bij $($r.Set) gebruiker(s)$(Get-OvLockedText $r.Locked)"
}

function Invoke-OvResetAll {
    Complete-OvEdit
    $n = 0
    foreach ($r in $script:Ov.Table.Rows) {
        if ([int]$r['Wijz'] -eq 0) { continue }
        foreach ($f in $script:Ov.Fields) {
            $p = $f.Prop
            if ([string]$r[$p] -cne [string]$r["o_$p"]) { $r[$p] = $r["o_$p"] }
        }
        Update-OvRowState $r
        $n++
    }
    Update-OvSummary "Wijzigingen bij $n gebruiker(s) ongedaan gemaakt"
}

function Add-OvToQueue {
    Complete-OvEdit
    $rows = @(foreach ($r in $script:Ov.Table.Rows) { if ([int]$r['Wijz'] -gt 0) { $r } })
    if ($rows.Count -eq 0) { Show-OvMsg "Er zijn geen wijzigingen om toe te voegen." | Out-Null; return }

    $errs = New-Object System.Collections.ArrayList
    foreach ($r in $rows) {
        foreach ($f in $script:Ov.Fields) {
            if ([string]$r["s_$($f.Prop)"] -ne 'E') { continue }
            $err = $null
            [void](ConvertTo-PatchValue -Prop $f.Prop -Text ([string]$r[$f.Prop]).Trim() -ErrorMessage ([ref]$err))
            [void]$errs.Add("$($r['Naam']) - $($f.Label): $err")
        }
    }
    if ($errs.Count) {
        $list = @($errs | Select-Object -First 15) -join "`n"
        $more = if ($errs.Count -gt 15) { "`n... en nog $($errs.Count - 15)" } else { "" }
        Show-OvMsg "Er zijn $($errs.Count) ongeldige waarde(n) (rood). Pas die eerst aan:`n`n$list$more" "Ongeldige invoer" "OK" "Warning" | Out-Null
        return
    }

    $n = 0
    foreach ($r in $rows) {
        $id = [string]$r['Id']
        $body = [ordered]@{}
        $old = @{}
        foreach ($f in $script:Ov.Fields) {
            $p = $f.Prop
            if ([string]$r["s_$p"] -ne 'C') { continue }
            $err = $null
            $body[$p] = ConvertTo-PatchValue -Prop $p -Text ([string]$r[$p]).Trim() -ErrorMessage ([ref]$err)
            $old[$p] = [string]$r["o_$p"]
        }
        if ($body.Count -eq 0) { continue }
        if (Add-QueueItem -Kind 'GraphPatch' -Target ([string]$r['Naam']) -Soort 'Entra (overzicht)' -Text (Format-BodyText $body $old) -Data @{ UserId = $id; Body = $body }) { $n++ }
        # Wat nu in de wachtrij staat wordt de nieuwe uitgangswaarde (blauw)
        foreach ($p in @($body.Keys)) {
            $r[$p] = ([string]$r[$p]).Trim()
            $r["o_$p"] = $r[$p]
            $script:Ov.Queued["$id|$p"] = $true
        }
        Update-OvRowState $r
    }
    Update-QueueSummary
    Write-Log "Overzicht: wijzigingen voor $n gebruiker(s) toegevoegd aan wachtrij" "info"
    Set-Status "Overzicht: $n gebruiker(s) toegevoegd aan de wachtrij" "success"
    Update-OvSummary "$n gebruiker(s) in de wachtrij gezet (blauw)."
    $a = Show-OvMsg "Wijzigingen voor $n gebruiker(s) staan in de wachtrij.`n`nVenster sluiten en de wachtrij nu uitvoeren?`n(Nee = verder bewerken in dit overzicht)" "Wachtrij" "YesNo" "Question"
    if ($a -eq 'Yes') { $script:Ov.RunQueue = $true; $script:Ov.Win.Close() }
}

function Show-OverviewDialog {
    if (-not $script:GraphConnected) { Show-Msg "Verbind eerst met Microsoft Graph." "Overzicht" "OK" "Warning" | Out-Null; return }
    if (-not $script:UserTable -or $script:UserTable.DefaultView.Count -eq 0) { Show-Msg "Laad eerst gebruikers in de lijst links." "Overzicht" "OK" "Warning" | Out-Null; return }
    if (-not (Confirm-LeaveUser)) { return }

    # Meerdere geselecteerd = alleen die; anders alle zichtbare gebruikers
    $picked = @($script:UserGrid.SelectedItems)
    $ids = if ($picked.Count -gt 1) { @($picked | ForEach-Object { [string]$_.Row['Id'] }) } else { @($script:UserTable.DefaultView | ForEach-Object { [string]$_.Row['Id'] }) }

    $users = $null
    Start-LongOp "Overzicht laden: $($ids.Count) gebruiker(s)..."
    try { $users = Get-OvUsers $ids }
    catch { Show-Msg "Gegevens ophalen mislukt:`n$_" "Fout" "OK" "Error" | Out-Null }
    finally { Stop-LongOp }
    if ($null -eq $users) { Set-Status "Overzicht laden mislukt" "error"; return }
    if ($script:StopRequested) { Set-Status "Overzicht laden gestopt"; return }
    if ($users.Count -eq 0) { Show-Msg "Er konden geen gebruikers worden opgehaald." "Overzicht" "OK" "Warning" | Out-Null; return }
    Set-Status "Overzicht: $($users.Count) gebruiker(s) geladen" "success"

    $w = ConvertFrom-Xaml ($script:OverviewXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $w.Owner = $script:Window
    $w.Title = "Overzicht bewerken - Entra eigenschappen ($($users.Count) gebruikers)"
    $fields = Get-OvFields
    $script:Ov = @{
        Win = $w; Fields = $fields
        Props = [System.Collections.Generic.HashSet[string]]::new([string[]]@($fields | ForEach-Object { $_.Prop }))
        Queued = @{}; Cols = @{}; Checks = @{}
        ValueFilter = $false; Refreshing = $false; Committing = $false; EditRow = $null; RunQueue = $false
    }
    foreach ($n in 'Grid', 'ColPanel', 'TxtFilter', 'ChkOnlyChanged', 'CmbCol', 'LblValues', 'LstValues', 'TxtReplace', 'BtnReplace', 'BtnShowAll',
                   'TxtSet', 'BtnSet', 'BtnClearSel', 'BtnRevertSel', 'BtnResetAll', 'BtnQueue', 'BtnClose', 'LblStatus', 'LblMsg') {
        $ctl = $w.FindName("OV_$n")
        if ($null -eq $ctl) { throw "XAML-element 'OV_$n' niet gevonden" }
        $script:Ov[$n] = $ctl
    }
    $script:Ov.Table = New-OvTable $users

    # Kolommen: vast (alleen-lezen) + per veld een bewerkbare kolom met celkleur
    $grid = $script:Ov.Grid
    $cellText = $w.FindResource('CellText')
    foreach ($c in @(@{ H = 'Naam'; P = 'Naam'; W = 190 }, @{ H = 'UPN'; P = 'UPN'; W = 230 }, @{ H = 'Bron'; P = 'Bron'; W = 55 }, @{ H = 'Mailbox'; P = 'Mailbox'; W = 80 }, @{ H = 'Wijz.'; P = 'Wijz'; W = 45 })) {
        $col = New-Object System.Windows.Controls.DataGridTextColumn
        $col.Header = $c.H
        $col.Binding = [System.Windows.Data.Binding]::new($c.P)
        $col.Width = [System.Windows.Controls.DataGridLength]::new($c.W)
        $col.IsReadOnly = $true
        $col.ElementStyle = $cellText
        $grid.Columns.Add($col)
    }
    foreach ($f in $fields) {
        $p = $f.Prop
        $show = ($script:OvDefaultProps -contains $p)
        $col = New-Object System.Windows.Controls.DataGridTextColumn
        $col.Header = $f.Label
        $col.Binding = [System.Windows.Data.Binding]::new($p)
        $col.Width = [System.Windows.Controls.DataGridLength]::new(150)
        $col.ElementStyle = $cellText
        $col.CellStyle = [System.Windows.Markup.XamlReader]::Parse($script:OvCellStyleXaml.Replace('{{P}}', $p))
        $col.Visibility = if ($show) { 'Visible' } else { 'Collapsed' }
        $grid.Columns.Add($col)
        $script:Ov.Cols[$p] = $col

        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $f.Label
        $cb.Tag = $p
        $cb.IsChecked = $show
        $cb.Margin = [System.Windows.Thickness]::new(0, 0, 14, 4)
        $cb.Add_Click({ Set-OvColumnVisible ([string]$this.Tag) ([bool]$this.IsChecked) })
        [void]$script:Ov.ColPanel.Children.Add($cb)
        $script:Ov.Checks[$p] = $cb
    }
    $grid.ItemsSource = $script:Ov.Table.DefaultView
    $script:Ov.CmbCol.ItemsSource = $fields
    $script:Ov.CmbCol.SelectedValue = 'department'

    # Bewerken in de tabel
    $grid.Add_BeginningEdit({
        $row = $_.Row.Item.Row
        if (-not (Test-OvEditable $row ([string]$_.Column.SortMemberPath))) {
            $_.Cancel = $true
            $script:Ov.LblMsg.Text = "$($row['Naam']) wordt gesynchroniseerd vanuit on-premises AD; pas dit veld aan in Active Directory."
        }
    })
    $grid.Add_CellEditEnding({
        # De waarde staat pas na het afronden van de rij in de DataTable; dat
        # gebeurt direct hierna via de Dispatcher, daarna kleur/tellers bijwerken.
        if ($_.EditAction -ne 'Commit' -or $script:Ov.Committing) { return }
        $script:Ov.Committing = $true
        $script:Ov.EditRow = $_.Row.Item
        [void]$script:Ov.Grid.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{
            if ($null -eq $script:Ov) { return }
            try { [void]$script:Ov.Grid.CommitEdit('Row', $true) } catch { }
            $script:Ov.Committing = $false
            if ($script:Ov.EditRow -is [System.Data.DataRowView]) { Update-OvRowState $script:Ov.EditRow.Row; Update-OvSummary }
        })
    })
    $grid.Add_PreviewKeyDown({
        if ($_.OriginalSource -is [System.Windows.Controls.TextBox]) { return }   # cel in bewerkmodus
        $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
        if ($ctrl -and $_.Key -eq 'V') { $_.Handled = $true; Invoke-OvPaste }
        elseif ($_.Key -eq 'Delete') { $_.Handled = $true; Invoke-OvSetSelected '' }
    })

    # Filter, waardenlijst en knoppen
    $script:Ov.TxtFilter.Add_TextChanged({ Update-OvFilter })
    $script:Ov.ChkOnlyChanged.Add_Click({ Update-OvFilter })
    $script:Ov.CmbCol.Add_SelectionChanged({
        $p = [string]$script:Ov.CmbCol.SelectedValue
        if ($p) { Set-OvColumnVisible $p $true }
        $script:Ov.TxtReplace.Text = ''
        Clear-OvValueFilter
    })
    $script:Ov.LstValues.Add_SelectionChanged({ Invoke-OvValueSelected })
    $script:Ov.BtnReplace.Add_Click({ Invoke-OvReplace })
    $script:Ov.TxtReplace.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-OvReplace } })
    $script:Ov.BtnShowAll.Add_Click({ Clear-OvValueFilter })
    $script:Ov.BtnSet.Add_Click({ Invoke-OvSetSelected $script:Ov.TxtSet.Text.Trim() })
    $script:Ov.TxtSet.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-OvSetSelected $script:Ov.TxtSet.Text.Trim() } })
    $script:Ov.BtnClearSel.Add_Click({ Invoke-OvSetSelected '' })
    $script:Ov.BtnRevertSel.Add_Click({ Invoke-OvSetSelected -Revert })
    $script:Ov.BtnResetAll.Add_Click({ Invoke-OvResetAll })
    $script:Ov.BtnQueue.Add_Click({ Add-OvToQueue })
    $script:Ov.BtnClose.Add_Click({ $script:Ov.Win.Close() })

    $w.Add_Closing({
        try { Complete-OvEdit } catch { }
        if ($script:Ov.RunQueue) { return }
        $n = @(foreach ($r in $script:Ov.Table.Rows) { if ([int]$r['Wijz'] -gt 0) { $r } }).Count
        if ($n) {
            $a = Show-OvMsg "Er zijn wijzigingen bij $n gebruiker(s) die nog niet in de wachtrij staan.`n`nToch sluiten? Deze wijzigingen gaan verloren." "Sluiten" "YesNo" "Warning"
            if ($a -ne 'Yes') { $_.Cancel = $true }
        }
    })
    $w.Add_ContentRendered({ [void]$script:Ov.Grid.Focus() })

    Update-OvSummary "Tip: kies links een kolom (bijv. Afdeling) om alle waarden en schrijfvarianten te zien."
    [void]$w.ShowDialog()

    $run = $script:Ov.RunQueue
    $script:Ov = $null
    if ($script:CurrentUser) { Set-EntraForm }
    if ($run) { Invoke-Queue }
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

function Get-FieldGroupXaml {
    param([string]$Group)
    $fields = @($script:EntraFields | Where-Object { $_.Group -eq $Group })
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<GroupBox Header=`"$Group`"><Grid>")
    [void]$sb.AppendLine('<Grid.ColumnDefinitions><ColumnDefinition Width="140"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>')
    [void]$sb.Append('<Grid.RowDefinitions>')
    foreach ($f in $fields) { [void]$sb.Append('<RowDefinition Height="Auto"/>') }
    [void]$sb.AppendLine('</Grid.RowDefinitions>')
    $i = 0
    foreach ($f in $fields) {
        [void]$sb.AppendLine("<TextBlock Grid.Row=`"$i`" Text=`"$($f.Label)`" Style=`"{StaticResource FieldLabel}`"/>")
        [void]$sb.AppendLine("<TextBox x:Name=`"F_$($f.Prop)`" Grid.Row=`"$i`" Grid.Column=`"1`" Tag=`"$($f.Prop)`" Style=`"{StaticResource FieldBox}`"/>")
        $i++
    }
    [void]$sb.AppendLine('</Grid></GroupBox>')
    return $sb.ToString()
}

function Get-BulkFieldsXaml {
    $fields = @($script:EntraFields | Where-Object { $_.Bulk })
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<Grid>')
    [void]$sb.AppendLine('<Grid.ColumnDefinitions><ColumnDefinition Width="28"/><ColumnDefinition Width="140"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>')
    [void]$sb.Append('<Grid.RowDefinitions>')
    foreach ($f in $fields) { [void]$sb.Append('<RowDefinition Height="Auto"/>') }
    [void]$sb.AppendLine('</Grid.RowDefinitions>')
    $i = 0
    foreach ($f in $fields) {
        [void]$sb.AppendLine("<CheckBox x:Name=`"BC_$($f.Prop)`" Grid.Row=`"$i`" VerticalAlignment=`"Center`" Margin=`"0,0,0,6`" ToolTip=`"Aanvinken = dit veld aanpassen`"/>")
        [void]$sb.AppendLine("<TextBlock Grid.Row=`"$i`" Grid.Column=`"1`" Text=`"$($f.Label)`" Style=`"{StaticResource FieldLabel}`"/>")
        [void]$sb.AppendLine("<TextBox x:Name=`"BT_$($f.Prop)`" Grid.Row=`"$i`" Grid.Column=`"2`" Tag=`"$($f.Prop)`" Style=`"{StaticResource FieldBox}`"/>")
        $i++
    }
    [void]$sb.AppendLine('</Grid>')
    return $sb.ToString()
}

$script:FolderComboItems = @'
<ComboBoxItem Content="Agenda" Tag="Calendar" IsSelected="True"/>
<ComboBoxItem Content="Contactpersonen" Tag="Contacts"/>
<ComboBoxItem Content="Postvak IN" Tag="Inbox"/>
'@

$script:QueueColXaml = '<DataGridTextColumn Header="Wachtrij" Binding="{Binding Wachtrij}" Width="170" ElementStyle="{StaticResource CellQueue}"/>'

$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Entra &amp; Exchange Gebruikersbeheer" Width="1500" Height="960" MinWidth="1200" MinHeight="720"
        WindowStartupLocation="CenterScreen" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="*" MinHeight="380"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="230" MinHeight="90"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0078D4">
      <DockPanel Margin="16,0,12,0">
        <StackPanel x:Name="HeaderPanel" DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Tenant (optioneel, GDAP):" Foreground="White" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="TxtTenant" Width="190" Height="26" VerticalContentAlignment="Center"
                   ToolTip="Leeg = eigen tenant. Voor een klant via GDAP: domein (klant.onmicrosoft.com) of tenant-ID."/>
          <StackPanel Margin="14,0,0,0" VerticalAlignment="Center">
            <TextBlock x:Name="LblGraph" Text="Graph: niet verbonden" Foreground="#FFC864"/>
            <TextBlock x:Name="LblExo" Text="Exchange: niet verbonden" Foreground="#FFC864"/>
          </StackPanel>
          <Button x:Name="BtnConnectGraph" Content="Verbinden met Graph" Style="{StaticResource BtnDark}"/>
          <Button x:Name="BtnConnectExo" Content="Verbinden met Exchange" Style="{StaticResource BtnDark}"/>
        </StackPanel>
        <TextBlock Text="Entra &amp; Exchange Gebruikersbeheer" Foreground="White" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </DockPanel>
    </Border>

    <Grid x:Name="MainArea" Grid.Row="1" Margin="6,6,6,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="5*" MinWidth="420"/>
        <ColumnDefinition Width="5"/>
        <ColumnDefinition Width="6*" MinWidth="640"/>
      </Grid.ColumnDefinitions>

      <DockPanel x:Name="UserPanel" Grid.Column="0">
        <Border DockPanel.Dock="Top" Background="White" Padding="8,8,8,2" BorderBrush="#E0E0E0" BorderThickness="1,1,1,0">
          <StackPanel>
            <DockPanel>
              <Button x:Name="BtnSearch" DockPanel.Dock="Right" Content="Laden" Margin="8,0,0,6"/>
              <TextBlock Text="Zoeken in Entra:" VerticalAlignment="Center" Margin="0,0,6,6" Width="95"/>
              <TextBox x:Name="TxtSearch" Height="28" VerticalContentAlignment="Center" Margin="0,0,0,6"
                       ToolTip="Zoekt op naam, UPN en e-mail (begint met). Leeg = ALLE gebruikers laden."/>
            </DockPanel>
            <DockPanel>
              <TextBlock Text="Filter in lijst:" VerticalAlignment="Center" Margin="0,0,6,6" Width="95"/>
              <TextBox x:Name="TxtFilter" Height="28" VerticalContentAlignment="Center" Margin="0,0,0,6"
                       ToolTip="Filtert direct op naam, UPN, afdeling, functie en bedrijf"/>
            </DockPanel>
            <WrapPanel>
              <CheckBox x:Name="ChkHideGuests"  Content="Gasten verbergen" IsChecked="True" VerticalAlignment="Center" Margin="0,0,14,6"/>
              <CheckBox x:Name="ChkOnlyEnabled" Content="Alleen actieve accounts" VerticalAlignment="Center" Margin="0,0,14,6"
                        ToolTip="Gedeelde, ruimte- en apparatuurmailboxen blijven zichtbaar (hun account is bewust uitgeschakeld)"/>
              <CheckBox x:Name="ChkOnlyFwd" Content="Alleen met doorsturen" VerticalAlignment="Center" Margin="0,0,14,6"
                        ToolTip="Kolom Doorsturen wordt gevuld zodra er met Exchange Online verbonden is"/>
              <TextBlock Text="Mailbox:" VerticalAlignment="Center" Margin="0,0,6,6"/>
              <ComboBox x:Name="CmbMbxType" Width="170" Margin="0,0,14,6"
                        ToolTip="Kolom Mailbox wordt gevuld zodra er met Exchange Online verbonden is">
                <ComboBoxItem Content="Alle" Tag="" IsSelected="True"/>
                <ComboBoxItem Content="Gebruikersmailboxen" Tag="User"/>
                <ComboBoxItem Content="Gedeelde mailboxen" Tag="Shared"/>
                <ComboBoxItem Content="Ruimte / apparatuur" Tag="Resource"/>
                <ComboBoxItem Content="Zonder mailbox" Tag="None"/>
              </ComboBox>
              <Button x:Name="BtnSelectAll" Content="Alles selecteren" Style="{StaticResource BtnGray}" Padding="10,4"/>
              <Button x:Name="BtnOverview" Content="Overzicht bewerken..." Style="{StaticResource BtnGreen}" Padding="10,4"
                      ToolTip="Tabel met Entra-velden (afdeling, functie, ...) van de geselecteerde gebruikers, of van alle zichtbare als er 0 of 1 geselecteerd is. Per cel aanpassen en in 1 keer naar de wachtrij."/>
              <Button x:Name="BtnCsv" Content="CSV import / export..." Padding="10,4"/>
              <Button x:Name="BtnFwdOverview" Content="Doorsturen-overzicht..." Padding="10,4"
                      ToolTip="Alle mailboxen in de tenant waarop doorsturen is ingesteld (intern en extern), met export en uitzetten voor selectie."/>
              <Button x:Name="BtnRcptOverview" Content="Alle ontvangers..." Padding="10,4"
                      ToolTip="Alles met een mailadres: mailboxen, Microsoft 365-groepen, distributielijsten, beveiligingsgroepen, contactpersonen en gasten. Met aliassen, leden per groep en export."/>
              <Button x:Name="BtnNewGroup" Content="Nieuwe groep / lijst..." Style="{StaticResource BtnGreen}" Padding="10,4"
                      ToolTip="Distributielijst, beveiligingsgroep met e-mail, Microsoft 365-groep, dynamische distributielijst of ruimtelijst aanmaken (via de wachtrij). Vereist Exchange-verbinding."/>
            </WrapPanel>
            <TextBlock x:Name="LblUserCount" Foreground="Gray" Margin="0,0,0,6"/>
          </StackPanel>
        </Border>
        <DataGrid x:Name="UserGrid" SelectionMode="Extended">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Naam"     Binding="{Binding Naam}"     Width="3*" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="UPN"      Binding="{Binding UPN}"      Width="3*" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Afdeling" Binding="{Binding Afdeling}" Width="2*" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Functie"  Binding="{Binding Functie}"  Width="2*" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Status"   Binding="{Binding Status}"   Width="85" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Mailbox"  Binding="{Binding Mailbox}"  Width="85" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Doorsturen" Binding="{Binding Doorsturen}" Width="2*" ElementStyle="{StaticResource CellQueue}"/>
          </DataGrid.Columns>
        </DataGrid>
      </DockPanel>

      <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#D0D0D0" ResizeDirection="Columns"/>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Background="White" Padding="12,8" BorderBrush="#E0E0E0" BorderThickness="1,1,1,0">
          <StackPanel>
            <TextBlock x:Name="LblUser" Text="Selecteer links een of meer gebruikers" FontSize="16" FontWeight="SemiBold"/>
            <TextBlock x:Name="LblUserSub" Foreground="Gray" TextWrapping="Wrap" Text="Tip: Ctrl/Shift-klik of 'Alles selecteren' voor bulk-aanpassingen"/>
            <TextBlock x:Name="LblUserQueue" Foreground="#0063B1" FontWeight="SemiBold"/>
          </StackPanel>
        </Border>
        <Border x:Name="SyncWarn" Grid.Row="1" Background="#FFF4CE" Padding="12,6" Visibility="Collapsed" BorderBrush="#E0C060" BorderThickness="1">
          <TextBlock TextWrapping="Wrap" Text="Deze gebruiker wordt gesynchroniseerd vanuit on-premises Active Directory. De meeste Entra-eigenschappen moeten in AD worden aangepast; Graph weigert die wijzigingen. Exchange-rechten werken wel gewoon."/>
        </Border>

        <!-- ===== 1 gebruiker ===== -->
        <TabControl x:Name="DetailTabs" Grid.Row="2" IsEnabled="False" Padding="0">

          <TabItem x:Name="TabEntra" Header="Entra eigenschappen" Padding="14,6">
            <DockPanel>
              <Border DockPanel.Dock="Bottom" Background="White" Padding="10,8,10,2" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0">
                <DockPanel LastChildFill="False">
                  <Button x:Name="BtnEntraSave"  Content="Toevoegen aan wachtrij" Style="{StaticResource BtnGreen}"/>
                  <Button x:Name="BtnEntraReset" Content="Ongedaan maken"         Style="{StaticResource BtnGray}"/>
                  <TextBlock x:Name="LblEntraDirty" Foreground="#B05A00" FontWeight="SemiBold" VerticalAlignment="Center" Margin="6,0,0,6"/>
                </DockPanel>
              </Border>
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <Grid Margin="8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0" Margin="0,0,6,0">
{{G_Identiteit}}
{{G_Organisatie}}
{{G_Contact}}
                  </StackPanel>
                  <StackPanel Grid.Column="1" Margin="6,0,0,0">
{{G_Adres}}
                    <GroupBox Header="Account">
                      <StackPanel>
                        <CheckBox x:Name="ChkAccountEnabled" Content="Account ingeschakeld (aanmelden toegestaan)" Margin="0,2,0,8"/>
                        <DockPanel>
                          <TextBlock Text="Manager" Width="140" Style="{StaticResource FieldLabel}"/>
                          <Button x:Name="BtnMgrClear" DockPanel.Dock="Right" Content="Wissen" Style="{StaticResource BtnGray}" Margin="6,0,0,6" Padding="10,4"/>
                          <Button x:Name="BtnMgrPick"  DockPanel.Dock="Right" Content="Kiezen..." Margin="6,0,0,6" Padding="10,4"/>
                          <TextBox x:Name="TxtManager" IsReadOnly="True" Background="#F7F7F7" Style="{StaticResource FieldBox}"/>
                        </DockPanel>
                      </StackPanel>
                    </GroupBox>
                    <GroupBox Header="Info (alleen lezen)">
                      <TextBox x:Name="TxtInfo" IsReadOnly="True" TextWrapping="Wrap" BorderThickness="0" FontFamily="Consolas" Background="White"/>
                    </GroupBox>
                    <TextBlock Style="{StaticResource Hint}" Text="Geel = gewijzigd, nog niet in de wachtrij. Blauw = staat al in de wachtrij (tooltip toont de huidige waarde in Entra)."/>
                  </StackPanel>
                </Grid>
              </ScrollViewer>
            </DockPanel>
          </TabItem>

          <TabItem x:Name="TabMbx" Header="Mailbox rechten" Padding="14,6">
            <Grid Margin="8">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <DockPanel LastChildFill="False">
                <Button x:Name="BtnMbxRefresh" Content="Vernieuwen"/>
                <TextBlock x:Name="LblMbxInfo" VerticalAlignment="Center" Margin="4,0,0,6" FontWeight="SemiBold"/>
              </DockPanel>
              <Grid x:Name="MbxContent" Grid.Row="1" IsEnabled="False">
                <Grid.RowDefinitions>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <GroupBox Grid.Row="0">
                  <GroupBox.Header><TextBlock Text="Volledige toegang (FullAccess) - wie mag deze mailbox openen" FontWeight="SemiBold"/></GroupBox.Header>
                  <DockPanel>
                    <WrapPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
                      <Button x:Name="BtnFullAdd" Content="+ Toevoegen" Style="{StaticResource BtnGreen}"/>
                      <CheckBox x:Name="ChkAutoMap" Content="Automapping (mailbox verschijnt vanzelf in Outlook)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,16,6"/>
                      <Button x:Name="BtnFullRemove" Content="Verwijderen" Style="{StaticResource BtnRed}"/>
                    </WrapPanel>
                    <DataGrid x:Name="MbxFullGrid">
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Gebruiker" Binding="{Binding Gebruiker}" Width="*"   ElementStyle="{StaticResource CellText}"/>
                        <DataGridTextColumn Header="Rechten"   Binding="{Binding Rechten}"   Width="160" ElementStyle="{StaticResource CellText}"/>
                        {{QCOL}}
                      </DataGrid.Columns>
                    </DataGrid>
                  </DockPanel>
                </GroupBox>
                <GroupBox Grid.Row="1">
                  <GroupBox.Header><TextBlock Text="Verzenden als (SendAs)" FontWeight="SemiBold"/></GroupBox.Header>
                  <DockPanel>
                    <WrapPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
                      <Button x:Name="BtnSendAsAdd"    Content="+ Toevoegen" Style="{StaticResource BtnGreen}"/>
                      <Button x:Name="BtnSendAsRemove" Content="Verwijderen" Style="{StaticResource BtnRed}"/>
                    </WrapPanel>
                    <DataGrid x:Name="MbxSendAsGrid">
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Gebruiker" Binding="{Binding Gebruiker}" Width="*" ElementStyle="{StaticResource CellText}"/>
                        {{QCOL}}
                      </DataGrid.Columns>
                    </DataGrid>
                  </DockPanel>
                </GroupBox>
                <GroupBox Grid.Row="2" Margin="0">
                  <GroupBox.Header><TextBlock Text="Verzenden namens (SendOnBehalf)" FontWeight="SemiBold"/></GroupBox.Header>
                  <DockPanel>
                    <WrapPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
                      <Button x:Name="BtnSobAdd"    Content="+ Toevoegen" Style="{StaticResource BtnGreen}"/>
                      <Button x:Name="BtnSobRemove" Content="Verwijderen" Style="{StaticResource BtnRed}"/>
                    </WrapPanel>
                    <DataGrid x:Name="MbxSobGrid">
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Gebruiker" Binding="{Binding Gebruiker}" Width="*" ElementStyle="{StaticResource CellText}"/>
                        {{QCOL}}
                      </DataGrid.Columns>
                    </DataGrid>
                  </DockPanel>
                </GroupBox>
              </Grid>
            </Grid>
          </TabItem>

          <TabItem x:Name="TabFwd" Header="Doorsturen" Padding="14,6">
            <Grid Margin="8">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <DockPanel LastChildFill="False">
                <Button x:Name="BtnFwdRefresh" Content="Vernieuwen"/>
                <TextBlock x:Name="LblFwdInfo" VerticalAlignment="Center" Margin="4,0,0,6" FontWeight="SemiBold"/>
              </DockPanel>
              <DockPanel x:Name="FwdContent" Grid.Row="1" IsEnabled="False">
                <StackPanel DockPanel.Dock="Top">
                  <GroupBox>
                    <GroupBox.Header><TextBlock Text="Huidige instelling" FontWeight="SemiBold"/></GroupBox.Header>
                    <StackPanel>
                      <Border x:Name="FwdStateBorder" Padding="10,8" CornerRadius="4" Background="#F0F0F0" Margin="0,0,0,4">
                        <StackPanel>
                          <TextBlock x:Name="LblFwdState" FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap"/>
                          <TextBlock x:Name="LblFwdDetail" TextWrapping="Wrap" Foreground="#333333"/>
                        </StackPanel>
                      </Border>
                      <TextBlock x:Name="LblFwdQueue" Foreground="#0063B1" FontWeight="SemiBold" TextWrapping="Wrap"/>
                    </StackPanel>
                  </GroupBox>
                  <GroupBox>
                    <GroupBox.Header><TextBlock Text="Wijzigen" FontWeight="SemiBold"/></GroupBox.Header>
                    <StackPanel>
                      <RadioButton x:Name="RbFwdOff" GroupName="Fwd" Content="Niet doorsturen"/>
                      <RadioButton x:Name="RbFwdInt" GroupName="Fwd" Content="Doorsturen naar een interne ontvanger (gebruiker, gedeelde mailbox, groep of contactpersoon)"/>
                      <DockPanel Margin="20,0,0,0">
                        <Button x:Name="BtnFwdPick" DockPanel.Dock="Right" Content="Kiezen..." Margin="6,0,0,6" Padding="10,4"/>
                        <TextBox x:Name="TxtFwdInt" IsReadOnly="True" Background="#F7F7F7" Style="{StaticResource FieldBox}"/>
                      </DockPanel>
                      <RadioButton x:Name="RbFwdExt" GroupName="Fwd" Content="Doorsturen naar een extern e-mailadres"/>
                      <TextBox x:Name="TxtFwdExt" Style="{StaticResource FieldBox}" Margin="20,0,0,6"/>
                      <CheckBox x:Name="ChkFwdKeep" Content="Kopie bewaren in deze mailbox (uit = mail gaat alleen naar het doorstuuradres)" IsChecked="True" Margin="0,4,0,8"/>
                      <WrapPanel>
                        <Button x:Name="BtnFwdSave"  Content="Toevoegen aan wachtrij" Style="{StaticResource BtnGreen}"/>
                        <Button x:Name="BtnFwdReset" Content="Ongedaan maken"         Style="{StaticResource BtnGray}"/>
                      </WrapPanel>
                      <TextBlock x:Name="LblFwdPolicy" Style="{StaticResource Hint}" Margin="0"/>
                    </StackPanel>
                  </GroupBox>
                </StackPanel>
                <GroupBox Margin="0">
                  <GroupBox.Header><TextBlock Text="Inbox-regels van de gebruiker die doorsturen (alleen lezen)" FontWeight="SemiBold"/></GroupBox.Header>
                  <DockPanel>
                    <WrapPanel DockPanel.Dock="Top">
                      <Button x:Name="BtnFwdRules" Content="Inbox-regels controleren"/>
                      <TextBlock x:Name="LblFwdRules" VerticalAlignment="Center" Margin="4,0,0,6" Foreground="Gray"/>
                    </WrapPanel>
                    <DataGrid x:Name="FwdRuleGrid" MinHeight="70">
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Regel"  Binding="{Binding Regel}"  Width="*"   ElementStyle="{StaticResource CellText}"/>
                        <DataGridTextColumn Header="Actief" Binding="{Binding Actief}" Width="60"  ElementStyle="{StaticResource CellText}"/>
                        <DataGridTextColumn Header="Actie"  Binding="{Binding Actie}"  Width="140" ElementStyle="{StaticResource CellText}"/>
                        <DataGridTextColumn Header="Naar"   Binding="{Binding Naar}"   Width="2*"  ElementStyle="{StaticResource CellText}"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </DockPanel>
                </GroupBox>
              </DockPanel>
            </Grid>
          </TabItem>

          <TabItem x:Name="TabFolder" Header="Map rechten (agenda)" Padding="14,6">
            <Grid Margin="8">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <WrapPanel>
                <TextBlock Text="Map:" VerticalAlignment="Center" Margin="0,0,6,6" FontWeight="SemiBold"/>
                <ComboBox x:Name="CmbFolder" Width="160" Margin="0,0,8,6">
{{FOLDER_ITEMS}}
                </ComboBox>
                <Button x:Name="BtnFolderRefresh" Content="Vernieuwen"/>
                <TextBlock x:Name="LblFolderPath" VerticalAlignment="Center" Margin="4,0,0,6" Foreground="Gray"/>
              </WrapPanel>
              <DockPanel x:Name="FolderContent" Grid.Row="1" IsEnabled="False">
                <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10,8,10,2" Margin="0,8,0,0">
                  <StackPanel>
                    <WrapPanel>
                      <TextBlock Text="Rechten:" VerticalAlignment="Center" Margin="0,0,6,6" FontWeight="SemiBold"/>
                      <ComboBox x:Name="CmbRights" Width="380" Margin="0,0,12,6" DisplayMemberPath="Label" SelectedValuePath="Value"/>
                      <CheckBox x:Name="ChkDelegate" Content="Gemachtigde (delegate) - ontvangt vergaderverzoeken" VerticalAlignment="Center" Margin="0,0,12,6"
                                ToolTip="Alleen mogelijk op de agenda in combinatie met Editor."/>
                      <CheckBox x:Name="ChkPrivate" Content="Mag priv&#233;-items zien" VerticalAlignment="Center" Margin="0,0,0,6"/>
                    </WrapPanel>
                    <WrapPanel>
                      <Button x:Name="BtnFolderAdd"    Content="+ Gebruiker toevoegen"         Style="{StaticResource BtnGreen}"/>
                      <Button x:Name="BtnFolderSet"    Content="Rechten van selectie wijzigen"/>
                      <Button x:Name="BtnFolderRemove" Content="Verwijderen"                   Style="{StaticResource BtnRed}"/>
                    </WrapPanel>
                    <TextBlock Style="{StaticResource Hint}"
                               Text="Default = iedereen in de organisatie, Anonymous = externe gebruikers. Die kunnen niet verwijderd worden, alleen gewijzigd (bijv. naar AvailabilityOnly of None). Postvak IN-rechten werken in Outlook pas als de gebruiker ook leesrechten op de bovenste map heeft; voor volledige toegang liever FullAccess gebruiken."/>
                  </StackPanel>
                </Border>
                <DataGrid x:Name="FolderGrid">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Gebruiker"   Binding="{Binding Gebruiker}"   Width="*"   ElementStyle="{StaticResource CellText}"/>
                    <DataGridTextColumn Header="Rechten"     Binding="{Binding Rechten}"     Width="180" ElementStyle="{StaticResource CellText}"/>
                    <DataGridTextColumn Header="Gemachtigde" Binding="{Binding Gemachtigde}" Width="100" ElementStyle="{StaticResource CellText}"/>
                    {{QCOL}}
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Grid>
          </TabItem>

          <TabItem x:Name="TabScan" Header="Toegang tot andere mailboxen" Padding="14,6">
            <DockPanel Margin="8">
              <StackPanel DockPanel.Dock="Top">
                <TextBlock Style="{StaticResource Hint}"
                           Text="Zoekt op welke mailboxen deze gebruiker FullAccess, SendAs of SendOnBehalf heeft. Dit controleert elke mailbox apart en kan bij grote tenants enkele minuten duren (Stop-knop rechtsonder). Agenda-/maprechten op andere mailboxen worden niet gecontroleerd."/>
                <WrapPanel>
                  <Button x:Name="BtnScan" Content="Zoeken"/>
                  <CheckBox x:Name="ChkScanShared" Content="Alleen gedeelde mailboxen (sneller)" IsChecked="True" VerticalAlignment="Center" Margin="8,0,16,6"/>
                  <CheckBox x:Name="ChkScanSob" Content="Ook SendOnBehalf controleren" IsChecked="True" VerticalAlignment="Center" Margin="0,0,0,6"/>
                </WrapPanel>
                <TextBlock x:Name="LblScan" Margin="0,0,0,6" FontWeight="SemiBold"/>
              </StackPanel>
              <WrapPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
                <Button x:Name="BtnScanRemove" Content="Geselecteerde rechten intrekken (naar wachtrij)" Style="{StaticResource BtnRed}"/>
              </WrapPanel>
              <DataGrid x:Name="ScanGrid">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Mailbox" Binding="{Binding Mailbox}" Width="*"   ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Naam"    Binding="{Binding Naam}"    Width="*"   ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Type"    Binding="{Binding Type}"    Width="120" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Recht"   Binding="{Binding Recht}"   Width="110" ElementStyle="{StaticResource CellText}"/>
                  {{QCOL}}
                </DataGrid.Columns>
              </DataGrid>
            </DockPanel>
          </TabItem>
        </TabControl>

        <!-- ===== Bulk (meerdere gebruikers) ===== -->
        <TabControl x:Name="BulkTabs" Grid.Row="2" Padding="0" Visibility="Collapsed">

          <TabItem Header="Bulk: Entra eigenschappen" Padding="14,6">
            <DockPanel>
              <Border DockPanel.Dock="Bottom" Background="White" Padding="10,8,10,2" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0">
                <WrapPanel>
                  <Button x:Name="BtnBulkEntraApply" Content="Toevoegen aan wachtrij voor selectie" Style="{StaticResource BtnGreen}"/>
                  <Button x:Name="BtnBulkEntraClear" Content="Formulier leegmaken" Style="{StaticResource BtnGray}"/>
                </WrapPanel>
              </Border>
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Margin="8">
                  <TextBlock Style="{StaticResource Hint}"
                             Text="Vink aan welke velden je wilt aanpassen; alleen aangevinkte velden worden gewijzigd (typen vinkt automatisch aan). Een aangevinkt LEEG veld wordt bij alle geselecteerde gebruikers leeggemaakt. Persoonlijke velden (naam, UPN, personeelsnummer, mobiel) pas je per gebruiker aan of in bulk via CSV import."/>
                  <GroupBox Header="Velden">
{{BULK_FIELDS}}
                  </GroupBox>
                  <GroupBox Header="Account en manager">
                    <StackPanel>
                      <DockPanel Margin="0,0,0,6">
                        <TextBlock Text="Account" Width="168" Style="{StaticResource FieldLabel}"/>
                        <ComboBox x:Name="CmbBulkAccount" Width="220" HorizontalAlignment="Left" Margin="0,0,0,6">
                          <ComboBoxItem Content="(niet wijzigen)" IsSelected="True"/>
                          <ComboBoxItem Content="Inschakelen"/>
                          <ComboBoxItem Content="Uitschakelen"/>
                        </ComboBox>
                      </DockPanel>
                      <DockPanel>
                        <CheckBox x:Name="ChkBulkMgr" Width="28" VerticalAlignment="Center" Margin="0,0,0,6" ToolTip="Aanvinken = manager aanpassen"/>
                        <TextBlock Text="Manager" Width="140" Style="{StaticResource FieldLabel}"/>
                        <Button x:Name="BtnBulkMgrClear" DockPanel.Dock="Right" Content="Leeg" Style="{StaticResource BtnGray}" Margin="6,0,0,6" Padding="10,4"/>
                        <Button x:Name="BtnBulkMgrPick"  DockPanel.Dock="Right" Content="Kiezen..." Margin="6,0,0,6" Padding="10,4"/>
                        <TextBox x:Name="TxtBulkMgr" IsReadOnly="True" Background="#F7F7F7" Style="{StaticResource FieldBox}" Text="(geen = manager verwijderen)"/>
                      </DockPanel>
                    </StackPanel>
                  </GroupBox>
                </StackPanel>
              </ScrollViewer>
            </DockPanel>
          </TabItem>

          <TabItem Header="Bulk: mailbox rechten" Padding="14,6">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <StackPanel Margin="12">
                <GroupBox Header="Richting">
                  <StackPanel>
                    <RadioButton x:Name="RbMbxA" GroupName="BMbx" IsChecked="True" Content="De geselecteerde gebruikers krijgen rechten OP een mailbox (bijv. gedeelde mailbox info@)"/>
                    <RadioButton x:Name="RbMbxB" GroupName="BMbx" Content="Iemand krijgt rechten op de mailboxen VAN alle geselecteerde gebruikers"/>
                  </StackPanel>
                </GroupBox>
                <GroupBox Header="Doel">
                  <DockPanel>
                    <TextBlock x:Name="LblBulkMbxTarget" Width="230" Style="{StaticResource FieldLabel}"/>
                    <Button x:Name="BtnBulkMbxPick" DockPanel.Dock="Right" Content="Kiezen..." Margin="6,0,0,6" Padding="10,4"/>
                    <TextBox x:Name="TxtBulkMbxTarget" IsReadOnly="True" Background="#F7F7F7" Style="{StaticResource FieldBox}"/>
                  </DockPanel>
                </GroupBox>
                <GroupBox Header="Rechten">
                  <WrapPanel>
                    <CheckBox x:Name="ChkBFull"   Content="Volledige toegang (FullAccess)" IsChecked="True" Margin="0,0,8,6"/>
                    <CheckBox x:Name="ChkBAuto"   Content="met automapping" IsChecked="True" Margin="0,0,24,6"/>
                    <CheckBox x:Name="ChkBSendAs" Content="Verzenden als (SendAs)" Margin="0,0,24,6"/>
                    <CheckBox x:Name="ChkBSob"    Content="Verzenden namens (SendOnBehalf)" Margin="0,0,0,6"/>
                  </WrapPanel>
                </GroupBox>
                <WrapPanel Margin="0,4,0,0">
                  <Button x:Name="BtnBulkMbxGrant"  Content="Toekennen (naar wachtrij)" Style="{StaticResource BtnGreen}"/>
                  <Button x:Name="BtnBulkMbxRevoke" Content="Intrekken (naar wachtrij)" Style="{StaticResource BtnRed}"/>
                </WrapPanel>
                <TextBlock Style="{StaticResource Hint}"
                           Text="Bij uitvoeren worden rechten die al bestaan (of bij intrekken al weg zijn) als 'was al zo' afgehandeld, niet als fout. Automapping wijzigen voor bestaande FullAccess: eerst intrekken, dan opnieuw toekennen."/>
              </StackPanel>
            </ScrollViewer>
          </TabItem>

          <TabItem Header="Bulk: map rechten (agenda)" Padding="14,6">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <StackPanel Margin="12">
                <GroupBox Header="Richting">
                  <StackPanel>
                    <RadioButton x:Name="RbFolA" GroupName="BFol" IsChecked="True" Content="De geselecteerde gebruikers krijgen rechten op de map van 1 mailbox (bijv. gedeelde agenda)"/>
                    <RadioButton x:Name="RbFolB" GroupName="BFol" Content="Iemand krijgt rechten op de map van ELKE geselecteerde gebruiker (bijv. secretaresse)"/>
                    <RadioButton x:Name="RbFolC" GroupName="BFol" Content="Standaardrecht (Default = iedereen in de organisatie) instellen op de map van elke geselecteerde gebruiker"/>
                  </StackPanel>
                </GroupBox>
                <GroupBox Header="Doel">
                  <DockPanel x:Name="BulkFolTargetPanel">
                    <TextBlock x:Name="LblBulkFolTarget" Width="230" Style="{StaticResource FieldLabel}"/>
                    <Button x:Name="BtnBulkFolPick" DockPanel.Dock="Right" Content="Kiezen..." Margin="6,0,0,6" Padding="10,4"/>
                    <TextBox x:Name="TxtBulkFolTarget" IsReadOnly="True" Background="#F7F7F7" Style="{StaticResource FieldBox}"/>
                  </DockPanel>
                </GroupBox>
                <GroupBox Header="Map en rechten">
                  <StackPanel>
                    <WrapPanel>
                      <TextBlock Text="Map:" VerticalAlignment="Center" Margin="0,0,6,6"/>
                      <ComboBox x:Name="CmbBulkFolder" Width="160" Margin="0,0,16,6">
{{FOLDER_ITEMS}}
                      </ComboBox>
                      <TextBlock Text="Rechten:" VerticalAlignment="Center" Margin="0,0,6,6"/>
                      <ComboBox x:Name="CmbBulkRights" Width="380" Margin="0,0,0,6" DisplayMemberPath="Label" SelectedValuePath="Value"/>
                    </WrapPanel>
                    <WrapPanel>
                      <CheckBox x:Name="ChkBulkDelegate" Content="Gemachtigde (delegate) - ontvangt vergaderverzoeken" Margin="0,0,12,6" ToolTip="Alleen bij Agenda + Editor"/>
                      <CheckBox x:Name="ChkBulkPrivate"  Content="Mag priv&#233;-items zien" Margin="0,0,0,6"/>
                    </WrapPanel>
                  </StackPanel>
                </GroupBox>
                <WrapPanel Margin="0,4,0,0">
                  <Button x:Name="BtnBulkFolGrant"  Content="Instellen (naar wachtrij)" Style="{StaticResource BtnGreen}"/>
                  <Button x:Name="BtnBulkFolRevoke" Content="Verwijderen (naar wachtrij)" Style="{StaticResource BtnRed}"/>
                </WrapPanel>
                <TextBlock Style="{StaticResource Hint}"
                           Text="Bij uitvoeren worden bestaande rechten bijgewerkt naar het gekozen niveau. Default kan niet verwijderd worden; kies daarvoor rechten None of AvailabilityOnly."/>
              </StackPanel>
            </ScrollViewer>
          </TabItem>
        </TabControl>
      </Grid>
    </Grid>

    <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="#3C3C3C" ResizeDirection="Rows"/>

    <!-- ===== Wachtrij + log ===== -->
    <Grid Grid.Row="3">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="3*" MinWidth="500"/>
        <ColumnDefinition Width="5"/>
        <ColumnDefinition Width="2*" MinWidth="250"/>
      </Grid.ColumnDefinitions>

      <DockPanel Grid.Column="0" Background="White">
        <Border DockPanel.Dock="Top" Background="#E8F1FB" BorderBrush="#C5DAF0" BorderThickness="0,0,0,1">
          <DockPanel>
            <StackPanel x:Name="QueueButtons" DockPanel.Dock="Right" Orientation="Horizontal">
              <Button x:Name="BtnQueueRemove"    Content="Selectie verwijderen" Style="{StaticResource BtnGray}" Padding="10,3" Margin="0,3,6,3"/>
              <Button x:Name="BtnQueueClearDone" Content="Gereed wissen"        Style="{StaticResource BtnGray}" Padding="10,3" Margin="0,3,6,3"/>
              <Button x:Name="BtnQueueClear"     Content="Leegmaken"            Style="{StaticResource BtnGray}" Padding="10,3" Margin="0,3,6,3"/>
              <Button x:Name="BtnRunQueue"       Content="Alles uitvoeren (0)"  Style="{StaticResource BtnGreen}" Padding="16,3" Margin="6,3,6,3" FontWeight="SemiBold"/>
            </StackPanel>
            <TextBlock Text="  Wachtrij" FontWeight="SemiBold" FontSize="13" VerticalAlignment="Center" Foreground="#004E8C"/>
            <TextBlock x:Name="LblQueue" VerticalAlignment="Center" Foreground="#004E8C"/>
          </DockPanel>
        </Border>
        <DataGrid x:Name="QueueGrid" BorderThickness="0">
          <DataGrid.Columns>
            <DataGridTextColumn Header="#"         Binding="{Binding Nr}"        Width="40"  ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Gebruiker" Binding="{Binding Doel}"      Width="*"   ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Soort"     Binding="{Binding Soort}"     Width="110" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Wijziging" Binding="{Binding Wijziging}" Width="3*"  ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Status"    Binding="{Binding Status}"    Width="*"   ElementStyle="{StaticResource CellStatus}"/>
          </DataGrid.Columns>
        </DataGrid>
      </DockPanel>

      <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#3C3C3C" ResizeDirection="Columns"/>

      <DockPanel Grid.Column="2" Background="#1E1E1E">
        <Border DockPanel.Dock="Top" Background="#2D2D2D" Height="32">
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
    </Grid>

    <Border x:Name="StatusBar" Grid.Row="4" Background="#0078D4">
      <DockPanel>
        <Button x:Name="BtnStop" DockPanel.Dock="Right" Content="Stoppen" Style="{StaticResource BtnRed}" Padding="14,2" Margin="0,3,8,3" Visibility="Collapsed"/>
        <TextBlock x:Name="LblStatus" Foreground="White" VerticalAlignment="Center" Margin="12,0" Text="Welkom - verbind eerst met Microsoft Graph (en voor rechten met Exchange)"/>
      </DockPanel>
    </Border>
  </Grid>
</Window>
'@

$script:PickerXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Zoeken" Width="560" Height="500" WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize" ShowInTaskbar="False" Background="White" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock x:Name="PK_Prompt" Margin="0,0,0,6"/>
    <DockPanel Grid.Row="1" Margin="0,0,0,8">
      <Button x:Name="PK_BtnSearch" DockPanel.Dock="Right" Content="Zoeken" Margin="8,0,0,0"/>
      <TextBox x:Name="PK_Txt" Height="28" VerticalContentAlignment="Center"/>
    </DockPanel>
    <ListBox x:Name="PK_List" Grid.Row="2"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="PK_BtnOk" Content="OK" MinWidth="90"/>
      <Button Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" MinWidth="90" Margin="0,0,0,6"/>
    </StackPanel>
  </Grid>
</Window>
'@

$script:RcptPickerXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ontvanger zoeken" Width="860" Height="600" MinWidth="560" MinHeight="380" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <DockPanel Margin="14">
    <TextBlock x:Name="RP_Prompt" DockPanel.Dock="Top" TextWrapping="Wrap" Margin="0,0,0,6"/>
    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,4">
      <Button x:Name="RP_BtnRefresh" DockPanel.Dock="Right" Content="Vernieuwen" Style="{StaticResource BtnGray}" Padding="10,4" Margin="8,0,0,6"
              ToolTip="Lijst opnieuw ophalen uit Exchange (bijv. na het aanmaken van nieuwe ontvangers)"/>
      <ComboBox x:Name="RP_CmbCat" DockPanel.Dock="Right" Width="190" Margin="8,0,0,6">
        <ComboBoxItem Content="Alle soorten" Tag="" IsSelected="True"/>
        <ComboBoxItem Content="Gebruikers" Tag="RType IN ('UserMailbox','MailUser')"/>
        <ComboBoxItem Content="Gedeelde mailboxen" Tag="RType = 'SharedMailbox'"/>
        <ComboBoxItem Content="Ruimtes en apparatuur" Tag="RType IN ('RoomMailbox','EquipmentMailbox')"/>
        <ComboBoxItem Content="Groepen en lijsten" Tag="Categorie = 'Groep'"/>
        <ComboBoxItem Content="Contactpersonen en gasten" Tag="RType IN ('MailContact','GuestMailUser')"/>
      </ComboBox>
      <TextBox x:Name="RP_Txt" Height="28" VerticalContentAlignment="Center" Margin="0,0,0,6"
               ToolTip="Zoekt live op elk deel van naam, e-mailadres en aliassen. Meerdere woorden: alle moeten voorkomen. Pijl omlaag = naar de lijst, Enter = toevoegen."/>
    </DockPanel>
    <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,6"
               Text="Typ een deel van de naam, het adres of het domein (bijv. 'jan', 'sales', '@klant.nl'). Meerdere kiezen met Ctrl/Shift-klik. Dubbelklik of Enter voegt toe."/>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button DockPanel.Dock="Right" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="8,0,0,6"/>
      <Button x:Name="RP_BtnOk" DockPanel.Dock="Right" Content="Toevoegen" Style="{StaticResource BtnGreen}" IsEnabled="False" Margin="8,0,0,6"/>
      <TextBlock x:Name="RP_LblCount" VerticalAlignment="Center" Foreground="Gray" Margin="0,0,0,6"/>
    </DockPanel>
    <DataGrid x:Name="RP_Grid" SelectionMode="Extended">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Naam"     Binding="{Binding Naam}"     Width="2*"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Adres"    Binding="{Binding Adres}"    Width="2*"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Type"     Binding="{Binding Type}"     Width="160" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Aliassen" Binding="{Binding Aliassen}" Width="*"   ElementStyle="{StaticResource CellText}"/>
      </DataGrid.Columns>
    </DataGrid>
  </DockPanel>
</Window>
'@

$script:FwdOverviewXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Doorsturen - alle mailboxen" Width="1150" Height="680" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <DockPanel Margin="14">
    <StackPanel DockPanel.Dock="Top">
      <TextBlock Style="{StaticResource Hint}"
                 Text="Alle mailboxen waarop doorsturen is ingesteld via de mailbox-instelling (ForwardingAddress / ForwardingSmtpAddress). Doorsturen via inbox-regels van gebruikers staat hier niet in; die zie je per gebruiker op het tabblad Doorsturen."/>
      <WrapPanel>
        <Button x:Name="FO_BtnLoad" Content="Vernieuwen"/>
        <CheckBox x:Name="FO_ChkExtern" Content="Alleen extern doorsturen" VerticalAlignment="Center" Margin="8,0,16,6"/>
        <TextBlock x:Name="FO_LblStatus" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,0,6"/>
      </WrapPanel>
      <TextBlock x:Name="FO_LblPolicy" Style="{StaticResource Hint}"/>
    </StackPanel>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button DockPanel.Dock="Right" Content="Sluiten" Style="{StaticResource BtnGray}" IsCancel="True" Margin="8,0,0,6"/>
      <Button x:Name="FO_BtnExport" DockPanel.Dock="Right" Content="Exporteren naar CSV..." Style="{StaticResource BtnGray}" Margin="8,0,0,6"/>
      <Button x:Name="FO_BtnOff" Content="Doorsturen uitzetten voor selectie (naar wachtrij)" Style="{StaticResource BtnRed}" HorizontalAlignment="Left"/>
    </DockPanel>
    <DataGrid x:Name="FO_Grid">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Mailbox"         Binding="{Binding Mailbox}" Width="2*"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Naam"            Binding="{Binding Naam}"    Width="2*"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Type"            Binding="{Binding Type}"    Width="110" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Doorsturen naar" Binding="{Binding Naar}"    Width="3*"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Soort"           Binding="{Binding Soort}"   Width="70"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Kopie bewaren"   Binding="{Binding Kopie}"   Width="95"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Wachtrij"        Binding="{Binding Wachtrij}" Width="110" ElementStyle="{StaticResource CellQueue}"/>
      </DataGrid.Columns>
    </DataGrid>
  </DockPanel>
</Window>
'@

$script:RcptOverviewXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ontvangers - alles met een mailadres" Width="1400" Height="760" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <DockPanel Margin="14">
    <StackPanel DockPanel.Dock="Top">
      <TextBlock Style="{StaticResource Hint}"
                 Text="Alle ontvangers in Exchange Online: gebruikers-, gedeelde en resource-mailboxen, Microsoft 365-groepen, distributielijsten (ook dynamisch), mail-enabled beveiligingsgroepen, contactpersonen, mailgebruikers en gasten. Dubbelklik op een groep voor de leden (rechts toevoegen/verwijderen), op een gebruiker/mailbox om hem in de gebruikerslijst te openen. Alle wijzigingen gaan via de wachtrij."/>
      <WrapPanel>
        <Button x:Name="RO_BtnLoad" Content="Vernieuwen"/>
        <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,6"/>
        <TextBox x:Name="RO_TxtFilter" Width="240" Height="26" VerticalContentAlignment="Center" Margin="0,0,14,6"
                 ToolTip="Filtert op naam, adres, aliassen, extern adres, type en beheerder"/>
        <TextBlock Text="Soort:" VerticalAlignment="Center" Margin="0,0,6,6"/>
        <ComboBox x:Name="RO_CmbCat" Width="200" Margin="0,0,14,6">
          <ComboBoxItem Content="Alles" Tag="" IsSelected="True"/>
          <ComboBoxItem Content="Mailboxen" Tag="Mailbox"/>
          <ComboBoxItem Content="Groepen en distributielijsten" Tag="Groep"/>
          <ComboBoxItem Content="Contacten, mailgebruikers, gasten" Tag="Contact"/>
          <ComboBoxItem Content="Overig" Tag="Overig"/>
        </ComboBox>
        <CheckBox x:Name="RO_ChkHidden" Content="Alleen verborgen in adresboek" VerticalAlignment="Center" Margin="0,0,16,6"/>
      </WrapPanel>
      <TextBlock x:Name="RO_LblStatus" FontWeight="SemiBold" Margin="0,0,0,6"/>
    </StackPanel>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button DockPanel.Dock="Right" Content="Sluiten" Style="{StaticResource BtnGray}" IsCancel="True" Margin="8,0,0,6"/>
      <Button x:Name="RO_BtnExport" DockPanel.Dock="Right" Content="Exporteren naar CSV..." Style="{StaticResource BtnGray}" Margin="8,0,0,6"/>
      <WrapPanel>
        <Button x:Name="RO_BtnMembers" Content="Leden tonen"/>
        <Button x:Name="RO_BtnSettings" Content="Groepsinstellingen..."
                ToolTip="Microsoft 365-groep: groepsmail in de inbox van leden, extern mailen, privacy, verbergen, MailTip en verzenden als/namens. Via de wachtrij."/>
        <Button x:Name="RO_BtnOpen" Content="Openen in gebruikerslijst" Style="{StaticResource BtnGreen}"/>
        <Button x:Name="RO_BtnNewGroup" Content="Nieuwe groep / lijst..." Style="{StaticResource BtnGreen}" Margin="0,0,8,6" Padding="12,6"
                ToolTip="Distributielijst, beveiligingsgroep met e-mail, Microsoft 365-groep, dynamische distributielijst of ruimtelijst aanmaken, met eigenaren en leden. Via de wachtrij."/>
        <Button x:Name="RO_BtnNewContact" Content="Nieuwe contactpersonen..." Style="{StaticResource BtnDark}" Margin="0,0,8,6" Padding="12,6"
                ToolTip="MailContact aanmaken voor externe adressen (los, zonder groep). Via de wachtrij."/>
      </WrapPanel>
    </DockPanel>
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="3*"/>
        <ColumnDefinition Width="6"/>
        <ColumnDefinition Width="2*" MinWidth="320"/>
      </Grid.ColumnDefinitions>
      <DataGrid x:Name="RO_Grid" Grid.Column="0" SelectionMode="Single">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Naam"          Binding="{Binding Naam}"       Width="2*"  ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Adres"         Binding="{Binding Adres}"      Width="2*"  ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Type"          Binding="{Binding Type}"       Width="150" ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Bron"          Binding="{Binding Bron}"       Width="50"  ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Aliassen"      Binding="{Binding Aliassen}"   Width="2*"  ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Extern adres"  Binding="{Binding Extern}"     Width="*"   ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Beheerder"     Binding="{Binding Beheerder}"  Width="*"   ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Verborgen"     Binding="{Binding Verborgen}"  Width="70"  ElementStyle="{StaticResource CellText}"/>
          <DataGridTextColumn Header="Doorsturen"    Binding="{Binding Doorsturen}" Width="*"   ElementStyle="{StaticResource CellQueue}"/>
          <DataGridTextColumn Header="Wachtrij"      Binding="{Binding Wachtrij}"   Width="85"  ElementStyle="{StaticResource CellQueue}"/>
        </DataGrid.Columns>
      </DataGrid>
      <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#D0D0D0" ResizeDirection="Columns"/>
      <DockPanel Grid.Column="2">
        <DockPanel DockPanel.Dock="Top">
          <Button x:Name="RO_BtnExportMembers" DockPanel.Dock="Right" Content="Export..." Style="{StaticResource BtnGray}" Padding="8,3" Margin="6,0,0,6"/>
          <TextBlock x:Name="RO_LblMembers" Text="Selecteer een groep en klik op 'Leden tonen'" TextWrapping="Wrap" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,0,6"/>
        </DockPanel>
        <WrapPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
          <Button x:Name="RO_BtnAddPick" Content="Lid toevoegen..." Style="{StaticResource BtnGreen}" Padding="10,4"
                  ToolTip="Bestaande gebruiker, mailbox, groep of contactpersoon zoeken en toevoegen"/>
          <Button x:Name="RO_BtnAddPaste" Content="Adressen plakken..." Style="{StaticResource BtnGreen}" Padding="10,4"
                  ToolTip="Meerdere adressen in 1 keer (uit Outlook/Excel). Externe adressen krijgen eerst automatisch een contactpersoon."/>
          <Button x:Name="RO_BtnRemove" Content="Verwijderen" Style="{StaticResource BtnRed}" Padding="10,4"
                  ToolTip="Geselecteerde leden uit de groep halen (via de wachtrij). Bij een lid met '+ toevoegen' wordt alleen het wachtrij-item geschrapt."/>
          <Button x:Name="RO_BtnSubOn" Content="Groepsmail in inbox aan" Padding="10,4" Visibility="Collapsed"
                  ToolTip="Geselecteerde leden krijgen de mail en afspraken van de M365-groep in hun eigen inbox (abonneren, via de wachtrij)"/>
          <Button x:Name="RO_BtnSubOff" Content="Groepsmail in inbox uit" Style="{StaticResource BtnGray}" Padding="10,4" Visibility="Collapsed"
                  ToolTip="Geselecteerde leden krijgen de groepsmail niet meer in hun eigen inbox (blijven wel lid)"/>
        </WrapPanel>
        <DataGrid x:Name="RO_MemberGrid" SelectionMode="Extended">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Naam"     Binding="{Binding Naam}"     Width="*"   ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Adres"    Binding="{Binding Adres}"    Width="*"   ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Type"     Binding="{Binding Type}"     Width="100" ElementStyle="{StaticResource CellText}"/>
            <DataGridTextColumn Header="Inbox"    Binding="{Binding Inbox}"    Width="85"  ElementStyle="{StaticResource CellText}" Visibility="Collapsed"/>
            <DataGridTextColumn Header="Wachtrij" Binding="{Binding Wachtrij}" Width="95"  ElementStyle="{StaticResource CellQueue}"/>
          </DataGrid.Columns>
        </DataGrid>
      </DockPanel>
    </Grid>
  </DockPanel>
</Window>
'@

$script:AddressXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Adressen" Width="640" Height="480" MinWidth="460" MinHeight="300" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <DockPanel Margin="14">
    <TextBlock x:Name="AD_Prompt" DockPanel.Dock="Top" TextWrapping="Wrap" Margin="0,0,0,8"/>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button DockPanel.Dock="Right" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="8,0,0,6"/>
      <Button x:Name="AD_BtnOk" DockPanel.Dock="Right" Content="Naar wachtrij" Style="{StaticResource BtnGreen}" Margin="8,0,0,6"/>
      <StackPanel>
        <CheckBox x:Name="AD_ChkHide" Content="Nieuwe contactpersonen verbergen in het adresboek"/>
        <TextBlock x:Name="AD_LblCount" Foreground="Gray" Margin="0,4,0,0"/>
      </StackPanel>
    </DockPanel>
    <TextBox x:Name="AD_Txt" AcceptsReturn="True" AcceptsTab="True" TextWrapping="NoWrap" FontFamily="Consolas"
             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
  </DockPanel>
</Window>
'@

$script:NewGroupXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Nieuwe groep / distributielijst" Width="920" Height="800" MinWidth="760" MinHeight="560" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <DockPanel Margin="14">
    <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}"
               Text="Nieuwe groep in Exchange Online. Er wordt nog niets aangemaakt: de groep en daarna de gekozen leden gaan naar de wachtrij en worden met 'Alles uitvoeren' in die volgorde uitgevoerd."/>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button DockPanel.Dock="Right" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="8,0,0,6"/>
      <Button x:Name="NG_BtnOk" DockPanel.Dock="Right" Content="Naar wachtrij" Style="{StaticResource BtnGreen}" Margin="8,0,0,6"/>
      <TextBlock x:Name="NG_LblAddr" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,0,6"/>
    </DockPanel>
    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
      <StackPanel>
        <GroupBox Header="Soort groep">
          <StackPanel>
            <RadioButton x:Name="NG_RbDL"   GroupName="NgType" IsChecked="True" Content="Distributielijst - mail naar 1 adres gaat naar alle leden"/>
            <RadioButton x:Name="NG_RbSec"  GroupName="NgType" Content="Beveiligingsgroep met e-mail - distributielijst die ook voor rechten te gebruiken is"/>
            <RadioButton x:Name="NG_RbM365" GroupName="NgType" Content="Microsoft 365-groep - gedeelde inbox, agenda en bestanden (SharePoint), optioneel Teams"/>
            <RadioButton x:Name="NG_RbDyn"  GroupName="NgType" Content="Dynamische distributielijst - leden volgen automatisch uit een filter (bijv. afdeling)"/>
            <RadioButton x:Name="NG_RbRoom" GroupName="NgType" Content="Ruimtelijst - groepeert vergaderruimtes voor de ruimtezoeker in Outlook"/>
            <TextBlock x:Name="NG_LblTypeInfo" Style="{StaticResource Hint}" Margin="0,2,0,0"/>
          </StackPanel>
        </GroupBox>
        <GroupBox Header="Gegevens">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="130"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="Weergavenaam" Style="{StaticResource FieldLabel}"/>
            <TextBox x:Name="NG_TxtName" Grid.Column="1" Style="{StaticResource FieldBox}"/>
            <TextBlock Grid.Row="1" Text="E-mailadres" Style="{StaticResource FieldLabel}"/>
            <DockPanel Grid.Row="1" Grid.Column="1">
              <ComboBox x:Name="NG_CmbDomain" DockPanel.Dock="Right" Width="280" Height="26" Margin="0,0,0,6"/>
              <TextBlock DockPanel.Dock="Right" Text="@" VerticalAlignment="Center" FontWeight="SemiBold" Margin="6,0,6,6"/>
              <TextBox x:Name="NG_TxtAlias" Style="{StaticResource FieldBox}"
                       ToolTip="Alias (deel voor de @). Wordt automatisch uit de naam gemaakt; zelf aanpassen kan."/>
            </DockPanel>
            <TextBlock Grid.Row="2" Text="Omschrijving" Style="{StaticResource FieldLabel}" VerticalAlignment="Top" Margin="0,4,8,6"/>
            <TextBox x:Name="NG_TxtDesc" Grid.Row="2" Grid.Column="1" Height="50" AcceptsReturn="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto" Padding="4,2" Margin="0,0,0,6" ToolTip="Optioneel"/>
          </Grid>
        </GroupBox>
        <GroupBox Header="Instellingen">
          <StackPanel>
            <CheckBox x:Name="NG_ChkExternal" Content="Mail van buiten de organisatie toestaan (externe afzenders)" Margin="0,0,0,6"/>
            <CheckBox x:Name="NG_ChkHide" Content="Verbergen in het adresboek (GAL)" Margin="0,0,0,6"/>
            <WrapPanel x:Name="NG_PnlJoin">
              <TextBlock Text="Lid worden:" Style="{StaticResource FieldLabel}"/>
              <ComboBox x:Name="NG_CmbJoin" Width="220" Margin="0,0,16,6">
                <ComboBoxItem Content="Gesloten (alleen eigenaren)" Tag="Closed" IsSelected="True"/>
                <ComboBoxItem Content="Na goedkeuring eigenaar" Tag="ApprovalRequired"/>
                <ComboBoxItem Content="Open (iedereen zelf)" Tag="Open"/>
              </ComboBox>
              <TextBlock Text="Opzeggen:" Style="{StaticResource FieldLabel}"/>
              <ComboBox x:Name="NG_CmbDepart" Width="220" Margin="0,0,0,6">
                <ComboBoxItem Content="Open (iedereen zelf)" Tag="Open" IsSelected="True"/>
                <ComboBoxItem Content="Gesloten (alleen eigenaren)" Tag="Closed"/>
              </ComboBox>
            </WrapPanel>
            <WrapPanel x:Name="NG_PnlM365">
              <TextBlock Text="Privacy:" Style="{StaticResource FieldLabel}"/>
              <ComboBox x:Name="NG_CmbPrivacy" Width="300" Margin="0,0,16,6">
                <ComboBoxItem Content="Priv&#233; - alleen leden zien de inhoud" Tag="Private" IsSelected="True"/>
                <ComboBoxItem Content="Openbaar - iedereen in de organisatie" Tag="Public"/>
              </ComboBox>
              <CheckBox x:Name="NG_ChkSubscribe" Content="Leden ontvangen groepsmail ook in hun eigen inbox" VerticalAlignment="Center" Margin="0,0,0,6"/>
            </WrapPanel>
            <Grid x:Name="NG_PnlDyn">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="130"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Text="Ontvangers" Style="{StaticResource FieldLabel}"/>
              <ComboBox x:Name="NG_CmbIncluded" Grid.Column="1" Height="26" Margin="0,0,0,6">
                <ComboBoxItem Content="Mailboxen (gebruikers en gedeelde mailboxen)" Tag="MailboxUsers" IsSelected="True"/>
                <ComboBoxItem Content="Mailboxen + e-mailgebruikers" Tag="MailboxUsers, MailUsers"/>
                <ComboBoxItem Content="Contactpersonen" Tag="MailContacts"/>
                <ComboBoxItem Content="Ruimtes en apparatuur" Tag="Resources"/>
                <ComboBoxItem Content="Alle ontvangers" Tag="AllRecipients"/>
              </ComboBox>
              <TextBlock Grid.Row="1" Text="Afdeling(en)" Style="{StaticResource FieldLabel}"/>
              <TextBox x:Name="NG_TxtDept" Grid.Row="1" Grid.Column="1" Style="{StaticResource FieldBox}" ToolTip="Leeg = alle afdelingen. Meerdere scheiden met ;"/>
              <TextBlock Grid.Row="2" Text="Bedrijf/bedrijven" Style="{StaticResource FieldLabel}"/>
              <TextBox x:Name="NG_TxtCompany" Grid.Row="2" Grid.Column="1" Style="{StaticResource FieldBox}" ToolTip="Leeg = alle bedrijven. Meerdere scheiden met ;"/>
              <TextBlock Grid.Row="3" Text="Eigen filter (OPATH)" Style="{StaticResource FieldLabel}"/>
              <TextBox x:Name="NG_TxtFilter" Grid.Row="3" Grid.Column="1" Style="{StaticResource FieldBox}" FontFamily="Consolas"
                       ToolTip="Optioneel, voor gevorderden. Bijv: (RecipientTypeDetails -eq 'UserMailbox') -and (Office -eq 'Utrecht')"/>
              <TextBlock Grid.Row="4" Grid.Column="1" Style="{StaticResource Hint}"
                         Text="Afdeling en Bedrijf komen uit Entra. Meerdere waarden scheiden met ; (een ontvanger hoeft er maar 1 te hebben). Met een eigen filter worden Ontvangers, Afdeling en Bedrijf genegeerd."/>
            </Grid>
          </StackPanel>
        </GroupBox>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="10"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <GroupBox x:Name="NG_GrpOwners" Header="Eigenaren">
            <DockPanel>
              <TextBlock x:Name="NG_LblOwners" DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,4"/>
              <WrapPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
                <Button x:Name="NG_BtnOwnerAdd" Content="Toevoegen..." Style="{StaticResource BtnGreen}" Padding="10,4"/>
                <Button x:Name="NG_BtnOwnerRemove" Content="Verwijderen" Style="{StaticResource BtnRed}" Padding="10,4"/>
              </WrapPanel>
              <ListBox x:Name="NG_LstOwners" Height="160" SelectionMode="Extended"/>
            </DockPanel>
          </GroupBox>
          <GroupBox x:Name="NG_GrpMembers" Grid.Column="2" Header="Leden">
            <DockPanel>
              <TextBlock x:Name="NG_LblMembers" DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,4"/>
              <WrapPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
                <Button x:Name="NG_BtnMemberAdd" Content="Toevoegen..." Style="{StaticResource BtnGreen}" Padding="10,4"/>
                <Button x:Name="NG_BtnMemberPaste" Content="Adressen plakken..." Style="{StaticResource BtnGreen}" Padding="10,4"
                        ToolTip="Meerdere adressen in 1 keer (uit Outlook/Excel). Externe adressen krijgen eerst automatisch een contactpersoon."/>
                <Button x:Name="NG_BtnMemberRemove" Content="Verwijderen" Style="{StaticResource BtnRed}" Padding="10,4"/>
              </WrapPanel>
              <ListBox x:Name="NG_LstMembers" Height="160" SelectionMode="Extended"/>
            </DockPanel>
          </GroupBox>
        </Grid>
      </StackPanel>
    </ScrollViewer>
  </DockPanel>
</Window>
'@

$script:GroupSettingsXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft 365-groep - instellingen" Width="780" Height="860" MinWidth="620" MinHeight="500" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
    <Style TargetType="CheckBox"><Setter Property="Margin" Value="0,0,0,6"/></Style>
  </Window.Resources>
  <DockPanel Margin="14">
    <StackPanel DockPanel.Dock="Top">
      <TextBlock x:Name="GS_LblTitle" FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap"/>
      <TextBlock x:Name="GS_LblInfo" Style="{StaticResource Hint}" Margin="0,2,0,8"/>
    </StackPanel>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button DockPanel.Dock="Right" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" Margin="8,0,0,6"/>
      <Button x:Name="GS_BtnOk" DockPanel.Dock="Right" Content="Naar wachtrij" Style="{StaticResource BtnGreen}" Margin="8,0,0,6"/>
      <TextBlock x:Name="GS_LblChanges" VerticalAlignment="Center" Foreground="#0063B1" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,6"/>
    </DockPanel>
    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
      <StackPanel Margin="0,0,6,0">
        <GroupBox Header="Groepsmail in de inbox van de leden">
          <StackPanel>
            <CheckBox x:Name="GS_ChkAutoSub" Content="Kopie&#235;n van groepsmail en -afspraken naar de inbox van nieuwe leden sturen"
                      ToolTip="AutoSubscribeNewMembers. Zelfde als 'Send copies of group conversations and events to group members inboxes' in het admin center."/>
            <CheckBox x:Name="GS_ChkCalSub" Content="Alleen groepsafspraken in de agenda van nieuwe leden (zonder de mail)"
                      ToolTip="AlwaysSubscribeMembersToCalendarEvents. Werkt alleen als de optie hierboven uit staat."/>
            <TextBlock x:Name="GS_LblSubs" Style="{StaticResource Hint}" Margin="20,0,0,6"/>
            <CheckBox x:Name="GS_ChkSubExisting" FontWeight="SemiBold"
                      ToolTip="Zet de huidige leden als abonnee (Add-UnifiedGroupLinks -LinkType Subscribers). Per lid aan/uit kan ook in het ledenpaneel."/>
            <CheckBox x:Name="GS_ChkWelcome" Content="Welkomstmail sturen naar nieuwe leden" ToolTip="UnifiedGroupWelcomeMessageEnabled"/>
            <CheckBox x:Name="GS_ChkSubEnabled" Content="Abonneren toegestaan (uit = niemand kan de groepsmail in de eigen inbox aanzetten)"
                      ToolTip="SubscriptionEnabled. Normaal aan laten staan."/>
          </StackPanel>
        </GroupBox>
        <GroupBox Header="Afzenders">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <CheckBox x:Name="GS_ChkExternal" Grid.ColumnSpan="2" Content="Mensen buiten de organisatie mogen naar de groep mailen"
                      ToolTip="RequireSenderAuthenticationEnabled uit. Standaard mag alleen intern."/>
            <TextBlock Grid.Row="1" Text="MailTip:" Style="{StaticResource FieldLabel}"/>
            <TextBox x:Name="GS_TxtMailTip" Grid.Row="1" Grid.Column="1" Style="{StaticResource FieldBox}"
                     ToolTip="Melding die afzenders in Outlook zien als ze naar deze groep mailen. Leeg = geen MailTip."/>
          </Grid>
        </GroupBox>
        <GroupBox Header="Privacy en zichtbaarheid">
          <StackPanel>
            <WrapPanel>
              <TextBlock Text="Privacy:" Style="{StaticResource FieldLabel}"/>
              <ComboBox x:Name="GS_CmbPrivacy" Width="380" Margin="0,0,0,6">
                <ComboBoxItem Content="Priv&#233; - alleen leden zien de inhoud, eigenaar keurt nieuwe leden goed" Tag="Private"/>
                <ComboBoxItem Content="Openbaar - iedereen in de organisatie kan de inhoud zien en lid worden" Tag="Public"/>
              </ComboBox>
            </WrapPanel>
            <CheckBox x:Name="GS_ChkHideGal" Content="Verbergen in het adresboek (GAL)"/>
            <CheckBox x:Name="GS_ChkHideOutlook" Content="Verbergen in Outlook (niet in de groepenlijst en niet te vinden; mail naar de groep komt wel binnen)"
                      ToolTip="HiddenFromExchangeClientsEnabled. Verbergt automatisch ook in het adresboek. Leden zien dan ook de abonneer-optie niet."/>
            <CheckBox x:Name="GS_ChkCalRO" Content="Leden hebben alleen leesrechten op de groepsagenda" ToolTip="CalendarMemberReadOnly"/>
          </StackPanel>
        </GroupBox>
        <GroupBox Header="Beschrijving">
          <TextBox x:Name="GS_TxtDesc" Height="54" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Padding="4,2"/>
        </GroupBox>
        <GroupBox Header="Verzenden als / namens de groep">
          <StackPanel>
            <TextBlock Style="{StaticResource Hint}"
                       Text="Verzenden als: de ontvanger ziet alleen de groep als afzender. Verzenden namens: 'Naam namens Groep'. Leden hebben dit niet automatisch; antwoorden komen in de groep."/>
            <DataGrid x:Name="GS_DelGrid" Height="140" SelectionMode="Extended">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Naam"     Binding="{Binding Naam}"     Width="*"   ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Recht"    Binding="{Binding Recht}"    Width="130" ElementStyle="{StaticResource CellText}"/>
                <DataGridTextColumn Header="Wachtrij" Binding="{Binding Wachtrij}" Width="100" ElementStyle="{StaticResource CellQueue}"/>
              </DataGrid.Columns>
            </DataGrid>
            <WrapPanel Margin="0,6,0,0">
              <Button x:Name="GS_BtnAddSendAs" Content="Verzenden als toevoegen..." Style="{StaticResource BtnGreen}" Padding="10,4"/>
              <Button x:Name="GS_BtnAddSob" Content="Verzenden namens toevoegen..." Style="{StaticResource BtnGreen}" Padding="10,4"/>
              <Button x:Name="GS_BtnDelRemove" Content="Verwijderen" Style="{StaticResource BtnRed}" Padding="10,4"
                      ToolTip="Geselecteerde rechten intrekken (via de wachtrij). Nog een keer klikken maakt het ongedaan."/>
            </WrapPanel>
          </StackPanel>
        </GroupBox>
      </StackPanel>
    </ScrollViewer>
  </DockPanel>
</Window>
'@

$script:CsvXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CSV import / export - Entra eigenschappen" Width="1100" Height="720" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <DockPanel Margin="14">
    <StackPanel DockPanel.Dock="Top">
      <GroupBox Header="1. Exporteren (om in Excel aan te passen)">
        <StackPanel>
          <WrapPanel>
            <Button x:Name="CV_BtnExportSel" Content="Exporteer selectie"/>
            <Button x:Name="CV_BtnExportAll" Content="Exporteer alle zichtbare gebruikers"/>
            <Button x:Name="CV_BtnTemplate"  Content="Leeg sjabloon" Style="{StaticResource BtnGray}"/>
          </WrapPanel>
          <TextBlock Style="{StaticResource Hint}" Margin="0"
                     Text="Kolommen: {{COLS}}"/>
        </StackPanel>
      </GroupBox>
      <GroupBox Header="2. Importeren">
        <StackPanel>
          <TextBlock Style="{StaticResource Hint}"
                     Text="UserPrincipalName is de sleutel (wordt niet gewijzigd). Lege cel = niet wijzigen. {{CLEAR}} = veld leegmaken / manager verwijderen. accountEnabled: TRUE of FALSE. managerUPN: UPN van de manager. Kolommen die je weglaat blijven ongewijzigd. Scheidingsteken ; of , wordt automatisch herkend. Een regel met een fout wordt in zijn geheel overgeslagen."/>
          <DockPanel>
            <Button x:Name="CV_BtnLoad" Content="CSV kiezen en vergelijken..."/>
            <TextBlock x:Name="CV_LblFile" VerticalAlignment="Center" Margin="4,0,0,6" Foreground="Gray"/>
          </DockPanel>
        </StackPanel>
      </GroupBox>
    </StackPanel>
    <Border DockPanel.Dock="Bottom" Padding="0,8,0,0">
      <DockPanel>
        <Button DockPanel.Dock="Right" Content="Sluiten" Style="{StaticResource BtnGray}" IsCancel="True" Margin="8,0,0,6"/>
        <Button x:Name="CV_BtnApply" DockPanel.Dock="Right" Content="3. Toevoegen aan wachtrij" Style="{StaticResource BtnGreen}" Margin="8,0,0,6"/>
        <CheckBox x:Name="CV_ChkOnlyChanges" DockPanel.Dock="Right" Content="Ongewijzigde gebruikers verbergen" IsChecked="True" VerticalAlignment="Center" Margin="8,0,8,6"/>
        <TextBlock x:Name="CV_LblStatus" VerticalAlignment="Center" TextWrapping="Wrap" FontWeight="SemiBold" Margin="0,0,0,6"/>
      </DockPanel>
    </Border>
    <DataGrid x:Name="CV_Grid">
      <DataGrid.Columns>
        <DataGridTextColumn Header="UPN"    Binding="{Binding UPN}"    Width="2*" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Veld"   Binding="{Binding Veld}"   Width="130" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Huidig" Binding="{Binding Huidig}" Width="2*" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Nieuw"  Binding="{Binding Nieuw}"  Width="2*" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="2*" ElementStyle="{StaticResource CellText}"/>
      </DataGrid.Columns>
    </DataGrid>
  </DockPanel>
</Window>
'@

# Celkleur per veldkolom in het overzicht; {{P}} = veldnaam
$script:OvCellStyleXaml = @'
<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="DataGridCell">
  <Style.Triggers>
    <DataTrigger Binding="{Binding s_{{P}}}" Value="Q"><Setter Property="Background" Value="#DCEBFF"/></DataTrigger>
    <DataTrigger Binding="{Binding s_{{P}}}" Value="C"><Setter Property="Background" Value="#FFF4CE"/></DataTrigger>
    <DataTrigger Binding="{Binding s_{{P}}}" Value="E"><Setter Property="Background" Value="#F8D7DA"/></DataTrigger>
    <Trigger Property="IsSelected" Value="True">
      <Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="White"/><Setter Property="BorderBrush" Value="#0078D4"/>
    </Trigger>
  </Style.Triggers>
</Style>
'@

$script:OverviewXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Overzicht bewerken" Width="1400" Height="820" MinWidth="1000" MinHeight="560" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <DockPanel Margin="12">
    <StackPanel DockPanel.Dock="Top">
      <TextBlock Style="{StaticResource Hint}"
                 Text="Pas waarden direct in de tabel aan: dubbelklik of begin te typen (Enter = klaar, Esc = annuleren). Ctrl+C / Ctrl+V werkt met Excel: 1 waarde plakken vult alle geselecteerde cellen, een blok plakt vanaf de aangeklikte cel. Delete maakt de geselecteerde cellen leeg. Geel = gewijzigd, rood = ongeldig, blauw = staat al in de wachtrij. Grijze rijen komen uit on-premises AD en zijn alleen-lezen (behalve Gebruikslocatie)."/>
      <GroupBox Header="Kolommen tonen">
        <WrapPanel x:Name="OV_ColPanel"/>
      </GroupBox>
    </StackPanel>
    <Border DockPanel.Dock="Bottom" Padding="0,8,0,0">
      <DockPanel>
        <Button x:Name="OV_BtnClose"    DockPanel.Dock="Right" Content="Sluiten" Style="{StaticResource BtnGray}" Margin="8,0,0,6"/>
        <Button x:Name="OV_BtnQueue"    DockPanel.Dock="Right" Content="Toevoegen aan wachtrij" Style="{StaticResource BtnGreen}" Margin="8,0,0,6"/>
        <Button x:Name="OV_BtnResetAll" DockPanel.Dock="Right" Content="Alles ongedaan maken" Style="{StaticResource BtnGray}" Margin="8,0,0,6"/>
        <StackPanel VerticalAlignment="Center" Margin="0,0,0,6">
          <TextBlock x:Name="OV_LblStatus" FontWeight="SemiBold" TextWrapping="Wrap"/>
          <TextBlock x:Name="OV_LblMsg" Foreground="#555555" TextWrapping="Wrap"/>
        </StackPanel>
      </DockPanel>
    </Border>
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="280" MinWidth="200"/>
        <ColumnDefinition Width="6"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border Grid.Column="0" Background="White" BorderBrush="#DDDDDD" BorderThickness="1" Padding="8">
        <DockPanel>
          <TextBlock DockPanel.Dock="Top" Text="Waarden per kolom" FontWeight="SemiBold" FontSize="13" Margin="0,0,0,6"/>
          <ComboBox x:Name="OV_CmbCol" DockPanel.Dock="Top" DisplayMemberPath="Label" SelectedValuePath="Prop" Margin="0,0,0,6"/>
          <TextBlock x:Name="OV_LblValues" DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,4"/>
          <GroupBox DockPanel.Dock="Bottom" Header="Geselecteerde cellen in de tabel" Margin="0">
            <StackPanel>
              <TextBox x:Name="OV_TxtSet" Style="{StaticResource FieldBox}" ToolTip="Nieuwe waarde voor alle geselecteerde cellen (Enter = instellen)"/>
              <WrapPanel>
                <Button x:Name="OV_BtnSet"       Content="Instellen" Padding="10,4"/>
                <Button x:Name="OV_BtnClearSel"  Content="Leegmaken" Style="{StaticResource BtnGray}" Padding="10,4"/>
                <Button x:Name="OV_BtnRevertSel" Content="Herstellen" Style="{StaticResource BtnGray}" Padding="10,4"
                        ToolTip="Zet de geselecteerde cellen terug naar de oorspronkelijke waarde"/>
              </WrapPanel>
            </StackPanel>
          </GroupBox>
          <GroupBox DockPanel.Dock="Bottom" Header="Aangeklikte waarde vervangen door" Margin="0,8,0,8">
            <StackPanel>
              <TextBox x:Name="OV_TxtReplace" Style="{StaticResource FieldBox}" ToolTip="Leeg = veld leegmaken (Enter = vervangen). Geldt voor alle gebruikers met deze waarde."/>
              <WrapPanel>
                <Button x:Name="OV_BtnReplace" Content="Vervangen" Padding="10,4"/>
                <Button x:Name="OV_BtnShowAll" Content="Filter wissen" Style="{StaticResource BtnGray}" Padding="10,4"/>
              </WrapPanel>
            </StackPanel>
          </GroupBox>
          <ListBox x:Name="OV_LstValues"/>
        </DockPanel>
      </Border>
      <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="Transparent" ResizeDirection="Columns"/>
      <DockPanel Grid.Column="2">
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
          <CheckBox x:Name="OV_ChkOnlyChanged" DockPanel.Dock="Right" Content="Alleen gewijzigde rijen" VerticalAlignment="Center" Margin="12,0,0,0"/>
          <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="OV_TxtFilter" Height="26" VerticalContentAlignment="Center" ToolTip="Filtert op naam, UPN en alle velden"/>
        </DockPanel>
        <DataGrid x:Name="OV_Grid" IsReadOnly="False" SelectionUnit="Cell" SelectionMode="Extended" FrozenColumnCount="2"
                  ClipboardCopyMode="ExcludeHeader" CanUserReorderColumns="False">
          <DataGrid.RowStyle>
            <Style TargetType="DataGridRow">
              <Style.Triggers>
                <DataTrigger Binding="{Binding Bron}" Value="AD">
                  <Setter Property="Foreground" Value="#8A8A8A"/>
                </DataTrigger>
              </Style.Triggers>
            </Style>
          </DataGrid.RowStyle>
        </DataGrid>
      </DockPanel>
    </Grid>
  </DockPanel>
</Window>
'@

function ConvertFrom-Xaml {
    param([string]$Text)
    [xml]$x = $Text
    return [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
}

#endregion

#region --- Hoofdvenster ---

function New-MainWindow {
    $xaml = $script:MainXaml.Replace('{{STYLES}}', $script:StylesXaml).
        Replace('{{FOLDER_ITEMS}}', $script:FolderComboItems).
        Replace('{{QCOL}}', $script:QueueColXaml).
        Replace('{{BULK_FIELDS}}', (Get-BulkFieldsXaml))
    foreach ($g in 'Identiteit', 'Organisatie', 'Contact', 'Adres') {
        $xaml = $xaml.Replace("{{G_$g}}", (Get-FieldGroupXaml $g))
    }
    $script:Window = ConvertFrom-Xaml $xaml

    $names = @(
        'HeaderPanel', 'TxtTenant', 'LblGraph', 'LblExo', 'BtnConnectGraph', 'BtnConnectExo',
        'MainArea', 'UserPanel', 'TxtSearch', 'BtnSearch', 'TxtFilter', 'ChkHideGuests', 'ChkOnlyEnabled', 'ChkOnlyFwd', 'CmbMbxType', 'BtnSelectAll', 'BtnOverview', 'BtnCsv', 'BtnFwdOverview', 'BtnRcptOverview', 'BtnNewGroup', 'LblUserCount', 'UserGrid',
        'LblUser', 'LblUserSub', 'LblUserQueue', 'SyncWarn', 'DetailTabs', 'TabEntra', 'TabMbx', 'TabFwd', 'TabFolder', 'TabScan',
        'BtnEntraSave', 'BtnEntraReset', 'LblEntraDirty', 'ChkAccountEnabled', 'TxtManager', 'BtnMgrPick', 'BtnMgrClear', 'TxtInfo',
        'BtnMbxRefresh', 'LblMbxInfo', 'MbxContent', 'MbxFullGrid', 'BtnFullAdd', 'ChkAutoMap', 'BtnFullRemove',
        'MbxSendAsGrid', 'BtnSendAsAdd', 'BtnSendAsRemove', 'MbxSobGrid', 'BtnSobAdd', 'BtnSobRemove',
        'BtnFwdRefresh', 'LblFwdInfo', 'FwdContent', 'FwdStateBorder', 'LblFwdState', 'LblFwdDetail', 'LblFwdQueue',
        'RbFwdOff', 'RbFwdInt', 'RbFwdExt', 'TxtFwdInt', 'BtnFwdPick', 'TxtFwdExt', 'ChkFwdKeep', 'BtnFwdSave', 'BtnFwdReset',
        'LblFwdPolicy', 'BtnFwdRules', 'LblFwdRules', 'FwdRuleGrid',
        'CmbFolder', 'BtnFolderRefresh', 'LblFolderPath', 'FolderContent', 'FolderGrid', 'CmbRights', 'ChkDelegate', 'ChkPrivate',
        'BtnFolderAdd', 'BtnFolderSet', 'BtnFolderRemove',
        'BtnScan', 'ChkScanShared', 'ChkScanSob', 'LblScan', 'ScanGrid', 'BtnScanRemove',
        'BulkTabs', 'BtnBulkEntraApply', 'BtnBulkEntraClear', 'CmbBulkAccount', 'ChkBulkMgr', 'TxtBulkMgr', 'BtnBulkMgrPick', 'BtnBulkMgrClear',
        'RbMbxA', 'RbMbxB', 'LblBulkMbxTarget', 'BtnBulkMbxPick', 'TxtBulkMbxTarget', 'ChkBFull', 'ChkBAuto', 'ChkBSendAs', 'ChkBSob', 'BtnBulkMbxGrant', 'BtnBulkMbxRevoke',
        'RbFolA', 'RbFolB', 'RbFolC', 'BulkFolTargetPanel', 'LblBulkFolTarget', 'BtnBulkFolPick', 'TxtBulkFolTarget', 'CmbBulkFolder', 'CmbBulkRights',
        'ChkBulkDelegate', 'ChkBulkPrivate', 'BtnBulkFolGrant', 'BtnBulkFolRevoke',
        'QueueGrid', 'QueueButtons', 'LblQueue', 'BtnRunQueue', 'BtnQueueRemove', 'BtnQueueClearDone', 'BtnQueueClear',
        'BtnClearLog', 'LogBox', 'LogPara', 'StatusBar', 'LblStatus', 'BtnStop'
    )
    foreach ($n in $names) {
        $ctl = $script:Window.FindName($n)
        if ($null -eq $ctl) { throw "XAML-element '$n' niet gevonden" }
        Set-Variable -Scope Script -Name $n -Value $ctl
    }

    $script:FieldBoxes = @{}
    foreach ($f in $script:EntraFields) {
        $box = $script:Window.FindName("F_$($f.Prop)")
        $box.Add_TextChanged({ Update-EntraDirty })
        $script:FieldBoxes[$f.Prop] = $box
    }
    $script:FieldBoxes['usageLocation'].MaxLength = 2
    $script:FieldBoxes['usageLocation'].CharacterCasing = 'Upper'

    $script:BulkBoxes = @{}
    foreach ($f in @($script:EntraFields | Where-Object { $_.Bulk })) {
        $chk = $script:Window.FindName("BC_$($f.Prop)")
        $box = $script:Window.FindName("BT_$($f.Prop)")
        # Typen in een bulkveld vinkt het automatisch aan
        $box.Add_TextChanged({ $script:BulkBoxes[$this.Tag].Chk.IsChecked = $true })
        $script:BulkBoxes[$f.Prop] = @{ Chk = $chk; Box = $box }
    }
    $script:BulkBoxes['usageLocation'].Box.MaxLength = 2
    $script:BulkBoxes['usageLocation'].Box.CharacterCasing = 'Upper'

    $script:CmbRights.ItemsSource = $script:FolderRights
    $script:CmbRights.SelectedValue = 'Reviewer'
    $script:CmbBulkRights.ItemsSource = $script:FolderRights
    $script:CmbBulkRights.SelectedValue = 'Reviewer'

    $script:QueueTable = New-Table @('Nr', 'Doel', 'Soort', 'Wijziging', 'Status', 'Kleur', 'Key')
    $script:QueueGrid.ItemsSource = $script:QueueTable.DefaultView

    Reset-ExoState
    Update-BulkLabels
    Update-QueueSummary

    if (-not $script:ExoAvailable) {
        $script:BtnConnectExo.IsEnabled = $false
        $script:BtnConnectExo.ToolTip = "Module ExchangeOnlineManagement niet gevonden"
    }

    # --- Verbinden ---
    $script:BtnConnectGraph.Add_Click({
        Set-Status "Verbinding maken met Graph - inlogvenster opent zo..."
        $script:Window.WindowState = 'Minimized'
        $ok = Connect-ToGraph
        $script:Window.WindowState = 'Normal'
        [void]$script:Window.Activate()
        if ($ok) {
            $ctx = Get-MgContext
            $script:GraphConnected = $true
            $script:LblGraph.Text = "Graph: $($ctx.Account)"
            $script:LblGraph.Foreground = New-Brush "#90EE90"
            $script:BtnConnectGraph.Content = "Graph opnieuw"
            Write-Log "Verbonden met Microsoft Graph als $($ctx.Account) (tenant $($ctx.TenantId))" "success"
            $missing = @($script:GraphScopes | Where-Object { $ctx.Scopes -notcontains $_ })
            if ($missing.Count -gt 0) { Write-Log "LET OP: scope(s) niet toegekend, mogelijk ontbreekt admin consent: $($missing -join ', ')" "warning" }
            Invoke-UserSearch
        }
        else {
            Set-Status "Verbinding met Graph mislukt" "error"
        }
    })

    $script:BtnConnectExo.Add_Click({
        # Niet minimaliseren (zoals bij Graph): het Exchange-inlogvenster
        # (zonder WAM) hangt aan het actieve venster en blijft anders onzichtbaar.
        Set-Status "Verbinding maken met Exchange Online - log in via de browser..."
        $script:HeaderPanel.IsEnabled = $false
        try { $ok = Connect-ToExchange }
        finally { $script:HeaderPanel.IsEnabled = $true }
        [void]$script:Window.Activate()
        if ($ok) {
            $ci = @(Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' }) | Select-Object -Last 1
            $script:ExoConnected = $true
            $script:FolderPathCache = @{}
            $script:AutoFwdText = $null
            $script:AcceptedDomains = $null
            $script:RcptCache = $null
            $org =if ($ci.DelegatedOrganization) { $ci.DelegatedOrganization } else { $ci.TenantID }
            $script:LblExo.Text = "Exchange: $($ci.UserPrincipalName)"
            $script:LblExo.Foreground = New-Brush "#90EE90"
            $script:BtnConnectExo.Content = "Exchange opnieuw"
            Write-Log "Verbonden met Exchange Online als $($ci.UserPrincipalName) (organisatie $org)" "success"
            Set-Status "Verbonden met Exchange Online" "success"
            Update-UserExoColumns
            if ($script:CurrentUser) { Reset-ExoState; Invoke-TabActivated }
        }
        else {
            Set-Status "Verbinding met Exchange mislukt" "error"
        }
    })

    # --- Gebruikers ---
    $script:BtnSearch.Add_Click({ Invoke-UserSearch })
    $script:TxtSearch.Add_KeyDown({ if ($_.Key -eq 'Return') { Invoke-UserSearch } })
    $script:TxtFilter.Add_TextChanged({ Update-UserFilter })
    $script:ChkHideGuests.Add_Click({ Update-UserFilter })
    $script:ChkOnlyEnabled.Add_Click({ Update-UserFilter })
    $script:ChkOnlyFwd.Add_Click({ Update-UserFilter })
    $script:CmbMbxType.Add_SelectionChanged({ Update-UserFilter })
    $script:BtnFwdOverview.Add_Click({ Show-ForwardOverview })
    $script:BtnRcptOverview.Add_Click({ Show-RecipientOverview })
    $script:BtnNewGroup.Add_Click({
        # Invoke-Queue vraagt zelf om bevestiging (Nee = later via 'Alles uitvoeren')
        if (Show-NewGroupDialog) { Invoke-Queue }
    })
    $script:BtnSelectAll.Add_Click({ $script:UserGrid.SelectAll(); [void]$script:UserGrid.Focus() })
    $script:BtnOverview.Add_Click({ Show-OverviewDialog })
    $script:BtnCsv.Add_Click({ Show-CsvDialog })
    $script:UserGrid.Add_SelectionChanged({
        if ($_.OriginalSource -ne $this) { return }
        On-UserSelectionChanged
    })
    $script:DetailTabs.Add_SelectionChanged({
        if ($_.OriginalSource -ne $this) { return }
        Invoke-TabActivated
    })

    # --- Entra (1 gebruiker) ---
    $script:ChkAccountEnabled.Add_Click({ Update-EntraDirty })
    $script:BtnMgrPick.Add_Click({ Invoke-PickManager })
    $script:BtnMgrClear.Add_Click({ Invoke-ClearManager })
    $script:BtnEntraSave.Add_Click({ [void](Add-EntraToQueue) })
    $script:BtnEntraReset.Add_Click({ if ($script:CurrentUser) { Set-EntraForm; Write-Log "Formulierwijzigingen ongedaan gemaakt" "info" } })

    # --- Mailbox rechten (1 gebruiker) ---
    $script:BtnMbxRefresh.Add_Click({ if ($script:CurrentUser) { Invoke-MbxReload } })
    $script:BtnFullAdd.Add_Click({      Invoke-MbxAdd    'FullAccess' })
    $script:BtnFullRemove.Add_Click({   Invoke-MbxRemove 'FullAccess' })
    $script:BtnSendAsAdd.Add_Click({    Invoke-MbxAdd    'SendAs' })
    $script:BtnSendAsRemove.Add_Click({ Invoke-MbxRemove 'SendAs' })
    $script:BtnSobAdd.Add_Click({       Invoke-MbxAdd    'SendOnBehalf' })
    $script:BtnSobRemove.Add_Click({    Invoke-MbxRemove 'SendOnBehalf' })

    # --- Doorsturen (1 gebruiker) ---
    $script:BtnFwdRefresh.Add_Click({ if ($script:CurrentUser) { $script:ExoState.FwdLoaded = $false; Load-ForwardTab } })
    foreach ($rb in $script:RbFwdOff, $script:RbFwdInt, $script:RbFwdExt) { $rb.Add_Click({ Update-ForwardOptions }) }
    $script:BtnFwdPick.Add_Click({ Invoke-PickForward })
    # Typen in het externe adres kiest automatisch 'extern'
    $script:TxtFwdExt.Add_TextChanged({ if ($this.IsKeyboardFocused) { $script:RbFwdExt.IsChecked = $true; Update-ForwardOptions } })
    $script:BtnFwdSave.Add_Click({ Add-ForwardToQueue })
    $script:BtnFwdReset.Add_Click({ Set-ForwardForm })
    $script:BtnFwdRules.Add_Click({ Invoke-ForwardRules })

    # --- Map rechten (1 gebruiker) ---
    $script:CmbFolder.Add_SelectionChanged({
        if ($script:CurrentUser -and $script:DetailTabs.SelectedItem -eq $script:TabFolder) { $script:ExoState.FolderLoaded = $false; Load-FolderTab }
    })
    $script:BtnFolderRefresh.Add_Click({ if ($script:CurrentUser) { $script:ExoState.FolderLoaded = $false; Load-FolderTab } })
    $script:FolderGrid.Add_SelectionChanged({ if ($_.OriginalSource -eq $this) { On-FolderRowSelected } })
    $script:CmbRights.Add_SelectionChanged({ Update-DelegateOptions })
    $script:ChkDelegate.Add_Click({ Update-DelegateOptions })
    $script:BtnFolderAdd.Add_Click({    Invoke-FolderAdd })
    $script:BtnFolderSet.Add_Click({    Invoke-FolderSet })
    $script:BtnFolderRemove.Add_Click({ Invoke-FolderRemove })

    # --- Toegang tot andere mailboxen ---
    $script:BtnScan.Add_Click({ if ($script:CurrentUser) { Invoke-AccessScan } })
    $script:BtnScanRemove.Add_Click({ Invoke-ScanRemove })

    # --- Bulk ---
    $script:BtnBulkEntraApply.Add_Click({ Invoke-BulkEntra })
    $script:BtnBulkEntraClear.Add_Click({
        foreach ($b in $script:BulkBoxes.Values) { $b.Box.Text = ''; $b.Chk.IsChecked = $false }
        $script:CmbBulkAccount.SelectedIndex = 0
        $script:ChkBulkMgr.IsChecked = $false
        $script:BulkManager = $null
        $script:TxtBulkMgr.Text = '(geen = manager verwijderen)'
    })
    $script:BtnBulkMgrPick.Add_Click({ Invoke-BulkPickManager })
    $script:BtnBulkMgrClear.Add_Click({
        $script:BulkManager = $null
        $script:TxtBulkMgr.Text = '(geen = manager verwijderen)'
    })
    foreach ($rb in $script:RbMbxA, $script:RbMbxB, $script:RbFolA, $script:RbFolB, $script:RbFolC) { $rb.Add_Click({ Update-BulkLabels }) }
    $script:BtnBulkMbxPick.Add_Click({   Invoke-BulkPickTarget 'Mbx' })
    $script:BtnBulkMbxGrant.Add_Click({  Invoke-BulkMbx -Grant $true })
    $script:BtnBulkMbxRevoke.Add_Click({ Invoke-BulkMbx -Grant $false })
    $script:BtnBulkFolPick.Add_Click({   Invoke-BulkPickTarget 'Fol' })
    $script:CmbBulkFolder.Add_SelectionChanged({ Update-BulkDelegateOptions })
    $script:CmbBulkRights.Add_SelectionChanged({ Update-BulkDelegateOptions })
    $script:ChkBulkDelegate.Add_Click({ Update-BulkDelegateOptions })
    $script:BtnBulkFolGrant.Add_Click({  Invoke-BulkFolder -Grant $true })
    $script:BtnBulkFolRevoke.Add_Click({ Invoke-BulkFolder -Grant $false })

    # --- Wachtrij ---
    $script:BtnRunQueue.Add_Click({ Invoke-Queue })
    $script:BtnQueueRemove.Add_Click({
        $keys = @(Get-SelectedRows $script:QueueGrid | ForEach-Object { [string]$_['Key'] })
        if ($keys.Count -eq 0) { Show-Msg "Selecteer eerst een of meer regels in de wachtrij." "Wachtrij" | Out-Null; return }
        $n = Remove-QueueItems $keys
        Update-QueueSummary
        if ($script:CurrentUser) { Set-EntraForm; Show-MailboxGrids; Show-FolderGrid; Show-ForwardState; Set-ForwardForm }
        Set-Status "$n item(s) uit de wachtrij verwijderd"
    })
    $script:BtnQueueClearDone.Add_Click({
        $keys = @($script:Queue | Where-Object { $_.Status -like 'Gereed*' } | ForEach-Object { $_.Key })
        [void](Remove-QueueItems $keys)
        Update-QueueSummary
    })
    $script:BtnQueueClear.Add_Click({
        if ($script:Queue.Count -eq 0) { return }
        $r = Show-Msg "Alle $($script:Queue.Count) items uit de wachtrij verwijderen? Niet-uitgevoerde wijzigingen gaan verloren." "Wachtrij leegmaken" "YesNo" "Warning"
        if ($r -ne 'Yes') { return }
        [void](Remove-QueueItems @($script:Queue | ForEach-Object { $_.Key }))
        Update-QueueSummary
        if ($script:CurrentUser) { Set-EntraForm; Show-MailboxGrids; Show-FolderGrid; Show-ForwardState; Set-ForwardForm }
    })

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
            return
        }
        $open = @($script:Queue | Where-Object { Test-QueueOpen $_ }).Count
        $msgs = @()
        if ($open) { $msgs += "$open wijziging(en) in de wachtrij zijn nog niet uitgevoerd." }
        if (Test-EntraDirty) { $msgs += "Er zijn formulierwijzigingen die nog niet in de wachtrij staan." }
        if ($msgs.Count) {
            $r = Show-Msg "$($msgs -join "`n")`n`nToch afsluiten? Deze wijzigingen gaan verloren." "Afsluiten" "YesNo" "Warning"
            if ($r -ne 'Yes') { $_.Cancel = $true }
        }
    })
    $script:Window.Add_Closed({
        if ($script:ExoConnected) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { } }
    })
}

function Show-MainWindow {
    New-MainWindow
    if (-not $script:ExoAvailable) {
        Write-Log "Module ExchangeOnlineManagement niet gevonden - Exchange-functies werken niet. Installeer met: Install-Module ExchangeOnlineManagement -Scope CurrentUser" "warning"
    }
    Write-Log "Alle wijzigingen gaan eerst naar de wachtrij (onder). Klik 'Alles uitvoeren' om ze in 1 keer door te voeren." "info"
    [void]$script:Window.ShowDialog()
}

#endregion

if (-not (Initialize-Modules)) {
    [System.Windows.MessageBox]::Show("Module Microsoft.Graph.Authentication niet gevonden.`n`nInstalleer met:`nInstall-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Scope CurrentUser", "Module ontbreekt", "OK", "Error") | Out-Null
    exit 1
}
Show-MainWindow
