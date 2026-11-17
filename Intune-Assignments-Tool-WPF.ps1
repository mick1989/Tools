#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Assignments Manager Tool (WPF).
.DESCRIPTION
    WPF-app met twee tabbladen:
      - Configuration Policies: Settings Catalog + klassieke Device Configuration
        profielen. Per policy groep(en) includen/excluden, of toewijzen aan
        Alle gebruikers / Alle apparaten.
      - Apps: alle Intune-apps (mobileApps). Per app groep(en) includen/excluden
        met een zichtbaar gekozen install intent (Required / Available /
        Uninstall / Available without enrollment). De intent van een bestaande
        toewijzing kan ook achteraf gewijzigd worden.
    Beide tabbladen ondersteunen het selecteren en bewerken van meerdere
    items tegelijk (bulk include/exclude).

    Policy-inhoud: dubbelklik op een policy (of 'Inhoud bekijken') toont alle
    instellingen. Bij meerdere geselecteerde policies worden de instellingen
    naast elkaar gezet (gelijk / conflict / uniek), te exporteren naar CSV.
    Settings Catalog policies met hetzelfde platform/technologie/template
    kunnen samengevoegd worden tot een nieuwe policy (bronnen blijven intact).

    Policies verwijderen (los of in bulk): na bevestiging wordt van elke
    policy eerst een JSON-backup (incl. instellingen en toewijzingen) gemaakt
    in de map 'Policy-backups' naast dit script; mislukt de backup, dan wordt
    die policy niet verwijderd.

    Bestaande assignments behouden bij opslaan hun oorspronkelijke target
    (incl. assignment filters) en app-instellingen (notificaties, delivery
    optimization, etc.).
.NOTES
    Vereist: Microsoft.Graph.Authentication module
    Installeren: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
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
$script:GraphBeta      = "https://graph.microsoft.com/beta"
$script:GroupNameCache = @{}
$script:Tabs           = @{}
$script:Window         = $null
$script:NameAllUsers   = "Alle gebruikers"
$script:NameAllDevices = "Alle apparaten"
$script:BackupRoot     = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Policy-backups' }
                         else { Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Intune-Policy-backups' }

$script:AppIntents = @(
    [pscustomobject]@{ Label = "Required";                     Value = "required" }
    [pscustomobject]@{ Label = "Available";                    Value = "available" }
    [pscustomobject]@{ Label = "Uninstall";                    Value = "uninstall" }
    [pscustomobject]@{ Label = "Available without enrollment"; Value = "availableWithoutEnrollment" }
)

$script:AppTypeLabels = @{
    'win32LobApp'              = 'Win32'
    'winGetApp'                = 'Microsoft Store (WinGet)'
    'officeSuiteApp'           = 'Microsoft 365 Apps'
    'windowsMicrosoftEdgeApp'  = 'Microsoft Edge'
    'windowsMobileMSI'         = 'MSI (LOB)'
    'windowsUniversalAppX'     = 'AppX/MSIX (LOB)'
    'windowsWebApp'            = 'Windows web link'
    'webApp'                   = 'Web link'
    'iosVppApp'                = 'iOS VPP'
    'iosStoreApp'              = 'iOS Store'
    'iosLobApp'                = 'iOS LOB'
    'managedIOSStoreApp'       = 'iOS Store (managed)'
    'androidManagedStoreApp'   = 'Android Managed Store'
    'androidStoreApp'          = 'Android Store'
    'androidLobApp'            = 'Android LOB'
    'managedAndroidStoreApp'   = 'Android Store (managed)'
    'macOSLobApp'              = 'macOS LOB'
    'macOSPkgApp'              = 'macOS PKG'
    'macOSDmgApp'              = 'macOS DMG'
    'macOSMicrosoftEdgeApp'    = 'macOS Edge'
    'macOSOfficeSuiteApp'      = 'macOS Office'
}

#endregion

#region --- UI basis-helpers ---

function New-Brush {
    param([string]$Hex)
    $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
    $b.Freeze()
    return $b
}

function Update-UI {
    # Laat WPF de wachtrij verwerken zodat statusbalk/log direct zichtbaar
    # worden tijdens langlopende Graph-calls op de UI-thread.
    if ($script:Window) {
        $script:Window.Dispatcher.Invoke([action]{ }, [System.Windows.Threading.DispatcherPriority]::Background)
    }
}

function Show-Msg {
    param([string]$Text, [string]$Title = "Info", [string]$Button = "OK", [string]$Icon = "Information", $Owner = $null)
    if ($Owner) { return [System.Windows.MessageBox]::Show($Owner, $Text, $Title, $Button, $Icon) }
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
    $script:MainTabs.IsEnabled   = $false
    $script:BtnConnect.IsEnabled = $false
    $script:Window.Cursor = [System.Windows.Input.Cursors]::Wait
    if ($Text) { Set-Status $Text }
}

function Exit-Busy {
    $script:MainTabs.IsEnabled   = $true
    $script:BtnConnect.IsEnabled = $true
    $script:Window.Cursor = $null
}

function New-Table {
    param([string[]]$Columns)
    $t = New-Object System.Data.DataTable
    foreach ($c in $Columns) { [void]$t.Columns.Add($c) }
    return ,$t
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

#region --- Graph verbinding ---

function Connect-ToGraph {
    param([string]$LoginHint)
    $scopes = @(
        "DeviceManagementConfiguration.ReadWrite.All",
        "DeviceManagementServiceConfig.ReadWrite.All",
        "DeviceManagementApps.ReadWrite.All",
        "Group.Read.All"
    )
    try {
        # Vanaf module-versie 2.34.0 dwingt Microsoft.Graph.Authentication WAM
        # (Windows Account Manager) af, waardoor Connect-MgGraph stilzwijgend het
        # al gecachte Windows-account gebruikt zonder inlogscherm of account-kiezer.
        # Versie 2.33.0 gebruikt nog de normale browser-login met account-kiezer,
        # dus die pinnen we hier expliciet vast.
        if (-not (Get-Module -Name Microsoft.Graph.Authentication)) {
            try {
                Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -ErrorAction Stop
            }
            catch {
                Write-Warning "Microsoft.Graph.Authentication 2.33.0 niet gevonden, gebruik de standaard versie (kan WAM afdwingen)."
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            }
        }
        # Ontkoppel een eventuele bestaande (mogelijk verlopen of met te weinig
        # scopes gecachte) sessie eerst, anders hergebruikt Connect-MgGraph soms
        # stilzwijgend een oud token zonder de hier gevraagde scopes.
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        $p = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($LoginHint) { $p.LoginHint = $LoginHint }
        Connect-MgGraph @p | Out-Null
        return $true
    }
    catch {
        Show-Msg "Verbinding mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" | Out-Null
        return $false
    }
}

function Invoke-Graph {
    param([string]$Uri, [string]$Method = "GET", $Body = $null)
    $short = $Uri -replace 'https://graph\.microsoft\.com/(v1\.0|beta)/', ''
    Write-Log "$Method $short" "api"
    $p = @{ Uri = $Uri; Method = $Method }
    # Settings Catalog instanties kunnen diep genest zijn (groep > keuze > kinderen ...)
    if ($Body) { $p.Body = ($Body | ConvertTo-Json -Depth 50); $p.ContentType = "application/json" }
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
        Write-Log "FOUT bij $Method $short`: $errMsg" "error"
        throw $errMsg
    }
}

function Get-AllGraphPages {
    param([string]$Uri)
    $all = @()
    do {
        $r = Invoke-Graph -Uri $Uri
        if ($r.value) { $all += $r.value }
        $Uri = $r.'@odata.nextLink'
    } while ($Uri)
    return $all
}

function Get-AllGraphPagesWithAssignments {
    # Probeert assignments direct mee te laden via $expand (1 call per pagina
    # i.p.v. 1 call per item). Valt terug op zonder $expand als de endpoint
    # dat niet ondersteunt; de assignments worden dan per item opgehaald.
    param([string]$BaseUri, [string]$Select)
    try {
        return Get-AllGraphPages -Uri "$BaseUri`?`$select=$Select&`$expand=assignments"
    }
    catch {
        Write-Log "`$expand=assignments niet ondersteund, assignments worden per item opgehaald" "warning"
        return Get-AllGraphPages -Uri "$BaseUri`?`$select=$Select"
    }
}

#endregion

#region --- Intune / Graph domein ---

function Get-GraphRootForKind {
    param([string]$Kind)
    # De Settings Catalog (configurationPolicies) resource bestaat alleen onder /beta;
    # klassieke deviceConfigurations profielen zitten wel gewoon op v1.0.
    if ($Kind -eq 'configurationPolicies') { return $script:GraphBeta }
    return $script:GraphRoot
}

function Get-AppIntentLabel {
    param([string]$Value)
    $m = $script:AppIntents | Where-Object { $_.Value -eq $Value } | Select-Object -First 1
    if ($m) { return $m.Label }
    return $Value
}

function Get-AllPolicies {
    $sc = Get-AllGraphPagesWithAssignments -BaseUri "$script:GraphBeta/deviceManagement/configurationPolicies" -Select "id,name" |
        ForEach-Object {
            [pscustomobject]@{
                Key = "configurationPolicies_$($_.id)"; Id = $_.id; Name = $_.name
                Kind = 'configurationPolicies'; TypeLabel = 'Settings Catalog'
                Assignments = @($_.assignments); AssignmentsLoaded = $_.ContainsKey('assignments')
            }
        }
    $dc = Get-AllGraphPagesWithAssignments -BaseUri "$script:GraphRoot/deviceManagement/deviceConfigurations" -Select "id,displayName" |
        ForEach-Object {
            [pscustomobject]@{
                Key = "deviceConfigurations_$($_.id)"; Id = $_.id; Name = $_.displayName
                Kind = 'deviceConfigurations'; TypeLabel = 'Device Config'
                Assignments = @($_.assignments); AssignmentsLoaded = $_.ContainsKey('assignments')
            }
        }
    return @($sc) + @($dc)
}

function Get-AllApps {
    $apps = Get-AllGraphPagesWithAssignments -BaseUri "$script:GraphRoot/deviceAppManagement/mobileApps" -Select "id,displayName"
    return @($apps | ForEach-Object {
        $type  = ($_.'@odata.type' -replace '#microsoft\.graph\.', '')
        $label = if ($script:AppTypeLabels.ContainsKey($type)) { $script:AppTypeLabels[$type] } else { $type }
        [pscustomobject]@{
            Key = $_.id; Id = $_.id; Name = $_.displayName; Kind = 'mobileApps'; TypeLabel = $label
            Assignments = @($_.assignments); AssignmentsLoaded = $_.ContainsKey('assignments')
        }
    })
}

function Get-ItemAssignments {
    param($Ctx, $Item)
    if ($Ctx.IsApp) {
        $r = Invoke-Graph -Uri "$script:GraphRoot/deviceAppManagement/mobileApps/$($Item.Id)/assignments"
    }
    else {
        $root = Get-GraphRootForKind -Kind $Item.Kind
        $r = Invoke-Graph -Uri "$root/deviceManagement/$($Item.Kind)/$($Item.Id)/assignments"
    }
    return @($r.value)
}

function New-Entry {
    # Een include/exclude-regel. Target is $null voor nieuw toegevoegde groepen
    # (wordt bij opslaan opgebouwd); bestaande regels bewaren het originele
    # target (incl. assignment filter) en eventuele app-settings.
    param([string]$Id, [string]$Name, [string]$Intent, [bool]$IsGroup = $true, $Target = $null, $Settings = $null)
    return [pscustomobject]@{ Id = $Id; Name = $Name; IsGroup = $IsGroup; Intent = $Intent; Target = $Target; Settings = $Settings }
}

function New-SpecialEntry {
    param([ValidateSet('allUsers', 'allDevices')][string]$Kind, [string]$Intent)
    if ($Kind -eq 'allUsers') {
        return New-Entry -Id 'allUsers' -Name $script:NameAllUsers -Intent $Intent -IsGroup $false `
            -Target @{ '@odata.type' = '#microsoft.graph.allLicensedUsersAssignmentTarget' }
    }
    return New-Entry -Id 'allDevices' -Name $script:NameAllDevices -Intent $Intent -IsGroup $false `
        -Target @{ '@odata.type' = '#microsoft.graph.allDevicesAssignmentTarget' }
}

function Split-Assignments {
    param($Assignments)
    $include = New-Object System.Collections.ArrayList
    $exclude = New-Object System.Collections.ArrayList
    $other   = New-Object System.Collections.ArrayList

    foreach ($a in @($Assignments)) {
        if (-not $a) { continue }
        $t = $a.target
        switch ($t.'@odata.type') {
            '#microsoft.graph.groupAssignmentTarget' {
                [void]$include.Add((New-Entry -Id $t.groupId -Intent $a.intent -Target $t -Settings $a.settings))
            }
            '#microsoft.graph.exclusionGroupAssignmentTarget' {
                [void]$exclude.Add((New-Entry -Id $t.groupId -Intent $a.intent -Target $t -Settings $a.settings))
            }
            '#microsoft.graph.allLicensedUsersAssignmentTarget' {
                [void]$include.Add((New-Entry -Id 'allUsers' -Name $script:NameAllUsers -IsGroup $false -Intent $a.intent -Target $t -Settings $a.settings))
            }
            '#microsoft.graph.allDevicesAssignmentTarget' {
                [void]$include.Add((New-Entry -Id 'allDevices' -Name $script:NameAllDevices -IsGroup $false -Intent $a.intent -Target $t -Settings $a.settings))
            }
            default { [void]$other.Add($a) }
        }
    }
    return [pscustomobject]@{ Include = $include; Exclude = $exclude; Other = $other }
}

function Resolve-GroupNames {
    param([string[]]$Ids)
    $unique  = @($Ids | Where-Object { $_ } | Select-Object -Unique)
    $toFetch = @($unique | Where-Object { -not $script:GroupNameCache.ContainsKey($_) })

    for ($i = 0; $i -lt $toFetch.Count; $i += 900) {
        $end   = [Math]::Min($i + 899, $toFetch.Count - 1)
        $chunk = @($toFetch[$i..$end])
        $body  = @{ ids = $chunk; types = @('group') }
        try {
            $r = Invoke-Graph -Uri "$script:GraphRoot/directoryObjects/getByIds" -Method POST -Body $body
            foreach ($obj in $r.value) { $script:GroupNameCache[$obj.id] = $obj.displayName }
        } catch { }
        foreach ($id in $chunk) {
            if (-not $script:GroupNameCache.ContainsKey($id)) { $script:GroupNameCache[$id] = "(onbekende groep)" }
        }
    }
}

function Set-EntryNames {
    param($DataList)
    $ids = foreach ($d in $DataList) {
        foreach ($e in (@($d.Include) + @($d.Exclude))) { if ($e.IsGroup) { $e.Id } }
    }
    Resolve-GroupNames -Ids @($ids)
    foreach ($d in $DataList) {
        foreach ($e in (@($d.Include) + @($d.Exclude))) {
            if ($e.IsGroup) { $e.Name = $script:GroupNameCache[$e.Id] }
        }
    }
}

function Search-Groups {
    param([string]$Query)
    $escaped = $Query -replace "'", "''"
    $uri = "$script:GraphRoot/groups?`$filter=startswith(displayName,'$escaped')&`$select=id,displayName&`$top=50"
    return @(Get-AllGraphPages -Uri $uri | Sort-Object displayName)
}

function New-EntryPayload {
    param($Entry, [bool]$Exclude, [bool]$IsApp)
    $target = $Entry.Target
    if (-not $target) {
        $type = if ($Exclude) { '#microsoft.graph.exclusionGroupAssignmentTarget' } else { '#microsoft.graph.groupAssignmentTarget' }
        $target = @{ '@odata.type' = $type; groupId = $Entry.Id }
    }
    $p = @{ target = $target }
    if ($IsApp) {
        $p.intent = if ($Entry.Intent) { $Entry.Intent } else { 'required' }
        if ($Entry.Settings) { $p.settings = $Entry.Settings }
    }
    return $p
}

function Save-ItemAssignments {
    param($Ctx, $Data)
    $payload = New-Object System.Collections.ArrayList
    foreach ($o in $Data.Other) {
        $p = @{ target = $o.target }
        if ($Ctx.IsApp) {
            $p.intent = $o.intent
            if ($o.settings) { $p.settings = $o.settings }
        }
        [void]$payload.Add($p)
    }
    foreach ($e in $Data.Include) { [void]$payload.Add((New-EntryPayload -Entry $e -Exclude $false -IsApp $Ctx.IsApp)) }
    foreach ($e in $Data.Exclude) { [void]$payload.Add((New-EntryPayload -Entry $e -Exclude $true  -IsApp $Ctx.IsApp)) }
    $arr = $payload.ToArray()

    $item = $Data.Item
    if ($Ctx.IsApp) {
        Invoke-Graph -Uri "$script:GraphRoot/deviceAppManagement/mobileApps/$($item.Id)/assign" -Method POST -Body @{ mobileAppAssignments = $arr } | Out-Null
    }
    else {
        $root = Get-GraphRootForKind -Kind $item.Kind
        Invoke-Graph -Uri "$root/deviceManagement/$($item.Kind)/$($item.Id)/assign" -Method POST -Body @{ assignments = $arr } | Out-Null
    }
}

function Backup-Policy {
    # Schrijft de volledige policy (incl. instellingen en toewijzingen) als JSON weg
    param($Item, [string]$Folder)
    if (-not (Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $root = Get-GraphRootForKind -Kind $Item.Kind
    $base = "$root/deviceManagement/$($Item.Kind)/$($Item.Id)"
    if ($Item.Kind -eq 'configurationPolicies') {
        $obj = Invoke-Graph -Uri $base
        $obj['settings'] = @(Get-AllGraphPages -Uri "$base/settings")
    }
    else {
        try { $obj = Invoke-Graph -Uri "$script:GraphBeta/deviceManagement/deviceConfigurations/$($Item.Id)" }
        catch { $obj = Invoke-Graph -Uri $base }
    }
    $obj['assignments'] = @((Invoke-Graph -Uri "$base/assignments").value)
    $safe = $Item.Name -replace '[\\/:*?"<>|]', '_'
    $file = Join-Path $Folder "$($safe)_$($Item.Id).json"
    ConvertTo-Json -InputObject $obj -Depth 50 | Set-Content -Path $file -Encoding UTF8
    return $file
}

function Remove-Policy {
    param($Item)
    $root = Get-GraphRootForKind -Kind $Item.Kind
    Invoke-Graph -Uri "$root/deviceManagement/$($Item.Kind)/$($Item.Id)" -Method DELETE | Out-Null
}

#endregion

#region --- Policy-inhoud (instellingen bekijken / vergelijken / samenvoegen) ---

# Metadata-eigenschappen van klassieke profielen die geen instelling zijn
$script:DcSkipProps = @(
    'id', 'createdDateTime', 'lastModifiedDateTime', 'displayName', 'description', 'version',
    'roleScopeTagIds', 'supportsScopeTags', 'assignments',
    'deviceManagementApplicabilityRuleOsEdition', 'deviceManagementApplicabilityRuleOsVersion',
    'deviceManagementApplicabilityRuleDeviceMode'
)

function Add-ContentRow {
    # Key is de vergelijkingssleutel (gelijk tussen policies = dezelfde instelling)
    param($Rows, [string]$Key, [string]$Setting, [string]$Path, [string]$Value, [bool]$Configured = $true)
    [void]$Rows.Add([pscustomobject]@{ Key = $Key; Setting = $Setting; Path = $Path; Value = $Value; Configured = $Configured })
}

function Get-SettingDefName {
    param($Defs, [string]$Id)
    $d = $Defs[$Id]
    if ($d -and $d.displayName) { return [string]$d.displayName }
    if ($d -and $d.name) { return [string]$d.name }
    return $Id
}

function Get-ChoiceLabel {
    param($Def, $Value)
    if ($null -eq $Value) { return '' }
    if ($Def) {
        foreach ($o in @($Def.options)) {
            if ($o -and $o.itemId -eq $Value) {
                if ($o.displayName) { return [string]$o.displayName }
                if ($o.name) { return [string]$o.name }
            }
        }
    }
    return [string]$Value
}

function Format-SimpleSettingValue {
    param($V)
    if (-not $V) { return '' }
    if ([string]$V.'@odata.type' -like '*SecretSettingValue') { return '(geheim)' }
    return [string]$V.value
}

function Add-SettingInstanceRows {
    # Zet een (geneste) Settings Catalog settingInstance om naar platte regels.
    # Groepen leveren zelf geen regel op, alleen hun kinderen.
    param($Inst, $Defs, [string]$ParentPath, [string]$ParentKey, $Rows)
    if (-not $Inst) { return }
    $defId = [string]$Inst.settingDefinitionId
    $def   = $Defs[$defId]
    $name  = Get-SettingDefName $Defs $defId
    $key   = if ($ParentKey) { "$ParentKey/$defId" } else { $defId }
    $here  = if ($ParentPath) { "$ParentPath > $name" } else { $name }
    $type  = ([string]$Inst.'@odata.type') -replace '#microsoft\.graph\.deviceManagementConfiguration', ''

    switch ($type) {
        'ChoiceSettingInstance' {
            $v = $Inst.choiceSettingValue
            Add-ContentRow $Rows $key $name $ParentPath (Get-ChoiceLabel $def $v.value)
            foreach ($ch in @($v.children)) { Add-SettingInstanceRows $ch $Defs $here $key $Rows }
        }
        'ChoiceSettingCollectionInstance' {
            $vals = @($Inst.choiceSettingCollectionValue | Where-Object { $_ })
            Add-ContentRow $Rows $key $name $ParentPath (@($vals | ForEach-Object { Get-ChoiceLabel $def $_.value }) -join '; ')
            foreach ($v in $vals) {
                foreach ($ch in @($v.children)) { Add-SettingInstanceRows $ch $Defs $here $key $Rows }
            }
        }
        'SimpleSettingInstance' {
            Add-ContentRow $Rows $key $name $ParentPath (Format-SimpleSettingValue $Inst.simpleSettingValue)
        }
        'SimpleSettingCollectionInstance' {
            $vals = @($Inst.simpleSettingCollectionValue | Where-Object { $_ } | ForEach-Object { Format-SimpleSettingValue $_ })
            Add-ContentRow $Rows $key $name $ParentPath ($vals -join '; ')
        }
        'GroupSettingInstance' {
            foreach ($ch in @($Inst.groupSettingValue.children)) { Add-SettingInstanceRows $ch $Defs $here $key $Rows }
        }
        'GroupSettingCollectionInstance' {
            $items = @($Inst.groupSettingCollectionValue | Where-Object { $_ })
            for ($i = 0; $i -lt $items.Count; $i++) {
                $p = $here; $k = $key
                if ($items.Count -gt 1) { $p = "$here [$($i + 1)]"; $k = "$key[$i]" }
                foreach ($ch in @($items[$i].children)) { Add-SettingInstanceRows $ch $Defs $p $k $Rows }
            }
        }
        default {
            Add-ContentRow $Rows $key $name $ParentPath (ConvertTo-Json -InputObject $Inst -Compress -Depth 20)
        }
    }
}

function ConvertTo-FriendlyName {
    param([string]$Name)
    $s = $Name -creplace '(?<=[a-z0-9])(?=[A-Z])', ' '
    if ($s.Length -gt 0) { $s = $s.Substring(0, 1).ToUpper() + $s.Substring(1) }
    return $s
}

function Test-ValueConfigured {
    # Klassieke profielen geven ALLE eigenschappen terug; null / '' /
    # notConfigured / false / lege lijst betekent in de praktijk "niet ingesteld".
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [string]) { return ($Value -ne '' -and $Value -ne 'notConfigured') }
    if ($Value -is [System.Collections.ICollection]) { return ($Value.Count -gt 0) }
    return $true
}

function Format-PlainValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { return $Value.ToString().ToLower() }
    if ($Value -is [System.Collections.IDictionary]) { return (ConvertTo-Json -InputObject $Value -Compress -Depth 10) }
    if ($Value -is [System.Array]) { return (@($Value | ForEach-Object { Format-PlainValue $_ }) -join '; ') }
    return [string]$Value
}

function Add-OmaSettingRows {
    param($OmaSettings, $Rows, [string]$PolicyId)
    foreach ($o in @($OmaSettings)) {
        if (-not $o) { continue }
        $val = $o.value
        if ($o.isEncrypted -and $o.secretReferenceValueId) {
            try {
                $r = Invoke-Graph -Uri "$script:GraphBeta/deviceManagement/deviceConfigurations/$PolicyId/getOmaSettingPlainTextValue(secretReferenceValueId='$($o.secretReferenceValueId)')"
                $val = $r.value
            }
            catch { $val = '(versleuteld - niet leesbaar)' }
        }
        if ([string]$o.'@odata.type' -like '*StringXml' -and $val) {
            try { $val = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$val)) } catch { }
        }
        $label = if ($o.displayName) { [string]$o.displayName } else { [string]$o.omaUri }
        Add-ContentRow $Rows "oma:$($o.omaUri)" $label "OMA-URI: $($o.omaUri)" (Format-PlainValue $val) $true
    }
}

function Add-DeviceConfigRows {
    param($Obj, [string]$KeyPrefix, [string]$PathPrefix, $Rows, [int]$Depth, [string]$PolicyId)
    foreach ($k in @($Obj.Keys | Sort-Object)) {
        if ($k -like '*@odata*') { continue }
        if ($Depth -eq 0 -and $script:DcSkipProps -contains $k) { continue }
        $v    = $Obj[$k]
        $name = ConvertTo-FriendlyName $k
        if ($Depth -eq 0 -and $k -eq 'omaSettings') {
            Add-OmaSettingRows -OmaSettings $v -Rows $Rows -PolicyId $PolicyId
            continue
        }
        if ($v -is [System.Collections.IDictionary] -and $Depth -lt 3) {
            $sub = if ($PathPrefix) { "$PathPrefix > $name" } else { $name }
            Add-DeviceConfigRows -Obj $v -KeyPrefix "$KeyPrefix$k." -PathPrefix $sub -Rows $Rows -Depth ($Depth + 1) -PolicyId $PolicyId
            continue
        }
        Add-ContentRow $Rows "dc:$KeyPrefix$k" $name $PathPrefix (Format-PlainValue $v) (Test-ValueConfigured $v)
    }
}

function Get-PolicyContent {
    # Haalt alle instellingen van 1 policy op als platte regels. Voor Settings
    # Catalog worden ook de ruwe settingInstances bewaard (nodig voor samenvoegen).
    param($Item)
    $rows = New-Object System.Collections.ArrayList
    $c = [pscustomobject]@{
        Item = $Item; Info = ''; Rows = $rows; IsSettingsCatalog = ($Item.Kind -eq 'configurationPolicies')
        Raw = @(); TopNames = @{}; Platforms = ''; Technologies = ''; TemplateId = ''; RoleScopeTagIds = @('0')
    }

    if ($c.IsSettingsCatalog) {
        $base = "$script:GraphBeta/deviceManagement/configurationPolicies/$($Item.Id)"
        $meta = Invoke-Graph -Uri $base
        $settings = Get-AllGraphPages -Uri "$base/settings?`$expand=settingDefinitions"
        $raw = New-Object System.Collections.ArrayList
        foreach ($s in $settings) {
            $inst = $s.settingInstance
            if (-not $inst) { continue }
            [void]$raw.Add($inst)
            $defs = @{}
            foreach ($d in @($s.settingDefinitions)) { if ($d) { $defs[$d.id] = $d } }
            $c.TopNames[[string]$inst.settingDefinitionId] = Get-SettingDefName $defs $inst.settingDefinitionId
            Add-SettingInstanceRows $inst $defs '' '' $rows
        }
        $c.Raw          = $raw.ToArray()
        $c.Platforms    = [string]$meta.platforms
        $c.Technologies = [string]$meta.technologies
        $tpl = $meta.templateReference
        if ($tpl -and $tpl.templateId) { $c.TemplateId = [string]$tpl.templateId }
        if ($meta.roleScopeTagIds) { $c.RoleScopeTagIds = @($meta.roleScopeTagIds) }
        $tplTxt = if ($tpl -and $tpl.templateDisplayName) { " | Template: $($tpl.templateDisplayName)" } else { '' }
        $c.Info = "Settings Catalog | Platform: $($c.Platforms) | Technologie: $($c.Technologies)$tplTxt | $($raw.Count) instelling(en)"
    }
    else {
        # Beta geeft voor veel profieltypes meer eigenschappen terug dan v1.0
        try { $obj = Invoke-Graph -Uri "$script:GraphBeta/deviceManagement/deviceConfigurations/$($Item.Id)" }
        catch { $obj = Invoke-Graph -Uri "$script:GraphRoot/deviceManagement/deviceConfigurations/$($Item.Id)" }
        $odType = ([string]$obj.'@odata.type') -replace '#microsoft\.graph\.', ''
        Add-DeviceConfigRows -Obj $obj -KeyPrefix '' -PathPrefix '' -Rows $rows -Depth 0 -PolicyId $Item.Id
        $conf = @($rows | Where-Object { $_.Configured }).Count
        $c.Info = "Device Configuration ($odType) | $conf geconfigureerd van $($rows.Count) eigenschappen"
    }
    return $c
}

function Get-UniqueNames {
    param([string[]]$Names)
    $seen = @{}
    $out = foreach ($n in $Names) {
        $b = $n; $i = 2
        while ($seen.ContainsKey($n)) { $n = "$b ($i)"; $i++ }
        $seen[$n] = 1
        $n
    }
    return @($out)
}

function Build-PolicyCompareTable {
    # 1 regel per instelling-sleutel, 1 waardekolom (P1..Pn) per policy.
    # State: same / conflict / unique (niet in alle policies) / single.
    param($Contents)
    $n = $Contents.Count
    $t = New-Table (@('Key', 'Instelling', 'Pad') + @(1..$n | ForEach-Object { "P$_" }) + @('Status', 'State', 'Configured'))
    $map = [ordered]@{}
    for ($i = 0; $i -lt $n; $i++) {
        foreach ($r in $Contents[$i].Rows) {
            $e = $map[$r.Key]
            if (-not $e) {
                $e = @{ Setting = $r.Setting; Path = $r.Path; Values = @{}; Configured = $false }
                $map[$r.Key] = $e
            }
            if ($e.Values.ContainsKey($i)) { $e.Values[$i] = "$($e.Values[$i]); $($r.Value)" }
            else { $e.Values[$i] = $r.Value }
            if ($r.Configured) { $e.Configured = $true }
        }
    }
    $names = @($Contents | ForEach-Object { $_.Item.Name })
    foreach ($k in $map.Keys) {
        $e = $map[$k]
        $present  = @($e.Values.Keys)
        $distinct = @($e.Values.Values | Select-Object -Unique)
        if ($n -eq 1)                   { $state = 'single';   $status = '' }
        elseif ($distinct.Count -gt 1)  { $state = 'conflict'; $status = 'Conflict' }
        elseif ($present.Count -lt $n) {
            $state  = 'unique'
            $status = if ($present.Count -eq 1) { "Alleen in: $($names[$present[0]])" } else { "In $($present.Count) van $n" }
        }
        else                            { $state = 'same';     $status = 'Gelijk' }
        $vals = New-Object System.Collections.ArrayList
        [void]$vals.AddRange(@($k, $e.Setting, $e.Path))
        for ($i = 0; $i -lt $n; $i++) { [void]$vals.Add([string]$e.Values[$i]) }
        [void]$vals.AddRange(@($status, $state, $(if ($e.Configured) { '1' } else { '0' })))
        [void]$t.Rows.Add([object[]]$vals.ToArray())
    }
    return ,$t
}

function Get-MergePlan {
    # Unie van alle top-level settingInstances. Identiek = 1x opnemen; zelfde
    # instelling met andere waarde = conflict, de policy met voorrang wint.
    param($Contents, [int]$PrioIndex = 0)
    $order = @($Contents[$PrioIndex]) + @(for ($i = 0; $i -lt $Contents.Count; $i++) { if ($i -ne $PrioIndex) { $Contents[$i] } })
    $byId      = @{}
    $settings  = New-Object System.Collections.ArrayList
    $conflicts = New-Object System.Collections.ArrayList
    $dup = 0
    foreach ($c in $order) {
        foreach ($inst in $c.Raw) {
            $id   = [string]$inst.settingDefinitionId
            $json = ConvertTo-Json -InputObject $inst -Depth 50 -Compress
            if ($byId.ContainsKey($id)) {
                if ($byId[$id].Json -ceq $json) { $dup++ }
                else { [void]$conflicts.Add("$($c.TopNames[$id]):  '$($byId[$id].Owner)' wint van '$($c.Item.Name)'") }
                continue
            }
            $byId[$id] = @{ Json = $json; Owner = $c.Item.Name }
            [void]$settings.Add($inst)
        }
    }
    return [pscustomobject]@{ Settings = $settings; Conflicts = $conflicts; Duplicates = $dup }
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
<Style x:Key="BtnPurple" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#5C2D91"/></Style>
<Style x:Key="BtnDark"  TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="#005A9E"/><Setter Property="Padding" Value="16,8"/><Setter Property="Margin" Value="0"/>
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
  <Setter Property="MinRowHeight" Value="26"/>
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
<Style x:Key="CellStatus" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
  <Setter Property="Foreground" Value="#C42B1C"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
</Style>
<Style TargetType="ComboBox">
  <Setter Property="Height" Value="28"/>
  <Setter Property="VerticalContentAlignment" Value="Center"/>
</Style>
'@

function Get-TabXaml {
    param([string]$P, [bool]$IsApp)
    $what     = if ($IsApp) { 'app' } else { 'policy' }
    $loadText = if ($IsApp) { 'Apps laden' } else { 'Policies laden' }

    $intentCol = ''
    $intentBar = ''
    $bulkIntent = ''
    $viewBtn = '<Button x:Name="' + $P + '_BtnView" Content="Inhoud bekijken" Style="{StaticResource BtnPurple}" ToolTip="Toon alle instellingen van deze policy (ook via dubbelklik in de lijst)"/>' +
               '<Button x:Name="' + $P + '_BtnDelete" Content="Policy verwijderen" Style="{StaticResource BtnRed}" ToolTip="Verwijdert de policy uit Intune (na bevestiging, met JSON-backup)"/>'
    $bulkView = @"
<WrapPanel Margin="0,0,0,12">
  <Button x:Name="${P}_BtnBulkView" Content="Inhoud vergelijken / samenvoegen" Style="{StaticResource BtnPurple}" ToolTip="Zet de instellingen van alle geselecteerde policies naast elkaar"/>
  <Button x:Name="${P}_BtnBulkDelete" Content="Geselecteerde policies verwijderen" Style="{StaticResource BtnRed}" ToolTip="Verwijdert alle geselecteerde policies uit Intune (na bevestiging, met JSON-backup)"/>
</WrapPanel>
"@
    $bulkHint = 'Kies een groep om als include of exclude toe te voegen aan alle geselecteerde policies. Een groep die al in de andere richting staat wordt automatisch verplaatst.'
    if ($IsApp) {
        $viewBtn  = ''
        $bulkView = ''
        $intentCol = '<DataGridTextColumn Header="Intent" Binding="{Binding Intent}" Width="190" ElementStyle="{StaticResource CellText}"/>'
        $intentBar = @"
<Border Grid.Row="0" Background="#F3F8FE" BorderBrush="#C7DDF5" BorderThickness="1" CornerRadius="4" Padding="10,8,10,4" Margin="0,4,0,6">
  <StackPanel>
    <WrapPanel>
      <TextBlock Text="Install intent:" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,8,6"/>
      <ComboBox x:Name="${P}_CmbIntent" Width="210" Margin="0,0,8,6"/>
      <Button x:Name="${P}_BtnChangeIntent" Content="Intent van selectie wijzigen" Style="{StaticResource BtnBlue}"/>
    </WrapPanel>
    <TextBlock Foreground="#555555" TextWrapping="Wrap" Margin="0,0,0,4" Text="Deze intent wordt gebruikt bij toevoegen (include en exclude). 'Available' werkt alleen voor gebruikersgroepen of Alle gebruikers, niet voor apparaatgroepen."/>
  </StackPanel>
</Border>
"@
        $bulkIntent = @"
<WrapPanel Margin="0,0,0,4">
  <TextBlock Text="Install intent:" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,8,6"/>
  <ComboBox x:Name="${P}_CmbBulkIntent" Width="210" Margin="0,0,8,6"/>
</WrapPanel>
"@
        $bulkHint = 'Kies de install intent en daarna een groep om als include of exclude toe te voegen aan alle geselecteerde apps. Een groep die al in de andere richting staat wordt automatisch verplaatst.'
    }

    return @"
<Grid>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="3*" MinWidth="450"/>
    <ColumnDefinition Width="5"/>
    <ColumnDefinition Width="2*" MinWidth="420"/>
  </Grid.ColumnDefinitions>

  <DockPanel Grid.Column="0">
    <Border DockPanel.Dock="Top" Background="White" Padding="8,8,8,2" BorderBrush="#E0E0E0" BorderThickness="0,0,0,1">
      <DockPanel LastChildFill="False">
        <Button x:Name="${P}_BtnLoad" Content="$loadText" Style="{StaticResource BtnBlue}"/>
        <TextBlock Text="Zoeken:" VerticalAlignment="Center" Margin="6,0,6,6"/>
        <TextBox x:Name="${P}_TxtFilter" Width="220" Height="28" VerticalContentAlignment="Center" Margin="0,0,10,6" ToolTip="Zoekt in naam en in include/exclude groepen"/>
        <TextBlock x:Name="${P}_LblCount" VerticalAlignment="Center" Foreground="Gray" Margin="0,0,0,6"/>
        <Button x:Name="${P}_BtnSaveAll" DockPanel.Dock="Right" Content="Alle wijzigingen opslaan" Style="{StaticResource BtnGreen}" Margin="0,0,0,6"/>
      </DockPanel>
    </Border>
    <DataGrid x:Name="${P}_Grid" SelectionMode="Extended" BorderThickness="0">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Type"    Binding="{Binding Type}"    Width="130" ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Naam"    Binding="{Binding Naam}"    Width="2*"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Include" Binding="{Binding Include}" Width="2*"  ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Exclude" Binding="{Binding Exclude}" Width="*"   ElementStyle="{StaticResource CellText}"/>
        <DataGridTextColumn Header="Status"  Binding="{Binding Status}"  Width="120" ElementStyle="{StaticResource CellStatus}"/>
      </DataGrid.Columns>
    </DataGrid>
  </DockPanel>

  <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#D0D0D0" ResizeDirection="Columns"/>

  <Border x:Name="${P}_Editor" Grid.Column="2" Background="White" Padding="12" IsEnabled="False">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <TextBlock x:Name="${P}_LblSelected" Text="Selecteer een $what" FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap"/>
      <TextBlock x:Name="${P}_LblDirty" Grid.Row="1" Foreground="#C42B1C" FontWeight="SemiBold" Margin="0,2,0,2"/>

      <Grid x:Name="${P}_SinglePanel" Grid.Row="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="3*"/>
          <RowDefinition Height="2*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        $intentBar
        <GroupBox Grid.Row="1" Margin="0,2,0,6" Padding="6" BorderBrush="#CCCCCC">
          <GroupBox.Header><TextBlock Text="Include - deze groep(en) krijgen de $what" FontWeight="SemiBold" Foreground="#107C10"/></GroupBox.Header>
          <DockPanel>
            <WrapPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
              <Button x:Name="${P}_BtnAddInclude"    Content="+ Groep"           Style="{StaticResource BtnGreen}"/>
              <Button x:Name="${P}_BtnAddAllUsers"   Content="+ Alle gebruikers" Style="{StaticResource BtnGreen}"/>
              <Button x:Name="${P}_BtnAddAllDevices" Content="+ Alle apparaten"  Style="{StaticResource BtnGreen}"/>
              <Button x:Name="${P}_BtnRemoveInclude" Content="Verwijderen"       Style="{StaticResource BtnGray}"/>
            </WrapPanel>
            <DataGrid x:Name="${P}_IncGrid" SelectionMode="Single" CanUserSortColumns="False">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Groep" Binding="{Binding Groep}" Width="*" ElementStyle="{StaticResource CellText}"/>
                $intentCol
              </DataGrid.Columns>
            </DataGrid>
          </DockPanel>
        </GroupBox>
        <GroupBox Grid.Row="2" Margin="0,0,0,6" Padding="6" BorderBrush="#CCCCCC">
          <GroupBox.Header><TextBlock Text="Exclude - deze groep(en) worden UITGESLOTEN" FontWeight="SemiBold" Foreground="#C42B1C"/></GroupBox.Header>
          <DockPanel>
            <WrapPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
              <Button x:Name="${P}_BtnAddExclude"    Content="+ Groep"     Style="{StaticResource BtnRed}"/>
              <Button x:Name="${P}_BtnRemoveExclude" Content="Verwijderen" Style="{StaticResource BtnGray}"/>
            </WrapPanel>
            <DataGrid x:Name="${P}_ExcGrid" SelectionMode="Single" CanUserSortColumns="False">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Groep" Binding="{Binding Groep}" Width="*" ElementStyle="{StaticResource CellText}"/>
                $intentCol
              </DataGrid.Columns>
            </DataGrid>
          </DockPanel>
        </GroupBox>
        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,4,0,0">
          <Button x:Name="${P}_BtnSave"   Content="Wijzigingen opslaan" Style="{StaticResource BtnBlue}"/>
          <Button x:Name="${P}_BtnCancel" Content="Ongedaan maken"      Style="{StaticResource BtnGray}"/>
          $viewBtn
        </StackPanel>
      </Grid>

      <StackPanel x:Name="${P}_BulkPanel" Grid.Row="2" Visibility="Collapsed" Margin="0,8,0,0">
        <TextBlock x:Name="${P}_LblBulkCount" Margin="0,0,0,12"/>
        $bulkIntent
        <WrapPanel Margin="0,0,0,6">
          <Button x:Name="${P}_BtnBulkInclude" Content="+ Groep INCLUDEN voor selectie" Style="{StaticResource BtnGreen}"/>
          <Button x:Name="${P}_BtnBulkExclude" Content="+ Groep EXCLUDEN voor selectie" Style="{StaticResource BtnRed}"/>
        </WrapPanel>
        <WrapPanel Margin="0,0,0,6">
          <Button x:Name="${P}_BtnBulkAllUsers"   Content="+ Alle gebruikers INCLUDEN voor selectie" Style="{StaticResource BtnGreen}"/>
          <Button x:Name="${P}_BtnBulkAllDevices" Content="+ Alle apparaten INCLUDEN voor selectie"  Style="{StaticResource BtnGreen}"/>
        </WrapPanel>
        <WrapPanel Margin="0,0,0,12">
          <Button x:Name="${P}_BtnBulkSave"   Content="Wijzigingen opslaan voor selectie" Style="{StaticResource BtnBlue}"/>
          <Button x:Name="${P}_BtnBulkCancel" Content="Ongedaan maken voor selectie"      Style="{StaticResource BtnGray}"/>
        </WrapPanel>
        $bulkView
        <TextBlock TextWrapping="Wrap" Foreground="Gray" Text="$bulkHint"/>
      </StackPanel>
    </Grid>
  </Border>
</Grid>
"@
}

$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Intune Assignments Manager" Width="1400" Height="860" MinWidth="1100" MinHeight="650"
        WindowStartupLocation="CenterScreen" Background="#F5F5F5" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="60"/>
      <RowDefinition Height="*" MinHeight="320"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="170" MinHeight="60"/>
      <RowDefinition Height="28"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0078D4">
      <DockPanel Margin="16,0,12,0">
        <Button x:Name="BtnConnect" DockPanel.Dock="Right" Content="Verbinden met Graph" Style="{StaticResource BtnDark}" VerticalAlignment="Center"/>
        <TextBlock x:Name="LblConn" DockPanel.Dock="Right" Text="Niet verbonden" Foreground="#FFC864" VerticalAlignment="Center" Margin="0,0,14,0"/>
        <TextBlock Text="Intune Assignments Manager" Foreground="White" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </DockPanel>
    </Border>

    <TabControl x:Name="MainTabs" Grid.Row="1" Margin="6,6,6,0" Padding="0">
      <TabItem Header="Configuration Policies" Padding="14,6">
{{TAB_P}}
      </TabItem>
      <TabItem Header="Apps" Padding="14,6">
{{TAB_A}}
      </TabItem>
    </TabControl>

    <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="#3C3C3C" ResizeDirection="Rows"/>

    <DockPanel Grid.Row="3" Background="#1E1E1E">
      <Border DockPanel.Dock="Top" Background="#2D2D2D" Height="28">
        <DockPanel>
          <Button x:Name="BtnClearLog" DockPanel.Dock="Right" Content="Leegmaken" Style="{StaticResource BtnLog}"/>
          <TextBlock Text="  Uitvoer / Log" Foreground="#B4B4B4" FontWeight="SemiBold" VerticalAlignment="Center"/>
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
      <TextBlock x:Name="LblStatus" Foreground="White" VerticalAlignment="Center" Margin="12,0" Text="Welkom - verbind eerst met Microsoft Graph"/>
    </Border>
  </Grid>
</Window>
'@

$script:GroupPickerXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Groep zoeken" Width="480" Height="500" WindowStartupLocation="CenterOwner"
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
    <TextBlock Text="Zoek op groepsnaam (begint met, min. 2 tekens):" Margin="0,0,0,6"/>
    <DockPanel Grid.Row="1" Margin="0,0,0,8">
      <Button x:Name="GP_BtnSearch" DockPanel.Dock="Right" Content="Zoeken" Margin="8,0,0,0"/>
      <TextBox x:Name="GP_Txt" Height="28" VerticalContentAlignment="Center"/>
    </DockPanel>
    <ListBox x:Name="GP_List" Grid.Row="2"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="GP_BtnOk" Content="OK" MinWidth="90"/>
      <Button x:Name="GP_BtnCancel" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" MinWidth="90" Margin="0,0,0,6"/>
    </StackPanel>
  </Grid>
</Window>
'@

$script:PolicyViewXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Policy inhoud" Width="1250" Height="800" MinWidth="800" MinHeight="450" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="White" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
<Style TargetType="DataGridCell">
  <Setter Property="BorderThickness" Value="0"/>
  <Style.Triggers>
    <DataTrigger Binding="{Binding State}" Value="conflict"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
    <DataTrigger Binding="{Binding State}" Value="unique"><Setter Property="Background" Value="#FFF4CE"/></DataTrigger>
    <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#CCE4F7"/><Setter Property="Foreground" Value="Black"/></Trigger>
  </Style.Triggers>
</Style>
  </Window.Resources>
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock x:Name="PV_LblTitle" FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap"/>
    <TextBlock x:Name="PV_LblInfo" Grid.Row="1" Foreground="#555555" TextWrapping="Wrap" Margin="0,2,0,8"/>
    <WrapPanel Grid.Row="2" Margin="0,0,0,6">
      <TextBlock Text="Zoeken:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="PV_TxtFilter" Width="260" Height="28" VerticalContentAlignment="Center" Margin="0,0,14,0" ToolTip="Zoekt in instelling, pad en waarden"/>
      <CheckBox x:Name="PV_ChkConfigured" Content="Alleen geconfigureerde waarden" IsChecked="True" VerticalAlignment="Center" Margin="0,0,14,0"
                ToolTip="Verbergt eigenschappen van klassieke profielen die leeg / notConfigured / false zijn"/>
      <CheckBox x:Name="PV_ChkDiff" Content="Alleen verschillen (conflict / niet overal)" VerticalAlignment="Center" Margin="0,0,14,0"/>
      <TextBlock x:Name="PV_LblCount" Foreground="Gray" VerticalAlignment="Center"/>
    </WrapPanel>
    <DataGrid x:Name="PV_Grid" Grid.Row="3" SelectionMode="Single" AlternatingRowBackground="White"/>
    <TextBox x:Name="PV_TxtDetail" Grid.Row="4" Height="120" Margin="0,8,0,0" IsReadOnly="True" TextWrapping="Wrap"
             VerticalScrollBarVisibility="Auto" FontFamily="Consolas" Background="#FAFAFA"/>
    <DockPanel Grid.Row="5" Margin="0,10,0,0" LastChildFill="False">
      <Button x:Name="PV_BtnExport" Content="Exporteren naar CSV" Style="{StaticResource BtnBlue}"/>
      <Button x:Name="PV_BtnMerge" Content="Samenvoegen tot nieuwe policy..." Style="{StaticResource BtnGreen}"
              ToolTip="Maakt een nieuwe Settings Catalog policy met alle instellingen van de geselecteerde policies"/>
      <Button x:Name="PV_BtnClose" DockPanel.Dock="Right" Content="Sluiten" Style="{StaticResource BtnGray}" IsCancel="True" Margin="0,0,0,6"/>
    </DockPanel>
  </Grid>
</Window>
'@

$script:PolicyMergeXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Policies samenvoegen" Width="620" Height="560" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="White" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
{{STYLES}}
  </Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Text="Naam nieuwe policy:" Margin="0,0,0,4"/>
    <TextBox x:Name="PM_TxtName" Grid.Row="1" Height="28" VerticalContentAlignment="Center"/>
    <TextBlock Grid.Row="2" Text="Beschrijving:" Margin="0,10,0,4"/>
    <TextBox x:Name="PM_TxtDesc" Grid.Row="3" Height="56" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/>
    <TextBlock x:Name="PM_LblPrio" Grid.Row="4" Text="Bij conflict voorrang geven aan:" Margin="0,10,0,4"/>
    <ComboBox x:Name="PM_CmbPrio" Grid.Row="5"/>
    <TextBox x:Name="PM_TxtSummary" Grid.Row="6" Margin="0,10,0,0" IsReadOnly="True" TextWrapping="Wrap"
             VerticalScrollBarVisibility="Auto" Background="#FAFAFA"/>
    <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="PM_BtnOk" Content="Policy aanmaken" Style="{StaticResource BtnGreen}" MinWidth="110"/>
      <Button x:Name="PM_BtnCancel" Content="Annuleren" Style="{StaticResource BtnGray}" IsCancel="True" MinWidth="90" Margin="0,0,0,6"/>
    </StackPanel>
  </Grid>
</Window>
'@

function ConvertFrom-Xaml {
    param([string]$Text)
    [xml]$x = $Text
    return [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
}

#endregion

#region --- Groep kiezen dialoog ---

function Show-GroupPicker {
    $script:GP_Win    = ConvertFrom-Xaml ($script:GroupPickerXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $script:GP_Win.Owner = $script:Window
    $script:GP_Txt    = $script:GP_Win.FindName('GP_Txt')
    $script:GP_List   = $script:GP_Win.FindName('GP_List')
    $script:GP_Map    = @{}
    $script:GP_Result = $null

    $script:GP_Search = {
        $q = $script:GP_Txt.Text.Trim()
        if ($q.Length -lt 2) {
            [System.Windows.MessageBox]::Show($script:GP_Win, "Voer minimaal 2 tekens in.", "Zoeken", "OK", "Information") | Out-Null
            return
        }
        $script:GP_List.Items.Clear(); $script:GP_Map.Clear()
        $script:GP_Win.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $results = Search-Groups -Query $q
            foreach ($r in $results) {
                $label = $r.displayName
                if ($script:GP_Map.ContainsKey($label)) { $label = "$($r.displayName)  [$($r.id)]" }
                $script:GP_Map[$label] = $r
                [void]$script:GP_List.Items.Add($label)
            }
            if ($results.Count -eq 0) { [void]$script:GP_List.Items.Add("(geen resultaten)") }
            elseif ($results.Count -eq 1) { $script:GP_List.SelectedIndex = 0 }
        }
        catch {
            [System.Windows.MessageBox]::Show($script:GP_Win, "Zoeken mislukt:`n$($_.Exception.Message)", "Fout", "OK", "Error") | Out-Null
        }
        finally { $script:GP_Win.Cursor = $null }
    }
    $script:GP_Accept = {
        $sel = [string]$script:GP_List.SelectedItem
        if ($sel -and $script:GP_Map.ContainsKey($sel)) {
            $script:GP_Result = $script:GP_Map[$sel]
            $script:GP_Win.DialogResult = $true
        }
    }

    $script:GP_Win.FindName('GP_BtnSearch').Add_Click({ & $script:GP_Search })
    $script:GP_Win.FindName('GP_BtnOk').Add_Click({ & $script:GP_Accept })
    $script:GP_List.Add_MouseDoubleClick({ & $script:GP_Accept })
    $script:GP_Txt.Add_PreviewKeyDown({
        if ($_.Key -eq 'Return') { & $script:GP_Search; $_.Handled = $true }
    })
    $script:GP_Win.Add_ContentRendered({ [void]$script:GP_Txt.Focus() })

    [void]$script:GP_Win.ShowDialog()
    return $script:GP_Result
}

#endregion

#region --- Policy-inhoud venster ---

function Add-GridColumn {
    param($Grid, [string]$Header, [string]$Binding, [double]$Star = 1, [double]$Pixel = 0, $ElementStyle = $null)
    $col = New-Object System.Windows.Controls.DataGridTextColumn
    # TextBlock als header, zodat underscores in policynamen niet als access key gelden
    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = $Header; $hdr.ToolTip = $Header; $hdr.TextTrimming = 'CharacterEllipsis'
    $col.Header  = $hdr
    $col.Binding = New-Object System.Windows.Data.Binding($Binding)
    $col.Width = if ($Pixel -gt 0) { New-Object System.Windows.Controls.DataGridLength($Pixel) }
                 else { New-Object System.Windows.Controls.DataGridLength($Star, [System.Windows.Controls.DataGridLengthUnitType]::Star) }
    $col.MinWidth = 110
    if ($ElementStyle) { $col.ElementStyle = $ElementStyle }
    [void]$Grid.Columns.Add($col)
}

function Test-IsRowClick {
    # True als de (dubbel)klik op een DataGrid-regel viel en niet op header/scrollbar
    param($Source)
    $d = $Source
    while ($d) {
        if ($d -is [System.Windows.Controls.DataGridRow]) { return $true }
        if ($d -is [System.Windows.Media.Visual]) { $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d) }
        elseif ($d -is [System.Windows.FrameworkContentElement]) { $d = $d.Parent }
        else { return $false }
    }
    return $false
}

function Show-PolicyContent {
    param($Ctx, [string[]]$Keys)
    $Keys = @($Keys | Where-Object { $_ })
    if ($Keys.Count -eq 0) { return }
    if ($Keys.Count -gt 15) {
        $r = Show-Msg "Je gaat $($Keys.Count) policies vergelijken, dat wordt een erg brede tabel. Doorgaan?" "Vergelijken" "YesNo" "Question"
        if ($r -ne 'Yes') { return }
    }
    $contents = New-Object System.Collections.ArrayList
    $failed = $null
    Enter-Busy "Policy-inhoud ophalen..."
    try {
        $i = 0
        foreach ($k in $Keys) {
            $i++
            $item = $Ctx.Data[$k].Item
            Set-Status "Inhoud ophalen ($i/$($Keys.Count)): $($item.Name)"
            [void]$contents.Add((Get-PolicyContent -Item $item))
        }
    }
    catch { $failed = $_.Exception.Message }
    finally { Exit-Busy }
    if ($failed) {
        Set-Status "Inhoud ophalen mislukt" "error"
        Show-Msg "Inhoud ophalen mislukt:`n$failed" "Fout" "OK" "Error" | Out-Null
        return
    }
    Set-Status "Inhoud van $($contents.Count) $(if ($contents.Count -eq 1) { 'policy' } else { 'policies' }) geladen" "success"

    $script:PV_Reload = $false
    $script:PV_DeleteSources = $false
    Show-PolicyViewWindow -Contents $contents.ToArray()
    if ($script:PV_DeleteSources) { Invoke-DeletePolicies $Ctx $Keys }
    if ($script:PV_Reload) { Invoke-LoadItems $Ctx }
}

function Show-PolicyViewWindow {
    param($Contents)
    $n = $Contents.Count
    $script:PV_Contents = $Contents
    $script:PV_Win = ConvertFrom-Xaml ($script:PolicyViewXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $script:PV_Win.Owner = $script:Window
    foreach ($nm in 'PV_LblTitle', 'PV_LblInfo', 'PV_TxtFilter', 'PV_ChkConfigured', 'PV_ChkDiff', 'PV_LblCount',
                    'PV_Grid', 'PV_TxtDetail', 'PV_BtnExport', 'PV_BtnMerge', 'PV_BtnClose') {
        Set-Variable -Scope Script -Name $nm -Value $script:PV_Win.FindName($nm)
    }
    $script:PV_Names     = @(Get-UniqueNames @($Contents | ForEach-Object { $_.Item.Name }))
    $script:PV_ValueCols = @(1..$n | ForEach-Object { "P$_" })
    $script:PV_Table     = Build-PolicyCompareTable -Contents $Contents

    $cellStyle = $script:PV_Win.FindResource('CellText')
    $g = $script:PV_Grid
    Add-GridColumn $g 'Instelling' 'Instelling' 2 -ElementStyle $cellStyle
    Add-GridColumn $g 'Pad' 'Pad' 1.5 -ElementStyle $cellStyle
    if ($n -eq 1) { Add-GridColumn $g 'Waarde' 'P1' 3 -ElementStyle $cellStyle }
    else {
        for ($i = 0; $i -lt $n; $i++) { Add-GridColumn $g $script:PV_Names[$i] "P$($i + 1)" 1.5 -ElementStyle $cellStyle }
        Add-GridColumn $g 'Status' 'Status' -Pixel 150 -ElementStyle $cellStyle
    }
    $g.ItemsSource = $script:PV_Table.DefaultView

    $hasDc = @($Contents | Where-Object { -not $_.IsSettingsCatalog }).Count -gt 0
    if (-not $hasDc) { $script:PV_ChkConfigured.Visibility = 'Collapsed' }
    $script:PV_BtnMerge.Visibility = if ($n -gt 1 -and -not $hasDc) { 'Visible' } else { 'Collapsed' }

    if ($n -eq 1) {
        $script:PV_Win.Title       = "Policy inhoud - $($Contents[0].Item.Name)"
        $script:PV_LblTitle.Text   = $Contents[0].Item.Name
        $script:PV_LblInfo.Text    = $Contents[0].Info
        $script:PV_ChkDiff.Visibility = 'Collapsed'
        $script:PV_Summary = ''
    }
    else {
        $script:PV_Win.Title     = "Policies vergelijken ($n)"
        $script:PV_LblTitle.Text = "$n policies vergelijken"
        $lines = for ($i = 0; $i -lt $n; $i++) { "$($i + 1). $($script:PV_Names[$i])  -  $($Contents[$i].Info)" }
        $script:PV_LblInfo.Text = (@($lines) -join "`n") + "`nRood = conflict (verschillende waarden), geel = staat niet in alle policies."
        $states = @($script:PV_Table.Rows | ForEach-Object { $_['State'] })
        $script:PV_Summary = "Gelijk: $(@($states | Where-Object { $_ -eq 'same' }).Count) | " +
                             "Conflict: $(@($states | Where-Object { $_ -eq 'conflict' }).Count) | " +
                             "Niet overal: $(@($states | Where-Object { $_ -eq 'unique' }).Count)"
    }

    $script:PV_TxtFilter.Add_TextChanged({ Update-PolicyViewFilter })
    $script:PV_ChkConfigured.Add_Click({ Update-PolicyViewFilter })
    $script:PV_ChkDiff.Add_Click({ Update-PolicyViewFilter })
    $script:PV_Grid.Add_SelectionChanged({ Update-PolicyViewDetail })
    $script:PV_BtnExport.Add_Click({ Export-PolicyView })
    $script:PV_BtnMerge.Add_Click({ Invoke-MergePolicies })
    $script:PV_BtnClose.Add_Click({ $script:PV_Win.Close() })
    $script:PV_Win.Add_ContentRendered({ [void]$script:PV_TxtFilter.Focus() })

    Update-PolicyViewFilter
    [void]$script:PV_Win.ShowDialog()
}

function Update-PolicyViewFilter {
    $parts = @()
    if ($script:PV_ChkConfigured.IsChecked -eq $true -and $script:PV_ChkConfigured.Visibility -eq 'Visible') { $parts += "Configured = '1'" }
    if ($script:PV_ChkDiff.IsChecked -eq $true -and $script:PV_ChkDiff.Visibility -eq 'Visible') { $parts += "State IN ('conflict','unique')" }
    $f = $script:PV_TxtFilter.Text.Trim()
    if ($f) {
        $lit = ConvertTo-LikeLiteral $f
        $cols = @('Instelling', 'Pad') + $script:PV_ValueCols
        $parts += '(' + ((@($cols | ForEach-Object { "$_ LIKE '*$lit*'" })) -join ' OR ') + ')'
    }
    $view = $script:PV_Table.DefaultView
    $view.RowFilter = $parts -join ' AND '
    $txt = "$($view.Count) van $($script:PV_Table.Rows.Count) instellingen"
    if ($script:PV_Summary) { $txt += "   ($script:PV_Summary)" }
    $script:PV_LblCount.Text = $txt
}

function Update-PolicyViewDetail {
    $rv = $script:PV_Grid.SelectedItem
    if (-not $rv) { $script:PV_TxtDetail.Text = ''; return }
    $row = $rv.Row
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("Instelling : $($row['Instelling'])")
    if ([string]$row['Pad']) { [void]$lines.Add("Pad        : $($row['Pad'])") }
    [void]$lines.Add("Sleutel    : $($row['Key'])")
    if ([string]$row['Status']) { [void]$lines.Add("Status     : $($row['Status'])") }
    [void]$lines.Add('')
    for ($i = 0; $i -lt $script:PV_Names.Count; $i++) {
        $v = [string]$row["P$($i + 1)"]
        if ($script:PV_Names.Count -gt 1) { [void]$lines.Add("[$($script:PV_Names[$i])]") } else { [void]$lines.Add("Waarde:") }
        [void]$lines.Add($(if ($v) { $v } else { '(niet aanwezig)' }))
        [void]$lines.Add('')
    }
    $script:PV_TxtDetail.Text = (@($lines) -join "`r`n").TrimEnd()
}

function Export-PolicyView {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "CSV-bestand (*.csv)|*.csv"
    $base = if ($script:PV_Contents.Count -eq 1) { $script:PV_Contents[0].Item.Name } else { "Policy-vergelijking" }
    $dlg.FileName = ($base -replace '[\\/:*?"<>|]', '_') + ".csv"
    if (-not $dlg.ShowDialog($script:PV_Win)) { return }
    $multi = $script:PV_Names.Count -gt 1
    $out = foreach ($rv in $script:PV_Table.DefaultView) {
        $o = [ordered]@{ Instelling = $rv.Row['Instelling']; Pad = $rv.Row['Pad'] }
        for ($i = 0; $i -lt $script:PV_Names.Count; $i++) {
            $col = if ($multi) { $script:PV_Names[$i] } else { 'Waarde' }
            $o[$col] = $rv.Row["P$($i + 1)"]
        }
        if ($multi) { $o['Status'] = $rv.Row['Status'] }
        [pscustomobject]$o
    }
    try {
        @($out) | Export-Csv -Path $dlg.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-Log "Policy-inhoud geexporteerd naar $($dlg.FileName)" "success"
        Show-Msg "Geexporteerd naar:`n$($dlg.FileName)" "Export" -Owner $script:PV_Win | Out-Null
    }
    catch {
        Show-Msg "Exporteren mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" -Owner $script:PV_Win | Out-Null
    }
}

function Update-MergeSummary {
    $prio = [Math]::Max(0, $script:PM_CmbPrio.SelectedIndex)
    $plan = Get-MergePlan -Contents $script:PV_Contents -PrioIndex $prio
    $txt = "$($plan.Settings.Count) instelling(en) in de nieuwe policy`r`n" +
           "$($plan.Duplicates) dubbele instelling(en) (identiek, 1x opgenomen)`r`n" +
           "$($plan.Conflicts.Count) conflict(en)"
    if ($plan.Conflicts.Count -gt 0) {
        $txt += "`r`n`r`nConflicten - de hele instelling (incl. onderliggende waarden) wordt overgenomen uit de policy met voorrang:`r`n- " +
                (@($plan.Conflicts) -join "`r`n- ")
    }
    $txt += "`r`n`r`nDe nieuwe policy krijgt GEEN toewijzingen. De originele policies blijven ongewijzigd."
    $script:PM_TxtSummary.Text = $txt
}

function Show-MergeDialog {
    $names = @($script:PV_Contents | ForEach-Object { $_.Item.Name })
    $script:PM_Win = ConvertFrom-Xaml ($script:PolicyMergeXaml.Replace('{{STYLES}}', $script:StylesXaml))
    $script:PM_Win.Owner = $script:PV_Win
    foreach ($nm in 'PM_TxtName', 'PM_TxtDesc', 'PM_CmbPrio', 'PM_LblPrio', 'PM_TxtSummary') {
        Set-Variable -Scope Script -Name $nm -Value $script:PM_Win.FindName($nm)
    }
    $script:PM_Result = $null
    $script:PM_TxtName.Text = "Samengevoegd - " + ($names -join ' + ')
    $script:PM_TxtDesc.Text = "Samengevoegd uit: " + ($names -join ', ')
    foreach ($nm in $names) { [void]$script:PM_CmbPrio.Items.Add($nm) }
    $script:PM_CmbPrio.SelectedIndex = 0
    Update-MergeSummary
    if (@((Get-MergePlan -Contents $script:PV_Contents).Conflicts).Count -eq 0) {
        $script:PM_CmbPrio.IsEnabled = $false
        $script:PM_LblPrio.Text = "Bij conflict voorrang geven aan: (niet nodig, geen conflicten)"
    }
    $script:PM_CmbPrio.Add_SelectionChanged({ Update-MergeSummary })
    $script:PM_Win.FindName('PM_BtnOk').Add_Click({
        $nm = $script:PM_TxtName.Text.Trim()
        if (-not $nm) {
            [System.Windows.MessageBox]::Show($script:PM_Win, "Vul een naam in.", "Samenvoegen", "OK", "Warning") | Out-Null
            return
        }
        $script:PM_Result = [pscustomobject]@{
            Name = $nm; Description = $script:PM_TxtDesc.Text.Trim(); Prio = [Math]::Max(0, $script:PM_CmbPrio.SelectedIndex)
        }
        $script:PM_Win.DialogResult = $true
    })
    [void]$script:PM_Win.ShowDialog()
    return $script:PM_Result
}

function Invoke-MergePolicies {
    $contents = $script:PV_Contents
    $first = $contents[0]
    $bad = @($contents | Where-Object {
        $_.Platforms -ne $first.Platforms -or $_.Technologies -ne $first.Technologies -or $_.TemplateId -ne $first.TemplateId
    })
    if ($bad.Count -gt 0) {
        $lines = foreach ($c in $contents) {
            $tpl = if ($c.TemplateId) { $c.TemplateId } else { 'geen' }
            "- $($c.Item.Name): $($c.Platforms) / $($c.Technologies) / template: $tpl"
        }
        Show-Msg ("Samenvoegen kan alleen als alle policies hetzelfde platform, dezelfde technologie en hetzelfde template hebben:`n`n" +
                  (@($lines) -join "`n")) "Samenvoegen niet mogelijk" "OK" "Warning" -Owner $script:PV_Win | Out-Null
        return
    }

    $opts = Show-MergeDialog
    if (-not $opts) { return }
    $plan = Get-MergePlan -Contents $contents -PrioIndex $opts.Prio
    $settings = @(foreach ($inst in $plan.Settings) {
        @{ '@odata.type' = '#microsoft.graph.deviceManagementConfigurationSetting'; settingInstance = $inst }
    })
    $body = @{
        name            = $opts.Name
        description     = $opts.Description
        platforms       = $first.Platforms
        technologies    = $first.Technologies
        roleScopeTagIds = @($first.RoleScopeTagIds)
        settings        = $settings
    }
    if ($first.TemplateId) { $body.templateReference = @{ templateId = $first.TemplateId } }

    $script:PV_Win.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $r = Invoke-Graph -Uri "$script:GraphBeta/deviceManagement/configurationPolicies" -Method POST -Body $body
        Write-Log "Nieuwe policy '$($opts.Name)' aangemaakt met $($plan.Settings.Count) instelling(en) (id $($r.id))" "success"
    }
    catch {
        $script:PV_Win.Cursor = $null
        Show-Msg "Aanmaken van de samengevoegde policy mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" -Owner $script:PV_Win | Out-Null
        return
    }
    $script:PV_Win.Cursor = $null
    $a = Show-Msg ("Policy '$($opts.Name)' is aangemaakt met $($plan.Settings.Count) instelling(en), nog zonder toewijzingen.`n" +
                   "De originele policies zijn niet gewijzigd.`n`n" +
                   "Wil je de $($contents.Count) bronpolicies nu verwijderen?`n" +
                   "(Je krijgt eerst nog een overzicht ter bevestiging en er wordt een backup gemaakt. " +
                   "Tip: wijs de nieuwe policy eerst toe voordat je de bronnen verwijdert.)`n`n" +
                   "Bij 'Nee' wordt alleen de policylijst opnieuw geladen.") "Samengevoegd" "YesNo" "Question" -Owner $script:PV_Win
    $script:PV_DeleteSources = ($a -eq 'Yes')
    $script:PV_Reload = $true
    $script:PV_Win.Close()
}

#endregion

#region --- Tab-logica (gedeeld door Policies en Apps) ---

function New-TabContext {
    param([string]$Prefix, [bool]$IsApp)
    $singular = 'policy'; $plural = 'policies'
    if ($IsApp) { $singular = 'app'; $plural = 'apps' }
    $c = @{
        Prefix = $Prefix; IsApp = $IsApp; Singular = $singular; Plural = $plural
        Data = @{}; Table = $null; SelectedKey = $null; SelectedKeys = @()
    }
    $names = 'Grid', 'TxtFilter', 'LblCount', 'BtnLoad', 'BtnSaveAll', 'Editor', 'LblSelected', 'LblDirty',
             'SinglePanel', 'BulkPanel', 'IncGrid', 'ExcGrid', 'BtnAddInclude', 'BtnAddAllUsers', 'BtnAddAllDevices',
             'BtnRemoveInclude', 'BtnAddExclude', 'BtnRemoveExclude', 'BtnSave', 'BtnCancel', 'LblBulkCount',
             'BtnBulkInclude', 'BtnBulkExclude', 'BtnBulkAllUsers', 'BtnBulkAllDevices', 'BtnBulkSave', 'BtnBulkCancel', 'CmbIntent', 'BtnChangeIntent', 'CmbBulkIntent',
             'BtnView', 'BtnBulkView', 'BtnDelete', 'BtnBulkDelete'
    foreach ($n in $names) { $c[$n] = $script:Window.FindName("$($Prefix)_$n") }

    foreach ($cmb in @($c.CmbIntent, $c.CmbBulkIntent)) {
        if (-not $cmb) { continue }
        foreach ($o in $script:AppIntents) { [void]$cmb.Items.Add($o.Label) }
        $cmb.SelectedIndex = 0
    }
    return $c
}

function Get-Ctx {
    param($Sender)
    return $script:Tabs[[string]$Sender.Tag]
}

function Get-SelectedIntent {
    param($Ctx, [switch]$Bulk)
    if (-not $Ctx.IsApp) { return $null }
    $cmb = if ($Bulk) { $Ctx.CmbBulkIntent } else { $Ctx.CmbIntent }
    $m = $script:AppIntents | Where-Object { $_.Label -eq [string]$cmb.SelectedItem } | Select-Object -First 1
    if ($m) { return $m.Value }
    return 'required'
}

function Format-EntryList {
    param($Ctx, $Entries)
    $parts = foreach ($e in @($Entries)) {
        if ($Ctx.IsApp) { "$($e.Name) ($(Get-AppIntentLabel $e.Intent))" } else { $e.Name }
    }
    return (@($parts) -join ", ")
}

function Build-MainTable {
    param($Ctx)
    $t = New-Table @('Key', 'Type', 'Naam', 'Include', 'Exclude', 'Status')
    $t.PrimaryKey = [System.Data.DataColumn[]]@($t.Columns['Key'])
    foreach ($kv in $Ctx.Data.GetEnumerator()) {
        $d = $kv.Value
        $status = if ($d.Dirty) { "* niet opgeslagen" } else { "" }
        [void]$t.Rows.Add($kv.Key, $d.Item.TypeLabel, $d.Item.Name,
            (Format-EntryList $Ctx $d.Include), (Format-EntryList $Ctx $d.Exclude), $status)
    }
    $t.DefaultView.Sort = 'Naam ASC'
    $Ctx.Table = $t
    $Ctx.Grid.ItemsSource = $t.DefaultView
    Update-Filter $Ctx
}

function Update-Filter {
    param($Ctx)
    if (-not $Ctx.Table) { return }
    $f = $Ctx.TxtFilter.Text.Trim()
    if ($f) {
        $lit = ConvertTo-LikeLiteral $f
        $Ctx.Table.DefaultView.RowFilter = "Naam LIKE '*$lit*' OR Include LIKE '*$lit*' OR Exclude LIKE '*$lit*'"
    }
    else { $Ctx.Table.DefaultView.RowFilter = '' }
    $Ctx.LblCount.Text = "$($Ctx.Table.DefaultView.Count) $($Ctx.Plural)"
}

function Update-GridRow {
    param($Ctx, $Key)
    if (-not $Ctx.Table) { return }
    $row = $Ctx.Table.Rows.Find($Key)
    if (-not $row) { return }
    $d = $Ctx.Data[$Key]
    $row['Include'] = Format-EntryList $Ctx $d.Include
    $row['Exclude'] = Format-EntryList $Ctx $d.Exclude
    $row['Status']  = if ($d.Dirty) { "* niet opgeslagen" } else { "" }
}

function New-EntryTable {
    param($Entries)
    $t = New-Table @('Idx', 'Groep', 'Intent')
    $i = 0
    foreach ($e in $Entries) {
        [void]$t.Rows.Add([string]$i, $e.Name, (Get-AppIntentLabel $e.Intent))
        $i++
    }
    return ,$t
}

function Update-Editor {
    param($Ctx, $Key)
    $d = $Ctx.Data[$Key]
    $Ctx.LblSelected.Text = "$($d.Item.Name)  [$($d.Item.TypeLabel)]"
    $Ctx.IncGrid.ItemsSource = (New-EntryTable $d.Include).DefaultView
    $Ctx.ExcGrid.ItemsSource = (New-EntryTable $d.Exclude).DefaultView
    $extra = if ($d.Other.Count -gt 0) { "  (+ $($d.Other.Count) overige toewijzing(en), blijven behouden)" } else { "" }
    $Ctx.LblDirty.Text = if ($d.Dirty) { "Niet opgeslagen wijzigingen$extra" } else { $extra.Trim() }
}

function Update-Both {
    param($Ctx, $Key)
    Update-Editor  $Ctx $Key
    Update-GridRow $Ctx $Key
}

function Reset-Editor {
    param($Ctx)
    $Ctx.SelectedKey  = $null
    $Ctx.SelectedKeys = @()
    $Ctx.Editor.IsEnabled = $false
    $Ctx.LblSelected.Text = "Selecteer een $($Ctx.Singular)"
    $Ctx.LblDirty.Text = ""
    $Ctx.IncGrid.ItemsSource = $null
    $Ctx.ExcGrid.ItemsSource = $null
    $Ctx.SinglePanel.Visibility = 'Visible'
    $Ctx.BulkPanel.Visibility   = 'Collapsed'
}

function Update-Selection {
    param($Ctx)
    $keys = @($Ctx.Grid.SelectedItems | ForEach-Object { [string]$_.Row['Key'] })
    $n = $keys.Count
    if ($n -eq 1) {
        $Ctx.SelectedKey  = $keys[0]
        $Ctx.SelectedKeys = $keys
        Update-Editor $Ctx $keys[0]
        $Ctx.SinglePanel.Visibility = 'Visible'
        $Ctx.BulkPanel.Visibility   = 'Collapsed'
        $Ctx.Editor.IsEnabled = $true
    }
    elseif ($n -gt 1) {
        $Ctx.SelectedKey  = $null
        $Ctx.SelectedKeys = $keys
        $Ctx.LblSelected.Text  = "$n $($Ctx.Plural) geselecteerd"
        $Ctx.LblDirty.Text     = ""
        $Ctx.LblBulkCount.Text = "$n $($Ctx.Plural) geselecteerd - kies hieronder een actie"
        $Ctx.SinglePanel.Visibility = 'Collapsed'
        $Ctx.BulkPanel.Visibility   = 'Visible'
        $Ctx.Editor.IsEnabled = $true
    }
    else { Reset-Editor $Ctx }
}

function Invoke-LoadItems {
    param($Ctx)
    Enter-Busy "$($Ctx.Plural) ophalen..."
    Write-Log "$($Ctx.Plural) ophalen..." "info"
    try {
        $items = if ($Ctx.IsApp) { Get-AllApps } else { Get-AllPolicies }
        $newData = @{}
        $i = 0
        foreach ($it in $items) {
            $i++
            if ($it.AssignmentsLoaded) { $assignments = $it.Assignments }
            else {
                Set-Status "Assignments ophalen ($i/$($items.Count))..."
                $assignments = Get-ItemAssignments -Ctx $Ctx -Item $it
            }
            $split = Split-Assignments -Assignments $assignments
            $newData[$it.Key] = [pscustomobject]@{ Item = $it; Include = $split.Include; Exclude = $split.Exclude; Other = $split.Other; Dirty = $false }
        }
        Set-Status "Groepsnamen ophalen..."
        Set-EntryNames -DataList $newData.Values

        $Ctx.Data = $newData
        Reset-Editor $Ctx
        Build-MainTable $Ctx
        Set-Status "$($items.Count) $($Ctx.Plural) geladen" "success"
        Write-Log "$($items.Count) $($Ctx.Plural) geladen" "success"
    }
    catch {
        Set-Status "Fout: $($_.Exception.Message)" "error"
        Write-Log "Fout bij ophalen $($Ctx.Plural): $($_.Exception.Message)" "error"
    }
    finally { Exit-Busy }
}

function Reload-OneItem {
    param($Ctx, $Key)
    $d = $Ctx.Data[$Key]
    $split = Split-Assignments -Assignments (Get-ItemAssignments -Ctx $Ctx -Item $d.Item)
    $d.Include = $split.Include
    $d.Exclude = $split.Exclude
    $d.Other   = $split.Other
    $d.Dirty   = $false
    Set-EntryNames -DataList @($d)
}

function Save-OneItem {
    param($Ctx, $Key)
    $d = $Ctx.Data[$Key]
    try {
        Save-ItemAssignments -Ctx $Ctx -Data $d
        $d.Dirty = $false
        Write-Log "Assignments opgeslagen voor '$($d.Item.Name)'" "success"
        # Herladen zodat de weergave exact de server-state toont
        try { Reload-OneItem -Ctx $Ctx -Key $Key } catch { }
    }
    catch {
        Write-Log "Opslaan mislukt voor '$($d.Item.Name)': $($_.Exception.Message)" "error"
    }
    Update-GridRow $Ctx $Key
    if ($Ctx.SelectedKey -eq $Key) { Update-Editor $Ctx $Key }
}

function Add-EntryToData {
    # Voegt een regel toe aan Include of Exclude en haalt dezelfde groep weg
    # uit de andere kant (een groep kan niet in beide tegelijk staan).
    param($Data, $Entry, [ValidateSet('Include', 'Exclude')][string]$Side, [switch]$Quiet)
    $otherSide = if ($Side -eq 'Include') { 'Exclude' } else { 'Include' }
    $own = $Data.$Side
    $oth = $Data.$otherSide
    if (@($own | ForEach-Object { $_.Id }) -contains $Entry.Id) { return $false }
    for ($i = $oth.Count - 1; $i -ge 0; $i--) {
        if ($oth[$i].Id -eq $Entry.Id) {
            $oth.RemoveAt($i)
            if (-not $Quiet) { Write-Log "'$($Entry.Name)' verplaatst van $otherSide naar $Side" "info" }
        }
    }
    [void]$own.Add($Entry)
    $Data.Dirty = $true
    return $true
}

function Invoke-AddGroup {
    param($Ctx, [string]$Side)
    if (-not $Ctx.SelectedKey) { return }
    $g = Show-GroupPicker
    if (-not $g) { return }
    $d = $Ctx.Data[$Ctx.SelectedKey]
    $e = New-Entry -Id $g.id -Name $g.displayName -Intent (Get-SelectedIntent $Ctx)
    if (-not (Add-EntryToData -Data $d -Entry $e -Side $Side)) {
        $hint = if ($Ctx.IsApp -and $Side -eq 'Include') { "`nGebruik 'Intent van selectie wijzigen' om de intent aan te passen." } else { "" }
        Show-Msg "Groep '$($g.displayName)' staat al in $Side.$hint" "Al aanwezig" | Out-Null
        return
    }
    Update-Both $Ctx $Ctx.SelectedKey
}

function Invoke-AddSpecial {
    param($Ctx, [string]$Kind)
    if (-not $Ctx.SelectedKey) { return }
    $d = $Ctx.Data[$Ctx.SelectedKey]
    $e = New-SpecialEntry -Kind $Kind -Intent (Get-SelectedIntent $Ctx)
    if (-not (Add-EntryToData -Data $d -Entry $e -Side 'Include')) {
        Show-Msg "'$($e.Name)' staat al in Include." "Al aanwezig" | Out-Null
        return
    }
    Update-Both $Ctx $Ctx.SelectedKey
}

function Invoke-RemoveEntry {
    param($Ctx, [string]$Side)
    if (-not $Ctx.SelectedKey) { return }
    $grid = if ($Side -eq 'Include') { $Ctx.IncGrid } else { $Ctx.ExcGrid }
    if (-not $grid.SelectedItem) { return }
    $idx = [int]$grid.SelectedItem.Row['Idx']
    $d = $Ctx.Data[$Ctx.SelectedKey]
    $removed = $d.$Side[$idx]
    $d.$Side.RemoveAt($idx)
    $d.Dirty = $true
    Write-Log "'$($removed.Name)' verwijderd uit $Side van '$($d.Item.Name)'" "info"
    Update-Both $Ctx $Ctx.SelectedKey
}

function Invoke-ChangeIntent {
    param($Ctx)
    if (-not $Ctx.SelectedKey) { return }
    $side = $null
    if ($Ctx.IncGrid.SelectedItem) { $side = 'Include'; $row = $Ctx.IncGrid.SelectedItem }
    elseif ($Ctx.ExcGrid.SelectedItem) { $side = 'Exclude'; $row = $Ctx.ExcGrid.SelectedItem }
    if (-not $side) {
        Show-Msg "Selecteer eerst een groep in Include of Exclude." "Intent wijzigen" | Out-Null
        return
    }
    $d = $Ctx.Data[$Ctx.SelectedKey]
    $e = $d.$side[[int]$row.Row['Idx']]
    $new = Get-SelectedIntent $Ctx
    if ($e.Intent -eq $new) { return }
    $old = $e.Intent
    $e.Intent = $new
    # Een installatie-deadline hoort alleen bij Required; bij andere intents
    # wordt die door Intune geweigerd, dus die halen we weg.
    if ($new -ne 'required' -and $e.Settings -is [System.Collections.IDictionary] -and $e.Settings.Contains('installTimeSettings')) {
        $e.Settings['installTimeSettings'] = $null
    }
    $d.Dirty = $true
    Write-Log "Intent van '$($e.Name)' ($side) gewijzigd: $(Get-AppIntentLabel $old) -> $(Get-AppIntentLabel $new)" "info"
    Update-Both $Ctx $Ctx.SelectedKey
}

function Invoke-SaveSingle {
    param($Ctx)
    if (-not $Ctx.SelectedKey) { return }
    $key = $Ctx.SelectedKey
    Enter-Busy "Opslaan..."
    try { Save-OneItem -Ctx $Ctx -Key $key } finally { Exit-Busy }
    $d = $Ctx.Data[$key]
    if ($d.Dirty) { Set-Status "Opslaan mislukt voor '$($d.Item.Name)'" "error" }
    else { Set-Status "Opgeslagen: $($d.Item.Name)" "success" }
}

function Invoke-SaveMany {
    param($Ctx, [switch]$All)
    $keys  = if ($All) { @($Ctx.Data.Keys) } else { @($Ctx.SelectedKeys) }
    $dirty = @($keys | Where-Object { $Ctx.Data[$_].Dirty })
    $scope = if ($All) { "" } else { " uit de selectie" }
    if ($dirty.Count -eq 0) {
        Show-Msg "Geen openstaande wijzigingen$scope." "Niets te doen" | Out-Null
        return
    }
    $confirm = Show-Msg "$($dirty.Count) $($Ctx.Plural)$scope met wijzigingen opslaan?" "Bevestigen" "YesNo" "Question"
    if ($confirm -ne 'Yes') { return }
    Enter-Busy "Opslaan..."
    $ok = 0; $fail = 0; $i = 0
    try {
        foreach ($k in $dirty) {
            $i++
            Set-Status "Opslaan ($i/$($dirty.Count))..."
            Save-OneItem -Ctx $Ctx -Key $k
            if ($Ctx.Data[$k].Dirty) { $fail++ } else { $ok++ }
        }
    }
    finally { Exit-Busy }
    Set-Status "$ok opgeslagen, $fail mislukt" $(if ($fail -eq 0) { "success" } else { "error" })
}

function Invoke-CancelSingle {
    param($Ctx)
    if (-not $Ctx.SelectedKey) { return }
    Enter-Busy "Herladen..."
    try {
        Reload-OneItem -Ctx $Ctx -Key $Ctx.SelectedKey
        Update-Both $Ctx $Ctx.SelectedKey
        Write-Log "Wijzigingen ongedaan gemaakt, opnieuw geladen vanuit Intune" "info"
        Set-Status "Opnieuw geladen" "success"
    }
    catch {
        Show-Msg "Herladen mislukt:`n$($_.Exception.Message)" "Fout" "OK" "Error" | Out-Null
    }
    finally { Exit-Busy }
}

function Invoke-CancelBulk {
    param($Ctx)
    if ($Ctx.SelectedKeys.Count -eq 0) { return }
    $errCount = 0
    Enter-Busy "Herladen..."
    try {
        foreach ($k in $Ctx.SelectedKeys) {
            try { Reload-OneItem -Ctx $Ctx -Key $k; Update-GridRow $Ctx $k }
            catch { $errCount++ }
        }
    }
    finally { Exit-Busy }
    Write-Log "Wijzigingen ongedaan gemaakt voor $($Ctx.SelectedKeys.Count - $errCount) van $($Ctx.SelectedKeys.Count) geselecteerde $($Ctx.Plural)" "info"
    Set-Status "Selectie opnieuw geladen" $(if ($errCount -eq 0) { "success" } else { "error" })
    if ($errCount -gt 0) { Show-Msg "Herladen mislukt voor $errCount $($Ctx.Plural)." "Fout" "OK" "Error" | Out-Null }
}

function Invoke-BulkAdd {
    param($Ctx, [string]$Side)
    if ($Ctx.SelectedKeys.Count -eq 0) { return }
    $g = Show-GroupPicker
    if (-not $g) { return }
    $intent = Get-SelectedIntent $Ctx -Bulk
    $changed = 0
    foreach ($k in $Ctx.SelectedKeys) {
        # Per item een eigen entry-object, zodat latere wijzigingen per item los staan
        $e = New-Entry -Id $g.id -Name $g.displayName -Intent $intent
        if (Add-EntryToData -Data $Ctx.Data[$k] -Entry $e -Side $Side -Quiet) {
            Update-GridRow $Ctx $k
            $changed++
        }
    }
    $intentTxt = if ($Ctx.IsApp) { " ($(Get-AppIntentLabel $intent))" } else { "" }
    Write-Log "Groep '$($g.displayName)' toegevoegd als $($Side.ToUpper())$intentTxt voor $changed van $($Ctx.SelectedKeys.Count) geselecteerde $($Ctx.Plural) (nog niet opgeslagen)" "success"
    Set-Status "Groep '$($g.displayName)' als $Side$intentTxt toegevoegd voor $changed $($Ctx.Plural) - nog niet opgeslagen" "success"
}

function Invoke-BulkAddSpecial {
    param($Ctx, [string]$Kind)
    if ($Ctx.SelectedKeys.Count -eq 0) { return }
    $intent = Get-SelectedIntent $Ctx -Bulk
    $changed = 0
    $name = ''
    foreach ($k in $Ctx.SelectedKeys) {
        $e = New-SpecialEntry -Kind $Kind -Intent $intent
        $name = $e.Name
        if (Add-EntryToData -Data $Ctx.Data[$k] -Entry $e -Side 'Include' -Quiet) {
            Update-GridRow $Ctx $k
            $changed++
        }
    }
    $intentTxt = if ($Ctx.IsApp) { " ($(Get-AppIntentLabel $intent))" } else { "" }
    Write-Log "'$name' toegevoegd als INCLUDE$intentTxt voor $changed van $($Ctx.SelectedKeys.Count) geselecteerde $($Ctx.Plural) (nog niet opgeslagen)" "success"
    Set-Status "'$name' als Include$intentTxt toegevoegd voor $changed $($Ctx.Plural) - nog niet opgeslagen" "success"
}

function Invoke-DeletePolicies {
    param($Ctx, [string[]]$Keys)
    $Keys = @($Keys | Where-Object { $_ -and $Ctx.Data.ContainsKey($_) })
    $n = $Keys.Count
    if ($n -eq 0) { return }
    $what = if ($n -eq 1) { 'policy' } else { 'policies' }

    $list = New-Object System.Collections.ArrayList
    $assigned = 0
    foreach ($k in $Keys) {
        $d = $Ctx.Data[$k]
        $isAssigned = ($d.Include.Count + $d.Exclude.Count + $d.Other.Count) -gt 0
        if ($isAssigned) { $assigned++ }
        if ($list.Count -lt 20) {
            [void]$list.Add("- $($d.Item.Name)  [$($d.Item.TypeLabel)]$(if ($isAssigned) { '  (TOEGEWEZEN)' })")
        }
    }
    if ($n -gt 20) { [void]$list.Add("... en nog $($n - 20)") }
    $warn = if ($assigned -gt 0) { "`n`nLET OP: $assigned hiervan is nog toegewezen. Apparaten/gebruikers verliezen deze instellingen." } else { '' }
    $msg = "Weet je zeker dat je $n $what DEFINITIEF uit Intune wilt verwijderen?`n`n$(@($list) -join "`n")$warn" +
           "`n`nVan elke policy wordt eerst een JSON-backup gemaakt in:`n$script:BackupRoot"
    if ((Show-Msg $msg "Policies verwijderen" "YesNo" "Warning") -ne 'Yes') { return }
    if ($n -gt 1 -or $assigned -gt 0) {
        if ((Show-Msg "Laatste controle: $n $what verwijderen? Dit kan niet ongedaan worden gemaakt in Intune." "Bevestigen" "YesNo" "Warning") -ne 'Yes') { return }
    }

    # Eerst de selectie leegmaken: verwijderde regels mogen niet meer geselecteerd zijn
    $Ctx.Grid.UnselectAll()
    Reset-Editor $Ctx
    $folder = Join-Path $script:BackupRoot (Get-Date -Format 'yyyy-MM-dd_HHmmss')
    $ok = 0; $fail = 0; $i = 0
    Enter-Busy "Verwijderen..."
    try {
        foreach ($k in $Keys) {
            $i++
            $d = $Ctx.Data[$k]
            Set-Status "Verwijderen ($i/$n): $($d.Item.Name)"
            try {
                $file = Backup-Policy -Item $d.Item -Folder $folder
                Write-Log "Backup van '$($d.Item.Name)' opgeslagen: $file" "info"
            }
            catch {
                Write-Log "Backup mislukt voor '$($d.Item.Name)', policy is NIET verwijderd: $($_.Exception.Message)" "error"
                $fail++
                continue
            }
            try {
                Remove-Policy -Item $d.Item
                Write-Log "Policy verwijderd: '$($d.Item.Name)'" "success"
                $row = if ($Ctx.Table) { $Ctx.Table.Rows.Find($k) }
                if ($row) { $row.Delete() }
                $Ctx.Data.Remove($k)
                $ok++
            }
            catch { $fail++ }
        }
    }
    finally {
        if ($Ctx.Table) { $Ctx.Table.AcceptChanges() }
        Exit-Busy
    }
    Update-Filter $Ctx
    Set-Status "$ok verwijderd, $fail mislukt" $(if ($fail -eq 0) { "success" } else { "error" })
    if ($fail -gt 0) { Show-Msg "$fail $what niet verwijderd, zie het log voor details." "Verwijderen" "OK" "Warning" | Out-Null }
}

function Register-TabEvents {
    param($Ctx)
    foreach ($k in @($Ctx.Keys)) {
        $ctl = $Ctx[$k]
        if ($ctl -is [System.Windows.FrameworkElement]) { $ctl.Tag = $Ctx.Prefix }
    }

    $Ctx.BtnLoad.Add_Click({ Invoke-LoadItems (Get-Ctx $this) })
    $Ctx.TxtFilter.Add_TextChanged({ Update-Filter (Get-Ctx $this) })
    $Ctx.Grid.Add_SelectionChanged({
        if ($_.OriginalSource -ne $this) { return }
        Update-Selection (Get-Ctx $this)
    })

    # Maar 1 van beide lijsten tegelijk geselecteerd, zodat "Intent wijzigen" eenduidig is
    $Ctx.IncGrid.Add_SelectionChanged({ $c = Get-Ctx $this; if ($c.IncGrid.SelectedItem) { $c.ExcGrid.UnselectAll() } })
    $Ctx.ExcGrid.Add_SelectionChanged({ $c = Get-Ctx $this; if ($c.ExcGrid.SelectedItem) { $c.IncGrid.UnselectAll() } })

    $Ctx.BtnAddInclude.Add_Click({    Invoke-AddGroup    (Get-Ctx $this) 'Include' })
    $Ctx.BtnAddExclude.Add_Click({    Invoke-AddGroup    (Get-Ctx $this) 'Exclude' })
    $Ctx.BtnAddAllUsers.Add_Click({   Invoke-AddSpecial  (Get-Ctx $this) 'allUsers' })
    $Ctx.BtnAddAllDevices.Add_Click({ Invoke-AddSpecial  (Get-Ctx $this) 'allDevices' })
    $Ctx.BtnRemoveInclude.Add_Click({ Invoke-RemoveEntry (Get-Ctx $this) 'Include' })
    $Ctx.BtnRemoveExclude.Add_Click({ Invoke-RemoveEntry (Get-Ctx $this) 'Exclude' })
    if ($Ctx.BtnChangeIntent) { $Ctx.BtnChangeIntent.Add_Click({ Invoke-ChangeIntent (Get-Ctx $this) }) }

    $Ctx.BtnSave.Add_Click({       Invoke-SaveSingle   (Get-Ctx $this) })
    $Ctx.BtnCancel.Add_Click({     Invoke-CancelSingle (Get-Ctx $this) })
    $Ctx.BtnSaveAll.Add_Click({    Invoke-SaveMany     (Get-Ctx $this) -All })
    $Ctx.BtnBulkSave.Add_Click({   Invoke-SaveMany     (Get-Ctx $this) })
    $Ctx.BtnBulkCancel.Add_Click({ Invoke-CancelBulk   (Get-Ctx $this) })
    $Ctx.BtnBulkInclude.Add_Click({ Invoke-BulkAdd (Get-Ctx $this) 'Include' })
    $Ctx.BtnBulkExclude.Add_Click({ Invoke-BulkAdd (Get-Ctx $this) 'Exclude' })
    $Ctx.BtnBulkAllUsers.Add_Click({   Invoke-BulkAddSpecial (Get-Ctx $this) 'allUsers' })
    $Ctx.BtnBulkAllDevices.Add_Click({ Invoke-BulkAddSpecial (Get-Ctx $this) 'allDevices' })

    if (-not $Ctx.IsApp) {
        $Ctx.BtnView.Add_Click({     $c = Get-Ctx $this; Show-PolicyContent $c @($c.SelectedKey) })
        $Ctx.BtnBulkView.Add_Click({ $c = Get-Ctx $this; Show-PolicyContent $c $c.SelectedKeys })
        $Ctx.BtnDelete.Add_Click({     $c = Get-Ctx $this; Invoke-DeletePolicies $c @($c.SelectedKey) })
        $Ctx.BtnBulkDelete.Add_Click({ $c = Get-Ctx $this; Invoke-DeletePolicies $c $c.SelectedKeys })
        $Ctx.Grid.Add_MouseDoubleClick({
            $c = Get-Ctx $this
            if ($c.SelectedKey -and (Test-IsRowClick $_.OriginalSource)) { Show-PolicyContent $c @($c.SelectedKey) }
        })
    }
}

#endregion

#region --- Hoofdvenster ---

function Show-MainWindow {
    $xaml = $script:MainXaml.Replace('{{STYLES}}', $script:StylesXaml).
        Replace('{{TAB_P}}', (Get-TabXaml -P 'P' -IsApp $false)).
        Replace('{{TAB_A}}', (Get-TabXaml -P 'A' -IsApp $true))
    $script:Window = ConvertFrom-Xaml $xaml

    foreach ($n in 'BtnConnect', 'LblConn', 'MainTabs', 'BtnClearLog', 'LogBox', 'LogPara', 'StatusBar', 'LblStatus') {
        Set-Variable -Scope Script -Name $n -Value $script:Window.FindName($n)
    }

    $script:Tabs['P'] = New-TabContext -Prefix 'P' -IsApp $false
    $script:Tabs['A'] = New-TabContext -Prefix 'A' -IsApp $true
    foreach ($ctx in $script:Tabs.Values) { Register-TabEvents $ctx }

    $script:BtnClearLog.Add_Click({ $script:LogPara.Inlines.Clear() })

    $script:BtnConnect.Add_Click({
        Set-Status "Verbinding maken - browser opent zo..."
        $script:BtnConnect.IsEnabled = $false
        $script:Window.WindowState = 'Minimized'
        # Geen LoginHint: laat gewoon het normale Microsoft-inlogscherm zien,
        # zodat je daar zelf het account voor de juiste tenant intypt.
        $ok = Connect-ToGraph
        $script:Window.WindowState = 'Normal'
        [void]$script:Window.Activate()
        $script:BtnConnect.IsEnabled = $true
        if ($ok) {
            $ctx = Get-MgContext
            $acc = $ctx.Account
            $script:LblConn.Text = $acc
            $script:LblConn.Foreground = New-Brush "#90EE90"
            $script:BtnConnect.Content = "Opnieuw verbinden"
            Set-Status "Verbonden als $acc" "success"
            Write-Log "Verbonden met Microsoft Graph als: $acc (tenant $($ctx.TenantId))" "success"
            Write-Log "Toegekende scopes: $($ctx.Scopes -join ', ')" "info"
            $missing = @("DeviceManagementConfiguration.ReadWrite.All", "DeviceManagementApps.ReadWrite.All", "Group.Read.All") | Where-Object { $ctx.Scopes -notcontains $_ }
            if (@($missing).Count -gt 0) {
                Write-Log "LET OP: deze scope(s) staan niet in de toegekende lijst, mogelijk ontbreekt admin consent: $($missing -join ', ')" "warning"
            }
        }
        else {
            Set-Status "Verbinding mislukt" "error"
            Write-Log "Verbinding met Microsoft Graph mislukt" "error"
        }
    })

    $script:Window.Add_Closing({
        $n = 0
        foreach ($c in $script:Tabs.Values) { $n += @($c.Data.Values | Where-Object { $_.Dirty }).Count }
        if ($n -gt 0) {
            $r = Show-Msg "Er zijn $n item(s) met niet-opgeslagen wijzigingen. Toch afsluiten?" "Afsluiten" "YesNo" "Warning"
            if ($r -ne 'Yes') { $_.Cancel = $true }
        }
    })

    [void]$script:Window.ShowDialog()
}

#endregion

Show-MainWindow
