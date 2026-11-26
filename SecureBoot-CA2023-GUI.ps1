<#
.SYNOPSIS
    WPF-GUI voor de Secure Boot CA 2023 scripts (Check_UEFI-CA2023, Update_UEFI-CA2023, Check_DBXUpdate.bin).

.DESCRIPTION
    - Toont dynamisch de systeemstatus (Windows-build incl. 26H2, UBR, Secure Boot, CA 2023 in KEK/DB,
      PCA 2011 in DBX, servicing-status, BitLocker).
    - Alles-in-1: Check_UEFI-CA2023, Update_UEFI-CA2023, Check_DBXUpdate.bin en README_UEFI.TXT zitten
      ingebed in dit script (regio 'Ingebedde bestanden') en worden bij start uitgepakt naar %TEMP%\SecureBootCA2023.
      Logbestanden (-Log) komen in de map van dit script.
    - Leest van de ingebedde scripts zelf de parameters + help uit (AST), zodat de opties in de GUI
      automatisch meegroeien als je een nieuwere scriptversie in de regio plakt.
    - Voert het gekozen script uit in een achtergrond-runspace en streamt de uitvoer (met kleuren) live.
#>

#region Elevatie
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
#endregion

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$script:Root       = $PSScriptRoot
$script:SelfName   = Split-Path -Leaf $PSCommandPath
$script:WorkDir    = Join-Path $env:TEMP 'SecureBootCA2023'
$script:Defs       = @()
$script:ParamCtrls = @()
$script:PS         = $null
$script:RS         = $null
$script:Async      = $null
$script:OutCol     = $null
$script:OutIdx     = 0
$script:ExitCode   = $null
$script:LastLineOpen = $false

#region XAML
[xml]$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Secure Boot CA 2023 - Beheer" Width="1240" Height="820" MinWidth="980" MinHeight="620"
        WindowStartupLocation="CenterScreen" Background="#F3F5F8" FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0F6CBD"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.85"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.4"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="6"/>
            <Setter Property="BorderBrush" Value="#DDE1E6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14"/>
        </Style>
    </Window.Resources>

    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="2,0,0,12">
            <TextBlock Text="Secure Boot CA 2023 - Beheer" FontSize="22" FontWeight="SemiBold" Foreground="#1B1B1B"/>
            <TextBlock x:Name="txtSub" Foreground="#5F6B7A" Margin="0,2,0,0"/>
        </StackPanel>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="370"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Systeemstatus -->
            <Border Grid.Column="0" Style="{StaticResource Card}">
                <DockPanel>
                    <TextBlock DockPanel.Dock="Top" Text="Systeemstatus" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <StackPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
                        <WrapPanel>
                            <Button x:Name="btnRefresh" Content="Vernieuwen" Margin="0,0,8,8"/>
                            <Button x:Name="btnTask" Content="Secure-Boot-Update taak starten" Background="#5B5FC7" Margin="0,0,8,8"/>
                            <Button x:Name="btnSuspend" Content="BitLocker 1 reboot pauzeren" Background="#C4314B" Margin="0,0,8,8"/>
                            <Button x:Name="btnReboot" Content="Herstarten" Background="#B4009E" Margin="0,0,8,8"/>
                        </WrapPanel>
                    </StackPanel>
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <ItemsControl x:Name="lstStatus">
                            <ItemsControl.ItemTemplate>
                                <DataTemplate>
                                    <Grid Margin="0,0,0,8">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="16"/>
                                            <ColumnDefinition Width="135"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Ellipse Width="9" Height="9" Fill="{Binding Kleur}" VerticalAlignment="Top" Margin="0,5,0,0"/>
                                        <TextBlock Grid.Column="1" Text="{Binding Naam}" Foreground="#5F6B7A" TextWrapping="Wrap" Margin="0,0,6,0"/>
                                        <TextBlock Grid.Column="2" Text="{Binding Waarde}" TextWrapping="Wrap" FontWeight="SemiBold"/>
                                    </Grid>
                                </DataTemplate>
                            </ItemsControl.ItemTemplate>
                        </ItemsControl>
                    </ScrollViewer>
                </DockPanel>
            </Border>

            <!-- Script + uitvoer -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="14"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" Style="{StaticResource Card}">
                    <StackPanel>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Script" FontSize="16" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,12,0"/>
                            <ComboBox x:Name="cmbScript" Grid.Column="1" Padding="6,4" DisplayMemberPath="Display"/>
                            <TextBlock x:Name="txtScriptVer" Grid.Column="2" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="12,0,0,0"/>
                        </Grid>
                        <TextBlock x:Name="txtSynopsis" TextWrapping="Wrap" Foreground="#333" Margin="0,10,0,0"/>
                        <TextBlock Text="Opties (beweeg over een optie voor uitleg)" Foreground="#5F6B7A" Margin="0,12,0,4"/>
                        <WrapPanel x:Name="pnlParams"/>
                        <WrapPanel Margin="0,12,0,0">
                            <Button x:Name="btnRun" Content="Uitvoeren" Background="#107C10"/>
                            <Button x:Name="btnStop" Content="Stoppen" Background="#8A8886" IsEnabled="False"/>
                            <Button x:Name="btnClear" Content="Uitvoer wissen" Background="#5F6B7A"/>
                            <Button x:Name="btnSave" Content="Uitvoer opslaan" Background="#5F6B7A"/>
                            <Button x:Name="btnFolder" Content="Map openen" Background="#5F6B7A"/>
                            <Button x:Name="btnReadme" Content="README (BIOS)" Background="#5F6B7A"/>
                        </WrapPanel>
                    </StackPanel>
                </Border>

                <Border Grid.Row="2" Background="#1E1E1E" CornerRadius="6" Padding="4">
                    <RichTextBox x:Name="rtbOut" IsReadOnly="True" Background="#1E1E1E" Foreground="#DCDCDC" BorderThickness="0"
                                 FontFamily="Cascadia Mono, Consolas" FontSize="12.5" VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto">
                        <FlowDocument PageWidth="2400">
                            <Paragraph x:Name="paraOut" Margin="4"/>
                        </FlowDocument>
                    </RichTextBox>
                </Border>
            </Grid>
        </Grid>

        <Border Grid.Row="2" Margin="0,10,0,0">
            <DockPanel>
                <ProgressBar x:Name="prg" DockPanel.Dock="Right" Width="160" Height="6" IsIndeterminate="False" Visibility="Hidden"/>
                <TextBlock x:Name="txtStatus" Foreground="#5F6B7A" Text="Gereed"/>
            </DockPanel>
        </Border>
    </Grid>
</Window>
'@
#endregion

$reader = New-Object System.Xml.XmlNodeReader $Xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)
$script:UI = @{}
foreach ($n in $Xaml.SelectNodes('//*[@*[local-name()="Name"]]')) {
    $name = $n.Attributes | Where-Object { $_.LocalName -eq 'Name' } | Select-Object -First 1 -ExpandProperty Value
    $script:UI[$name] = $script:Window.FindName($name)
}

#region Helpers: build-ondersteuning / systeemstatus
function Get-BuildSupport {
    param([int]$Build, [int]$UBR)
    # Zelfde logica als Confirm-MinimumUBR in de scripts (incl. 26H2 = 26300)
    switch ($Build) {
        { $_ -in 19044, 19045 }        { if ($UBR -lt 6456) { return @('Fout', 'KB5066791 (okt 2025) of later nodig') } ; return @('OK', 'Ondersteund') }
        22000                          { if ($UBR -lt 3260) { return @('Fout', 'KB5044280 (okt 2025) of later nodig') } ; return @('OK', 'Ondersteund') }
        { $_ -in 22621, 22631 }        { if ($UBR -lt 6060) { return @('Fout', 'KB5066793 (okt 2025) of later nodig') } ; return @('OK', 'Ondersteund') }
        { $_ -in 26100, 26200, 26300 } { if ($UBR -lt 6899) { return @('Fout', 'KB5066835 (okt 2025) of later nodig') } ; return @('OK', 'Ondersteund') }
        { $_ -gt 26300 }               { return @('OK', 'Ondersteund (nieuwere build)') }
        default                        { return @('Fout', 'Build niet ondersteund') }
    }
}

function New-StatusItem {
    param([string]$Naam, [string]$Waarde, [ValidateSet('OK', 'Waarschuwing', 'Fout', 'Info')][string]$Status = 'Info')
    $kleur = switch ($Status) { 'OK' { '#107C10' } 'Waarschuwing' { '#D88B00' } 'Fout' { '#C4314B' } default { '#8A8886' } }
    [pscustomobject]@{ Naam = $Naam; Waarde = $Waarde; Kleur = $kleur }
}

function Test-UefiVarContains {
    param([string]$Var, [string]$Text)
    try {
        $bytes = (Get-SecureBootUEFI -Name $Var -ErrorAction Stop).Bytes
        return ([Text.Encoding]::ASCII.GetString($bytes) -match [regex]::Escape($Text))
    }
    catch { return $null }
}

function Get-SystemStatus {
    $items = New-Object System.Collections.Generic.List[object]
    $cv = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion'
    $build = [int]$cv.CurrentBuildNumber
    $ubr = [int]$cv.UBR
    $sup = Get-BuildSupport -Build $build -UBR $ubr

    $cs = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_BIOS

    $items.Add((New-StatusItem 'Apparaat' ('{0} {1}' -f $cs.Manufacturer, $cs.Model)))
    $items.Add((New-StatusItem 'BIOS' ('{0} ({1})' -f $bios.SMBIOSBIOSVersion, $(if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString('yyyy-MM-dd') }))))
    $prod = [string]$cv.ProductName
    if ($build -ge 22000) { $prod = $prod -replace 'Windows 10', 'Windows 11' }
    $items.Add((New-StatusItem 'Windows' ('{0} {1}' -f $prod, $cv.DisplayVersion)))
    $items.Add((New-StatusItem 'Build' ('{0}.{1}  -  {2}' -f $build, $ubr, $sup[1]) $sup[0]))

    $sb = $null
    try { $sb = Confirm-SecureBootUEFI -ErrorAction Stop } catch { $sb = $null }
    if ($sb -eq $true) { $items.Add((New-StatusItem 'Secure Boot' 'Aan' 'OK')) }
    elseif ($sb -eq $false) { $items.Add((New-StatusItem 'Secure Boot' 'Uit (UEFI, wel ondersteund)' 'Waarschuwing')) }
    else { $items.Add((New-StatusItem 'Secure Boot' 'Niet ondersteund / legacy BIOS' 'Fout')) }

    if ($null -ne $sb) {
        $kek = Test-UefiVarContains 'KEK' 'Microsoft Corporation KEK 2K CA 2023'
        $db  = Test-UefiVarContains 'db'  'Windows UEFI CA 2023'
        $dbo = Test-UefiVarContains 'db'  'Microsoft UEFI CA 2023'
        $dbx = Test-UefiVarContains 'dbx' 'Microsoft Windows Production PCA 2011'
        $items.Add((New-StatusItem 'KEK CA 2023' $(if ($kek) { 'Aanwezig' } else { 'Ontbreekt' }) $(if ($kek) { 'OK' } else { 'Fout' })))
        $items.Add((New-StatusItem 'DB Windows CA 2023' $(if ($db) { 'Aanwezig' } else { 'Ontbreekt' }) $(if ($db) { 'OK' } else { 'Fout' })))
        $items.Add((New-StatusItem 'DB UEFI CA 2023' $(if ($dbo) { 'Aanwezig' } else { 'Ontbreekt (3rd-party/optie-ROM)' }) $(if ($dbo) { 'OK' } else { 'Waarschuwing' })))
        $items.Add((New-StatusItem 'PCA 2011 in DBX' $(if ($dbx) { 'Ingetrokken' } else { 'Nog niet ingetrokken' }) $(if ($dbx) { 'OK' } else { 'Info' })))
    }

    $svc = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing' -ErrorAction SilentlyContinue
    $sbk = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot' -ErrorAction SilentlyContinue
    if ($svc -and $null -ne $svc.UEFICA2023Status) {
        $st = [string]$svc.UEFICA2023Status
        $items.Add((New-StatusItem 'Servicing-status' $st $(if ($st -eq 'Updated') { 'OK' } elseif ($st -eq 'InProgress') { 'Waarschuwing' } else { 'Info' })))
    }
    if ($svc -and $null -ne $svc.WindowsUEFICA2023Capable) {
        $cap = switch ([int]$svc.WindowsUEFICA2023Capable) { 0 { 'Nee (0)' } 1 { 'Cert in DB (1)' } 2 { 'Boot met CA 2023 (2)' } default { [string]$_ } }
        $items.Add((New-StatusItem 'CA2023 capable' $cap $(if ([int]$svc.WindowsUEFICA2023Capable -eq 2) { 'OK' } else { 'Info' })))
    }
    if ($svc -and $svc.UEFICA2023Error) {
        $items.Add((New-StatusItem 'Servicing-fout' ('0x{0:X}' -f [int64]$svc.UEFICA2023Error) 'Fout'))
    }
    if ($sbk -and $null -ne $sbk.AvailableUpdates) {
        $items.Add((New-StatusItem 'AvailableUpdates' ('0x{0:X}' -f [int64]$sbk.AvailableUpdates) $(if ([int64]$sbk.AvailableUpdates -eq 0 -or [int64]$sbk.AvailableUpdates -eq 0x4000) { 'OK' } else { 'Waarschuwing' })))
    }

    try {
        $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        $prot = [string]$bl.ProtectionStatus
        $txt = '{0} ({1})' -f $(if ($prot -eq 'On') { 'Beveiligd' } else { 'Niet beveiligd / gepauzeerd' }), $bl.VolumeStatus
        $items.Add((New-StatusItem 'BitLocker' $txt $(if ($prot -eq 'On') { 'Waarschuwing' } else { 'Info' })))
    }
    catch { $items.Add((New-StatusItem 'BitLocker' 'Onbekend' 'Info')) }

    # HP: 2023-certificaatinstellingen in de BIOS (dynamisch, naam verschilt per model)
    if ($cs.Manufacturer -match 'HP|Hewlett') {
        try {
            $hp = @(Get-CimInstance -Namespace 'root/HP/InstrumentedBIOS' -ClassName HP_BIOSEnumeration -ErrorAction Stop | Where-Object { $_.Name -match 'CA\s*2023' } | Sort-Object Name)
            if ($hp.Count -eq 0) { $items.Add((New-StatusItem 'HP BIOS 2023' 'Geen CA 2023-instellingen (BIOS updaten?)' 'Waarschuwing')) }
            foreach ($h in $hp) {
                $on = ($h.CurrentValue -match '^Enable')
                $items.Add((New-StatusItem ('HP: ' + ($h.Name -replace '^Microsoft ', 'MS ')) ([string]$h.CurrentValue) $(if ($on) { 'OK' } else { 'Waarschuwing' })))
            }
        }
        catch { $items.Add((New-StatusItem 'HP BIOS 2023' 'HP BIOS WMI niet beschikbaar' 'Info')) }
    }

    return , $items.ToArray()
}

function Update-SystemStatus {
    $script:UI.txtStatus.Text = 'Systeemstatus ophalen...'
    $script:Window.Dispatcher.Invoke([action] {}, 'Render')
    try {
        $script:UI.lstStatus.ItemsSource = Get-SystemStatus
        $cv = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion'
        $script:UI.txtSub.Text = '{0}  |  Windows {1} (build {2}.{3})  |  PowerShell {4}' -f $env:COMPUTERNAME, $cv.DisplayVersion, $cv.CurrentBuildNumber, $cv.UBR, $PSVersionTable.PSVersion
        $script:UI.txtStatus.Text = 'Systeemstatus bijgewerkt om {0:HH:mm:ss}' -f (Get-Date)
    }
    catch { $script:UI.txtStatus.Text = "Fout bij ophalen status: $($_.Exception.Message)" }
}
#endregion

#region Helpers: scripts dynamisch inlezen
function Export-EmbeddedFiles {
    # Ingebedde scripts/README uitpakken naar een werkmap. Logbestanden (-Log) komen naast deze GUI.
    if (-not (Test-Path $script:WorkDir)) { New-Item -ItemType Directory -Path $script:WorkDir -Force | Out-Null }
    Get-ChildItem -Path $script:WorkDir -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    $logRoot = "'" + ($script:Root -replace "'", "''") + "'"
    foreach ($k in $script:Embedded.Keys) {
        $content = $script:Embedded[$k]
        if ($k -like '*.ps1') { $content = $content.Replace('$PSScriptRoot', $logRoot) }
        [IO.File]::WriteAllText((Join-Path $script:WorkDir $k), $content, (New-Object Text.UTF8Encoding $true))
    }
}

function Get-ScriptDefinitions {
    $list = @()
    foreach ($f in Get-ChildItem -Path $script:WorkDir -Filter '*.ps1' -File | Sort-Object Name) {
        $tokens = $null; $errs = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errs)
        $help = $ast.GetHelpContent()
        $ver = ''
        $m = [regex]::Match($ast.Extent.Text, '(?m)^\.VERSION\s+(\S+)')
        if ($m.Success) { $ver = $m.Groups[1].Value }

        $helpParams = @{}
        if ($help -and $help.Parameters) { foreach ($k in $help.Parameters.Keys) { $helpParams[$k.ToUpper()] = ($help.Parameters[$k] -replace '\s+', ' ').Trim() } }

        $params = @()
        if ($ast.ParamBlock) {
            foreach ($p in $ast.ParamBlock.Parameters) {
                $name = $p.Name.VariablePath.UserPath
                if ($name -in 'Version', 'ignored') { continue }
                $t = $p.StaticType
                $kind = if ($t -eq [switch]) { 'Switch' } elseif ($t -eq [string[]]) { 'StringArray' } else { 'String' }
                $default = ''
                if ($p.DefaultValue) { $default = $p.DefaultValue.Extent.Text.Trim('"', "'") }
                $params += [pscustomobject]@{ Name = $name; Kind = $kind; Default = $default; Help = $helpParams[$name.ToUpper()] }
            }
            # -Verbose is een common parameter; tonen als het script hem in de help documenteert
            if ($helpParams.ContainsKey('VERBOSE') -and -not ($params.Name -contains 'Verbose')) {
                $params = @([pscustomobject]@{ Name = 'Verbose'; Kind = 'Switch'; Default = ''; Help = $helpParams['VERBOSE'] }) + $params
            }
        }

        $syn = if ($help) { (($help.Synopsis, $help.Description) -join ' ' -replace '\s+', ' ').Trim() } else { '' }
        $list += [pscustomobject]@{
            Name     = $f.BaseName
            Path     = $f.FullName
            Version  = $ver
            Synopsis = $syn
            Params   = $params
            Display  = $f.BaseName
        }
    }
    return , $list
}

function Show-ScriptParams {
    $def = $script:UI.cmbScript.SelectedItem
    $pnl = $script:UI.pnlParams
    $pnl.Children.Clear()
    $script:ParamCtrls = @()
    if (-not $def) { return }

    $script:UI.txtSynopsis.Text = $def.Synopsis
    $script:UI.txtScriptVer.Text = $(if ($def.Version) { "versie $($def.Version)" } else { '' })

    foreach ($p in $def.Params) {
        if ($p.Kind -eq 'Switch') {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $p.Name
            $cb.Margin = '0,4,22,4'
            $cb.VerticalContentAlignment = 'Center'
            if ($p.Help) { $cb.ToolTip = (New-ToolTip $p.Help) }
            [void]$pnl.Children.Add($cb)
            $script:ParamCtrls += [pscustomobject]@{ Name = $p.Name; Kind = $p.Kind; Control = $cb }
        }
        else {
            $sp = New-Object System.Windows.Controls.StackPanel
            $sp.Orientation = 'Horizontal'
            $sp.Margin = '0,4,22,4'
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = "$($p.Name):"
            $lbl.VerticalAlignment = 'Center'
            $lbl.Margin = '0,0,6,0'
            $tb = New-Object System.Windows.Controls.TextBox
            $tb.Width = 330
            $tb.Padding = '4,3'
            $tb.ToolTip = (New-ToolTip ($(if ($p.Help) { $p.Help } else { $p.Name }) + $(if ($p.Default) { "`n`nStandaard: $($p.Default)" } else { '' }) + $(if ($p.Kind -eq 'StringArray') { "`nMeerdere waarden scheiden met ;" } else { '' })))
            [void]$sp.Children.Add($lbl)
            [void]$sp.Children.Add($tb)
            if ($p.Name -match 'Folder|Path') {
                $btn = New-Object System.Windows.Controls.Button
                $btn.Content = '...'
                $btn.Padding = '10,3'
                $btn.Margin = '6,0,0,0'
                $btn.Tag = $tb
                $btn.Add_Click({
                    param($s, $e)
                    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
                    if ($dlg.ShowDialog() -eq 'OK') {
                        $target = $s.Tag
                        if ($target.Text) { $target.Text = $target.Text + ';' + $dlg.SelectedPath } else { $target.Text = $dlg.SelectedPath }
                    }
                })
                [void]$sp.Children.Add($btn)
            }
            [void]$pnl.Children.Add($sp)
            $script:ParamCtrls += [pscustomobject]@{ Name = $p.Name; Kind = $p.Kind; Control = $tb }
        }
    }
}

function New-ToolTip([string]$Text) {
    $tt = New-Object System.Windows.Controls.ToolTip
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Text
    $tb.TextWrapping = 'Wrap'
    $tb.MaxWidth = 460
    $tt.Content = $tb
    return $tt
}
#endregion

#region Helpers: uitvoer
$script:ColorMap = @{
    'Black' = '#8A8A8A'; 'DarkBlue' = '#569CD6'; 'DarkGreen' = '#4EC94E'; 'DarkCyan' = '#3FB6C9'
    'DarkRed' = '#E05252'; 'DarkMagenta' = '#C586C0'; 'DarkYellow' = '#D7BA7D'; 'Gray' = '#DCDCDC'
    'DarkGray' = '#9A9A9A'; 'Blue' = '#6CB6FF'; 'Green' = '#6BD968'; 'Cyan' = '#6BD3F2'
    'Red' = '#FF6B6B'; 'Magenta' = '#E28BE2'; 'Yellow' = '#F2D16B'; 'White' = '#FFFFFF'
}
$script:BrushConv = New-Object System.Windows.Media.BrushConverter

function Add-Output {
    param([string]$Text, [string]$Color = '#DCDCDC', [switch]$NoNewLine)
    $para = $script:UI.paraOut
    $lines = $Text -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $run = New-Object System.Windows.Documents.Run $lines[$i]
        $run.Foreground = $script:BrushConv.ConvertFromString($Color)
        $para.Inlines.Add($run)
        $isLast = ($i -eq $lines.Count - 1)
        if (-not $isLast -or -not $NoNewLine) { $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak)) }
    }
}

function Get-LineColor([string]$Line) {
    if ($Line -match '^\s*(ERROR|FOUT|FAILED)') { return '#FF6B6B' }
    if ($Line -match '^\s*(SUCCESS|PASS)') { return '#6BD968' }
    if ($Line -match '^\s*(WARNING|WAARSCHUWING|REQUIRED ACTION)') { return '#F2D16B' }
    if ($Line -match '^[=\-]{6,}\s*$') { return '#6A6A6A' }
    return '#DCDCDC'
}

function Write-OutputItem($item) {
    if ($null -eq $item) { return }
    $obj = $item
    if ($item -is [psobject]) { $obj = $item.PSObject.BaseObject }

    if ($obj -is [System.Management.Automation.InformationRecord]) {
        $md = $obj.MessageData
        if ($md -is [System.Management.Automation.HostInformationMessage]) {
            $col = '#DCDCDC'
            if ($null -ne $md.ForegroundColor -and $script:ColorMap.ContainsKey([string]$md.ForegroundColor)) { $col = $script:ColorMap[[string]$md.ForegroundColor] }
            if ($md.NoNewLine) { Add-Output -Text ([string]$md.Message) -Color $col -NoNewLine } else { Add-Output -Text ([string]$md.Message) -Color $col }
        }
        else { Add-Output -Text ([string]$md) }
    }
    elseif ($obj -is [System.Management.Automation.ErrorRecord]) { Add-Output -Text ("ERROR: " + $obj.ToString()) -Color '#FF6B6B' }
    elseif ($obj -is [System.Management.Automation.WarningRecord]) { Add-Output -Text ("WARNING: " + $obj.Message) -Color '#F2D16B' }
    elseif ($obj -is [System.Management.Automation.VerboseRecord]) { Add-Output -Text ("VERBOSE: " + $obj.Message) -Color '#6CB6FF' }
    elseif ($obj -is [System.Management.Automation.DebugRecord]) { Add-Output -Text ("DEBUG: " + $obj.Message) -Color '#9A9A9A' }
    else {
        $s = [string]$obj
        if ($s -match '^__EXITCODE__:(.*)$') { $script:ExitCode = $Matches[1]; return }
        foreach ($l in ($s -split "`r?`n")) { Add-Output -Text $l -Color (Get-LineColor $l) }
    }
}
#endregion

#region Uitvoeren in runspace
$script:Wrapper = {
    param($ScriptPath, $Params)
    Set-Location -LiteralPath (Split-Path -Parent $ScriptPath)
    $ErrorActionPreference = 'Continue'
    $global:LASTEXITCODE = 0
    & $ScriptPath @Params *>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.InformationRecord] -or $_ -is [System.Management.Automation.ErrorRecord] -or
            $_ -is [System.Management.Automation.WarningRecord] -or $_ -is [System.Management.Automation.VerboseRecord] -or
            $_ -is [System.Management.Automation.DebugRecord] -or $_ -is [string]) { $_ }
        else { ($_ | Out-String -Width 300).TrimEnd() }
    }
    "__EXITCODE__:$LASTEXITCODE"
}

function Set-Running([bool]$Running) {
    $script:UI.btnRun.IsEnabled = -not $Running
    $script:UI.btnStop.IsEnabled = $Running
    $script:UI.cmbScript.IsEnabled = -not $Running
    $script:UI.pnlParams.IsEnabled = -not $Running
    $script:UI.btnRefresh.IsEnabled = -not $Running
    $script:UI.prg.IsIndeterminate = $Running
    $script:UI.prg.Visibility = $(if ($Running) { 'Visible' } else { 'Hidden' })
}

function Start-SelectedScript {
    $def = $script:UI.cmbScript.SelectedItem
    if (-not $def) { return }

    $splat = @{}
    $argText = @()
    foreach ($pc in $script:ParamCtrls) {
        switch ($pc.Kind) {
            'Switch' { if ($pc.Control.IsChecked) { $splat[$pc.Name] = $true; $argText += "-$($pc.Name)" } }
            'String' { $v = $pc.Control.Text.Trim(); if ($v) { $splat[$pc.Name] = $v; $argText += "-$($pc.Name) `"$v`"" } }
            'StringArray' {
                $v = @($pc.Control.Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                if ($v.Count) { $splat[$pc.Name] = [string[]]$v; $argText += "-$($pc.Name) " + (($v | ForEach-Object { "`"$_`"" }) -join ',') }
            }
        }
    }

    # Extra bevestiging voor ingrijpende acties
    if ($def.Name -match '^(Update_|HP_)' -and -not $splat.ContainsKey('Audit')) {
        if ($def.Name -like 'HP_*') { $msg = "Je gaat '$($def.Name)' uitvoeren zonder -Audit. Dit wijzigt HP BIOS-instellingen (2023-certificaten aan)." }
        else { $msg = "Je gaat '$($def.Name)' uitvoeren zonder -Audit. Dit wijzigt UEFI Secure Boot-variabelen/bootbestanden." }
        if ($splat.ContainsKey('Revoke')) { $msg += "`n`nLET OP: -Revoke trekt PCA 2011 in. Oude bootmedia (USB/ISO/PXE) starten daarna niet meer." }
        $bl = $null; try { $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop } catch {}
        if ($bl -and [string]$bl.ProtectionStatus -eq 'On') { $msg += "`n`nBitLocker is actief: pauzeer BitLocker eerst (knop links) om een herstelsleutel-prompt te voorkomen." }
        $msg += "`n`nDoorgaan?"
        if ([System.Windows.MessageBox]::Show($msg, 'Bevestigen', 'YesNo', 'Warning') -ne 'Yes') { return }
    }

    Add-Output -Text ('=' * 100) -Color '#6A6A6A'
    Add-Output -Text ("{0:yyyy-MM-dd HH:mm:ss}  >  {1}.ps1 {2}" -f (Get-Date), $def.Name, ($argText -join ' ')) -Color '#6BD3F2'
    Add-Output -Text ('=' * 100) -Color '#6A6A6A'

    $script:RS = [runspacefactory]::CreateRunspace()
    $script:RS.ApartmentState = 'STA'
    $script:RS.ThreadOptions = 'ReuseThread'
    $script:RS.Open()
    $script:PS = [powershell]::Create()
    $script:PS.Runspace = $script:RS
    [void]$script:PS.AddScript($script:Wrapper.ToString()).AddArgument($def.Path).AddArgument($splat)

    $inCol = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $inCol.Complete()
    $script:OutCol = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $script:OutIdx = 0
    $script:ExitCode = $null
    $script:Async = $script:PS.BeginInvoke($inCol, $script:OutCol)

    Set-Running $true
    $script:UI.txtStatus.Text = "Bezig: $($def.Name)..."
    $script:Timer.Start()
}

function Read-NewOutput {
    $col = $script:OutCol
    if (-not $col) { return }
    $added = $false
    while ($script:OutIdx -lt $col.Count) {
        Write-OutputItem $col[$script:OutIdx]
        $script:OutIdx++
        $added = $true
    }
    if ($added) { $script:UI.rtbOut.ScrollToEnd() }
}

function Complete-Run {
    $script:Timer.Stop()
    Read-NewOutput
    try { [void]$script:PS.EndInvoke($script:Async) }
    catch { Add-Output -Text ("ERROR: " + $_.Exception.InnerException.Message) -Color '#FF6B6B' }
    foreach ($e in $script:PS.Streams.Error) { Add-Output -Text ("ERROR: " + $e.ToString()) -Color '#FF6B6B' }

    $state = $script:PS.InvocationStateInfo.State
    $code = $script:ExitCode
    try { $script:PS.Dispose(); $script:RS.Close(); $script:RS.Dispose() } catch {}
    $script:PS = $null; $script:RS = $null; $script:OutCol = $null

    $endColor = $(if ($state -eq 'Stopped') { '#F2D16B' } elseif ($code -and $code -ne '0') { '#FF6B6B' } else { '#6BD968' })
    $endText = $(if ($state -eq 'Stopped') { 'Gestopt door gebruiker' } else { "Klaar (exitcode: $(if ($code) { $code } else { '0' }))" })
    Add-Output -Text ''
    Add-Output -Text ("--- {0} om {1:HH:mm:ss} ---" -f $endText, (Get-Date)) -Color $endColor
    Add-Output -Text ''
    $script:UI.rtbOut.ScrollToEnd()

    Set-Running $false
    Update-SystemStatus
    $script:UI.txtStatus.Text = $endText
}

$script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$script:Timer.Interval = [TimeSpan]::FromMilliseconds(200)
$script:Timer.Add_Tick({
    Read-NewOutput
    if ($script:Async -and $script:Async.IsCompleted) { Complete-Run }
})
#endregion

#region Events
$script:UI.cmbScript.Add_SelectionChanged({ Show-ScriptParams })
$script:UI.btnRun.Add_Click({ Start-SelectedScript })
$script:UI.btnStop.Add_Click({
    if ($script:PS) { $script:UI.txtStatus.Text = 'Stoppen...'; [void]$script:PS.BeginStop($null, $null) }
})
$script:UI.btnClear.Add_Click({ $script:UI.paraOut.Inlines.Clear() })
$script:UI.btnRefresh.Add_Click({ Update-SystemStatus })
$script:UI.btnFolder.Add_Click({ Start-Process explorer.exe -ArgumentList "`"$script:Root`"" })
$script:UI.btnReadme.Add_Click({
    $rm = Join-Path $script:WorkDir 'README_UEFI.TXT'
    if (Test-Path $rm) { Start-Process notepad.exe -ArgumentList "`"$rm`"" } else { [System.Windows.MessageBox]::Show('README_UEFI.TXT niet gevonden.') | Out-Null }
})
$script:UI.btnSave.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Logbestand (*.log)|*.log|Tekst (*.txt)|*.txt'
    $dlg.InitialDirectory = $script:Root
    $dlg.FileName = '{0:yyyy-MM-dd_HHmm} {1} SecureBoot.log' -f (Get-Date), $env:COMPUTERNAME
    if ($dlg.ShowDialog()) {
        $doc = $script:UI.rtbOut.Document
        $range = New-Object System.Windows.Documents.TextRange($doc.ContentStart, $doc.ContentEnd)
        [IO.File]::WriteAllText($dlg.FileName, $range.Text, [Text.Encoding]::UTF8)
        $script:UI.txtStatus.Text = "Uitvoer opgeslagen: $($dlg.FileName)"
    }
})
$script:UI.btnTask.Add_Click({
    try {
        Start-ScheduledTask -TaskName '\Microsoft\Windows\PI\Secure-Boot-Update' -ErrorAction Stop
        Add-Output -Text ("{0:HH:mm:ss}  Geplande taak 'Secure-Boot-Update' gestart." -f (Get-Date)) -Color '#6BD968'
        $script:UI.txtStatus.Text = 'Secure-Boot-Update taak gestart; vernieuw de status over enkele seconden.'
    }
    catch { Add-Output -Text ("ERROR: Taak starten mislukt: " + $_.Exception.Message) -Color '#FF6B6B' }
})
$script:UI.btnSuspend.Add_Click({
    $r = [System.Windows.MessageBox]::Show("BitLocker op $env:SystemDrive pauzeren voor 1 herstart?`nNa de herstart wordt de beveiliging automatisch hervat.", 'BitLocker pauzeren', 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    try {
        Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount 1 -ErrorAction Stop | Out-Null
        Add-Output -Text ("{0:HH:mm:ss}  BitLocker gepauzeerd op {1} voor 1 herstart." -f (Get-Date), $env:SystemDrive) -Color '#6BD968'
        Update-SystemStatus
    }
    catch { Add-Output -Text ("ERROR: BitLocker pauzeren mislukt: " + $_.Exception.Message) -Color '#FF6B6B' }
})
$script:UI.btnReboot.Add_Click({
    if ($script:PS) {
        [System.Windows.MessageBox]::Show('Er draait nog een script. Wacht tot het klaar is of stop het eerst.', 'Herstarten', 'OK', 'Warning') | Out-Null
        return
    }
    $bl = $null; try { $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop } catch {}
    $blOn = ($bl -and [string]$bl.ProtectionStatus -eq 'On')

    if ($blOn) {
        $r = [System.Windows.MessageBox]::Show("BitLocker is actief op $env:SystemDrive.`n`nJa = BitLocker 1 herstart pauzeren en daarna herstarten (aanbevolen na UEFI-wijzigingen)`nNee = herstarten zonder BitLocker te pauzeren`nAnnuleren = niet herstarten", 'Herstarten', 'YesNoCancel', 'Question')
        if ($r -eq 'Cancel') { return }
        if ($r -eq 'Yes') {
            try { Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount 1 -ErrorAction Stop | Out-Null }
            catch {
                [System.Windows.MessageBox]::Show("BitLocker pauzeren mislukt: $($_.Exception.Message)`n`nHerstart afgebroken.", 'Herstarten', 'OK', 'Error') | Out-Null
                return
            }
        }
    }
    else {
        $r = [System.Windows.MessageBox]::Show("Computer $env:COMPUTERNAME nu herstarten?", 'Herstarten', 'YesNo', 'Question')
        if ($r -ne 'Yes') { return }
    }

    Add-Output -Text ("{0:HH:mm:ss}  Computer wordt herstart..." -f (Get-Date)) -Color '#F2D16B'
    Restart-Computer -Force
})
$script:Window.Add_Closing({
    param($s, $e)
    if ($script:PS) {
        $r = [System.Windows.MessageBox]::Show('Er draait nog een script. Toch afsluiten?', 'Afsluiten', 'YesNo', 'Warning')
        if ($r -ne 'Yes') { $e.Cancel = $true; return }
        try { $script:PS.Stop() } catch {}
    }
})
$script:Window.Add_ContentRendered({
    Update-SystemStatus
})
#endregion

#region Ingebedde bestanden (niet bewerken tenzij je een nieuwe scriptversie plakt)
$script:Embedded = [ordered]@{}
$script:Embedded['Check_UEFI-CA2023.ps1'] = @'
<#PSScriptInfo

.VERSION 2026.01.18

.GUID 240507af-7454-491f-8e42-acb2a40ae3ef

.AUTHOR garlin

.COPYRIGHT

.TAGS UEFI, Secure Boot, CA 2023, PK, KEK, DB, DBX, SVN, Windows Boot Manager

.RELEASENOTES

#>

<#
.SYNOPSIS
    Script to identify Secure Boot certificates installed in the UEFI variables, and signing certs for Windows boot files.

.DESCRIPTION
    Run this script to check Windows compliance with Secure Boot CA 2023 updates, and CA 2011 revocation.

.PARAMETER Version
    Print the script's version number and exit.

.PARAMETER Verbose
    Identify extra details including the Windows build version, BIOS version details, UEFI Platform Key, and Factory Defaults for PK, KEK, DB and DBX variables.
    Download "kek_update_map.json" from Microsoft's Secure Boot Objects GitHub, and check if vendor signed KEK CA 2023 update is available.
    Windows Boot Manager SVN will be reported, if present in DBX.

.PARAMETER Audit
    Perform an audit report of the UEFI variables and Windows Boot Manager version.  Identify any missing UEFI certs, and validate if current boot file is
    allowed by enabling Secure Boot mode.

    Identify all required actions to bring system into compliance for upcoming Windows CA 2023 changes.

    If Secure Boot is currently disabled, audit report will simulate conditions where Secure Boot is enabled.

.PARAMETER BootMedia
    Search all mounted removable media (DVD & USB drives), for Windows boot files and install images.  Validate if boot file and install image are allowed by
    current Secure Boot settings.

.PARAMETER NoSkip
    When checking Windows install files on removable media, examine every image in the install WIM/ESD file.
    By default, -BootMedia parameter stops checking after the first image in the install file to improve script reporting time.

.PARAMETER Log
    Save script output to a file named "YYYY-MM-DD [Model] Check UEFI.log"

.EXAMPLE
    Check_UEFI-CA2023.ps1
.EXAMPLE
    Check_UEFI-CA2023.ps1 -Audit
.EXAMPLE
    Check_UEFI-CA2023.ps1 -Verbose -Audit -BootMedia -Log
#>

[CmdletBinding(DefaultParameterSetName='Default')]
param (
    [Parameter(Mandatory=$false,ParameterSetName='Version')]
    [switch]$Version,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$Audit,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$BootMedia,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$NoSkip,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$Log,

    [Parameter(Mandatory=$false,ParameterSetName='Default',DontShow,ValueFromRemainingArguments=$true)]
    [string[]]$ignored
)

$ScriptVersion = '2026.01.18'

# https://github.com/microsoft/secureboot_objects/blob/main/Archived/dbx_info_msft_4_09_24_svns.csv
$EFI_BOOTMGR_DBXSVN_GUID = '01612B139DD5598843AB1C185C3CB2EB92'
$EFI_CDBOOT_DBXSVN_GUID =  '019D2EF8E827E15841A4884C18ABE2F284'
$EFI_WDSMGR_DBXSVN_GUID =  '01C2CA99C9FE7F6F4981279E2A8A535976'

$VMWARE_GUID = 'a3d5e95b-0a8f-4753-8735-445afb708f62'

$CN_Regex = '(CN=)([^,]+)'

$KEKUpdateMap_URL = 'https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PostSignedObjects/KEK/kek_update_map.json'

$Tab4 = ' ' * 4
$Tab8 = ' ' * 8
$Tab12 = ' ' * 12

if ($Version) {
    '{0} version ({1}){2}' -f $MyInvocation.MyCommand.Name, $ScriptVersion, $(if ($MyInvocation.Line -ne '') { "`n" })
    exit 0
}

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if ($PSVersionTable.PSVersion.Major -gt 5) {
        $PS = 'pwsh'
    }
    else {
        $PS = 'powershell'
    }

    $args = ($MyInvocation.BoundParameters.Keys.GetEnumerator() | where { $_ -notmatch 'ignored' } | foreach { '-{0}' -f $_ }) -join ' '

    Start-Process $PS -ArgumentList "-nop -ep bypass -NoLogo -NoExit -f $($MyInvocation.MyCommand.Path) $args" -Verb RunAs
    exit 0
}

if ($PSBoundParameters['Verbose']) {
    $Verbose = $true
    $VerbosePreference = 'SilentlyContinue'
}

function Get-UefiDatabaseSignatures {
    <#
        .SYNOPSIS
        Parses UEFI Signature Databases into logical Powershell objects
        # https://github.com/cjee21/Check-UEFISecureBootVariables

        .DESCRIPTION
        Original Author: Matthew Graeber (@mattifestation)
        Modified By: Jeremiah Cox (@int0x6)
        Modified By: Joel Roth (@nafai)
        Modified By: garlin (@garlin-cant-code)
        Additional Source: https://gist.github.com/mattifestation/991a0bea355ec1dc19402cef1b0e3b6f
        Additional Source: https://www.powershellgallery.com/packages/SplitDbxContent/1.0
        License: BSD 3-Clause

        .PARAMETER Variable
        Specifies an UEFI variable, an instance of which is returned by calling the Get-SecureBootUEFI cmdlet.

        .PARAMETER BytesIn
        Specifies a byte array consisting of the PKDefault, KEKDefault, dbDefault, dbxDefault, PK, KEK, db, or dbx UEFI variable contents.

        .EXAMPLE
        $DbxBytes = [IO.File]::ReadAllBytes('.\dbx.bin')
        Get-UEFIDatabaseSignatures -BytesIn $DbxBytes

        .EXAMPLE
        Get-UEFIDatabaseSignatures -Filename ".\DBXUpdate-20230314.x64.bin"

        .EXAMPLE
        Get-SecureBootUEFI -Name db | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name dbx | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name pk | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name kek | Get-UEFIDatabaseSignatures

        .INPUTS
        Microsoft.SecureBoot.Commands.UEFIEnvironmentVariable
        Accepts the output of Get-SecureBootUEFI over the pipeline.

        .OUTPUTS
        UefiSignatureDatabase
        Outputs an array of custom powershell objects describing a UEFI Signature Database. "77fa9abd-0359-4d32-bd60-28f4e78f784b" refers to Microsoft as the owner.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'UEFIVariable')]
        [ValidateScript({ ($_.GetType().Fullname -eq 'Microsoft.SecureBoot.Commands.UEFIEnvironmentVariable') -and ($_.Name -in 'PKDefault','KEKDefault','dbDefault','dbxDefault','pk','kek','db','dbx') })]
        $Variable,

        [Parameter(Mandatory, ParameterSetName = 'ByteArray')]
        [Byte[]]
        [ValidateNotNullOrEmpty()]
        $BytesIn,

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]
        [ValidateScript({ (Resolve-Path "$_").where({Test-Path $_}).Path })]
        $Filename
    )

    $PSVersion = $PSVersionTable.PSVersion.Major

    $SignatureTypeMapping = @{
        'C1C41626-504C-4092-ACA9-41F936934328' = 'EFI_CERT_SHA256_GUID' # Most often used for dbx
        'A5C059A1-94E4-4AA7-87B5-AB155C2BF072' = 'EFI_CERT_X509_GUID'   # Most often used for db
    }

    $Bytes = $null

    if ($Filename)
    {
        if ($PSVersion -gt 5) {
            $Bytes = Get-Content -AsByteStream $Filename -ErrorAction Stop
        }
        else {
            $Bytes = Get-Content -Encoding Byte $Filename  -ErrorAction Stop
        }
    }
    elseif ($Variable)
    {
        $Bytes = $Variable.Bytes
    }
    else
    {
        $Bytes = $BytesIn
    }

    # Modified from Split-Dbx
    if (($Bytes[40] -eq 0x30) -and ($Bytes[41] -eq 0x82 ))
    {
        Write-Debug "Removing signature."

        # Signature is known to be ASN size plus header of 4 bytes
        $sig_length = $Bytes[42] * 256 + $Bytes[43] + 4
        if ($sig_length -gt ($Bytes.Length + 40)) {
            Write-Error "Signature longer than file size!" -ErrorAction Stop
        }

        ## Unsigned db store
        [System.Byte[]]$Bytes = @($Bytes[($sig_length+40)..($Bytes.Length - 1)].Clone())
    }
    else
    {
        Write-Debug "Signature not found. Assuming it's already split."
    }

    try
    {
        $MemoryStream = New-Object -TypeName IO.MemoryStream -ArgumentList @(,$Bytes)
        $BinaryReader = New-Object -TypeName IO.BinaryReader -ArgumentList $MemoryStream, ([Text.Encoding]::Unicode)
    }
    catch
    {
        throw $_
        return
    }

    # What follows will be an array of EFI_SIGNATURE_LIST structs

    while ($BinaryReader.PeekChar() -ne -1) {
        $SignatureType = $SignatureTypeMapping[([Guid][Byte[]] $BinaryReader.ReadBytes(16)).Guid]
        $SignatureListSize = $BinaryReader.ReadUInt32()
        $SignatureHeaderSize = $BinaryReader.ReadUInt32()
        $SignatureSize = $BinaryReader.ReadUInt32()

        $SignatureHeader = $BinaryReader.ReadBytes($SignatureHeaderSize)

        # 0x1C is the size of the EFI_SIGNATURE_LIST header
        $SignatureCount = ($SignatureListSize - 0x1C) / $SignatureSize

        $SignatureList = 1..$SignatureCount | ForEach-Object {
            $SignatureDataBytes = $BinaryReader.ReadBytes($SignatureSize)

            $SignatureOwner = [Guid][Byte[]] $SignatureDataBytes[0..15]

            switch ($SignatureType) {
                'EFI_CERT_SHA256_GUID' {
                    $SignatureData = ([Byte[]] $SignatureDataBytes[0x10..0x2F] | ForEach-Object { $_.ToString('X2') }) -join ''
                }

                'EFI_CERT_X509_GUID' {
                    try {
                        $SignatureData = New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,([Byte[]] $SignatureDataBytes[16..($SignatureDataBytes.Count - 1)]))
                    }
                    catch {
                        Write-Host "Skipping an invalid $Variable X509 certificate."
                    }
                }
            }

            [PSCustomObject] @{
                PSTypeName = 'EFI.SignatureData'
                SignatureOwner = $SignatureOwner
                SignatureData = $SignatureData
            }
        }

        [PSCustomObject] @{
            PSTypeName = 'EFI.SignatureList'
            SignatureType = $SignatureType
            SignatureList = $SignatureList
        }
    }
}

function Get-UEFICert {
    param (
        [Parameter(Mandatory)]
        [ValidateSet('PKDefault','KEKDefault','dbDefault','dbxDefault','PK','KEK','db','dbx')]
        [string]$Variable
    )

    try {
        $SignatureList = (Get-SecureBootUEFI $Variable | Get-UefiDatabaseSignatures).SignatureList
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return @()
        }
        else {
            throw $_.Exception.Message
        }
    }

    $Subject = $SignatureList.SignatureData.Subject

    if ($Verbose) {
        $Certs = $Subject | where { $_ -match '\s' } | foreach { $null = $_ -match $CN_Regex; $Matches[2] }
    }
    else {
        $Certs = $Subject | where { $_ -match 'Microsoft|Mosby' } | foreach { $null = $_ -match $CN_Regex; $Matches[2] }
    }

    if ($Variable -match 'PK') {
        if ($SignatureList.SignatureData -eq $null -and $SignatureList.SignatureOwner.Guid -eq $VMWARE_GUID) {
            $Certs = @('VMware Default PK')
        }
        elseif ($Subject -match 'VirtualBox') {
            $Certs = @('VirtualBox UEFI PK')
        }
    }

    return $Certs
}

function Print-Header {
    param (
        [Parameter(Mandatory=$false)]
        [switch]$Bold,

        [Parameter(Mandatory)]
        [string]$Header
    )

    if ($Bold) {
        $Separator = '='
    }
    else {
        $Separator = '-'
    }

    return ("`n{0}`n{1}" -f $Header, ($Header -replace "`n" -replace '(.)',$Separator))
}

function Print-UEFICerts {
    param (
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ref]$CertArray
    )

    if ($Name -match 'Default') {
        $Header = '{0} Cert{1}' -f ($Name -replace 'Default','Factory Default UEFI'), $(if ($Name -notmatch 'PK') { 's' })
    }
    else {
        $Header = 'UEFI {0} Cert{1}' -f $Name, $(if ($Name -notmatch 'PK') { 's' })
    }

    Print-Header $Header

    if (($CertArray.Value).Count) {
        $SortedArray = @(foreach ($item in $CertArray.Value) {
            [PSCustomObject] @{
                Date = ($item -split ' ')[-1]
                Name = $item
            }
        }) | sort Date, Name | select -ExpandProperty Name

        $SortedArray | foreach { '{0}{1}' -f $Tab4, $_ }
    }
    else {
        '{0}(NONE)' -f $Tab4
    }
}

function Check-TrustedPK {
    try {
        $PKSignatureList = (Get-SecureBootUEFI PK | Get-UefiDatabaseSignatures).SignatureList
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $false
        }
        else {
            throw $_.Exception.Message
        }
    }

    if ($PKSignatureList.SignatureData.Subject -notmatch 'DO NOT |Example') {
        return $true
    }
    else {
        return $false
    }
}

function Check-KEKUpdateMap {
    try {
        $JSON = (Invoke-WebRequest -UseBasicParsing -Uri $KEKUpdateMap_URL).Content | ConvertFrom-Json
    }
    catch {
        return (($_.Exception.Message -split "`n") | select -First 1)
    }

    try {
        $PK_Thumbprint = (Get-UefiDatabaseSignatures -BytesIn (Get-SecureBootUEFI PK).Bytes).SignatureList.SignatureData.Thumbprint
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $null
        }
        else {
            throw $_.Exception.Message
        }
    }

    if ($JSON.$PK_Thumbprint.KEKUpdate -ne $null) {
        $script:SignedKEK = $true
        return $JSON.$PK_Thumbprint.KEKUpdate
    }
    else {
        return $null
    }
}

function Get-SignatureDataSVN {
    param (
        [Parameter(Mandatory)]
        [string]$SignatureData
    )

    # https://github.com/microsoft/secureboot_objects/blob/main/scripts/utility_functions.py
    $SVN = '{0}.{1}' -f [int]::Parse($SignatureData.Substring(36,4), [System.Globalization.NumberStyles]::HexNumber), [int]::Parse($SignatureData.Substring(40,4), [System.Globalization.NumberStyles]::HexNumber)

    return $SVN
}

function Get-SecureBootUEFI_DBXSVN {
    param (
        [Parameter(Mandatory)]
        [string]$DBXSVN
    )

    try {
        $SignatureData = (Get-SecureBootUEFI dbx | Get-UEFIDatabaseSignatures).SignatureList.SignatureData
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $null
        }
        else {
            throw $_.Exception.Message
        }
    }

    $LastSig = $SignatureData -match "^$DBXSVN" | sort | select -Last 1

    if ($LastSig.Count) {
        $SVN = Get-SignatureDataSVN $LastSig
    }
    else {
        $SVN = $null
    }

    return $SVN
}

function Get-WindowsUpdate_DBXSVN {
    $DBXUpdateSVN_File = "$env:SystemRoot\System32\SecureBootUpdates\DBXUpdateSVN.bin"

    try {
        $Signatures = Get-UEFIDatabaseSignatures -BytesIn ([IO.File]::ReadAllBytes($DBXUpdateSVN_File)) | where { $_.SignatureType -eq 'EFI_CERT_SHA256_GUID' }
    }
    catch {
        $_.Exception.Message
        exit 1
    }

    $SignatureData = $Signatures.SignatureList.SignatureData -match "^$EFI_BOOTMGR_DBXSVN_GUID"
    $Count = $SignatureData.Count

    if ($Count -eq 0) {
        return $null
    }

    return (Get-SignatureDataSVN $($SignatureData))
}

function Match-DBXSignatureData {
    <#
        .SYNOPSIS
        Parses EFI signatures from a DBX Update .bin file and compares the entire list against the current UEFI DBX.

        .DESCRIPTION
        From https://gist.github.com/out0xb2/f8e0bae94214889a89ac67fceb37f8c0#file-check-dbx-ps1
        Modified by github.com/cjee21
        Modified by github.com/garlin-cant-code

        .PARAMETER DBXUpdateFile
        Specifies a filename containing signed DBX Update signatures

        .OUTPUTS
        $true or $false
    #>

    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$DBXUpdateFile
    )

    if (-not (Test-Path $DBXUpdateFile)) {
        Write-Host "DBX update file `"$DBXUpdateFile`" not found." -Foreground Red
        exit 1
    }

    try {
        $RequiredSignatures = Get-UEFIDatabaseSignatures -BytesIn ([IO.File]::ReadAllBytes($DBXUpdateFile)) | where { $_.SignatureType -eq 'EFI_CERT_SHA256_GUID' }
    }
    catch {
        Write-Host "No EFI_CERT_SHA256 signatures in $DBXUpdateFile" -Foreground Red
        return $true
    }

    try {
        $DBXSignatureData = (Get-SecureBootUEFI dbx | Get-UEFIDatabaseSignatures).SignatureList.SignatureData
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $false
        }
        else {
            throw $_.Exception.Message
        }
    }

    $RequiredSignatureData = $RequiredSignatures.SignatureList.SignatureData
    $RequiredCount = $RequiredSignatureData.Count

    if ($RequiredCount -eq 0) {
        Write-Host "No DBX signatures in $DBXUpdateFile" -Foreground Red
        return $true
    }

    $Matched = 0

    foreach ($RequiredSig in $RequiredSignatureData) {
        if ($DBXSignatureData -contains $RequiredSig) {
            $Matched++
        }
        else {
            $RequiredSVN = Get-SignatureDataSVN $RequiredSig

            switch ($RequiredSig) {
                { $_ -match "^$EFI_BOOTMGR_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_BOOTMGR_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                }

                { $_ -match "^$EFI_CDBOOT_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_CDBOOT_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                }

                { $_ -match "^$EFI_WDSMGR_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_WDSMGR_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                }
            }
        }
    }

    if ($Matched -eq $RequiredCount) {
        return $true
    }
    else {
        return $false
    }
}

function Get-PFXCert {
    param (
        [Parameter(Mandatory)]
        [string]$FileName
    )

    try {
        $Issuer = (Get-PfxCertificate -LiteralPath $FileName).Issuer
    }
    catch {
        $_.Exception.Message
        exit 1
    }

    if ($Issuer -match $CN_Regex) {
        return $Matches[2]
    }
    else {
        return $Issuer
    }
}

function Validate-PFXCert {
    param (
        [Parameter(Mandatory)]
        [string]$CertName
    )

    if ($SecureBoot -eq $false) {
        return 'ALLOWED'
    }

    switch ($CertName) {
        { $_ -match '2011' } {
            if ($KEK_Certs -contains 'Microsoft Corporation KEK CA 2011' -and $db_Certs -contains $CertName -and $dbx_Certs -notcontains $CertName) {
               return 'ALLOWED'
            }
            else {
               return 'BANNED'
            }
        }

        { $_ -match '2023' } {
            if ($KEK_Certs -contains 'Microsoft Corporation KEK 2K CA 2023' -and $db_Certs -contains $CertName -and $dbx_Certs -notcontains $CertName) {
               return 'ALLOWED'
            }
            else {
               return 'BANNED'
            }
        }

        default { return 'UNKNOWN' }
    }
}

function Confirm-MinimumUBR {
    $Build = [int]$CurrentVersion.CurrentBuildNumber
    $UBR = [int]$CurrentVersion.UBR
    $Release = $CurrentVersion.DisplayVersion

    switch ($Build) {
        { $_ -in 19044,19045 } {
            if ($UBR -lt 6456) {
                return "Update W10 $Release to KB5066791 (Oct 2025) or later"
            }
        }

        22000 {
            if ($UBR -lt 3260) {
                return "Update W11 21H2 to KB5044280 (Oct 2025) or later"
            }
        }

        { $_ -in 22621,22631 } {
            if ($UBR -lt 6060) {
                return "Update W11 $Release to KB5066793 (Oct 2025) or later"
            }
        }

        # 24H2 (26100), 25H2 (26200) en 26H2 (26300) delen dezelfde servicing-branch/UBR
        { $_ -in 26100,26200,26300 } {
            if ($UBR -lt 6899) {
                return "Update W11 $Release to KB5066835 (Oct 2025) or later"
            }
        }

        # Nieuwere builds (26H1/28000 en later) bevatten de CA 2023-ondersteuning standaard
        { $_ -gt 26300 } {
            return $true
        }

        default {
            return ('Build {0}.{1} is unsupported' -f $Build, $UBR)
        }
    }

    return $true
}

function Audit-UEFI {
    $CheckList = $null
    $index = 1
    $script:UpdateFlags = $script:RevokeFlags = 0

    $Result = Confirm-MinimumUBR

    if ($Result -ne $true) {
        $CheckList += "{0,-3} {1}`n" -f ('{0}.' -f $index++), $Result
        $NotMinimumUBR = $true
    }

    if (-not $SetupMode -and -not (Confirm-SecureBootUEFI)) {
        $CheckList += "{0,-3} Secure Boot is DISABLED`n" -f ('{0}.' -f $index++)
    }

    if ($SetupMode) {
        $CheckList += "{0,-3} UEFI is in Setup Mode`n" -f ('{0}.' -f $index++)
    }

    if ($PK_Cert.Count -and -not $PK_Trusted) {
        $CheckList += "{0,-3} [{1}] is UNTRUSTED`n" -f ('{0}.' -f $index++), $PK_Cert
    }

    # https://support.microsoft.com/en-us/topic/secure-boot-certificate-updates-guidance-for-it-professionals-and-organizations-e2b43f9f-b424-42df-bc6a-8476db65ab2f#bkmk_troubleshooting

    if ('Microsoft Corporation KEK 2K CA 2023' -notin $KEK_Certs) {
        $CheckList += "{0,-3} [Microsoft Corporation KEK 2K CA 2023] is missing from UEFI KEK`n" -f ('{0}.' -f $index++)
        $script:UpdateFlags = $script:UpdateFlags -bor 0x4
    }

    if ('Windows UEFI CA 2023' -notin $db_Certs) {
        $CheckList += "{0,-3} [Windows UEFI CA 2023] is missing from UEFI DB`n" -f ('{0}.' -f $index++)
        $script:UpdateFlags = $script:UpdateFlags -bor 0x40
    }

    if ('Microsoft UEFI CA 2023' -notin $db_Certs) {
        $CheckList += "{0,-3} [Microsoft UEFI CA 2023] is missing from UEFI DB`n" -f ('{0}.' -f $index++)
        $script:UpdateFlags = $script:UpdateFlags -bor 0x1000 -bor 0x4000
    }

    if ('Microsoft Option ROM UEFI CA 2023' -notin $db_Certs) {
        $CheckList += "{0,-3} [Microsoft Option ROM UEFI CA 2023] is missing from UEFI DB`n" -f ('{0}.' -f $index++)
        $script:UpdateFlags = $script:UpdateFlags -bor 0x800 -bor 0x4000
    }

    if ('Microsoft Windows Production PCA 2011' -notin $dbx_Certs) {
        $CheckList += "{0,-3} [Production PCA 2011] is missing from UEFI DBX`n" -f ('{0}.' -f $index++)
        $script:RevokeFlags = $script:RevokeFlags -bor 0x80
    }

    if (($dbx_BytesCount -eq 0) -or -not (Match-DBXSignatureData "$env:SystemRoot\System32\SecureBootUpdates\dbxupdate.bin")) {
        $CheckList += "{0,-3} DBX Updates are missing from UEFI DBX`n" -f ('{0}.' -f $index++)
        $script:RevokeFlags = $script:RevokeFlags -bor 0x2
    }

    $UEFI_DBXSVN = Get-SecureBootUEFI_DBXSVN $EFI_BOOTMGR_DBXSVN_GUID

    if ($UEFI_DBXSVN -eq $null) {
        $CheckList += "{0,-3} Windows BootMgr SVN is missing from UEFI DBX`n" -f ('{0}.' -f $index++)
        $script:RevokeFlags = $script:RevokeFlags -bor 0x200
    }
    elseif ((Get-WindowsUpdate_DBXSVN) -gt $UEFI_DBXSVN) {
        $CheckList += "{0,-3} SecureBootUpdates SVN is higher than UEFI DBX`n" -f ('{0}.' -f $index++)
        $script:RevokeFlags = $script:RevokeFlags -bor 0x200
    }

    $BootMgr_File_Hash = (Get-FileHash -LiteralPath $BootMgr_File).Hash
    $BootMgrEX_File_Hash = (Get-FileHash $BootMgrEX_File).Hash

    if (($PFXCert -notmatch 'Windows UEFI CA 2023') -or ((Get-WindowsUpdate_DBXSVN) -gt $UEFI_DBXSVN) -and ($BootMgr_File_Hash -ne $BootMgrEX_File_Hash)) {
        $CheckList += "{0,-3} Windows Boot Manager [{1}] is wrong version`n" -f ('{0}.' -f $index++), ($PFXCert -replace 'Microsoft Windows ')
        $script:UpdateFlags = $script:UpdateFlags -bor 0x100
    }

    if ($VBS_Enabled) {
        if ((Test-Path -LiteralPath $EFI_SkuSiPolicy_File)) {
            $SkuSiPolicy_File_Hash = (Get-FileHash $SkuSiPolicy_File).Hash
            $EFI_SkuSiPolicy_File_Hash = (Get-FileHash -LiteralPath $EFI_SkuSiPolicy_File).Hash

            if ($EFI_SkuSiPolicy_File_Hash -ne $SkuSiPolicy_File_Hash) {
                $CheckList += "{0,-3} SkuSiPolicy.p7b (for VBS) is not updated`n" -f ('{0}.' -f $index++)
                $script:UpdateSkuSiPolicy = $true
            }
        }
        else {
            $CheckList += "{0,-3} SkuSiPolicy.p7b (for VBS) is missing`n" -f ('{0}.' -f $index++)
            $script:UpdateSkuSiPolicy = $true
        }
    }

    return $CheckList
}

function Check-BootManager {
    param (
        [Parameter(Mandatory)]
        [string]$WIM_File,

        [Parameter(Mandatory)]
        [int]$Index
    )

    $WIM_Image = '{0}:{1}' -f (Split-Path $WIM_File -Leaf), $Index

    if (((& dism /list-image /imagefile:$WIM_File /index:$Index) -match '\\bootmgfw_EX.efi').Count) {
        '{0}{1,-13} Boot Manager [Windows UEFI CA 2023] is PRESENT.' -f $Tab8, $WIM_Image
    }
    else {
        if ($SecureBoot -eq $false -or $dbx_Certs -notcontains 'Microsoft Windows Production PCA 2011') {
            '{0}{1,-13} Boot Manager [Production PCA 2011] {2} ALLOWED.' -f $Tab8, $WIM_Image, $Verb
        }
        else {
            '{0}{1,-13} Boot Manager [Production PCA 2011] {2} BANNED.' -f $Tab8, $WIM_Image, $Verb
        }
    }
}

function Get-ProductVersion {
    param (
        [Parameter(Mandatory)]
        [string]$File
    )

    $ProductVersion = (Get-Item -LiteralPath $File).VersionInfo.ProductVersion -replace '^10.0.'
    return $ProductVersion
}

function Check-BootMedia {
    $RemovableDrives = Get-Volume | where { $_.DriveType -in 'CD-ROM','Removable' -and $_.DriveLetter -ne $null } | sort DriveLetter

    if ($RemovableDrives.Count -eq 0) {
        return
    }

    Print-Header 'Bootable Media'
    foreach ($Volume in $RemovableDrives) {
        $DriveLetter = $Volume.DriveLetter

        $EFI_BootFile = "${DriveLetter}:\EFI\boot\boot${Arch}.efi"
        $Boot_WIM = "${DriveLetter}:\sources\boot.wim"
        $WIM_Formats = @('wim','esd','swm')

        if ($Volume.DriveType -eq 'Removable') {
            $DriveType = 'USB'
        }
        else {
            $DriveType = 'DVD'
        }

        $Label = $Volume.FileSystemLabel

        if ($Label -ne '') {
            '{0}{1} Drive {2}: "{3}"' -f $Tab4, $DriveType, $DriveLetter, $Label
        }
        else {
            '{0}{1} Drive {2}:' -f $Tab4, $DriveType, $DriveLetter
        }

        if (Test-Path $EFI_BootFile) {
            $PFXCert = Get-PFXCert $EFI_BootFile

            if ((Validate-PFXCert $PFXCert) -eq 'BANNED') {
                '{0}Boot File [{1}] {2} BANNED.' -f $Tab8, ($PFXCert -replace 'Microsoft Windows '), $Verb
            }
            else {
                $BootMgrEX_File_Hash = (Get-FileHash $BootMgrEX_File).Hash
                $EFI_BootFile_Hash = (Get-FileHash -LiteralPath $EFI_BootFile).Hash

                if ($UEFI_DBXSVN -and ($EFI_BootFile_Hash -ne $BootMgrEX_File_Hash)) {
                    '{0}Boot File [{1}] {2} BANNED.' -f $Tab8, ($PFXCert -replace 'Microsoft Windows '), $Verb
                }
                else {
                    '{0}Boot File [{1}] {2} ALLOWED.' -f $Tab8, ($PFXCert -replace 'Microsoft Windows '), $Verb
                }
            }

            if ($Verbose) {
                "{0}boot${Arch}.efi File version: {1}`n" -f $Tab12, (Get-ProductVersion $EFI_BootFile)
            }
        }

        if (Test-Path $Boot_WIM) {
            try {
                $Index = (Get-WindowsImage -ImagePath $Boot_WIM -Name *Setup*).ImageIndex

                if ($Index -eq $null) {
                    $Index = (Get-WindowsImage -ImagePath $Boot_WIM).Count
                }

                Check-BootManager -WIM_File $Boot_WIM -Index $Index
            }
            catch {
                $ErrorMessage = $_.Exception.Message

                if ($ErrorMessage -ne 'There is no matching image.') {
                    $ErrorMessage
                }
            }
        }

        foreach ($Format in $WIM_Formats) {
            $ImageFile = "${DriveLetter}:\sources\install.$Format"

            if (Test-Path $ImageFile) {
                $ImageCount = (Get-WindowsImage -ImagePath $ImageFile).Count

                if ($NoSkip) {
                    $Count = $ImageCount
                }
                else {
                    $Count = 1
                }

                try {
                    for ($i = 1; $i -le $Count; $i++) {
                        Check-BootManager -WIM_File $ImageFile -Index $i
                    }
                }
                catch {
                    $ErrorMessage = $_.Exception.Message

                    if ($ErrorMessage -ne 'There is no matching image.') {
                        $ErrorMessage
                    }
                }

                if ($ImageCount -gt 1 -and -not $NoSkip) {
                    '{0}Skipping checks on next {1} install.{2} images.' -f $Tab12, --$ImageCount, $Format
                }
            }
        }
    }
}

$ScriptBlock = {
    switch ($env:PROCESSOR_ARCHITECTURE) {
        'amd64' { $Arch = 'x64' }
        'x86'   { $Arch = 'x86' }
        'arm64' { $Arch = 'aa64' }
        'arm'   { $Arch = 'aa32' }
    }

    $System = Get-CimInstance -ClassName Win32_ComputerSystem
    $BIOS = Get-CimInstance -ClassName Win32_BIOS
    $CurrentVersion = Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion'

    $SystemDrive = (Get-CimInstance Win32_OperatingSystem).SystemDrive

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'

    # Force refresh of reg key 'WindowsUEFICA2023Capable'
    Start-ScheduledTask -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update"

    if ($Verbose) {
        $CurrentBuild = $CurrentVersion.CurrentBuildNumber
        "Windows {0} {1} ({2}.{3})`n" -f $(if ($CurrentBuild -lt 22000) { '10' } else { '11' }), $CurrentVersion.DisplayVersion, $CurrentBuild, $CurrentVersion.UBR
    }

    $SecureBoot = Confirm-SecureBootUEFI
    $Verb = 'is'

    switch ($SecureBoot) {
        $true { 'Secure Boot: ON' }

        $false {
            if ($Audit) {
                'Secure Boot: OFF (Audit Report runs as ON)'
                $SecureBoot = $true
                $Verb = 'will be'
            }
            else {
                'Secure Boot: OFF'
            }
        }

        default {
            "ERROR: This PC doesn't support Secure Boot.`n"
            exit 1
        }
    }

    $VBS_Status = (Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard).VirtualizationBasedSecurityStatus

    if ($VBS_Status -gt 0) {
        'Virtualization Based Security: ON'
        $VBS_Enabled = $true
    }
    else {
        'Virtualization Based Security: OFF'
    }

    try {
        $ProtectionStatus = ([string](Get-BitLockerVolume -MountPoint $SystemDrive).ProtectionStatus).ToUpper()
        $ManageBDECount = (Get-CimInstance -Namespace 'ROOT/CIMV2/Security/MicrosoftVolumeEncryption' -Class Win32_EncryptableVolume -Filter "DriveLetter=`"$SystemDrive`"" | Invoke-CimMethod -MethodName 'GetSuspendCount').SuspendCount

        if ($ManageBDECount) {
            "`nBitLocker on ({0}) {1}`n{2}SUSPENDED for {3} reboot{4}." -f $SystemDrive, $ProtectionStatus, $Tab4, $ManageBDECount, $(if ($ManageBDECount -gt 1) { 's' })
        }
        else {
            'BitLocker on ({0}) {1}' -f $SystemDrive, $ProtectionStatus

            if ($ProtectionStatus -eq 'On') {
                $BitLocker_Enabled = $true
            }
       }

    }
    catch {
        $_.Exception.Message
    }

    if ($Verbose) {
        $Model = '{0} {1}' -f ($System.Manufacturer -split ',')[0], $System.Model
        $BIOS_Version = $BIOS.SMBIOSBIOSVersion
        $BIOS_Date = $BIOS.ReleaseDate.ToString('yyyy-MM-dd')

        Print-Header 'BIOS Firmware'
        '{0}{1}' -f $Tab4, $Model
        '{0}Version: {1}' -f $Tab4, $BIOS_Version
        '{0}Date: {1}' -f $Tab4, $BIOS_Date
    }

    foreach ($Variable in 'PK','KEK','db','dbx') {
        try {
            $Count = (Get-SecureBootUEFI $Variable).Bytes.Count
        }
        catch {
            if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
                $Count = 0
            }
        }

        New-Variable -Name "${Variable}_BytesCount" -Value $Count
    }

    if ((((Get-SecureBootUEFI SetupMode).Bytes -Join '') -eq 1) -or ($PK_BytesCount -eq 0 -and $KEK_BytesCount -eq 0 -and $db_BytesCount -eq 0 -and $dbx_BytesCount -eq 0)) {
        if (-not $Verbose) {
            "`nUEFI is in Setup Mode (NO CERTS)"
        }

        $SetupMode = $true
    }

    try {
        $PKDefault_Cert = Get-UEFICert PKDefault
        $KEKDefault_Certs = Get-UEFICert KEKDefault
        $dbDefault_Certs = Get-UEFICert dbDefault
        $dbxDefault_Certs = Get-UEFICert dbxDefault

        $PK_Cert = Get-UEFICert PK
        $KEK_Certs = Get-UEFICert KEK
        $db_Certs = Get-UEFICert db
        $dbx_Certs = Get-UEFICert dbx
    }
    catch {
        Write-Host 'ERROR: Failed to read UEFI Secure Boot settings.' -Foreground Red
        $_.Exception.Message
        exit 1
    }

    $PK_Trusted = Check-TrustedPK

    if ($Verbose) {
        Print-UEFICerts -Name 'Default PK' -CertArray ([ref]$PKDefault_Cert)
    }

    if ((-not $SetupMode -and -not $PK_Trusted) -or $Verbose) {
        Print-UEFICerts -Name 'PK' -CertArray ([ref]$PK_Cert)

        if ($PK_Cert -ne $null -and -not $PK_Trusted) {
            '{0}Platform Key is UNTRUSTED.' -f $Tab8
        }
    }

    $KEKUpdate = Check-KEKUpdateMap

    if ($Verbose -and 'Microsoft Corporation KEK 2K CA 2023' -notin $KEK_Certs) {
        switch ($KEKUpdate) {
            { $_ -match '\.bin' } {
                if ($Verbose) {
                    $Vendor = ($_ -split '/')[0]

                    if ($Vendor -ne 'Microsoft') {
                        '{0}[KEK CA 2023] Update is available from {1} or Microsoft.' -f $Tab8, $Vendor
                    }
                    else {
                        '{0}[KEK CA 2023] Update is available from Microsoft.' -f $Tab8
                    }
                }
            }

            $null {
                if (-not $SetupMode) {
                    '{0}Manual update of [KEK CA 2023] is REQUIRED.' -f $Tab8
                }
            }

            default {
                "`nERROR: Unable to parse Microsoft's KEK update map."
                Write-Host $_ -Foreground Red
            }
        }
    }

    if ($Verbose) {
        Print-UEFICerts -Name 'Default KEK' -CertArray ([ref]$KEKDefault_Certs)
    }

    if (-not $SetupMode -or $Verbose) {
        Print-UEFICerts -Name 'KEK' -CertArray ([ref]$KEK_Certs)
    }

    if ($Verbose) {
        Print-UEFICerts -Name 'Default DB' -CertArray ([ref]$dbDefault_Certs)
    }

    if (-not $SetupMode -or $Verbose) {
        Print-UEFICerts -Name 'DB' -CertArray ([ref]$db_Certs)
    }

    if ($Verbose) {
        Print-UEFICerts -Name 'Default DBX' -CertArray ([ref]$dbxDefault_Certs)
        '{0}EFI_CERT_SHA256_GUID Signatures: {1}' -f $Tab4, (Get-SecureBootUEFI -Name dbxDefault | Get-UEFIDatabaseSignatures | where { $_.SignatureType -eq 'EFI_CERT_SHA256_GUID' }).SignatureList.Count
    }

    if (-not $SetupMode -or $Verbose) {
        Print-UEFICerts -Name 'DBX' -CertArray ([ref]$dbx_Certs)
    }

    $DBX_BootMgrSVN = Get-SecureBootUEFI_DBXSVN $EFI_BOOTMGR_DBXSVN_GUID

    if ($DBX_BootMgrSVN -ne $null) {
        '{0}Windows BootMgr SVN {1}' -f $Tab4, $DBX_BootMgrSVN
    }

    if ($Verbose) {
        if ($DBX_BootMgrSVN -eq $null) {
            '{0}Windows BootMgr SVN is MISSING.' -f $Tab4
        }

        if ($dbx_BytesCount -ne 0) {
            '{0}EFI_CERT_SHA256_GUID Signatures: {1}' -f $Tab4, (Get-SecureBootUEFI -Name dbx | Get-UEFIDatabaseSignatures | where { $_.SignatureType -eq 'EFI_CERT_SHA256_GUID' }).SignatureList.Count
        }
        else {
            '{0}EFI_CERT_SHA256_GUID Signatures: 0' -f $Tab4
        }
    }

    $SystemDisk = (Get-Disk | Where-Object {$_.IsSystem -eq $true}).Number
    $GUID = (Get-Partition -DiskNumber $SystemDisk | Where-Object { $_.Type -eq 'System' }).Guid

    $EFI_Path = "\\?\Volume$GUID\EFI"

    $BootMgrEX_File = "$env:SystemRoot\Boot\EFI_EX\bootmgfw_EX.efi"
    $SkuSiPolicy_File = "$env:SystemRoot\System32\SecureBootUpdates\SkuSiPolicy.p7b"

    $BootMgr_File = "$EFI_Path\Microsoft\Boot\bootmgfw.efi"
    $EFI_SkuSiPolicy_File = "$EFI_Path\Microsoft\Boot\SkuSiPolicy.p7b"

    $PFXCert = Get-PFXCert $BootMgr_File

    Print-Header 'EFI Files'

    if ((Validate-PFXCert $PFXCert) -eq 'BANNED') {
        '{0}Disk {1}: Windows Boot Manager [{2}] {3} BANNED.' -f $Tab4, $SystemDisk, ($PFXCert -replace 'Microsoft Windows '), $Verb
    }
    else {
        $BootMgrEX_File_Hash = (Get-FileHash $BootMgrEX_File).Hash
        $BootMgr_File_Hash = (Get-FileHash -LiteralPath $BootMgr_File).Hash

        if ($UEFI_DBXSVN -and ($BootMgr_File_Hash -ne $BootMgrEX_File_Hash)) {
            '{0}Disk {1}: Windows Boot Manager [{2}] {3} BANNED.' -f $Tab4, $SystemDisk, ($PFXCert -replace 'Microsoft Windows '), $Verb
        }
        else {
            '{0}Disk {1}: Windows Boot Manager [{2}] {3} ALLOWED.' -f $Tab4, $SystemDisk, ($PFXCert -replace 'Microsoft Windows '), $Verb
        }
    }

    if ($Verbose) {
        "{0}bootmgfw.efi File version: {1}" -f $Tab8, (Get-ProductVersion $BootMgr_File)
    }

    $WindowsUEFICA2023Capable = Get-ItemPropertyValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing' -Name WindowsUEFICA2023Capable -ErrorAction SilentlyContinue

    if ($WindowsUEFICA2023Capable -ne $null) {
        "`n{0}Registry: WindowsUEFICA2023Capable = {1}" -f $Tab4, $WindowsUEFICA2023Capable

        switch ($WindowsUEFICA2023Capable) {
            0  { '{0}[Windows UEFI CA 2023] not in UEFI DB.' -f $Tab8 }
            1  { '{0}[Windows UEFI CA 2023] in UEFI DB.' -f $Tab8 }
            2  { '{0}[Windows UEFI CA 2023] in UEFI DB, and Windows starting from CA 2023 Boot Manager.' -f $Tab8 }
            default { '{0}Unknown status.' -f $Tab8 }
        }
    }

    if ($VBS_Enabled) {
        if ((Test-Path -LiteralPath $EFI_SkuSiPolicy_File)) {
            $SkuSiPolicy_File_Hash = (Get-FileHash $SkuSiPolicy_File).Hash
            $EFI_SkuSiPolicy_File_Hash = (Get-FileHash -LiteralPath $EFI_SkuSiPolicy_File).Hash

            if ($EFI_SkuSiPolicy_File_Hash -eq $SkuSiPolicy_File_Hash) {
                "`n{0}Disk {1}: SkuSiPolicy.p7b (for VBS) is CURRENT." -f $Tab4, $SystemDisk
            }
            else {
                "`n{0}Disk {1}: SkuSiPolicy.p7b (for VBS) is WRONG VERSION." -f $Tab4, $SystemDisk
            }
        }
        else {
            "`n{0}Disk {1}: SkuSiPolicy.p7b (for VBS) is NOT PRESENT." -f $Tab4, $SystemDisk
        }
    }

    if ($BootMedia) {
        Check-BootMedia
    }

    $CheckList = Audit-UEFI

    if ($Audit) {
        Print-Header -Bold "`nAUDIT REPORT"
    
        if ($CheckList -ne $null) {
            $CheckList.TrimEnd("`n")
        }
        else {
            Write-Output ''
        }
    }

    switch ($UpdateFlags) {
        0x100 {
            $UpdateMessage = 'To install Windows Boot Manager [UEFI CA 2023]'
        }
        default {
            if ($RevokeFlags) {
                $UpdateMessage = 'To install [UEFI CA 2023] certs WITHOUT REVOKING the [PCA 2011] cert'
            }
            else {
                $UpdateMessage = 'To install [UEFI CA 2023] certs'
            }
        }
    }

    if ($RevokeFlags) {
        if ($UpdateFlags) {
            $RevokeMessage = 'To install [UEFI CA 2023] certs and REVOKE the [PCA 2011] cert'
        }
        else {
            $RevokeMessage = 'To revoke the [PCA 2011] cert, run the commands'
        }
    }

    if ($UpdateFlags -or $RevokeFlags -or $UpdateSkuSiPolicy) {
        if ($BitLocker_Enabled -and $UpdateFlags -ne 0x100) {
            $DeviceGuard_Running = (Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard).SecurityServicesRunning

            if ($DeviceGuard_Running -eq 1) {
                $ManageBDE = "manage-bde -Protectors -Disable $SystemDrive -RebootCount 3"
            }
            else {
                $ManageBDE = "manage-bde -Protectors -Disable $SystemDrive -RebootCount 1"
            }
        }

        Print-Header -Bold "`nREQUIRED ACTION"

        if (('Microsoft Corporation KEK 2K CA 2023' -notin $KEK_Certs) -and ('Windows UEFI CA 2023' -in $db_Certs)) {
            "`nRun the command:`n{0}Update_UEFI-CA2023.ps1{1}`n" -f $Tab4, $(if ($RevokeFlags) { ' -Revoke' })

            if (-not $PK_Trusted) {
                "Finish the UEFI steps to manually add the Platform Key (PK) cert, if the script provided instructions.`n"
            }

            "Finish the UEFI steps to manually add the [KEK CA 2023] cert, if the script provided instructions.`n"

            break
        }

        if ($PK_Trusted -and (('Microsoft Corporation KEK 2K CA 2023' -in $KEK_Certs) -or $SignedKEK)) {
            $MergedFlags = $UpdateFlags -bor $RevokeFlags

            if ($UpdateFlags -and $RevokeFlags) {
                "`nOPTION 1:  DO NOTHING.  Windows will apply the UEFI updates in 2026 (supported BIOS)."

                "`nOPTION 2:  {0}, run the commands:`n" -f $UpdateMessage

                if ($ManageBDE -ne $null) { '{0}{1}' -f $Tab4, $ManageBDE }

                '{0}reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Secureboot /v AvailableUpdates /t REG_DWORD /d 0x{1:x} /f' -f $Tab4, $UpdateFlags
                '{0}powershell Start-ScheduledTask -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update"' -f $Tab4

                "`nOPTION 3:  {0}, run the commands:`n" -f $RevokeMessage

                if ($ManageBDE -ne $null) { '{0}{1}' -f $Tab4, $ManageBDE }

                if ($UpdateFlags -eq 0x100) {
                    '{0}reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Secureboot /v AvailableUpdates /t REG_DWORD /d 0x{1:x} /f' -f $Tab4, $RevokeFlags
                }
                else {
                    '{0}reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Secureboot /v AvailableUpdates /t REG_DWORD /d 0x{1:x} /f' -f $Tab4, $MergedFlags
                }

                '{0}powershell Start-ScheduledTask -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update"' -f $Tab4
            }
            elseif ($UpdateFlags) {
                "`n{0}, run the commands:`n" -f $UpdateMessage

                if ($ManageBDE -ne $null) { '{0}{1}' -f $Tab4, $ManageBDE }

                '{0}reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Secureboot /v AvailableUpdates /t REG_DWORD /d 0x{1:x} /f' -f $Tab4, $UpdateFlags
                '{0}powershell Start-ScheduledTask -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update"' -f $Tab4
            }
            elseif ($RevokeFlags) {
                "`n{0}, run the commands:`n" -f $RevokeMessage

                if ($ManageBDE -ne $null) { '{0}{1}' -f $Tab4, $ManageBDE }

                '{0}reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Secureboot /v AvailableUpdates /t REG_DWORD /d 0x{1:x} /f' -f $Tab4, $RevokeFlags
                '{0}powershell Start-ScheduledTask -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update"' -f $Tab4
            }

            if ($UpdateSkuSiPolicy) {
                "`nTo install SkuSiPolicy.p7b, run the command:"
                '{0}Update_UEFI-CA2023.ps1 -SkuSiPolicy' -f $Tab4
            }

        }
        else {
            if (-not $SetupMode) {
                "`nMANUAL UPDATE of the BIOS is required.`n"

                "Enter the BIOS menu, and search for User or Custom Mode option of updating the UEFI PK or KEK keys."
                "If your BIOS doesn't support this feature, select Setup Mode to clear all certs."
            }

            "`nOPTION 1:  To install [UEFI CA 2023] certs WITHOUT REVOKING the [PCA 2011] cert, run the command:`n"
            '{0}Update_UEFI-CA2023.ps1' -f $Tab4

            "`n`nOPTION 2:  To install [UEFI CA 2023] certs and REVOKE the [PCA 2011] cert, run the command:`n"
            '{0}Update_UEFI-CA2023.ps1 -Revoke' -f $Tab4
        }
    }
    else {
        Print-Header 'STATUS REPORT'
        $UEFICA2023Status = Get-ItemPropertyValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing' -Name UEFICA2023Status -ErrorAction SilentlyContinue

        if ($UEFICA2023Status -ne $null) {
            "{0}Registry: UEFICA2023Status = {1}`n" -f $Tab4, $UEFICA2023Status
        }

        'SUCCESS: NO UPDATES ARE REQUIRED.'
    }
}

if ($Log) {
    $System = Get-CimInstance -ClassName Win32_ComputerSystem
    $LogFile = '{0}\{1} {2} Check-UEFI.log' -f $PSScriptRoot, (Get-Date -Format 'yyyy-MM-dd'), $System.Model.ToUpper()

    & $ScriptBlock | Tee-Object $LogFile
    "`nLog file saved as `"{0}`"`n" -f $LogFile
}
else {
    & $ScriptBlock
    Write-Output ''
}
'@

$script:Embedded['Update_UEFI-CA2023.ps1'] = @'
<#PSScriptInfo

.VERSION 2026.01.18

.GUID 7c7848ed-3952-4726-8f23-8644881c2c91

.AUTHOR garlin

.COPYRIGHT

.TAGS UEFI, Secure Boot, PK, KEK, DB, DBX, SVN, Windows Boot Manager

.RELEASENOTES

#>

<#
.SYNOPSIS
    Script to install Secure Boot CA 2023 certificates in UEFI, and revoke PCA 2011 certificate if needed.

.DESCRIPTION
    Run this script to bring Windows into compliance with UEFI CA 2023 changes.

.PARAMETER Version
    Print the script's version number and exit.

.PARAMETER UpdatesFolder
    Provide a different source folder for the Post-Signed object (.bin) files.

.PARAMETER Audit
    Perform an audit report of the UEFI variables and Windows Boot Manager version.  Identify any missing UEFI certs, and validate if current boot files are
    allowed by enabling Secure Boot mode.

    Identify all required actions to bring system into compliance for upcoming CA 2023 changes.

    If Secure Boot is currently disabled, audit report will simulate conditions where Secure Boot is enabled

.PARAMETER Revoke
    Revoke [Microsoft Windows Production PCA 2011] certificate by adding the cert to the UEFI DBX.
    To allow dual-booting of [Production PCA 2011] & [UEFI CA 2023] media, do not use the -Revoke option.

.PARAMETER Latest
    Download latest version of DBXUpdate.bin and DBXUpdateSVN.bin, from Microsoft's Secure Boot Objects GitHub before proceeding.

.PARAMETER SkuSiPolicy
    Deploy \Windows\System32\SecureBootUpdates\SkuSiPolicy.p7b to EFI partition.

.PARAMETER SBAT
    Apply optional Secure Boot Advanced Targeting (SBAT) update, when sharing UEFI with Linux OS'es.

.PARAMETER BootMedia
    Check boot files on all mounted removable media, and replace with [UEFI CA 2023] version if needed.

.PARAMETER Log
    Save script output to a file named "YYYY-MM-DD [Model] Update UEFI.log"

.EXAMPLE
    Update_UEFI-CA2023.ps1
.EXAMPLE
    Update_UEFI-CA2023.ps1 -Audit
.EXAMPLE
    Update_UEFI-CA2023.ps1 -Revoke
.EXAMPLE
    Update_UEFI-CA2023.ps1 -Revoke -Latest -Log
#>

[CmdletBinding(DefaultParameterSetName='Default')]
param (
    [Parameter(Mandatory=$false,ParameterSetName='Version')]
    [switch]$Version,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [ValidateScript({ if (Test-Path $_ ) { $true } else { throw "Folder `"$_`" not found." } })]
    [string]$UpdatesFolder = "$env:SystemRoot\System32\SecureBootUpdates",

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$Audit,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$Latest,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$Revoke,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$SkuSiPolicy,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$SBAT,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$BootMedia,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$Log,

    [Parameter(Mandatory=$false,ParameterSetName='Default',DontShow,ValueFromRemainingArguments=$true)]
    [string[]]$ignored
)

$ScriptVersion = '2026.01.18'

# https://github.com/microsoft/secureboot_objects/blob/main/Archived/dbx_info_msft_4_09_24_svns.csv
$EFI_BOOTMGR_DBXSVN_GUID = '01612B139DD5598843AB1C185C3CB2EB92'
$EFI_CDBOOT_DBXSVN_GUID =  '019D2EF8E827E15841A4884C18ABE2F284'
$EFI_WDSMGR_DBXSVN_GUID =  '01C2CA99C9FE7F6F4981279E2A8A535976'

$VMWARE_GUID = 'a3d5e95b-0a8f-4753-8735-445afb708f62'

$CN_Regex = '(CN=)([^,]+)'

$Arch = $env:PROCESSOR_ARCHITECTURE.ToLower()

switch ($Arch) {
    'amd64' { $EDK2_Arch = 'x64' }
    'x86'   { $EDK2_Arch = 'ia32' }
    'arm64' { $EDK2_Arch = 'aarch64' }
    'arm'   { $EDK2_Arch = 'arm' }
}

$EDK2bin_URL = "https://github.com/microsoft/secureboot_objects/releases/download/v1.6.2/edk2-${EDK2_Arch}-secureboot-binaries.zip"
$PK_DER_URL = 'https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PreSignedObjects/PK/Certificate/WindowsOEMDevicesPK.der'

$KEKUpdateMap_URL = 'https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PostSignedObjects/KEK/kek_update_map.json'
$KEK_DER_URL = 'https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PreSignedObjects/KEK/Certificates/microsoft%20corporation%20kek%202k%20ca%202023.der'

$DBXUpdate_bin_URL = "https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PostSignedObjects/DBX/$Arch/DBXUpdate.bin"
$DBXUpdateSVN_bin_URL = "https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PostSignedObjects/Optional/DBX/$Arch/DBXUpdateSVN.bin"

$Tab4 = ' ' * 4

if ($Version) {
    '{0} version ({1}){2}' -f $MyInvocation.MyCommand.Name, $ScriptVersion, $(if ($MyInvocation.Line -ne '') { "`n" })
    exit 0
}

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if ($PSVersionTable.PSVersion.Major -gt 5) {
        $PS = 'pwsh'
    }
    else {
        $PS = 'powershell'
    }

    $args = ($MyInvocation.BoundParameters.Keys.GetEnumerator() | where { $_ -notmatch 'UpdatesFolder|DBXupdate|ignored' } | foreach { '-{0}' -f $_ }) -join ' '

    if ($MyInvocation.BoundParameters.'UpdatesFolder' -ne $null) {
        $args += ' -UpdatesFolder "{0}"' -f (Get-Item $MyInvocation.BoundParameters.'UpdatesFolder' -ErrorAction SilentlyContinue).FullName
    }

    Start-Process $PS -ArgumentList "-nop -ep bypass -NoLogo -NoExit -f $($MyInvocation.MyCommand.Path) $args" -Verb RunAs
    exit 0
}

function Confirm-MinimumUBR {
    $Build = [int]$CurrentVersion.CurrentBuildNumber
    $UBR = [int]$CurrentVersion.UBR
    $Release = $CurrentVersion.DisplayVersion

    switch ($Build) {
        { $_ -in 19044,19045 } {
            if ($UBR -lt 6456) {
                return "Update W10 $Release to KB5066791 (Oct 2025) or later"
            }
        }

        22000 {
            if ($UBR -lt 3260) {
                return "Update W11 21H2 to KB5044280 (Oct 2025) or later"
            }
        }

        { $_ -in 22621,22631 } {
            if ($UBR -lt 6060) {
                return "Update W11 $Release to KB5066793 (Oct 2025) or later"
            }
        }

        # 24H2 (26100), 25H2 (26200) en 26H2 (26300) delen dezelfde servicing-branch/UBR
        { $_ -in 26100,26200,26300 } {
            if ($UBR -lt 6899) {
                return "Update W11 $Release to KB5066835 (Oct 2025) or later"
            }
        }

        # Nieuwere builds (26H1/28000 en later) bevatten de CA 2023-ondersteuning standaard
        { $_ -gt 26300 } {
            return $true
        }

        default {
            return ('Build {0}.{1} is unsupported' -f $Build, $UBR)
        }
    }

    return $true
}

function Get-UefiDatabaseSignatures {
    <#
        .SYNOPSIS
        Parses UEFI Signature Databases into logical Powershell objects
        # https://github.com/cjee21/Check-UEFISecureBootVariables

        .DESCRIPTION
        Original Author: Matthew Graeber (@mattifestation)
        Modified By: Jeremiah Cox (@int0x6)
        Modified By: Joel Roth (@nafai)
        Modified By: garlin (@garlin-cant-code)
        Additional Source: https://gist.github.com/mattifestation/991a0bea355ec1dc19402cef1b0e3b6f
        Additional Source: https://www.powershellgallery.com/packages/SplitDbxContent/1.0
        License: BSD 3-Clause

        .PARAMETER Variable
        Specifies an UEFI variable, an instance of which is returned by calling the Get-SecureBootUEFI cmdlet.

        .PARAMETER BytesIn
        Specifies a byte array consisting of the PKDefault, KEKDefault, dbDefault, dbxDefault, PK, KEK, db, or dbx UEFI variable contents.

        .EXAMPLE
        $DbxBytes = [IO.File]::ReadAllBytes('.\dbx.bin')
        Get-UEFIDatabaseSignatures -BytesIn $DbxBytes

        .EXAMPLE
        Get-UEFIDatabaseSignatures -Filename ".\DBXUpdate-20230314.x64.bin"

        .EXAMPLE
        Get-SecureBootUEFI -Name db | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name dbx | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name pk | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name kek | Get-UEFIDatabaseSignatures

        .INPUTS
        Microsoft.SecureBoot.Commands.UEFIEnvironmentVariable
        Accepts the output of Get-SecureBootUEFI over the pipeline.

        .OUTPUTS
        UefiSignatureDatabase
        Outputs an array of custom powershell objects describing a UEFI Signature Database. "77fa9abd-0359-4d32-bd60-28f4e78f784b" refers to Microsoft as the owner.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'UEFIVariable')]
        [ValidateScript({ ($_.GetType().Fullname -eq 'Microsoft.SecureBoot.Commands.UEFIEnvironmentVariable') -and ($_.Name -in 'PKDefault','KEKDefault','dbDefault','dbxDefault','pk','kek','db','dbx') })]
        $Variable,

        [Parameter(Mandatory, ParameterSetName = 'ByteArray')]
        [Byte[]]
        [ValidateNotNullOrEmpty()]
        $BytesIn,

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]
        [ValidateScript({ (Resolve-Path "$_").where({Test-Path $_}).Path })]
        $Filename
    )

    $PSVersion = $PSVersionTable.PSVersion.Major

    $SignatureTypeMapping = @{
        'C1C41626-504C-4092-ACA9-41F936934328' = 'EFI_CERT_SHA256_GUID' # Most often used for dbx
        'A5C059A1-94E4-4AA7-87B5-AB155C2BF072' = 'EFI_CERT_X509_GUID'   # Most often used for db
    }

    $Bytes = $null

    if ($Filename)
    {
        if ($PSVersion -gt 5) {
            $Bytes = Get-Content -AsByteStream $Filename -ErrorAction Stop
        }
        else {
            $Bytes = Get-Content -Encoding Byte $Filename  -ErrorAction Stop
        }
    }
    elseif ($Variable)
    {
        $Bytes = $Variable.Bytes
    }
    else
    {
        $Bytes = $BytesIn
    }

    # Modified from Split-Dbx
    if (($Bytes[40] -eq 0x30) -and ($Bytes[41] -eq 0x82 ))
    {
        Write-Debug "Removing signature."

        # Signature is known to be ASN size plus header of 4 bytes
        $sig_length = $Bytes[42] * 256 + $Bytes[43] + 4
        if ($sig_length -gt ($Bytes.Length + 40)) {
            Write-Error "Signature longer than file size!" -ErrorAction Stop
        }

        ## Unsigned db store
        [System.Byte[]]$Bytes = @($Bytes[($sig_length+40)..($Bytes.Length - 1)].Clone())
    }
    else
    {
        Write-Debug "Signature not found. Assuming it's already split."
    }

    try
    {
        $MemoryStream = New-Object -TypeName IO.MemoryStream -ArgumentList @(,$Bytes)
        $BinaryReader = New-Object -TypeName IO.BinaryReader -ArgumentList $MemoryStream, ([Text.Encoding]::Unicode)
    }
    catch
    {
        throw $_
        return
    }

    # What follows will be an array of EFI_SIGNATURE_LIST structs

    while ($BinaryReader.PeekChar() -ne -1) {
        $SignatureType = $SignatureTypeMapping[([Guid][Byte[]] $BinaryReader.ReadBytes(16)).Guid]
        $SignatureListSize = $BinaryReader.ReadUInt32()
        $SignatureHeaderSize = $BinaryReader.ReadUInt32()
        $SignatureSize = $BinaryReader.ReadUInt32()

        $SignatureHeader = $BinaryReader.ReadBytes($SignatureHeaderSize)

        # 0x1C is the size of the EFI_SIGNATURE_LIST header
        $SignatureCount = ($SignatureListSize - 0x1C) / $SignatureSize

        $SignatureList = 1..$SignatureCount | ForEach-Object {
            $SignatureDataBytes = $BinaryReader.ReadBytes($SignatureSize)

            $SignatureOwner = [Guid][Byte[]] $SignatureDataBytes[0..15]

            switch ($SignatureType) {
                'EFI_CERT_SHA256_GUID' {
                    $SignatureData = ([Byte[]] $SignatureDataBytes[0x10..0x2F] | ForEach-Object { $_.ToString('X2') }) -join ''
                }

                'EFI_CERT_X509_GUID' {
                    try {
                        $SignatureData = New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,([Byte[]] $SignatureDataBytes[16..($SignatureDataBytes.Count - 1)]))
                    }
                    catch {
                        Write-Host "Skipping an invalid $Variable X509 certificate."
                    }
                }
            }

            [PSCustomObject] @{
                PSTypeName = 'EFI.SignatureData'
                SignatureOwner = $SignatureOwner
                SignatureData = $SignatureData
            }
        }

        [PSCustomObject] @{
            PSTypeName = 'EFI.SignatureList'
            SignatureType = $SignatureType
            SignatureList = $SignatureList
        }
    }
}

function Get-UEFICert {
    param (
        [Parameter(Mandatory)]
        [ValidateSet('PKDefault','KEKDefault','dbDefault','dbxDefault','PK','KEK','db','dbx')]
        [string]$Variable
    )

    try {
        $SignatureList = (Get-SecureBootUEFI $Variable | Get-UefiDatabaseSignatures).SignatureList
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return @()
        }
        else {
            throw $_.Exception.Message
        }
    }

    $Subject = $SignatureList.SignatureData.Subject

    if ($Verbose) {
        $Certs = $Subject | where { $_ -match '\s' } | foreach { $null = $_ -match $CN_Regex; $Matches[2] }
    }
    else {
        $Certs = $Subject | where { $_ -match 'Microsoft|Mosby' } | foreach { $null = $_ -match $CN_Regex; $Matches[2] }
    }

    if ($Variable -match 'PK') {
        if ($SignatureList.SignatureData -eq $null -and $SignatureList.SignatureOwner.Guid -eq $VMWARE_GUID) {
            $Certs = @('VMware Default PK')
        }
        elseif ($Subject -match 'VirtualBox') {
            $Certs = @('VirtualBox UEFI PK')
        }
    }

    return $Certs
}

function Check-TrustedPK {
    try {
        $PKSignatureList = (Get-SecureBootUEFI PK | Get-UefiDatabaseSignatures).SignatureList
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $false
        }
        else {
            throw $_.Exception.Message
        }
    }

    if ($PKSignatureList.SignatureData.Subject -notmatch 'DO NOT |Example') {
        return $true
    }
    else {
        return $false
    }
}

function Get-SignatureDataSVN {
    param (
        [Parameter(Mandatory)]
        [string]$SignatureData
    )

    # https://github.com/microsoft/secureboot_objects/blob/main/scripts/utility_functions.py
    $SVN = '{0}.{1}' -f [int]::Parse($SignatureData.Substring(36,4), [System.Globalization.NumberStyles]::HexNumber), [int]::Parse($SignatureData.Substring(40,4), [System.Globalization.NumberStyles]::HexNumber)

    return $SVN
}

function Get-SecureBootUEFI_DBXSVN {
    param (
        [Parameter(Mandatory)]
        [string]$DBXSVN
    )

    try {
        $SignatureData = (Get-SecureBootUEFI dbx | Get-UEFIDatabaseSignatures).SignatureList.SignatureData
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $null
        }
        else {
            throw $_.Exception.Message
        }
    }

    $LastSig = $SignatureData -match "^$DBXSVN" | sort | select -Last 1

    if ($LastSig.Count) {
        $SVN = Get-SignatureDataSVN $LastSig
    }
    else {
        $SVN = $null
    }

    return $SVN
}

function Get-WindowsUpdate_DBXSVN {
    $DBXSVN_File = "$env:SystemRoot\System32\SecureBootUpdates\DBXUpdateSVN.bin"

    try {
        $Signatures = Get-UEFIDatabaseSignatures -BytesIn ([IO.File]::ReadAllBytes($DBXSVN_File)) | where { $_.SignatureType -eq 'EFI_CERT_SHA256_GUID' }
    }
    catch {
        $_.Exception.Message
        exit 1
    }

    $SignatureData = $Signatures.SignatureList.SignatureData -match "^$EFI_BOOTMGR_DBXSVN_GUID"
    $Count = $SignatureData.Count

    if ($Count -eq 0) {
        return $null
    }

    return (Get-SignatureDataSVN $($SignatureData))
}

function Audit-UEFI {
    $CheckList = $null
    $index = 1

    if ($SetupMode) {
        $CheckList += "{0,-3} UEFI is in Setup Mode`n" -f ('{0}.' -f $index++)
    }

    if ($PK_Cert.Count -and -not $PK_Trusted) {
        $CheckList += "{0,-3} [{1}] is UNTRUSTED`n" -f ('{0}.' -f $index++), $PK_Cert
    }

    if ('Microsoft Corporation KEK 2K CA 2023' -notin $KEK_Certs) {
        $CheckList += "{0,-3} [Microsoft Corporation KEK 2K CA 2023] missing from UEFI KEK`n" -f ('{0}.' -f $index++)
    }

    if ('Windows UEFI CA 2023' -notin $db_Certs) {
        $CheckList += "{0,-3} [Windows UEFI CA 2023] missing from UEFI DB (dbupdate2024.bin)`n" -f ('{0}.' -f $index++)
    }

    if ('Microsoft UEFI CA 2023' -notin $db_Certs) {
        $CheckList += "{0,-3} [Microsoft UEFI CA 2023] missing from UEFI DB (DBUpdate3P2023.bin)`n" -f ('{0}.' -f $index++)
    }

    if ('Microsoft Option ROM UEFI CA 2023' -notin $db_Certs) {
        $CheckList += "{0,-3} [Microsoft Option ROM UEFI CA 2023] missing from UEFI DB (DBUpdateOROM2023.bin)`n" -f ('{0}.' -f $index++)
    }

    if ('Microsoft Windows Production PCA 2011' -notin $dbx_Certs) {
        $CheckList += "{0,-3} [Production PCA 2011] missing from UEFI DBX (DBXUpdate2024.bin)`n" -f ('{0}.' -f $index++)
    }

    if (($dbx_BytesCount -eq 0) -or -not (Match-DBXSignatureData "$env:SystemRoot\System32\SecureBootUpdates\dbxupdate.bin")) {
        $CheckList += "{0,-3} DBX Updates are missing from UEFI DBX (dbxupdate.bin)`n" -f ('{0}.' -f $index++)
    }

    $UEFI_DBXSVN = Get-SecureBootUEFI_DBXSVN $EFI_BOOTMGR_DBXSVN_GUID

    if ($UEFI_DBXSVN -eq $null) {
        $CheckList += "{0,-3} Windows BootMgr SVN is missing from UEFI DBX (DBXUpdateSVN.bin)`n" -f ('{0}.' -f $index++)
    }
    elseif ((Get-WindowsUpdate_DBXSVN) -gt $UEFI_DBXSVN) {
        $CheckList += "{0,-3} SecureBootUpdates SVN is higher than UEFI DBX`n" -f ('{0}.' -f $index++)
    }

    $null = (Get-PfxCertificate -LiteralPath $BootMgr_File).Issuer -match $CN_Regex
    $PFXCert = $Matches[2]

    $BootMgr_File_Hash = (Get-FileHash -LiteralPath $BootMgr_File).Hash
    $BootMgrEX_File_Hash = (Get-FileHash $BootMgrEX_File).Hash

    if (($PFXCert -notmatch 'Windows UEFI CA 2023') -or ((Get-WindowsUpdate_DBXSVN) -gt $UEFI_DBXSVN) -and ($BootMgr_File_Hash -ne $BootMgrEX_File_Hash)) {
        $CheckList += "{0,-3} Windows Boot Manager [{1}] is wrong version`n" -f ('{0}.' -f $index++), ($PFXCert -replace 'Microsoft Windows ')
    }

    if ($VBS_Enabled) {
        if ((Test-Path -LiteralPath $EFI_SkuSiPolicy_File)) {
            $SkuSiPolicy_File_Hash = (Get-FileHash $SkuSiPolicy_File).Hash
            $EFI_SkuSiPolicy_File_Hash = (Get-FileHash -LiteralPath $EFI_SkuSiPolicy_File).Hash

            if ($EFI_SkuSiPolicy_File_Hash -ne $SkuSiPolicy_File_Hash) {
                $CheckList += "{0,-3} SkuSiPolicy.p7b (for VBS) is not updated`n" -f ('{0}.' -f $index++)
                $script:UpdateSkuSiPolicy = $true
            }
        }
        else {
            $CheckList += "{0,-3} SkuSiPolicy.p7b (for VBS) is missing`n" -f ('{0}.' -f $index++)
            $script:UpdateSkuSiPolicy = $true
        }
    }

    return $CheckList
}

function Download-EDK2bin {
    $ZIP_File = "$EDK2_Folder\edk2-secureboot-binaries.zip"

    if (-not (Test-Path $EDK2_Folder)) {
        $null = New-Item -Path $EDK2_Folder -Type Directory -Force
    }

    try {
        'Downloading "{0}" from GitHub.' -f ($EDK2bin_URL -split '/')[-1]
        Invoke-WebRequest -UseBasicParsing -Uri $EDK2bin_URL -OutFile $ZIP_File
    }
    catch {
        $_.Exception.Message
        exit 1
    }

    $DefaultBin_Files = @('Default3PDb.bin', 'DefaultDbx.bin', 'DefaultKek.bin', 'DefaultPk.bin')

    $objShell = New-Object -ComObject 'Shell.Application'
    $objFolder = $objShell.NameSpace($EDK2_Folder)

    foreach ($File in $DefaultBin_Files) {
        $objFolder.CopyHere("$ZIP_File\LegacyFirmwareDefaults\Firmware\$File", 0x14)
    }
}

function Suspend-Bitlocker {
    $ProtectionStatus = (Get-BitLockerVolume -MountPoint $SystemDrive).ProtectionStatus

    if ($ProtectionStatus -eq 'On') {
        $DeviceGuard_Running = (Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard).SecurityServicesRunning

        if ($DeviceGuard_Running -eq 1) {
            'Suspending BitLocker for two reboots (Device Guard).'
            $RebootCount = 3
        }
        else {
            'Suspending BitLocker for one reboot.'
            $RebootCount = 1
        }

        try {
            $null = Suspend-Bitlocker -MountPoint $SystemDrive -RebootCount $RebootCount
        }
        catch {
            $_.Exception.Message
            exit 1
        }
    }
}

function Set-SecureBootSignedFile {
    param (
        [Parameter(Mandatory)]
        [ValidateSet('PKDefault','KEKDefault','dbDefault','dbxDefault','PK','KEK','db','dbx')]
        [string]$Variable,

        [Parameter(Mandatory)]
        [string]$Filename
    )

    if (-not (Test-Path $Filename)) {
        "$Filename not found."
        exit 1
    }

    try {
        # https://github.com/microsoft/secureboot_objects/blob/main/scripts/windows/InstallSecureBootKeys.ps1
        $null = Set-SecureBootUEFI -Name $Variable -ContentFilePath $Filename -Time '2015-08-28T00:00:00Z'
    }
    catch {
        $ErrorMessage = 'ERROR: Failed to write "{0}" to UEFI {1}.' -f (Split-Path $Filename -Leaf), $Variable.ToUpper()
        Write-Host $ErrorMessage -Foreground Red

        if ($_.Exception.Message -match 'Incorrect authentication') {
            Write-Host 'Wrong signature for this UEFI variable.' -Foreground Red
        }
        else {
            $_.Exception.Message
        }

        exit 1
    }

    'Successfully wrote "{0}" to UEFI {1}.' -f (Split-Path $Filename -Leaf), $Variable
    $script:UEFI_Updated = $true

    Suspend-Bitlocker
}

function Append-SecureBootSignedFile {
    <#
        .SYNOPSIS
        Appends a signed UEFI update package to an UEFI variable

        .DESCRIPTION
        Original Author: Microsoft Secure Boot Team, https://www.powershellgallery.com/packages/SplitDbxContent/1.0
        Modified By: ManubrioTenorio
        Modified By: garlin

        .PARAMETER Variable
        Specifies an UEFI variable, an instance of which is returned by calling the Get-SecureBootUEFI cmdlet.

        .PARAMETER Filename
        Specifies a signed UEFI update package.

        .EXAMPLE
        Append-SecureBootSignedFile -Variable db -Filename ".\DBXUpdate-20230314.x64.bin"
    #>

    param (
        [Parameter(Mandatory)]
        [ValidateSet('PKDefault','KEKDefault','dbDefault','dbxDefault','PK','KEK','db','dbx')]
        [string]$Variable,

        [Parameter(Mandatory)]
        [string]$Filename
    )

    $PSVersion = $PSVersionTable.PSVersion.Major

    $CertName = (Split-Path $Filename -Leaf) -replace '.bin'

    if (-not (Test-Path $Filename)) {
        "$Filename not found."
        exit 1
    }

    if ($PSVersion -gt 5) {
        $Bytes = Get-Content -AsByteStream $Filename
    }
    else {
        $Bytes = Get-Content -Encoding Byte $Filename
    }

    # Identify file signature
    if (($Bytes[40] -ne 0x30) -or ($Bytes[41] -ne 0x82 )) {
        Write-Error "Cannot find signature!" -ErrorAction Stop
    }

    # Signature is known to be ASN size plus header of 4 bytes
    $sig_length = $Bytes[42] * 256 + $Bytes[43] + 4

    if ($sig_length -gt ($Bytes.Length + 40)) {
        Write-Error "Signature longer than file size!" -ErrorAction Stop
    }

    $SigFile = '{0}\{1}.signature.p7' -f $env:TEMP, $CertName
    $ContentFile = '{0}\{1}.content.bin' -f $env:TEMP, $CertName

    # Build and write signature output file
    if ($PSVersion -ge 6) {
        Set-Content -AsByteStream -Path $SigFile -Value ([Byte[]] $Bytes[40..($sig_length + 40 - 1)]) -ErrorAction Stop
    }
    else {
        Set-Content -Encoding Byte -Path $SigFile -Value ([Byte[]] $Bytes[40..($sig_length + 40 - 1)]) -ErrorAction Stop
    }

    # Build and write variable content output file
    if ($PSVersion -ge 6) {
        Set-Content -AsByteStream -Path $ContentFile -Value ([Byte[]] $Bytes[($sig_length + 40)..($Bytes.Length - 1)]) -ErrorAction Stop
    }
    else {
        Set-Content -Encoding Byte -Path $ContentFile -Value ([Byte[]] $Bytes[($sig_length + 40)..($Bytes.Length - 1)]) -ErrorAction Stop
    }

    try {
        # https://github.com/microsoft/secureboot_objects/discussions/158
        $null = Set-SecureBootUEFI -Name $Variable -ContentFilePath $ContentFile -SignedFilePath $SigFile -Time '2010-03-06T19:17:21Z' -AppendWrite
    }
    catch {
        $ErrorMessage = 'ERROR: Failed to append "{0}.bin" to UEFI {1}.' -f $CertName, $Variable.ToUpper()
        Write-Host $ErrorMessage -Foreground Red

        if ($_.Exception.Message -match 'Incorrect authentication') {
            Write-Host 'Wrong signature for this UEFI variable.' -Foreground Red
        }
        else {
            $_.Exception.Message
        }

        exit 1
    }

    'Successfully appended "{0}" to UEFI {1}.' -f (Split-Path $Filename -Leaf), $Variable.ToUpper()
    $script:UEFI_Updated = $true

    Suspend-Bitlocker
    Remove-Item $SigFile,$ContentFile -Force
}

function Match-DBXSignatureData {
    <#
        .SYNOPSIS
        Parses EFI signatures from a DBX Update .bin file and compares the entire list against the current UEFI DBX.

        .DESCRIPTION
        From https://gist.github.com/out0xb2/f8e0bae94214889a89ac67fceb37f8c0#file-check-dbx-ps1
        Modified by github.com/cjee21
        Modified by github.com/garlin-cant-code

        .PARAMETER DBXUpdateFile
        Specifies a filename containing signed DBX Update signatures

        .OUTPUTS
        $true or $false
    #>

    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$DBXUpdateFile
    )

    if (-not (Test-Path $DBXUpdateFile)) {
        Write-Host "DBX update file `"$DBXUpdateFile`" not found." -Foreground Red
        exit 1
    }

    try {
        $RequiredSignatures = Get-UEFIDatabaseSignatures -BytesIn ([IO.File]::ReadAllBytes($DBXUpdateFile)) | where { $_.SignatureType -eq 'EFI_CERT_SHA256_GUID' }
    }
    catch {
        Write-Host "No EFI_CERT_SHA256 signatures in $DBXUpdateFile" -Foreground Red
        return $true
    }

    try {
        $DBXSignatureData = (Get-SecureBootUEFI dbx | Get-UEFIDatabaseSignatures).SignatureList.SignatureData
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $false
        }
        else {
            throw $_.Exception.Message
        }
    }

    $RequiredSignatureData = $RequiredSignatures.SignatureList.SignatureData
    $RequiredCount = $RequiredSignatureData.Count

    if ($RequiredCount -eq 0) {
        Write-Host "No DBX signatures in $DBXUpdateFile" -Foreground Red
        return $true
    }

    $Matched = 0

    foreach ($RequiredSig in $RequiredSignatureData) {
        if ($DBXSignatureData -contains $RequiredSig) {
            $Matched++
        }
        else {
            $RequiredSVN = Get-SignatureDataSVN $RequiredSig

            switch ($RequiredSig) {
                { $_ -match "^$EFI_BOOTMGR_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_BOOTMGR_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                }

                { $_ -match "^$EFI_CDBOOT_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_CDBOOT_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                }

                { $_ -match "^$EFI_WDSMGR_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_WDSMGR_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                }
            }
        }
    }

    if ($Matched -eq $RequiredCount) {
        return $true
    }
    else {
        return $false
    }
}

function Update-PK_Cert {
    # Pre-signed object for Windows OEM Devices PK

    $CertFile = 'WindowsOEMDevicesPK.der'
    $PreSignedObj_File = "$env:TEMP\$CertFile"

    if (-not (Test-Path -LiteralPath "$EFI_FolderPath\$CertFile")) {
        try {
            'Downloading "{0}" from GitHub.' -f $CertFile
            Invoke-WebRequest -UseBasicParsing -Uri $PK_DER_URL -OutFile $PreSignedObj_File
        }
        catch {
            $_.Exception.Message
            exit 1
        }

        if (-not (Test-Path -LiteralPath $EFI_FolderPath)) {
            $null = New-Item -Path $EFI_FolderPath -Type Directory -Force
        }

        'Copying "{0}" to EFI.' -f $CertFile
        Copy-Item -Path $PreSignedObj_File -Destination $EFI_FolderPath -Force

        Remove-Item $PreSignedObj_File -Force
    }

    $script:PK_README = $true
}

function Update-KEK_Cert {
    try {
        $JSON = (Invoke-WebRequest -UseBasicParsing -Uri $KEKUpdateMap_URL).Content | ConvertFrom-Json
    }
    catch {
        "`nERROR: Unable to parse Microsoft's KEK update map."
        Write-Host (($_.Exception.Message -split "`n") | select -First 1) -Foreground Red
        exit 1
    }

    try {
        $PK_Thumbprint = (Get-UefiDatabaseSignatures -BytesIn (Get-SecureBootUEFI PK).Bytes).SignatureList.SignatureData.Thumbprint
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $null
        }
        else {
            throw $_.Exception.Message
        }
    }

    if ($JSON.$PK_Thumbprint.KEKUpdate -ne $null) {
        # Post-signed object for KEK 2K CA 2023

        $array = @($JSON.$PK_Thumbprint.KEKUpdate -split '/')
        $Vendor = $array[0]
        $KEK_Update = $array[1]

        $KEK_BIN_URL = "https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PostSignedObjects/KEK/$Vendor/$KEK_Update"
        $PostSignedObj_File = "$env:TEMP\$KEK_Update"

        try {
            'Downloading "{0}" from GitHub.' -f $KEK_Update
            Invoke-WebRequest -UseBasicParsing -Uri $KEK_BIN_URL -OutFile $PostSignedObj_File
        }
        catch {
            Write-Host $_.Exception.Message -Foreground Red
            exit 1
        }

        Append-SecureBootSignedFile -Variable KEK -Filename $PostSignedObj_File

        Remove-Item $PostSignedObj_File -Force
    }
    else {
        # Pre-signed object for KEK 2K CA 2023

        $CertFile = 'Microsoft Corporation KEK 2K CA 2023.der'
        $PreSignedObj_File = "$env:TEMP\$CertFile"

        if (-not (Test-Path -LiteralPath "$EFI_FolderPath\$CertFile")) {
            try {
                'Downloading "{0}" from GitHub.' -f $CertFile
                Invoke-WebRequest -UseBasicParsing -Uri $KEK_DER_URL -OutFile $PreSignedObj_File
            }
            catch {
                Write-Host $_.Exception.Message -Foreground Red
                exit 1
            }

            if (-not (Test-Path -LiteralPath $EFI_FolderPath)) {
                $null = New-Item -Path $EFI_FolderPath -Type Directory -Force
            }

            'Copying "{0}" to EFI.' -f $CertFile
            Copy-Item -Path $PreSignedObj_File -Destination $EFI_FolderPath -Force
            Copy-Item -Path $PreSignedObj_File -Destination "$EFI_FolderPath\$($CertFile -replace '\.der','.crt')" -Force

            Remove-Item $PreSignedObj_File -Force
        }

        $script:KEK_README = $true
    }
}

function Print-Header {
    param (
        [Parameter(Mandatory=$false)]
        [switch]$Bold,

        [Parameter(Mandatory)]
        [string]$Header
    )

    if ($Bold) {
        $Separator = '='
    }
    else {
        $Separator = '-'
    }

    return ("{0}`n{1}" -f $Header, ($Header -replace "`n" -replace '(.)',$Separator))
}

$ScriptBlock = {
    $CurrentVersion = Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion'

    $SystemDrive = (Get-CimInstance Win32_OperatingSystem).SystemDrive

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'

    $Result = Confirm-MinimumUBR

    if ($Result -ne $true) {
        "ERROR: $Result.`n"
        exit 1
    }

    $SecureBoot = Confirm-SecureBootUEFI

    if ($SecureBoot -isnot [bool]) {
        "ERROR: This PC doesn't support Secure Boot.`n"
        exit 1
    }

    $VBS_Status = (Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard).VirtualizationBasedSecurityStatus

    if ($VBS_Status -gt 0) {
        $VBS_Enabled = $true
    }

    foreach ($Variable in 'PK','KEK','db','dbx') {
        try {
            $Count = (Get-SecureBootUEFI $Variable).Bytes.Count
        }
        catch {
            if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
                $Count = 0
            }
        }

        New-Variable -Name "${Variable}_BytesCount" -Value $Count
    }

    if (((Get-SecureBootUEFI SetupMode).Bytes -Join '') -eq 1) {
        $SetupMode = $true
    }

    try {
        $PK_Cert = Get-UEFICert PK
        $KEK_Certs = Get-UEFICert KEK
        $db_Certs = Get-UEFICert db
        $dbx_Certs = Get-UEFICert dbx
    }
    catch {
        Write-Host 'ERROR: Failed to read UEFI Secure Boot settings.' -Foreground Red
        exit 1
    }

    $PK_Trusted = Check-TrustedPK

    $SystemDisk = (Get-Disk | Where-Object {$_.IsSystem -eq $true}).Number
    $GUID = (Get-Partition -DiskNumber $SystemDisk | Where-Object { $_.Type -eq 'System' }).Guid

    $EFI_Path = "\\?\Volume$GUID\EFI"
    $EFI_FolderPath = "$EFI_Path\Certs"

    $BootMgrEX_File = "$env:SystemRoot\Boot\EFI_EX\bootmgfw_EX.efi"
    $SkuSiPolicy_File = "$env:SystemRoot\System32\SecureBootUpdates\SkuSiPolicy.p7b"

    $BootMgr_File = "$EFI_Path\Microsoft\Boot\bootmgfw.efi"
    $EFI_SkuSiPolicy_File = "$EFI_Path\Microsoft\Boot\SkuSiPolicy.p7b"

    $CheckList = Audit-UEFI

    if ($Audit) {
        Print-Header -Bold "`nAUDIT REPORT"

        if ($CheckList -eq $null) {
            'No action is required.'
        }
        else {
            $CheckList.TrimEnd("`n")
        }

        return
    }

    if ($PK_BytesCount -eq 0 -and ($KEK_BytesCount -eq 0 -or $db_BytesCount -eq 0 -or $dbx_BytesCount -eq 0)) {
        $EDK2_Folder = "$env:TEMP\EDK2_bin"
        Download-EDK2bin

        if ($db_BytesCount -eq 0) {
            Set-SecureBootSignedFile -Variable db -Filename "$EDK2_Folder\Default3PDb.bin"
        }

        if ($dbx_BytesCount -eq 0) {
            Set-SecureBootSignedFile -Variable dbx -Filename "$EDK2_Folder\DefaultDbx.bin"
        }

        if ($KEK_BytesCount -eq 0) {
            Set-SecureBootSignedFile -Variable KEK -Filename "$EDK2_Folder\DefaultKek.bin"
        }

        if ($PK_BytesCount -eq 0) {
            Set-SecureBootSignedFile -Variable PK -Filename "$EDK2_Folder\DefaultPk.bin"
        }

        try {
            $PK_Cert = Get-UEFICert PK
            $KEK_Certs = Get-UEFICert KEK
            $db_Certs = Get-UEFICert db
            $dbx_Certs = Get-UEFICert dbx
        }
        catch {
            Write-Host 'ERROR: Failed to read UEFI Secure Boot settings.' -Foreground Red
            exit 1
        }

        Remove-Item $EDK2_Folder -Recurse -Force
    }

    if (-not $PK_Trusted) {
        Update-PK_Cert
    }

    if ('Microsoft Corporation KEK 2K CA 2023' -notin $KEK_Certs) {
        Update-KEK_Cert
    }

    if ('Windows UEFI CA 2023' -notin $db_Certs) {
        Append-SecureBootSignedFile -Variable db -Filename "$UpdatesFolder\dbupdate2024.bin"
    }

    if ('Microsoft UEFI CA 2023' -notin $db_Certs) {
        Append-SecureBootSignedFile -Variable db -Filename "$UpdatesFolder\DBUpdate3P2023.bin"
    }

    if ('Microsoft Option ROM UEFI CA 2023' -notin $db_Certs) {
        Append-SecureBootSignedFile -Variable db -Filename "$UpdatesFolder\DBUpdateOROM2023.bin"
    }

    if ($Revoke) {
        if ($SecureBoot -and $KEK_Certs -notcontains 'Microsoft Corporation KEK 2K CA 2023') {
            'WARNING: Disable Secure Boot, before attempting to use -Revoke option.  No [KEK 2K CA 2023] cert is currently enrolled.'
            '{0}System will fail to boot due to a security violation.' -f $Tab4
            exit 1
        }

        if ($Latest) {
            try {
                Invoke-WebRequest -UseBasicParsing -Uri $DBXUpdate_bin_URL -OutFile "$env:TEMP\DBXUpdate.bin"
                Invoke-WebRequest -UseBasicParsing -Uri $DBXUpdateSVN_bin_URL -OutFile "$env:TEMP\DBXUpdateSVN.bin"
            }
            catch {
                $_.Exception.Message
                exit 1
            }

            $DBXUpdate_bin = "$env:TEMP\dbxupdate.bin"
            $DBXUpdateSVN_bin = "$env:TEMP\DBXUpdateSVN.bin"
        }
        else {
            $DBXUpdate_bin = "$UpdatesFolder\dbxupdate.bin"
            $DBXUpdateSVN_bin = "$UpdatesFolder\DBXUpdateSVN.bin"
        }

        $DBXSignatureData = (Get-SecureBootUEFI dbx | Get-UEFIDatabaseSignatures).SignatureList.SignatureData

        if (-not $(Match-DBXSignatureData $DBXUpdate_bin)) {
            Append-SecureBootSignedFile -Variable dbx -Filename $DBXUpdate_bin
        }
        elseif ($Latest) {
            '"dbxupdate.bin" is not a newer version of file.'
        }

        if ('Microsoft Windows Production PCA 2011' -notin (Get-UEFICert dbx)) {
            Append-SecureBootSignedFile -Variable dbx -Filename "$UpdatesFolder\DBXUpdate2024.bin"
        }

        if ('Microsoft Windows Production PCA 2011' -in (Get-UEFICert dbx)) {
            if (-not $(Match-DBXSignatureData $DBXUpdateSVN_bin)) {
                $Result = Append-SecureBootSignedFile -Variable dbx -Filename $DBXUpdateSVN_bin
                $SVN = Get-SecureBootUEFI_DBXSVN $EFI_BOOTMGR_DBXSVN_GUID

                $Result -replace ' to'," (SVN $SVN) to"
                $UEFI_Updated = $true
            }
            elseif ($Latest) {
                '"DBXUpdateSVN.bin" is not a newer version of file.'
            }
        }
    }

    if (($Revoke -and $VBS_Enabled) -or $SkuSiPolicy) {
        if ((Test-Path -LiteralPath $EFI_SkuSiPolicy_File)) {
            $SkuSiPolicy_File_Hash = (Get-FileHash $SkuSiPolicy_File).Hash
            $EFI_SkuSiPolicy_File_Hash = (Get-FileHash -LiteralPath $EFI_SkuSiPolicy_File).Hash

            if ($EFI_SkuSiPolicy_File_Hash -ne $SkuSiPolicy_File_Hash) {
                Copy-Item $SkuSiPolicy_File "$EFI_SkuSiPolicy_File" -Force

                'Deployed SkuSiPolicy.p7b (for VBS).'
                $UEFI_Updated = $true
            }
        }
        else {
            Copy-Item $SkuSiPolicy_File "$EFI_SkuSiPolicy_File" -Force

            'Deployed SkuSiPolicy.p7b (for VBS).'
            $UEFI_Updated = $true
        }
    }

    if ($Revoke -and $SBAT) {
        $null = Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot' -Name AvailableUpdates -Value 0x400
        Start-ScheduledTask -TaskName "\Microsoft\Windows\PI\Secure-Boot-Update"

        'Applying SBAT update for Linux.'
        $UEFI_Updated = $true
    }

    if ($Latest) {
        Remove-Item $DBXUpdate_bin,$DBXUpdateSVN_bin -Force
    }

    $BootMgrEX_File = "$env:SystemRoot\Boot\EFI_EX\bootmgfw_EX.efi"

    if ('Windows UEFI CA 2023' -in (Get-UEFICert db)) {
        $BootMgrEX_File_Hash = (Get-FileHash $BootMgrEX_File).Hash
        $BootMgr_File_Hash = (Get-FileHash -LiteralPath $BootMgr_File).Hash

        if ($BootMgr_File_Hash -ne $BootMgrEX_File_Hash) {
            'Copying EFI boot files.'

            $RE_info = reagentc /info

            if (($RE_info -match 'RE status:' -split ' ')[-1] -eq 'Enabled') {
                $WinRE = $true

                $WinRE_Path = ($RE_info -match 'RE location:' -split ' ')[-1]
                $WinRE_GUID = ($RE_info -match 'identifier:' -split ' ')[-1]
            }

            $EFI_DriveLetter = (& mountvol) -split "`n" | foreach { if ($_ -match '(.*mounted at )(.*)(\\)') { $Matches[2] } }

            if ($EFI_DriveLetter -eq $null) {
                $EFI_DriveLetter = ((68..89 | foreach { [char]$_ + ':' }) | where { (Get-CimInstance Win32_LogicalDisk).DeviceID -notcontains $_ }) | select -First 1

                if ($EFI_DriveLetter -eq $null) {
                    'ERROR: Unable to assign drive letter for EFI partition.'
                    exit 1
                }

                try {
                    Start-Process 'mountvol' -ArgumentList "$EFI_DriveLetter /s" -NoNewWindow -Wait
                    Start-Process 'bcdboot' -ArgumentList "$env:SystemRoot /s $EFI_DriveLetter /f UEFI /bootex" -NoNewWindow -Wait
                    Start-Process 'mountvol' -ArgumentList "$EFI_DriveLetter /d" -NoNewWindow -Wait
                }
                catch {
                    $_.Exception.Message
                    exit 1
                }
            }
            else {
                try {
                    Start-Process 'bcdboot' -ArgumentList "$env:SystemRoot /s $EFI_DriveLetter /f UEFI /bootex" -NoNewWindow -Wait
                }
                catch {
                    $_.Exception.Message
                    exit 1
                }
            }

            if ($WinRE) {
                try {
                    Start-Process 'reagentc' -ArgumentList "/setreimage /path $WinRE_Path" -NoNewWindow -Wait
                    Start-Process 'bcdedit' -ArgumentList "/set {default} recoverysequence {$WinRE_GUID}" -NoNewWindow -Wait
                }
                catch {
                    $_.Exception.Message
                    exit 1
                }

                $null = New-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name Enable_WinRE -Value 'conhost --headless C:\Windows\System32\reagentc.exe /enable' -Force
            }

            $UEFI_Updated = $true
        }

        if ($BootMedia) {
            $RemovableDrives = Get-Volume | where { $_.DriveType -eq 'Removable' -and $_.DriveLetter -ne $null } | sort DriveLetter

            if ($RemovableDrives.Count -eq 0) {
                "No USB removable media found.`n"
            }
            else {
                foreach ($Volume in $RemovableDrives) {
                    $DriveLetter = $Volume.DriveLetter
                    $EFI_BootFile = "${DriveLetter}:\EFI\boot\boot${EDK2_Arch}.efi"

                    if (-not (Test-Path $EFI_BootFile)) {
                        continue
                    }

                    $EFI_BootFile_Hash = (Get-FileHash -LiteralPath $EFI_BootFile).Hash

                    if ($EFI_BootFile_Hash -ne $BootMgrEX_File_Hash) {
                        $Label = (Get-Volume -DriveLetter $DriveLetter).FileSystemLabel

                        if ($Label -ne '') {
                            '{0}Copying EFI boot files to USB Drive {1}: "{2}"' -f $Tab4, $DriveLetter, $Label
                        }
                        else {
                            '{0}Copying EFI boot files to USB Drive {1}:' -f $Tab4, $DriveLetter
                        }

                         try {
                             Copy-Item "${DriveLetter}:\EFI\Microsoft\Boot\BCD" $env:TEMP -Force
                             Start-Process 'bcdboot' -ArgumentList "$env:SystemRoot /f UEFI /s $DriveLetter /bootex" -NoNewWindow -Wait
                             Copy-Item "$env:TEMP\BCD" "${DriveLetter}:\EFI\Microsoft\Boot\BCD" -Force
                             Remove-Item "$env:TEMP\BCD" -Force
                         }
                         catch {
                             $_.Exception.Message
                             exit 1
                         }

                        $Media_Updated = $true
                    }
                }

                if ($Media_Updated) { '' }
            }
        }
    }

    if ($UEFI_Updated -or $PK_README -or $KEK_README) {
        if ($UEFI_Updated -or $Latest) {
            Write-Output ''
        }

        Print-Header 'REQUIRED ACTION'

        if ($PK_README -or $KEK_README) {
            if ($PK_README -or $KEK_README) {
                if ($PK_README -and $KEK_README) {
                    $CertName = 'PK and [KEK CA 2023] certs'
                }
                elseif ($PK_README) {
                    $CertName = 'PK cert'
                }
                else {
                    $CertName = '[KEK CA 2023] cert'
                }

                "Please follow the README_UEFI.TXT instructions, for installing the {0} from BIOS.`n" -f $CertName
            }
        }

        'Restart Windows, for UEFI updates to take effect.'
    }
    else {
        if ($Latest) {
            Write-Output ''
        }

       'SUCCESS: NO UPDATES ARE REQUIRED.'
    }
}

if ($Log) {
    $System = Get-CimInstance -ClassName Win32_ComputerSystem
    $LogFile = '{0}\{1} {2} Update-UEFI.log' -f $PSScriptRoot, (Get-Date -Format 'yyyy-MM-dd'), $System.Model.ToUpper()

    & $ScriptBlock | Tee-Object $LogFile
    "`nLog file saved as `"{0}`"`n" -f $LogFile
}
else {
    & $ScriptBlock
    Write-Output ''
}
'@

$script:Embedded['Check_DBXUpdate.bin.ps1'] = @'
<#PSScriptInfo

.VERSION 2026.01.18

.GUID dbcc69b3-3e30-4e71-a1a9-29ef49f06afc

.AUTHOR garlin

.COPYRIGHT

.TAGS UEFI, Secure Boot, DBX, SVN, Windows Boot Manager

.RELEASENOTES

#>

<#
.SYNOPSIS
    Script to confirm every EFI_CERT_SHA256 signature read from a DBX Update .bin file, is contained in the UEFI DBX variable.

.DESCRIPTION
    Run this script to check if DBX Update files have been fully applied to UEFI DBX.

.PARAMETER Version
    Print the script's version number and exit.

.PARAMETER Verbose
    Download "dbx_info_msft_latest.json" from Microsoft's Secure Boot Objects GitHub, and identify unmatched EFI_CERT_SHA256 signatures
    by their signature hash, filename, vendor and revocation date.

    Identify if unmatched signatures contain a higher DBX SVN, than currently stored in UEFI DBX.

.PARAMETER Log
    Save script output to a file named "YYYY-MM-DD [Model] Check DBX.log"

.PARAMETER Paths
    Search a list of provided folder paths or individual filenames, for DBX Update files and check each file for confirmation that UEFI DBX
    contains every EFI_CERT_SHA256 signature in the file.

    Folder path defaults to "C:\Windows\System32\SecureBootUpdates".

.EXAMPLE
    Check_DBXUpdate.bin.ps1
.EXAMPLE
    Check_DBXUpdate.bin.ps1 \path\Folder1 \path\DBXUpdate2.bin
.EXAMPLE
    Check_DBXUpdate.bin.ps1 -Verbose -Log
#>

[CmdletBinding(DefaultParameterSetName='Default')]
param (
    [Parameter(Mandatory=$false,ParameterSetName='Version')]
    [switch]$Version,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [switch]$Log,

    [Parameter(Mandatory=$false,ParameterSetName='Default',DontShow,ValueFromRemainingArguments=$true)]
    [string[]]$Paths = @()
)

$ScriptVersion = '2026.01.18'

# https://github.com/microsoft/secureboot_objects/blob/main/Archived/dbx_info_msft_4_09_24_svns.csv
$EFI_BOOTMGR_DBXSVN_GUID = '01612B139DD5598843AB1C185C3CB2EB92'
$EFI_CDBOOT_DBXSVN_GUID =  '019D2EF8E827E15841A4884C18ABE2F284'
$EFI_WDSMGR_DBXSVN_GUID =  '01C2CA99C9FE7F6F4981279E2A8A535976'

$DBXinfo_URL = 'https://raw.githubusercontent.com/microsoft/secureboot_objects/main/PreSignedObjects/DBX/dbx_info_msft_latest.json'

$Tab4 = ' ' * 4

if ($Version) {
    '{0} version ({1}){2}' -f $MyInvocation.MyCommand.Name, $ScriptVersion, $(if ($MyInvocation.Line -ne '') { "`n" })
    exit 0
}

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if ($PSVersionTable.PSVersion.Major -gt 5) {
        $PS = 'pwsh'
    }
    else {
        $PS = 'powershell'
    }

    $args = ($MyInvocation.BoundParameters.Keys.GetEnumerator() | where { $_ -ne 'Paths' } | foreach { '-{0}' -f $_ }) -join ' '

    if ($MyInvocation.BoundParameters.'Paths' -ne $null) {
        $args += ' ' + ($MyInvocation.BoundParameters.'Paths' | foreach { '"{0}"' -f (Get-Item $_ -ErrorAction SilentlyContinue).FullName }) -join ' '
    }

    Start-Process $PS -ArgumentList "-nop -ep bypass -NoLogo -NoExit -f $($MyInvocation.MyCommand.Path) $args" -Verb RunAs
    exit 0
}

if ($PSBoundParameters['Verbose']) {
    $Verbose = $true
    $VerbosePreference = 'SilentlyContinue'
}

function Get-UefiDatabaseSignatures {
    <#
        .SYNOPSIS
        Parses UEFI Signature Databases into logical Powershell objects
        # https://github.com/cjee21/Check-UEFISecureBootVariables

        .DESCRIPTION
        Original Author: Matthew Graeber (@mattifestation)
        Modified By: Jeremiah Cox (@int0x6)
        Modified By: Joel Roth (@nafai)
        Modified By: garlin (@garlin-cant-code)
        Additional Source: https://gist.github.com/mattifestation/991a0bea355ec1dc19402cef1b0e3b6f
        Additional Source: https://www.powershellgallery.com/packages/SplitDbxContent/1.0
        License: BSD 3-Clause

        .PARAMETER Variable
        Specifies an UEFI variable, an instance of which is returned by calling the Get-SecureBootUEFI cmdlet.

        .PARAMETER BytesIn
        Specifies a byte array consisting of the PKDefault, KEKDefault, dbDefault, dbxDefault, PK, KEK, db, or dbx UEFI variable contents.

        .EXAMPLE
        $DbxBytes = [IO.File]::ReadAllBytes('.\dbx.bin')
        Get-UEFIDatabaseSignatures -BytesIn $DbxBytes

        .EXAMPLE
        Get-UEFIDatabaseSignatures -Filename ".\DBXUpdate-20230314.x64.bin"

        .EXAMPLE
        Get-SecureBootUEFI -Name db | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name dbx | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name pk | Get-UEFIDatabaseSignatures

        .EXAMPLE
        Get-SecureBootUEFI -Name kek | Get-UEFIDatabaseSignatures

        .INPUTS
        Microsoft.SecureBoot.Commands.UEFIEnvironmentVariable
        Accepts the output of Get-SecureBootUEFI over the pipeline.

        .OUTPUTS
        UefiSignatureDatabase
        Outputs an array of custom powershell objects describing a UEFI Signature Database. "77fa9abd-0359-4d32-bd60-28f4e78f784b" refers to Microsoft as the owner.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'UEFIVariable')]
        [ValidateScript({ ($_.GetType().Fullname -eq 'Microsoft.SecureBoot.Commands.UEFIEnvironmentVariable') -and ($_.Name -in 'PKDefault','KEKDefault','dbDefault','dbxDefault','pk','kek','db','dbx') })]
        $Variable,

        [Parameter(Mandatory, ParameterSetName = 'ByteArray')]
        [Byte[]]
        [ValidateNotNullOrEmpty()]
        $BytesIn,

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]
        [ValidateScript({ (Resolve-Path "$_").where({Test-Path $_}).Path })]
        $Filename
    )

    $PSVersion = $PSVersionTable.PSVersion.Major

    $SignatureTypeMapping = @{
        'C1C41626-504C-4092-ACA9-41F936934328' = 'EFI_CERT_SHA256_GUID' # Most often used for dbx
        'A5C059A1-94E4-4AA7-87B5-AB155C2BF072' = 'EFI_CERT_X509_GUID'   # Most often used for db
    }

    $Bytes = $null

    if ($Filename)
    {
        if ($PSVersion -gt 5) {
            $Bytes = Get-Content -AsByteStream $Filename -ErrorAction Stop
        }
        else {
            $Bytes = Get-Content -Encoding Byte $Filename  -ErrorAction Stop
        }
    }
    elseif ($Variable)
    {
        $Bytes = $Variable.Bytes
    }
    else
    {
        $Bytes = $BytesIn
    }

    # Modified from Split-Dbx
    if (($Bytes[40] -eq 0x30) -and ($Bytes[41] -eq 0x82 ))
    {
        Write-Debug "Removing signature."

        # Signature is known to be ASN size plus header of 4 bytes
        $sig_length = $Bytes[42] * 256 + $Bytes[43] + 4
        if ($sig_length -gt ($Bytes.Length + 40)) {
            Write-Error "Signature longer than file size!" -ErrorAction Stop
        }

        ## Unsigned db store
        [System.Byte[]]$Bytes = @($Bytes[($sig_length+40)..($Bytes.Length - 1)].Clone())
    }
    else
    {
        Write-Debug "Signature not found. Assuming it's already split."
    }

    try
    {
        $MemoryStream = New-Object -TypeName IO.MemoryStream -ArgumentList @(,$Bytes)
        $BinaryReader = New-Object -TypeName IO.BinaryReader -ArgumentList $MemoryStream, ([Text.Encoding]::Unicode)
    }
    catch
    {
        throw $_
        return
    }

    # What follows will be an array of EFI_SIGNATURE_LIST structs

    while ($BinaryReader.PeekChar() -ne -1) {
        $SignatureType = $SignatureTypeMapping[([Guid][Byte[]] $BinaryReader.ReadBytes(16)).Guid]
        $SignatureListSize = $BinaryReader.ReadUInt32()
        $SignatureHeaderSize = $BinaryReader.ReadUInt32()
        $SignatureSize = $BinaryReader.ReadUInt32()

        $SignatureHeader = $BinaryReader.ReadBytes($SignatureHeaderSize)

        # 0x1C is the size of the EFI_SIGNATURE_LIST header
        $SignatureCount = ($SignatureListSize - 0x1C) / $SignatureSize

        $SignatureList = 1..$SignatureCount | ForEach-Object {
            $SignatureDataBytes = $BinaryReader.ReadBytes($SignatureSize)

            $SignatureOwner = [Guid][Byte[]] $SignatureDataBytes[0..15]

            switch ($SignatureType) {
                'EFI_CERT_SHA256_GUID' {
                    $SignatureData = ([Byte[]] $SignatureDataBytes[0x10..0x2F] | ForEach-Object { $_.ToString('X2') }) -join ''
                }

                'EFI_CERT_X509_GUID' {
                    try {
                        $SignatureData = New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,([Byte[]] $SignatureDataBytes[16..($SignatureDataBytes.Count - 1)]))
                    }
                    catch {
                        Write-Host "Skipping an invalid $Variable X509 certificate."
                    }
                }
            }

            [PSCustomObject] @{
                PSTypeName = 'EFI.SignatureData'
                SignatureOwner = $SignatureOwner
                SignatureData = $SignatureData
            }
        }

        [PSCustomObject] @{
            PSTypeName = 'EFI.SignatureList'
            SignatureType = $SignatureType
            SignatureList = $SignatureList
        }
    }
}

function Get-SignatureDataSVN {
    param (
        [Parameter(Mandatory)]
        [string]$SignatureData
    )

    # https://github.com/microsoft/secureboot_objects/blob/main/scripts/utility_functions.py
    $SVN = '{0}.{1}' -f [int]::Parse($SignatureData.Substring(36,4), [System.Globalization.NumberStyles]::HexNumber), [int]::Parse($SignatureData.Substring(40,4), [System.Globalization.NumberStyles]::HexNumber)

    return $SVN
}

function Get-SecureBootUEFI_DBXSVN {
    param (
        [Parameter(Mandatory)]
        [string]$DBXSVN
    )

    try {
        $SignatureData = (Get-SecureBootUEFI dbx | Get-UEFIDatabaseSignatures).SignatureList.SignatureData
    }
    catch {
        if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
            return $null
        }
        else {
            throw $_.Exception.Message
        }
    }

    $LastSig = $SignatureData -match "^$DBXSVN" | sort | select -Last 1

    if ($LastSig.Count) {
        $SVN = Get-SignatureDataSVN $LastSig
    }
    else {
        $SVN = $null
    }

    return $SVN
}

function Compare-DBXSignatureData {
    <#
        .SYNOPSIS
        Parses EFI signatures from a DBX Update .bin file and compares the entire list against the current UEFI DBX.

        .DESCRIPTION
        From https://gist.github.com/out0xb2/f8e0bae94214889a89ac67fceb37f8c0#file-check-dbx-ps1
        Modified by github.com/cjee21
        Modified by github.com/garlin-cant-code

        .PARAMETER InputObject
        [PSCustomObject] containing list of signed DBX Update filenames

        .OUTPUTS
        List of unmatched DBX signatures
    #>

    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]$InputObject
    )

    $DBXUpdateFile = $InputObject.FullPath

    if ($HidePath) {
        $Filename = '"{0}"' -f (Split-Path $InputObject.FullPath -Leaf)
    }
    else {
        $Filename = '"{0}"' -f $InputObject.RelativePath
    }

    if (-not (Test-Path $DBXUpdateFile)) {
        Write-Host "DBX update file `"$DBXUpdateFile`" not found." -Foreground Red
    }

    try {
        $RequiredSignatures = Get-UEFIDatabaseSignatures -BytesIn ([IO.File]::ReadAllBytes($DBXUpdateFile)) | where { $_.SignatureType -eq 'EFI_CERT_SHA256_GUID' }
    }
    catch {
        Write-Host "No EFI_CERT_SHA256 signatures in $DBXUpdateFile" -Foreground Red
        return $null
    }

    $RequiredSignatureData = $RequiredSignatures.SignatureList.SignatureData
    $RequiredCount = $RequiredSignatureData.Count

    if ($RequiredCount -eq 0) {
        Write-Host "No DBX signatures in $DBXUpdateFile" -Foreground Red
        return $null
    }

    $Matched = 0
    $MissingSigList = $null

    foreach ($RequiredSig in $RequiredSignatureData) {
        if ($DBXSignatureData -contains $RequiredSig) {
            $Matched++

            switch ($RequiredSig) {
                { $_ -match "^$EFI_BOOTMGR_DBXSVN_GUID" } { $SVN_SigCount++ }
                { $_ -match "^$EFI_CDBOOT_DBXSVN_GUID" }  { $SVN_SigCount++ }
                { $_ -match "^$EFI_WDSMGR_DBXSVN_GUID" }  { $SVN_SigCount++ }
                default { $EFI_SigCount++ }
             }
        }
        else {
            $RequiredSVN = Get-SignatureDataSVN $RequiredSig

            switch ($RequiredSig) {
                { $_ -match "^$EFI_BOOTMGR_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_BOOTMGR_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                    else {
                        $MissingSigList += "{0}Missing [{1}] bootmgfw.efi SVN {2}`n" -f $Tab4, $RequiredSig, (Get-SignatureDataSVN $RequiredSig)
                    }
                }

                { $_ -match "^$EFI_CDBOOT_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_CDBOOT_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                    else {
                        $MissingSigList += "{0}Missing [{1}] cdboot.efi SVN {2}`n" -f $Tab4, $RequiredSig, (Get-SignatureDataSVN $RequiredSig)
                    }
                }

                { $_ -match "^$EFI_WDSMGR_DBXSVN_GUID" } {
                    $CurrentSVN = Get-SecureBootUEFI_DBXSVN $EFI_WDSMGR_DBXSVN_GUID

                    if ($CurrentSVN -ge $RequiredSVN) {
                        $Matched++
                    }
                    else {
                        $MissingSigList += "{0}Missing [{1}] wdsmgfw.efi SVN {2}`n" -f $Tab4, $RequiredSig, (Get-SignatureDataSVN $RequiredSig)
                    }
                }

                default {
                    if ($Verbose) {
                        $MissingSig = $JSON.images.$Arch | where { $_.authenticodeHash -eq $RequiredSig }

                        if ($MissingSig -ne $null) {
                            if ($MissingSig.filename -eq '') {
                                $MissingSig.filename = '(none)'
                            }

                            $Columns = ('{0} {1} {2}' -f $MissingSig.filename, $MissingSig.companyName, $MissingSig.dateOfAddition) -replace '  '
                            $MissingSigList += "{0}Missing [{1}] {2}`n" -f $Tab4, $MissingSig.authenticodeHash, $Columns
                        }
                    }
                }
            }
        }
    }


    if ($EFI_SigCount -and $SVN_SigCount) {
        $SigType = 'EFI/SVN'
    }
    elseif ($EFI_SigCount) {
        $SigType = 'EFI'
    }
    else {
        $SigType = 'SVN'
    }

    if ($Matched -eq $RequiredCount) {
        $Result = 'SUCCESS: Matched {0}/{1} {2} signatures from {3}' -f $Matched, $RequiredCount, $SigType, $Filename
        Write-Host $Result -ForegroundColor Green

        if ($Log) {
            $Result | Add-Content $LogFile
        }
    }
    else {
        $Result = 'FAILED: Missing {0}/{1} {2} signatures from {3}' -f ($RequiredCount - $Matched), $RequiredCount, $SigType, $Filename
        Write-Host $Result -ForegroundColor Red

        if ($Verbose) {
            $MissingSigList
        }

        if ($Log) {
            @($Result; if ($Verbose) { $MissingSigList }) | Add-Content $LogFile
        }
    }
}


switch ($env:PROCESSOR_ARCHITECTURE) {
    'amd64' { $Arch = 'x64' }
    'x86'   { $Arch = 'x86' }
    'arm64' { $Arch = 'aa64' }
    'arm'   { $Arch = 'aa32' }
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'

if ($Verbose) {
    try {
        $JSON = (Invoke-WebRequest -UseBasicParsing -Uri $DBXinfo_URL).Content | ConvertFrom-Json
    }
    catch {
        $_.Exception.Message
        exit 1
    }
}

$System = Get-CimInstance -ClassName Win32_ComputerSystem
$LogFile = '{0}\{1} {2} Check-DBXUpdate.log' -f $PSScriptRoot, (Get-Date -Format 'yyyy-MM-dd'), $System.Model.ToUpper()

if (Test-Path $LogFile) {
    Remove-Item $LogFile -Force
}

if ($Paths.Count -eq 0) {
    $Paths = @("$env:SystemRoot\System32\SecureBootUpdates")
    $HidePath = $true
}

try {
    $DBXSignatureData = (Get-SecureBootUEFI dbx | Get-UEFIDatabaseSignatures).SignatureList.SignatureData
}
catch {
    if ($_.Exception.Message -eq 'Variable is currently undefined: 0xC0000100') {
        $Result = "FAILED: UEFI DBX variable is currently empty.`n"
        Write-Host $Result -Foreground Red

        if ($Log) {
            $Result | Add-Content $LogFile
        }

        exit 1
    }
    else {
        throw $_.Exception.Message
    }
}

$DBX_Files = @()
$SortKey = 0

foreach ($item in $Paths) {
    if ($item -match '^\.') {
        $Path = Resolve-Path $item -Relative

        if (Test-Path $Path -PathType Container) {
            foreach ($File in (Resolve-Path (Get-ChildItem $Path -File).FullName -Relative)) {
                if ($File -match 'dbx.*bin$') {
                    $DBX_Files += [PSCustomObject]@{
                        RelativePath = $File
                        FullPath = (Get-Item $File).FullName
                        SortKey = $SortKey++
                    }
                }
            }
        }
        else {
            $DBX_Files += [PSCustomObject]@{
                RelativePath = $Path
                FullPath = (Get-Item $Path).FullName
                SortKey = $SortKey++
            }
        }
    }
    else {
        $Path = (Resolve-Path $item).Path

        if (Test-Path $Path -PathType Container) {
            foreach ($File in (Get-ChildItem $Path -File).FullName) {
                if ($File -match 'dbx.*bin$') {
                    $DBX_Files += [PSCustomObject]@{
                        RelativePath = $File
                        FullPath = $File
                        SortKey = $SortKey++
                    }
                }
            }
        }
        else {
            $DBX_Files += [PSCustomObject]@{
                RelativePath = $Path
                FullPath = (Get-Item $Path).FullName
                SortKey = $SortKey++
            }
        }
    }
}

foreach ($File in ($DBX_Files | sort FullPath -Unique | sort SortKey)) {
    Compare-DBXSignatureData $File
}

if ($Log) {
    "`nLog file saved as `"{0}`"`n" -f $LogFile
}
else {
    Write-Output ''
}
'@

$script:Embedded['HP_Enable-CA2023-BIOS.ps1'] = @'
<#PSScriptInfo

.VERSION 2026.11.26

.TAGS HP, BIOS, UEFI, Secure Boot, CA 2023

#>

<#
.SYNOPSIS
    HP: zet de Secure Boot 2023-certificaatinstellingen in de BIOS aan via HP BIOS WMI (geen HP CMSL nodig).

.DESCRIPTION
    Zoekt dynamisch alle HP BIOS-instellingen met 'CA 2023' in de naam (bv. Windows UEFI CA 2023, Microsoft UEFI CA 2023,
    Microsoft Option ROM UEFI CA 2023) en zet ze op Enable. De wijziging is pas actief na een herstart.

.PARAMETER Audit
    Alleen de huidige waarden van de HP 2023-instellingen tonen, niets wijzigen.

.PARAMETER BiosPassword
    BIOS-adminwachtwoord, alleen nodig als er een BIOS-wachtwoord is ingesteld.

.PARAMETER SuspendBitLocker
    BitLocker op de systeemschijf 1 herstart pauzeren als er een BIOS-instelling is gewijzigd (voorkomt herstelsleutel-prompt).
#>

[CmdletBinding()]
param (
    [switch]$Audit,
    [string]$BiosPassword,
    [switch]$SuspendBitLocker
)

$Namespace = 'root/HP/InstrumentedBIOS'
$ReturnCodes = @{
    0 = 'Success'
    1 = 'Not Supported'
    2 = 'Unspecified Error'
    3 = 'Timeout'
    4 = 'Failed'
    5 = 'Invalid Parameter'
    6 = 'Access Denied (verkeerd of ontbrekend BIOS-wachtwoord?)'
}

function Write-Status {
    param([string]$Message, [string]$Color = 'Gray')
    Write-Host "[HP CA2023] $Message" -ForegroundColor $Color
}

$cs = Get-CimInstance -ClassName Win32_ComputerSystem
if ($cs.Manufacturer -notmatch 'HP|Hewlett') {
    Write-Status "Geen HP-apparaat (fabrikant: $($cs.Manufacturer)). Gestopt." 'Red'
    exit 2
}
Write-Status ("Apparaat: {0} {1}" -f $cs.Manufacturer, $cs.Model) 'Cyan'

if ($env:firmware_type -and $env:firmware_type -ne 'UEFI') {
    Write-Status "Windows draait in $($env:firmware_type)-modus; Secure Boot-certificaten zijn niet van toepassing. Gestopt." 'Red'
    exit 1
}

try {
    $Enum      = @(Get-CimInstance -Namespace $Namespace -ClassName HP_BIOSEnumeration -ErrorAction Stop)
    $Interface = Get-CimInstance -Namespace $Namespace -ClassName HP_BIOSSettingInterface -ErrorAction Stop
}
catch {
    Write-Status 'HP BIOS WMI-namespace niet beschikbaar (geen HP business-model of BIOS te oud).' 'Red'
    exit 1
}

$Settings = @($Enum | Where-Object { $_.Name -match 'CA\s*2023' } | Sort-Object Name)
if ($Settings.Count -eq 0) {
    Write-Status 'Geen BIOS-instellingen met "CA 2023" gevonden. Update eerst de HP BIOS naar een versie met CA 2023-ondersteuning.' 'Yellow'
    exit 1
}

''
Write-Status 'Gevonden 2023-instellingen:' 'Cyan'
foreach ($s in $Settings) {
    $col = if ($s.CurrentValue -match '^Enable') { 'Green' } else { 'Yellow' }
    Write-Host ("    {0,-40} {1}" -f $s.Name, $s.CurrentValue) -ForegroundColor $col
}
''

if ($Audit) {
    $todo = @($Settings | Where-Object { $_.CurrentValue -notmatch '^Enable' })
    if ($todo.Count) { Write-Status "AUDIT: $($todo.Count) instelling(en) staan nog niet op Enable." 'Yellow' }
    else { Write-Status 'AUDIT: alle 2023-instellingen staan al op Enable.' 'Green' }
    exit 0
}

$Password = '<utf-16/>' + $BiosPassword
$Changed = 0
$Failed = 0

foreach ($s in $Settings) {
    $enableValue = @($s.PossibleValues | Where-Object { $_ -match '^Enable' }) | Select-Object -First 1
    if (-not $enableValue) { $enableValue = 'Enable' }

    if ($s.CurrentValue -eq $enableValue) {
        Write-Status "'$($s.Name)' staat al op '$enableValue'." 'Gray'
        continue
    }
    if ($s.IsReadOnly -eq 1) {
        Write-Status "'$($s.Name)' is alleen-lezen in deze BIOS; overgeslagen." 'Yellow'
        continue
    }

    $r = Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{ Name = $s.Name; Value = $enableValue; Password = $Password }
    $code = [int]$r.Return
    $txt = if ($ReturnCodes.ContainsKey($code)) { $ReturnCodes[$code] } else { "Onbekende code $code" }
    if ($code -eq 0) {
        Write-Status "'$($s.Name)' -> '$enableValue': OK" 'Green'
        $Changed++
    }
    else {
        Write-Status "'$($s.Name)' -> '$enableValue': MISLUKT ($txt)" 'Red'
        $Failed++
    }
}

''
Write-Status 'Controle (waarden na wijziging, actief na herstart):' 'Cyan'
$Enum = @(Get-CimInstance -Namespace $Namespace -ClassName HP_BIOSEnumeration)
foreach ($s in $Settings) {
    $e = $Enum | Where-Object Name -eq $s.Name
    $val = if ($e.CurrentValue) { $e.CurrentValue } else { $e.Value }
    Write-Host ("    {0,-40} {1}" -f $s.Name, $val) -ForegroundColor $(if ($val -match 'Enable') { 'Green' } else { 'Yellow' })
}
''

if ($Changed -gt 0 -and $SuspendBitLocker) {
    try {
        $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        if ([string]$bl.ProtectionStatus -eq 'On') {
            Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount 1 -ErrorAction Stop | Out-Null
            Write-Status "BitLocker gepauzeerd op $env:SystemDrive voor 1 herstart." 'Green'
        }
    }
    catch { Write-Status "BitLocker pauzeren mislukt: $($_.Exception.Message)" 'Red' }
}

if ($Failed -gt 0) {
    'ERROR: Niet alle 2023-instellingen konden worden aangezet.'
    exit 1
}
if ($Changed -gt 0) {
    'SUCCESS: HP 2023-certificaatinstellingen aangezet. Herstart nodig om ze toe te passen.'
    if (-not $BiosPassword) {
        'WARNING: Zonder BIOS-wachtwoord kan HP bij de volgende boot om lokale bevestiging (code) vragen.'
    }
    'Draai na de herstart Check_UEFI-CA2023 / Update_UEFI-CA2023 om de certificaten in KEK/DB te controleren.'
}
else {
    'SUCCESS: Alle HP 2023-instellingen stonden al aan.'
}
exit 0
'@

$script:Embedded['README_UEFI.TXT'] = @'
Manual installation of [Windows OEM Devices PK]
===============================================

1. Shutdown Windows, and enter your UEFI's Secure Boot menu.

2. Enter "PK Options / Enroll PK / Enroll PK Using File" or "Key Management / PK Management / Set Key".
   The menu options may be different for your BIOS.

   - Browse the system drive's EFI partition
   - Enter the <EFI> folder
   - Enter the <Certs> sub-folder

3. Find the file "WindowsOEMDevicesPK.der".  Add the certificate.

4. Enter "KEK Options / Enroll KEK / Enroll KEK Using File" or "Key Management / KEK Management / Append Key".
   The menu options may be different for your BIOS.

   - Browse the system drive's EFI partition
   - Enter the <EFI> folder
   - Enter the <Updates> sub-folder

5. Find the file "Microsoft Corporation KEK 2K CA 2023.der".  Add this certificate.

6. Save changes and exit.

7. Start Windows, and re-run the 'Update-UEFI_CA2023.ps1' script.


Manual installation of [KEK 2K CA 2023]
=======================================

1. Shutdown Windows, and enter your UEFI's Secure Boot menu.

2. Enter "KEK Options / Enroll KEK / Enroll KEK Using File" or "Key Management / KEK Management / Append Key".
   The menu options may be different for your BIOS.

   - Browse the system drive's EFI partition
   - Enter the <EFI> folder
   - Enter the <Certs> sub-folder

3. Find the file "Microsoft Corporation KEK 2K CA 2023.der".  Add this certificate.
   If you encounter an error, try the file "Microsoft Corporation KEK 2K CA 2023.crt".

4. Save changes and exit.

5. Start Windows, and re-run the 'Update-UEFI_CA2023.ps1' script.
'@

#endregion

#region Start
Export-EmbeddedFiles
$script:Defs = Get-ScriptDefinitions
$script:UI.cmbScript.ItemsSource = $script:Defs
$pref = $script:Defs | Where-Object { $_.Name -like 'Check_UEFI*' } | Select-Object -First 1
if ($pref) { $script:UI.cmbScript.SelectedItem = $pref } elseif ($script:Defs.Count) { $script:UI.cmbScript.SelectedIndex = 0 }

[void]$script:Window.ShowDialog()
#endregion
