<#
.SYNOPSIS
    WPF GUI om een testmail via SMTP te versturen.
.DESCRIPTION
    Interactief SMTP-server, poort, credentials, afzender, ontvanger, onderwerp en body invullen
    en een testmail versturen. Gebruikt System.Net.Mail.SmtpClient (STARTTLS op 587/25).
#>

# WPF vereist STA; herstart indien nodig
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="EOO - SMTP Testmail versturings" Width="560" Height="660"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        FontFamily="Segoe UI" FontSize="13" Background="#F5F6F8">
    <Window.Resources>
        <Style TargetType="TextBlock" x:Key="Lbl">
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Height" Value="28"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="4,0"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Height" Value="28"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="4,0"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Height" Value="28"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="Background" Value="White"/>
        </Style>
    </Window.Resources>

    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="90"/>
        </Grid.RowDefinitions>

        <!-- Server -->
        <GroupBox Grid.Row="0" Header="SMTP server">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="80"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition/>
                    <RowDefinition/>
                    <RowDefinition/>
                    <RowDefinition/>
                </Grid.RowDefinitions>

                <TextBlock Style="{StaticResource Lbl}" Text="Server"/>
                <ComboBox Grid.Column="1" x:Name="cbServer" IsEditable="True">
                    <ComboBoxItem Content="c23008-p1.proxhost.nl"/>
                    <ComboBoxItem Content="smtp.eoo.support"/>
                </ComboBox>
                <TextBlock Grid.Column="2" Style="{StaticResource Lbl}" Text="Poort" Margin="10,0,10,0"/>
                <ComboBox Grid.Column="3" x:Name="cbPort" IsEditable="True">
                    <ComboBoxItem Content="587"/>
                    <ComboBoxItem Content="25"/>
                    <ComboBoxItem Content="2525"/>
                </ComboBox>

                <CheckBox Grid.Row="1" Grid.Column="1" x:Name="chkSsl" Content="SSL/TLS gebruiken (STARTTLS)" IsChecked="True" Margin="0,6"/>
                <CheckBox Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="3" x:Name="chkAuth" Content="Authenticatie" IsChecked="True" Margin="230,6,0,6"/>

                <TextBlock Grid.Row="2" Style="{StaticResource Lbl}" Text="Gebruikersnaam"/>
                <TextBox Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="3" x:Name="tbUser"/>

                <TextBlock Grid.Row="3" Style="{StaticResource Lbl}" Text="Wachtwoord"/>
                <PasswordBox Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="3" x:Name="pbPass"/>
            </Grid>
        </GroupBox>

        <!-- Adressen -->
        <GroupBox Grid.Row="1" Header="Bericht">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition/>
                    <RowDefinition/>
                    <RowDefinition/>
                </Grid.RowDefinitions>
                <TextBlock Style="{StaticResource Lbl}" Text="Van"/>
                <TextBox Grid.Column="1" x:Name="tbFrom" ToolTip="Leeg = gebruikersnaam"/>
                <TextBlock Grid.Row="1" Style="{StaticResource Lbl}" Text="Aan"/>
                <TextBox Grid.Row="1" Grid.Column="1" x:Name="tbTo" ToolTip="Meerdere adressen scheiden met ; of ,"/>
                <TextBlock Grid.Row="2" Style="{StaticResource Lbl}" Text="Onderwerp"/>
                <TextBox Grid.Row="2" Grid.Column="1" x:Name="tbSubject" Text="EOO testmail versturings"/>
            </Grid>
        </GroupBox>

        <!-- Body -->
        <GroupBox Grid.Row="2" Header="Inhoud">
            <TextBox x:Name="tbBody" Height="Auto" AcceptsReturn="True" TextWrapping="Wrap"
                     VerticalContentAlignment="Top" VerticalScrollBarVisibility="Auto" Padding="4"
                     Text="Dit is een testmail van EOO."/>
        </GroupBox>

        <!-- Knoppen -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,10">
            <Button x:Name="btnClear" Content="Log wissen" Width="100" Height="32" Margin="0,0,8,0"/>
            <Button x:Name="btnSend" Content="Verstuur testmail" Width="150" Height="32"
                    Background="#0B5CAD" Foreground="White" FontWeight="SemiBold" IsDefault="True"/>
        </StackPanel>

        <!-- Log -->
        <TextBox Grid.Row="4" x:Name="tbLog" Height="Auto" IsReadOnly="True" TextWrapping="Wrap"
                 VerticalContentAlignment="Top" VerticalScrollBarVisibility="Auto"
                 FontFamily="Consolas" FontSize="12" Background="#1E1E1E" Foreground="#D4D4D4" Padding="4"/>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$ui = @{}
$xaml.SelectNodes('//*[@*[local-name()="Name"]]') | ForEach-Object {
    $ui[$_.Name] = $window.FindName($_.Name)
}

$ui.cbServer.SelectedIndex = 0
$ui.cbPort.SelectedIndex   = 0

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Message
    $ui.tbLog.AppendText($line)
    $ui.tbLog.ScrollToEnd()
}

# Authenticatie aan/uit schakelt velden
$ui.chkAuth.Add_Click({
    $enabled = [bool]$ui.chkAuth.IsChecked
    $ui.tbUser.IsEnabled = $enabled
    $ui.pbPass.IsEnabled = $enabled
})

$ui.btnClear.Add_Click({ $ui.tbLog.Clear() })

$ui.btnSend.Add_Click({
    $server  = $ui.cbServer.Text.Trim()
    $portTxt = $ui.cbPort.Text.Trim()
    $useAuth = [bool]$ui.chkAuth.IsChecked
    $user    = $ui.tbUser.Text.Trim()
    $from    = $ui.tbFrom.Text.Trim()
    if (-not $from) { $from = $user }
    $toList  = $ui.tbTo.Text -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    # Validatie
    $port = 0
    $errors = @()
    if (-not $server)                                 { $errors += 'SMTP server is leeg.' }
    if (-not [int]::TryParse($portTxt, [ref]$port))   { $errors += 'Poort is geen geldig getal.' }
    if ($useAuth -and -not $user)                     { $errors += 'Gebruikersnaam is leeg.' }
    if (-not $from)                                   { $errors += 'Afzender (Van) is leeg.' }
    if (-not $toList)                                 { $errors += 'Ontvanger (Aan) is leeg.' }
    if ($errors) {
        [System.Windows.MessageBox]::Show(($errors -join "`n"), 'Controleer invoer', 'OK', 'Warning') | Out-Null
        return
    }

    if ($port -eq 465) {
        Write-Log 'Let op: poort 465 (implicit SSL) wordt niet ondersteund door SmtpClient. Gebruik 587 (STARTTLS).'
    }

    $ui.btnSend.IsEnabled = $false
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    Write-Log "Verbinden met ${server}:$port (SSL: $([bool]$ui.chkSsl.IsChecked), Auth: $useAuth)..."

    $msg  = $null
    $smtp = $null
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = $from
        foreach ($t in $toList) { $msg.To.Add($t) }
        $msg.Subject = $ui.tbSubject.Text
        $msg.Body    = $ui.tbBody.Text
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
        $msg.BodyEncoding    = [System.Text.Encoding]::UTF8

        $smtp = New-Object System.Net.Mail.SmtpClient($server, $port)
        $smtp.EnableSsl = [bool]$ui.chkSsl.IsChecked
        $smtp.Timeout   = 30000
        if ($useAuth) {
            $smtp.UseDefaultCredentials = $false
            $smtp.Credentials = New-Object System.Net.NetworkCredential($user, $ui.pbPass.Password)
        }

        $smtp.Send($msg)
        Write-Log "Testmail verstuurd van $from naar $($toList -join ', ')."
        [System.Windows.MessageBox]::Show('Testmail succesvol verstuurd.', 'Gelukt', 'OK', 'Information') | Out-Null
    }
    catch {
        $ex = $_.Exception
        Write-Log "FOUT: $($ex.Message)"
        while ($ex.InnerException) {
            $ex = $ex.InnerException
            Write-Log "  -> $($ex.Message)"
        }
        [System.Windows.MessageBox]::Show("Versturen mislukt:`n$($_.Exception.Message)", 'Fout', 'OK', 'Error') | Out-Null
    }
    finally {
        if ($msg)  { $msg.Dispose() }
        if ($smtp) { $smtp.Dispose() }
        $window.Cursor = $null
        $ui.btnSend.IsEnabled = $true
    }
})

$window.ShowDialog() | Out-Null
