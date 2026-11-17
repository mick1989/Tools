#Requires -Version 5.1
<#
.SYNOPSIS
    WPF-tool om een netwerkprinter (inclusief driver, wachtrijen en lade-instellingen) te verpakken als Intune Win32-app.

.DESCRIPTION
    De tool vraagt interactief om:
      - printeradres, poort en protocol (RAW/LPR, SNMP)
      - de map met uitgepakte driverbestanden, het INF-bestand en de drivernaam
      - installeerbare opties (bijv. extra lades) en een of meer printerwachtrijen met standaardlade,
        dubbelzijdig, kleur en papierformaat
    Daarna wordt gevraagd waar het pakket opgeslagen moet worden en wordt automatisch gemaakt:
      <map>\<Packagenaam>\Source\install.ps1 + uninstall.ps1 + Drivers\
      <map>\<Packagenaam>\Output\install.intunewin
      <map>\<Packagenaam>\detect.ps1, info.txt, <Packagenaam>_logo.png, PrinterPackage.json

    "Lades/opties uitlezen" installeert de driver tijdelijk lokaal (administrator nodig) om de
    beschikbare lades en installeerbare opties van de driver op te halen.

.PARAMETER LoadConfig
    Optioneel pad naar een eerder opgeslagen PrinterPackage.json.
#>
[CmdletBinding()]
param(
    [string]$LoadConfig
)

# WPF vereist STA; gebruik altijd Windows PowerShell 5.1 (PrintManagement + WPF)
if ($PSVersionTable.PSEdition -eq 'Core' -or [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`"")
    if ($LoadConfig) { $argList += @('-LoadConfig', "`"$LoadConfig`"") }
    Start-Process -FilePath (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -ArgumentList $argList
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

if (-not ('EooPrinterPkg.QueueRow' -as [type])) {
    Add-Type -TypeDefinition @'
namespace EooPrinterPkg {
    public class QueueRow {
        public string Name { get; set; }
        public string Location { get; set; }
        public string Comment { get; set; }
        public string Tray { get; set; }
        public string Duplex { get; set; }
        public string Color { get; set; }
        public string PaperSize { get; set; }
        public QueueRow() { Name = ""; Location = ""; Comment = ""; Tray = ""; Duplex = ""; Color = ""; PaperSize = ""; }
    }
    public class OptionRow {
        public string Property { get; set; }
        public string Default { get; set; }
        public string Value { get; set; }
    }
    public class Choice {
        public string Display { get; set; }
        public string Value { get; set; }
        public Choice() { }
        public Choice(string display, string value) { Display = display; Value = value; }
        public override string ToString() { return Display; }
    }
}
'@
}

$script:ToolPath = $PSCommandPath
$script:ToolDir  = $PSScriptRoot
$script:PsfUri   = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework'
$script:PskUri   = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords'
$script:RegBase  = 'SOFTWARE\EOO\IntunePrinters'
$script:Utf8Bom  = New-Object System.Text.UTF8Encoding($true)
$script:LastAutoPort = ''
$script:LastOutput   = ''

#region Generated script templates

$script:InstallTemplate = @'
<#
    Printer package : __PACKAGE_NAME__
    Versie          : __VERSION__
    Gegenereerd     : __DATE__ (New-PrinterIntunePackage.ps1)

    Install command : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
    Log             : %ProgramData%\Microsoft\IntuneManagementExtension\Logs\__PACKAGE_NAME__-install.log
#>

__CONFIG_BLOCK__

# Herstart in 64-bit PowerShell wanneer Intune het script als 32-bit proces start
if ($env:PROCESSOR_ARCHITEW6432 -eq 'AMD64') {
    $ps64 = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $ps64) {
        & $ps64 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
        exit $LASTEXITCODE
    }
}

$PackageName = $Config.PackageName
$LogFolder   = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
Start-Transcript -Path (Join-Path $LogFolder "$PackageName-install.log") -Append | Out-Null

$PsfUri = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework'
$PskUri = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords'

function New-PsfNamespaceManager {
    param([System.Xml.XmlDocument]$Document)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
    $nsm.AddNamespace('psf', $PsfUri)
    return ,$nsm
}

function Get-LocalName {
    param([string]$QName)
    return ($QName -split ':')[-1]
}

function Format-ErrorDetail {
    param($ErrorRecord)
    $e = $ErrorRecord.Exception
    $text = $e.Message
    if ($e.InnerException) { $text += " | $($e.InnerException.Message)" }
    $text += ' (HResult 0x{0:X8})' -f $e.HResult
    if ($e.PSObject.Properties['NativeErrorCode']) { $text += " NativeErrorCode=$($e.NativeErrorCode)" }
    if ($e.PSObject.Properties['ErrorData'] -and $e.ErrorData) {
        foreach ($p in $e.ErrorData.CimInstanceProperties | Where-Object { $_.Value -and $_.Name -match 'Message|Code' }) {
            $text += " $($p.Name)=$($p.Value)"
        }
    }
    return $text
}

function Resolve-QName {
    param([System.Xml.XmlElement]$Node, [string]$QName)
    if ($QName.Contains(':')) {
        $prefix, $local = $QName -split ':', 2
        return [pscustomobject]@{ Uri = $Node.GetNamespaceOfPrefix($prefix); Local = $local }
    }
    return [pscustomobject]@{ Uri = ''; Local = $QName }
}

function Test-OptionMatch {
    param([System.Xml.XmlElement]$Option, [string]$Wanted)
    $qname = $Option.GetAttribute('name')
    if (-not $qname) { return $false }
    $resolved = Resolve-QName -Node $Option -QName $qname
    if ($Wanted -match '^\{(?<ns>[^}]*)\}(?<name>.+)$') {
        return ($resolved.Local -eq $Matches['name'] -and $resolved.Uri -eq $Matches['ns'])
    }
    return ($resolved.Local -eq $Wanted)
}

function Find-CapabilityChoice {
    # Zoekt in de PrintCapabilities van de driver de feature(s) en de best passende optie
    param(
        [System.Xml.XmlDocument]$Caps,
        [string]$Label,
        [string[]]$FeatureNames,
        [string[]]$WantedOptions
    )
    $nsm = New-PsfNamespaceManager $Caps
    $features = @($Caps.DocumentElement.SelectNodes('psf:Feature', $nsm) |
        Where-Object { $FeatureNames -contains (Get-LocalName $_.GetAttribute('name')) })
    if ($features.Count -eq 0) {
        Write-Warning "    ${Label}: de driver ondersteunt geen $($FeatureNames -join '/')."
        return
    }
    foreach ($feature in $features) {
        $options = @($feature.SelectNodes('psf:Option', $nsm))
        $match = $null
        foreach ($wanted in $WantedOptions) {
            $match = $options | Where-Object { Test-OptionMatch -Option $_ -Wanted $wanted } | Select-Object -First 1
            if ($match) { break }
        }
        $featureLocal = Get-LocalName $feature.GetAttribute('name')
        if (-not $match) {
            Write-Warning ("    {0}: '{1}' niet beschikbaar in {2}. Driver biedt: {3}" -f $Label, ($WantedOptions -join ' / '), $featureLocal,
                (($options | ForEach-Object { $_.GetAttribute('name') }) -join ', '))
            continue
        }
        Write-Host "    ${Label}: $featureLocal = $($match.GetAttribute('name'))"
        [pscustomobject]@{
            Label   = $Label
            Feature = Resolve-QName -Node $feature -QName $feature.GetAttribute('name')
            Option  = Resolve-QName -Node $match -QName $match.GetAttribute('name')
        }
    }
}

function New-DeltaTicketXml {
    # Minimaal PrintTicket met alleen de gewenste features; de driver vult de rest aan bij validatie
    param($Choices)
    $prefixes = [ordered]@{}
    $prefixes[$PsfUri] = 'psf'
    $prefixes[$PskUri] = 'psk'
    $qualify = {
        param($Resolved)
        if (-not $Resolved.Uri) { return $Resolved.Local }
        if (-not $prefixes.Contains($Resolved.Uri)) { $prefixes[$Resolved.Uri] = "ns$($prefixes.Count)" }
        return "$($prefixes[$Resolved.Uri]):$($Resolved.Local)"
    }
    $body = foreach ($c in $Choices) {
        '  <psf:Feature name="{0}"><psf:Option name="{1}"/></psf:Feature>' -f (& $qualify $c.Feature), (& $qualify $c.Option)
    }
    $declarations = ($prefixes.GetEnumerator() | ForEach-Object { 'xmlns:{0}="{1}"' -f $_.Value, $_.Key }) -join ' '
    return "<psf:PrintTicket version=`"1`" $declarations xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`">`r`n$($body -join "`r`n")`r`n</psf:PrintTicket>"
}

function Set-DefaultTicketSystemPrinting {
    # Zelfde route als het venster 'Standaardinstellingen voor afdrukken': driver valideert en zet om naar DEVMODE
    param([string]$PrinterName, [string]$DeltaXml)
    Add-Type -AssemblyName System.Printing, ReachFramework
    $server = New-Object System.Printing.LocalPrintServer
    $queue = New-Object System.Printing.PrintQueue($server, $PrinterName, [System.Printing.PrintSystemDesiredAccess]::AdministratePrinter)
    try {
        $stream = New-Object System.IO.MemoryStream(, [System.Text.Encoding]::UTF8.GetBytes($DeltaXml))
        $delta = New-Object System.Printing.PrintTicket($stream)
        $result = $queue.MergeAndValidatePrintTicket($queue.DefaultPrintTicket, $delta)
        Write-Host "    Validatie door driver: $($result.ConflictStatus)"
        $queue.DefaultPrintTicket = $result.ValidatedPrintTicket
        $queue.Commit()
    }
    finally {
        $queue.Dispose()
        $server.Dispose()
    }
}

function Set-SingleTicketOption {
    # Fallback: een feature per keer in het huidige ticket zetten via Set-PrintConfiguration
    param([string]$PrinterName, $Choice)
    $config = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop
    $ticket = New-Object System.Xml.XmlDocument
    $ticket.LoadXml($config.PrintTicketXML)
    $root = $ticket.DocumentElement
    $nsm = New-PsfNamespaceManager $ticket

    $getPrefix = {
        param([string]$Uri)
        $p = $root.GetPrefixOfNamespace($Uri)
        if ([string]::IsNullOrEmpty($p)) {
            $i = 0
            do { $p = "eoo$i"; $i++ } while ($root.GetNamespaceOfPrefix($p))
            $attribute = $ticket.CreateAttribute('xmlns', $p, 'http://www.w3.org/2000/xmlns/')
            $attribute.Value = $Uri
            [void]$root.Attributes.Append($attribute)
        }
        return $p
    }
    $psfPrefix   = & $getPrefix $PsfUri
    $featureName = if ($Choice.Feature.Uri) { "$(& $getPrefix $Choice.Feature.Uri):$($Choice.Feature.Local)" } else { $Choice.Feature.Local }
    $optionName  = if ($Choice.Option.Uri) { "$(& $getPrefix $Choice.Option.Uri):$($Choice.Option.Local)" } else { $Choice.Option.Local }

    $feature = @($root.SelectNodes('psf:Feature', $nsm)) |
        Where-Object { (Get-LocalName $_.GetAttribute('name')) -eq $Choice.Feature.Local } | Select-Object -First 1
    if (-not $feature) {
        $feature = $ticket.CreateElement($psfPrefix, 'Feature', $PsfUri)
        $feature.SetAttribute('name', $featureName)
        [void]$root.AppendChild($feature)
    }
    foreach ($old in @($feature.SelectNodes('psf:Option', $nsm))) { [void]$feature.RemoveChild($old) }
    $option = $ticket.CreateElement($psfPrefix, 'Option', $PsfUri)
    $option.SetAttribute('name', $optionName)
    [void]$feature.AppendChild($option)

    Set-PrintConfiguration -PrinterName $PrinterName -PrintTicketXml $root.OuterXml -ErrorAction Stop
}

function Get-TicketOption {
    param([System.Xml.XmlDocument]$Ticket, [string[]]$FeatureNames)
    $nsm = New-PsfNamespaceManager $Ticket
    $feature = @($Ticket.DocumentElement.SelectNodes('psf:Feature', $nsm)) |
        Where-Object { $FeatureNames -contains (Get-LocalName $_.GetAttribute('name')) } | Select-Object -First 1
    if (-not $feature) { return '(niet in ticket)' }
    $option = $feature.SelectSingleNode('psf:Option', $nsm)
    if ($option) { return $option.GetAttribute('name') }
    return '(leeg)'
}

function Set-PrinterDefaults {
    param($Queue)

    $name = $Queue.Name
    $paperMap = @{ A4 = 'ISOA4'; A3 = 'ISOA3'; A5 = 'ISOA5'; Letter = 'NorthAmericaLetter'; Legal = 'NorthAmericaLegal' }
    $binFeatures = @('PageInputBin', 'JobInputBin', 'DocumentInputBin')

    # Direct na Add-Printer is de configuratie soms nog niet beschikbaar
    $printConfig = $null
    for ($i = 1; $i -le 10 -and -not $printConfig; $i++) {
        try { $printConfig = Get-PrintConfiguration -PrinterName $name -ErrorAction Stop }
        catch { Start-Sleep -Seconds 3 }
    }
    if (-not $printConfig) { throw "Printerconfiguratie van '$name' is niet leesbaar." }

    $caps = New-Object System.Xml.XmlDocument
    $caps.LoadXml($printConfig.PrintCapabilitiesXML)

    $choices = @()
    if ($Queue.Duplex)            { $choices += @(Find-CapabilityChoice -Caps $caps -Label 'Dubbelzijdig' -FeatureNames 'JobDuplexAllDocumentsContiguously', 'DocumentDuplex' -WantedOptions $Queue.Duplex) }
    if ($Queue.Color -eq 'Color') { $choices += @(Find-CapabilityChoice -Caps $caps -Label 'Kleur' -FeatureNames 'PageOutputColor' -WantedOptions 'Color') }
    if ($Queue.Color -eq 'Mono')  { $choices += @(Find-CapabilityChoice -Caps $caps -Label 'Kleur' -FeatureNames 'PageOutputColor' -WantedOptions 'Monochrome', 'Grayscale') }
    if ($Queue.PaperSize)         { $choices += @(Find-CapabilityChoice -Caps $caps -Label 'Papier' -FeatureNames 'PageMediaSize' -WantedOptions $paperMap[$Queue.PaperSize]) }
    if ($Queue.Tray)              { $choices += @(Find-CapabilityChoice -Caps $caps -Label 'Standaardlade' -FeatureNames $binFeatures -WantedOptions $Queue.Tray) }

    if ($choices.Count) {
        $deltaXml = New-DeltaTicketXml -Choices $choices
        $done = $false
        try {
            Set-DefaultTicketSystemPrinting -PrinterName $name -DeltaXml $deltaXml
            $done = $true
            Write-Host '    Afdrukvoorkeuren opgeslagen via System.Printing.'
        }
        catch {
            Write-Host "    System.Printing mislukt: $(Format-ErrorDetail $_)"
            Write-Host "    Ticket:`r`n$deltaXml"
        }
        if (-not $done) {
            foreach ($choice in $choices) {
                try {
                    Set-SingleTicketOption -PrinterName $name -Choice $choice
                    Write-Host "    $($choice.Label) los ingesteld via Set-PrintConfiguration."
                }
                catch { Write-Host "    $($choice.Label) los instellen mislukt: $(Format-ErrorDetail $_)" }
            }
        }
    }

    # Controle en laatste fallback via de standaard parameters van Set-PrintConfiguration
    $after = Get-PrintConfiguration -PrinterName $name -ErrorAction Stop
    $wantColor = $Queue.Color -eq 'Color'
    $fallback = @()
    if ($Queue.Duplex -and "$($after.DuplexingMode)" -ne $Queue.Duplex) { $fallback += , @('Dubbelzijdig', @{ DuplexingMode = $Queue.Duplex }) }
    if ($Queue.Color -and [bool]$after.Color -ne $wantColor)              { $fallback += , @('Kleur', @{ Color = $wantColor }) }
    if ($Queue.PaperSize -and "$($after.PaperSize)" -ne $Queue.PaperSize)  { $fallback += , @('Papier', @{ PaperSize = $Queue.PaperSize }) }
    foreach ($f in $fallback) {
        $params = $f[1]
        Write-Host "    $($f[0]) nog niet goed; fallback Set-PrintConfiguration $(($params.GetEnumerator() | ForEach-Object { "-$($_.Key) $($_.Value)" }) -join ' ')"
        try { Set-PrintConfiguration -PrinterName $name @params -ErrorAction Stop }
        catch { Write-Host "    Fallback mislukt: $(Format-ErrorDetail $_)" }
    }

    $final = Get-PrintConfiguration -PrinterName $name -ErrorAction Stop
    $finalTicket = New-Object System.Xml.XmlDocument
    $finalTicket.LoadXml($final.PrintTicketXML)
    $finalBin = Get-TicketOption -Ticket $finalTicket -FeatureNames $binFeatures
    Write-Host ("    Resultaat: Dubbelzijdig={0}, Kleur={1}, Papier={2}, Lade={3}" -f $final.DuplexingMode, $(if ($final.Color) { 'Kleur' } else { 'Zwart-wit' }), $final.PaperSize, $finalBin)

    if ($Queue.Duplex -and "$($final.DuplexingMode)" -ne $Queue.Duplex) { Write-Warning "    Dubbelzijdig is NIET ingesteld zoals gewenst ($($Queue.Duplex))." }
    if ($Queue.Color -and [bool]$final.Color -ne $wantColor) { Write-Warning "    Kleur/zwart-wit is NIET ingesteld zoals gewenst ($($Queue.Color))." }
    if ($Queue.PaperSize -and "$($final.PaperSize)" -ne $Queue.PaperSize) { Write-Warning "    Papier is NIET ingesteld zoals gewenst ($($Queue.PaperSize))." }
    if ($Queue.Tray) {
        $wantedBin = $Queue.Tray -replace '^\{[^}]*\}', ''
        if ((Get-LocalName $finalBin) -ne $wantedBin) { Write-Warning "    Standaardlade is NIET ingesteld zoals gewenst ($wantedBin)." }
    }
}

function Clear-UserPrintPreferences {
    # Oude per-gebruiker voorkeuren (van een eerdere installatie) gaan voor op de printerstandaard
    param([string]$PrinterName)
    foreach ($hive in Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^S-1-5-21-[\d-]+$' }) {
        foreach ($sub in 'Printers\DevModePerUser', 'Printers\DevModes2') {
            $path = "Registry::HKEY_USERS\$($hive.PSChildName)\$sub"
            $key = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
            if ($key -and ($key.GetValueNames() -contains $PrinterName)) {
                Remove-ItemProperty -LiteralPath $path -Name $PrinterName -ErrorAction SilentlyContinue
                Write-Host "    Oude gebruikersvoorkeuren verwijderd: $($hive.PSChildName)\$sub"
            }
        }
    }
}

function Invoke-WithTimeout {
    param([string]$FilePath, [string]$Arguments, [int]$TimeoutSeconds = 300)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $process = [System.Diagnostics.Process]::Start($psi)
    $stdOut = $process.StandardOutput.ReadToEndAsync()
    $stdErr = $process.StandardError.ReadToEndAsync()
    $timedOut = -not $process.WaitForExit($TimeoutSeconds * 1000)
    if ($timedOut) { try { $process.Kill() } catch { } }
    $process.WaitForExit()
    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        Output   = "$($stdOut.Result)$($stdErr.Result)"
        TimedOut = $timedOut
    }
}

function Add-DriverPublisherTrust {
    param([string]$Path)
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('TrustedPublisher', 'LocalMachine')
    $store.Open('ReadWrite')
    try {
        $seen = @{}
        foreach ($cat in Get-ChildItem -Path $Path -Filter '*.cat' -Recurse -File -ErrorAction SilentlyContinue) {
            $signature = Get-AuthenticodeSignature -FilePath $cat.FullName -ErrorAction SilentlyContinue
            $cert = $signature.SignerCertificate
            if (-not $cert -or $seen.ContainsKey($cert.Thumbprint)) { continue }
            $seen[$cert.Thumbprint] = $true
            Write-Host "Catalogus $($cat.Name): $($signature.Status), ondertekend door $($cert.Subject)"
            if ($cert.Subject -match 'Microsoft Windows Hardware Compatibility') { continue }
            if ($store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }) {
                Write-Host '    Uitgever staat al in TrustedPublisher.'
                continue
            }
            $store.Add($cert)
            Write-Host "    Uitgever toegevoegd aan TrustedPublisher (LocalMachine): $($cert.Subject)"
        }
        if ($seen.Count -eq 0) { Write-Warning 'Geen ondertekende .cat-bestanden gevonden; Windows kan om bevestiging vragen of de driver weigeren.' }
    }
    finally {
        $store.Close()
    }
}

function Get-InfDriverVersion {
    # Leest DriverVer (datum,versie) uit het INF-bestand, ook als die via %strings% is opgegeven
    param([string]$Path)
    $lines = Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue
    $line = $lines | Where-Object { $_ -match '^\s*DriverVer\s*=' } | Select-Object -First 1
    if (-not $line) { return $null }
    $value = (($line -split '=', 2)[1] -split ';')[0]
    $versionPart = (($value -split ',')[-1]).Trim().Trim('"')
    if ($versionPart -match '^%(.+)%$') {
        $key = [regex]::Escape($Matches[1])
        $def = $lines | Where-Object { $_ -match "^\s*$key\s*=" } | Select-Object -First 1
        if ($def) { $versionPart = ((($def -split '=', 2)[1] -split ';')[0]).Trim().Trim('"') }
    }
    $version = $null
    if ([version]::TryParse($versionPart, [ref]$version)) { return $version }
    return $null
}

function ConvertTo-DriverVersion {
    # Get-PrinterDriver geeft DriverVersion als UInt64 (4 x 16 bit)
    param($Value)
    if ($null -eq $Value) { return $null }
    $v = [uint64]$Value
    return [version]('{0}.{1}.{2}.{3}' -f (($v -shr 48) -band 0xFFFF), (($v -shr 32) -band 0xFFFF), (($v -shr 16) -band 0xFFFF), ($v -band 0xFFFF))
}

function Find-DriverStoreInf {
    param([string]$InfName, [version]$Version)
    $repo = Join-Path $env:WINDIR 'System32\DriverStore\FileRepository'
    $candidates = @(Get-ChildItem -Path $repo -Directory -Filter "$InfName*" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object { Join-Path $_.FullName $InfName } |
        Where-Object { Test-Path $_ })
    if ($Version) {
        $match = $candidates | Where-Object { (Get-InfDriverVersion -Path $_) -eq $Version } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $candidates | Select-Object -First 1
}

function Install-Driver {
    $driverRoot = Join-Path $PSScriptRoot 'Drivers'
    $pnputil    = Join-Path $env:WINDIR 'System32\pnputil.exe'

    if ($Config.StageAllInfs) {
        $infFiles = @(Get-ChildItem -Path $driverRoot -Filter '*.inf' -Recurse -File)
    }
    else {
        $infFiles = @(Get-Item -LiteralPath (Join-Path $driverRoot $Config.InfPath) -ErrorAction Stop)
    }

    # Uitgever(s) van de driver vertrouwen, anders vraagt Windows om bevestiging (kan niet onder SYSTEM/Intune)
    $trustRoot = if ($Config.StageAllInfs) { $driverRoot } else { Split-Path (Join-Path $driverRoot $Config.InfPath) -Parent }
    Add-DriverPublisherTrust -Path $trustRoot

    foreach ($inf in $infFiles) {
        Write-Host "Driver toevoegen aan driver store: $($inf.FullName)"
        $result = Invoke-WithTimeout -FilePath $pnputil -Arguments "/add-driver `"$($inf.FullName)`"" -TimeoutSeconds 300
        $result.Output -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Host "    $_" }
        Write-Host "    pnputil exitcode: $($result.ExitCode)"
        if ($result.TimedOut) { throw "pnputil reageert niet (wacht waarschijnlijk op een bevestiging); driver niet toegevoegd." }
    }

    # Bestaat de driver al (bijv. van een eerdere installatie of andere printer)? Dan versie vergelijken.
    $infName        = Split-Path $Config.InfPath -Leaf
    $packageVersion = Get-InfDriverVersion -Path (Join-Path $driverRoot $Config.InfPath)
    $existing       = Get-PrinterDriver -Name $Config.DriverName -ErrorAction SilentlyContinue
    if ($existing) {
        $installedVersion = ConvertTo-DriverVersion $existing.DriverVersion
        Write-Host "Printerdriver '$($Config.DriverName)' is al aanwezig: versie $installedVersion (package: $packageVersion)."
        if (-not $packageVersion -or ($installedVersion -and $installedVersion -ge $packageVersion)) {
            Write-Host '    Zelfde of nieuwere versie; bestaande driver wordt gebruikt.'
            return
        }
        Write-Host '    Oudere versie; driver wordt bijgewerkt naar de versie uit het package.'
    }

    # Installeren/bijwerken vanuit precies de versie die net in de driver store is gezet
    $storeInf = Find-DriverStoreInf -InfName $infName -Version $packageVersion
    if ($storeInf) {
        Write-Host "    Driver store INF: $storeInf"
        Add-PrinterDriver -Name $Config.DriverName -InfPath $storeInf -ErrorAction Stop
    }
    else {
        Write-Host "    INF '$infName' (versie $packageVersion) niet gevonden in de driver store; installeren op naam."
        Add-PrinterDriver -Name $Config.DriverName -ErrorAction Stop
    }
    $installed = Get-PrinterDriver -Name $Config.DriverName -ErrorAction Stop
    Write-Host "    Geinstalleerde versie: $(ConvertTo-DriverVersion $installed.DriverVersion)"
    Write-Host "Printerdriver '$($Config.DriverName)' geinstalleerd."
}

function Install-Port {
    $address = $Config.PrinterAddress
    $port    = Get-PrinterPort -Name $Config.PortName -ErrorAction SilentlyContinue

    if (-not $port) {
        $params = @{ Name = $Config.PortName; ErrorAction = 'Stop' }
        if ($Config.Protocol -eq 'LPR') {
            $params.LprHostAddress  = $address
            $params.LprQueueName    = $Config.LprQueue
            $params.LprByteCounting = $true
        }
        else {
            $params.PrinterHostAddress = $address
            $params.PortNumber         = [uint32]$Config.PortNumber
        }
        if ($Config.SnmpEnabled) {
            $params.SNMP          = 1
            $params.SNMPCommunity = $Config.SnmpCommunity
        }
        Add-PrinterPort @params
        Write-Host "Printerpoort '$($Config.PortName)' aangemaakt ($address)."
        return
    }

    $filterName = $Config.PortName -replace "'", "\'"
    $cimPort = Get-CimInstance -ClassName Win32_TCPIPPrinterPort -Filter "Name='$filterName'" -ErrorAction SilentlyContinue
    if ($cimPort -and $cimPort.HostAddress -ne $address) {
        Set-CimInstance -InputObject $cimPort -Property @{ HostAddress = $address } -ErrorAction Stop
        Write-Host "Printerpoort '$($Config.PortName)' bijgewerkt naar $address."
    }
    else {
        Write-Host "Printerpoort '$($Config.PortName)' bestaat al."
    }
}

function Install-Queue {
    param($Queue)

    $name = $Queue.Name
    if (Get-Printer -Name $name -ErrorAction SilentlyContinue) {
        Set-Printer -Name $name -DriverName $Config.DriverName -PortName $Config.PortName -ErrorAction Stop
        Write-Host "Printer '$name' bestond al; driver en poort bijgewerkt."
    }
    else {
        Add-Printer -Name $name -DriverName $Config.DriverName -PortName $Config.PortName -ErrorAction Stop
        Write-Host "Printer '$name' toegevoegd."
    }

    $printerParams = @{}
    if ($Queue.Location) { $printerParams.Location = $Queue.Location }
    if ($Queue.Comment)  { $printerParams.Comment  = $Queue.Comment }
    if ($printerParams.Count) {
        try { Set-Printer -Name $name @printerParams -ErrorAction Stop }
        catch { Write-Warning "Locatie/opmerking instellen mislukt: $($_.Exception.Message)" }
    }

    # Installeerbare opties (apparaatinstellingen, bijv. extra lades)
    foreach ($option in @($Config.DeviceOptions)) {
        try {
            Set-PrinterProperty -PrinterName $name -PropertyName $option.Property -Value $option.Value -ErrorAction Stop
            Write-Host "    Optie $($option.Property) = $($option.Value)"
        }
        catch { Write-Warning "Optie $($option.Property) instellen mislukt: $($_.Exception.Message)" }
    }

    # Standaard afdrukvoorkeuren (dubbelzijdig, kleur, papier, lade) in een PrintTicket
    if ($Queue.Duplex -or $Queue.Color -or $Queue.PaperSize -or $Queue.Tray) {
        Write-Host "    Afdrukvoorkeuren instellen (gewenst: Dubbelzijdig=$($Queue.Duplex), Kleur=$($Queue.Color), Papier=$($Queue.PaperSize), Lade=$($Queue.Tray))"
        try { Set-PrinterDefaults -Queue $Queue }
        catch { Write-Warning "Afdrukvoorkeuren instellen mislukt: $($_.Exception.Message)" }
        Clear-UserPrintPreferences -PrinterName $name
    }
}

$exitCode = 0
try {
    Write-Host "=== Installatie $PackageName (versie $($Config.Version)) ==="
    Install-Driver
    Install-Port
    foreach ($queue in @($Config.Queues)) { Install-Queue -Queue $queue }

    $regKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').CreateSubKey("__REG_BASE__\$PackageName")
    $regKey.SetValue('Version', [string]$Config.Version)
    $regKey.SetValue('InstallDate', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $regKey.Close()

    Write-Host "=== Installatie geslaagd ==="
}
catch {
    Write-Error "Installatie mislukt: $($_.Exception.Message)"
    $exitCode = 1
}
finally {
    Stop-Transcript | Out-Null
}
exit $exitCode
'@

$script:UninstallTemplate = @'
<#
    Printer package : __PACKAGE_NAME__
    Versie          : __VERSION__
    Gegenereerd     : __DATE__ (New-PrinterIntunePackage.ps1)

    Uninstall command : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1
    Log               : %ProgramData%\Microsoft\IntuneManagementExtension\Logs\__PACKAGE_NAME__-uninstall.log
#>

__CONFIG_BLOCK__

if ($env:PROCESSOR_ARCHITEW6432 -eq 'AMD64') {
    $ps64 = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $ps64) {
        & $ps64 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
        exit $LASTEXITCODE
    }
}

$PackageName = $Config.PackageName
$LogFolder   = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
Start-Transcript -Path (Join-Path $LogFolder "$PackageName-uninstall.log") -Append | Out-Null

$exitCode = 0
try {
    Write-Host "=== Verwijderen $PackageName ==="

    foreach ($queue in @($Config.Queues)) {
        if (Get-Printer -Name $queue.Name -ErrorAction SilentlyContinue) {
            Remove-Printer -Name $queue.Name -Confirm:$false -ErrorAction Stop
            Write-Host "Printer '$($queue.Name)' verwijderd."
        }
        else {
            Write-Host "Printer '$($queue.Name)' niet gevonden."
        }
    }

    # Poort verwijderen als geen andere printer hem nog gebruikt (spooler kan hem even vasthouden)
    for ($i = 1; $i -le 5; $i++) {
        if (-not (Get-PrinterPort -Name $Config.PortName -ErrorAction SilentlyContinue)) { break }
        if (Get-Printer | Where-Object { $_.PortName -eq $Config.PortName }) {
            Write-Host "Poort '$($Config.PortName)' is nog in gebruik door een andere printer; blijft staan."
            break
        }
        try {
            Remove-PrinterPort -Name $Config.PortName -ErrorAction Stop
            Write-Host "Poort '$($Config.PortName)' verwijderd."
            break
        }
        catch {
            Write-Host "Poort verwijderen poging $i mislukt: $($_.Exception.Message)"
            Start-Sleep -Seconds 5
        }
    }

    if ($Config.RemoveDriverOnUninstall) {
        if (Get-Printer | Where-Object { $_.DriverName -eq $Config.DriverName }) {
            Write-Host "Driver '$($Config.DriverName)' is nog in gebruik; blijft staan."
        }
        else {
            try {
                Remove-PrinterDriver -Name $Config.DriverName -ErrorAction Stop
                Write-Host "Driver '$($Config.DriverName)' verwijderd."
            }
            catch { Write-Warning "Driver verwijderen mislukt: $($_.Exception.Message)" }
        }
    }

    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64')
    $hklm.DeleteSubKeyTree("__REG_BASE__\$PackageName", $false)
    $hklm.Close()

    Write-Host "=== Verwijderen geslaagd ==="
}
catch {
    Write-Error "Verwijderen mislukt: $($_.Exception.Message)"
    $exitCode = 1
}
finally {
    Stop-Transcript | Out-Null
}
exit $exitCode
'@

$script:DetectTemplate = @'
<#
    Detectiescript voor printer package : __PACKAGE_NAME__ (versie __VERSION__)
    Intune: "Use a custom detection script", Run script as 32-bit process: No, Enforce signature check: No
#>

__CONFIG_BLOCK__

$detected = $true
foreach ($queue in @($Config.Queues)) {
    $printer = Get-Printer -Name $queue.Name -ErrorAction SilentlyContinue
    if (-not $printer -or $printer.DriverName -ne $Config.DriverName -or $printer.PortName -ne $Config.PortName) {
        $detected = $false
    }
}

$regKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey("__REG_BASE__\$($Config.PackageName)")
if (-not $regKey -or [string]$regKey.GetValue('Version') -ne [string]$Config.Version) { $detected = $false }
if ($regKey) { $regKey.Close() }

if ($detected) {
    Write-Output "Printer package $($Config.PackageName) $($Config.Version) gedetecteerd"
    exit 0
}
exit 1
'@

#endregion

# Driverversie-functies uit de install-template ook in de tool gebruiken (een bron)
$helperStart = $script:InstallTemplate.IndexOf('function Get-InfDriverVersion')
$helperEnd   = $script:InstallTemplate.IndexOf('function Install-Driver')
. ([scriptblock]::Create($script:InstallTemplate.Substring($helperStart, $helperEnd - $helperStart)))

#region XAML

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Printer Intune Packager" Width="1040" Height="900" MinWidth="900" MinHeight="720"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13" Background="#F3F4F6">
    <Window.Resources>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="5,4"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Padding" Value="6,5"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,6"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="12,5"/>
            <Setter Property="Margin" Value="8,4,0,4"/>
            <Setter Property="MinWidth" Value="90"/>
        </Style>
        <Style TargetType="TextBlock" x:Key="FieldLabel">
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,0,12,0"/>
        </Style>
        <Style TargetType="TextBlock" x:Key="Hint">
            <Setter Property="Foreground" Value="#5F6670"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,2,0,6"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,12"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="Button" x:Key="Primary" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#0F6CBD"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#0F6CBD"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="18,7"/>
        </Style>
    </Window.Resources>

    <DockPanel>
        <Border DockPanel.Dock="Top" Background="#0F6CBD" Padding="20,14">
            <Grid>
                <StackPanel>
                    <TextBlock Text="Printer naar Intune package" Foreground="White" FontSize="21" FontWeight="SemiBold"/>
                    <TextBlock Text="Vul de printergegevens in, kies de driver, stel wachtrijen en lades in en maak het .intunewin-pakket." Foreground="#DCEBFA" Margin="0,3,0,0"/>
                </StackPanel>
                <TextBlock x:Name="txtAdminStatus" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="White" FontSize="12"/>
            </Grid>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#D9DCE1" BorderThickness="0,1,0,0" Padding="16,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="140"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" LastChildFill="False">
                    <Button x:Name="btnLoad" Content="Configuratie laden..." DockPanel.Dock="Left" Margin="0,4,0,4"/>
                    <Button x:Name="btnSave" Content="Configuratie opslaan..." DockPanel.Dock="Left"/>
                    <Button x:Name="btnBuild" Content="Package maken..." DockPanel.Dock="Right" Style="{StaticResource Primary}"/>
                </DockPanel>
                <TextBox x:Name="txtLog" Grid.Row="1" Margin="0,8,0,0" IsReadOnly="True" TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto" VerticalContentAlignment="Top"
                         FontFamily="Consolas" FontSize="12" Background="#FAFBFC"/>
            </Grid>
        </Border>

        <TabControl x:Name="tabs" Margin="14,12,14,10">
            <!-- ===== Tab 1 ===== -->
            <TabItem Header="  1. Printer &amp; driver  ">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="10">
                        <GroupBox Header="Printer en poort">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                                </Grid.RowDefinitions>

                                <TextBlock Grid.Row="0" Text="IP-adres / hostnaam" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtAddress"/>
                                <Button Grid.Row="0" Grid.Column="2" x:Name="btnPing" Content="Test verbinding"/>

                                <TextBlock Grid.Row="1" Text="Poortnaam" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtPortName" ToolTip="Wordt automatisch IP_&lt;adres&gt;; mag aangepast worden"/>

                                <TextBlock Grid.Row="2" Text="Protocol" Style="{StaticResource FieldLabel}"/>
                                <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal">
                                    <ComboBox x:Name="cmbProtocol" Width="160" SelectedIndex="0">
                                        <ComboBoxItem Content="RAW (TCP/IP)"/>
                                        <ComboBoxItem Content="LPR"/>
                                    </ComboBox>
                                    <TextBlock Text="Poortnummer" Margin="20,0,10,0" VerticalAlignment="Center"/>
                                    <TextBox x:Name="txtPortNumber" Width="80" Text="9100"/>
                                </StackPanel>

                                <TextBlock Grid.Row="3" Text="LPR-wachtrijnaam" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="3" Grid.Column="1" x:Name="txtLprQueue" IsEnabled="False"/>

                                <TextBlock Grid.Row="4" Text="SNMP" Style="{StaticResource FieldLabel}"/>
                                <StackPanel Grid.Row="4" Grid.Column="1" Orientation="Horizontal">
                                    <CheckBox x:Name="chkSnmp" Content="SNMP-status ingeschakeld"/>
                                    <TextBlock Text="Community" Margin="20,0,10,0" VerticalAlignment="Center"/>
                                    <TextBox x:Name="txtSnmpCommunity" Width="140" Text="public" IsEnabled="False"/>
                                </StackPanel>

                                <TextBlock Grid.Row="5" Grid.Column="1" Grid.ColumnSpan="2" Style="{StaticResource Hint}"
                                           Text="Gebruik bij voorkeur een vast IP-adres of DNS-naam. Bestaat de poort al op een client, dan wordt het adres bijgewerkt."/>
                            </Grid>
                        </GroupBox>

                        <GroupBox Header="Driver">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                                </Grid.RowDefinitions>

                                <TextBlock Grid.Row="0" Text="Drivermap" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtDriverFolder"/>
                                <Button Grid.Row="0" Grid.Column="2" x:Name="btnBrowseDriver" Content="Bladeren..."/>

                                <TextBlock Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" x:Name="txtDriverInfo" Style="{StaticResource Hint}"
                                           Text="Kies de map met uitgepakte driverbestanden (.inf, .cat, .cab ...). Pak een .exe/.zip-driver eerst uit."/>

                                <TextBlock Grid.Row="2" Text="Printermodel zoeken" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtModelSearch"
                                         ToolTip="Typ (een deel van) het model, bijv. 'PA2600cx' of 'ECOSYS M2540'. Alle woorden moeten voorkomen."/>
                                <Button Grid.Row="2" Grid.Column="2" x:Name="btnClearSearch" Content="Wissen"/>

                                <ListBox Grid.Row="3" Grid.Column="1" x:Name="lstModels" Height="170" Margin="0,2,0,2" DisplayMemberPath="Display"
                                         ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>

                                <TextBlock Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2" x:Name="txtModelCount" Style="{StaticResource Hint}"
                                           Text="Alle printermodellen uit alle INF-bestanden in de drivermap. Klik op een model om INF-bestand en drivernaam in te vullen."/>

                                <TextBlock Grid.Row="5" Text="INF-bestand" Style="{StaticResource FieldLabel}"/>
                                <ComboBox Grid.Row="5" Grid.Column="1" x:Name="cmbInf"/>

                                <TextBlock Grid.Row="6" Text="Drivernaam (model)" Style="{StaticResource FieldLabel}"/>
                                <ComboBox Grid.Row="6" Grid.Column="1" x:Name="cmbDriverName" IsEditable="True"/>

                                <TextBlock Grid.Row="7" Grid.Column="1" Grid.ColumnSpan="2" Style="{StaticResource Hint}"
                                           Text="De drivernaam moet exact overeenkomen met het model in het INF-bestand, bijv. 'Kyocera ECOSYS PA2600cx KX' of 'HP Universal Printing PCL 6'."/>

                                <CheckBox Grid.Row="8" Grid.Column="1" x:Name="chkStageAll" Content="Alle INF-bestanden uit de drivermap toevoegen aan de driver store (nodig bij sommige drivers met afhankelijkheden)"/>
                                <CheckBox Grid.Row="9" Grid.Column="1" x:Name="chkRemoveDriver" Content="Driver verwijderen bij uninstall (alleen als geen andere printer hem gebruikt)"/>
                            </Grid>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- ===== Tab 2 ===== -->
            <TabItem Header="  2. Wachtrijen &amp; lades  ">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="1.3*"/>
                    </Grid.RowDefinitions>

                    <GroupBox Grid.Row="0" Header="Lades en opties uit de driver ophalen">
                        <StackPanel>
                            <DockPanel>
                                <Button x:Name="btnReadOptions" Content="Lades/opties uitlezen" DockPanel.Dock="Left" Margin="0,4,12,4"/>
                                <TextBlock x:Name="txtReadStatus" VerticalAlignment="Center" TextWrapping="Wrap" Foreground="#5F6670"
                                           Text="Nog niet uitgelezen. Zonder uitlezen zijn alleen de standaard Print Schema-lades beschikbaar."/>
                            </DockPanel>
                            <TextBlock Style="{StaticResource Hint}"
                                       Text="Installeert de gekozen driver tijdelijk op deze pc (administrator nodig), maakt een tijdelijke printer aan en leest de beschikbare lades en installeerbare opties uit. Zet je hieronder een optie aan (bijv. extra lade = Installed), lees dan opnieuw uit zodat die lade ook als standaardlade gekozen kan worden."/>
                        </StackPanel>
                    </GroupBox>

                    <GroupBox Grid.Row="1" Header="Installeerbare opties (apparaatinstellingen, gelden voor alle wachtrijen)">
                        <DockPanel>
                            <StackPanel DockPanel.Dock="Right" Margin="8,0,0,0">
                                <Button x:Name="btnAddOption" Content="Toevoegen" Margin="0,0,0,4"/>
                                <Button x:Name="btnRemoveOption" Content="Verwijderen" Margin="0,0,0,4"/>
                            </StackPanel>
                            <TextBlock DockPanel.Dock="Bottom" Style="{StaticResource Hint}"
                                       Text="Alleen opties waarvan de gewenste waarde afwijkt van de driverstandaard worden bij installatie ingesteld (Set-PrinterProperty). Voorbeeld: Config:InstalledTray3 = Installed."/>
                            <DataGrid x:Name="dgOptions" AutoGenerateColumns="False" CanUserAddRows="True" HeadersVisibility="Column"
                                      SelectionMode="Single" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E5E7EB" MinHeight="90">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Eigenschap" Width="2*" Binding="{Binding Property, UpdateSourceTrigger=PropertyChanged}"/>
                                    <DataGridTextColumn Header="Driverstandaard" Width="*" IsReadOnly="True" Binding="{Binding Default}"/>
                                    <DataGridTextColumn Header="Gewenste waarde" Width="*" Binding="{Binding Value, UpdateSourceTrigger=PropertyChanged}"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </DockPanel>
                    </GroupBox>

                    <GroupBox Grid.Row="2" Header="Printerwachtrijen (bijv. dezelfde printer als 'Blanco' en 'Briefpapier' met elk een eigen lade)">
                        <DockPanel>
                            <StackPanel DockPanel.Dock="Right" Margin="8,0,0,0">
                                <Button x:Name="btnAddQueue" Content="Toevoegen" Margin="0,0,0,4"/>
                                <Button x:Name="btnRemoveQueue" Content="Verwijderen" Margin="0,0,0,4"/>
                            </StackPanel>
                            <DataGrid x:Name="dgQueues" AutoGenerateColumns="False" CanUserAddRows="False" HeadersVisibility="Column"
                                      SelectionMode="Single" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E5E7EB" MinHeight="110">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Printernaam" Width="1.4*" Binding="{Binding Name, UpdateSourceTrigger=PropertyChanged}"/>
                                    <DataGridTextColumn Header="Locatie" Width="*" Binding="{Binding Location, UpdateSourceTrigger=PropertyChanged}"/>
                                    <DataGridTextColumn Header="Opmerking" Width="*" Binding="{Binding Comment, UpdateSourceTrigger=PropertyChanged}"/>
                                    <DataGridComboBoxColumn Header="Standaardlade" Width="1.6*" SelectedValuePath="Value" DisplayMemberPath="Display"
                                                            SelectedValueBinding="{Binding Tray, UpdateSourceTrigger=PropertyChanged}"/>
                                    <DataGridComboBoxColumn Header="Dubbelzijdig" Width="1.2*" SelectedValuePath="Value" DisplayMemberPath="Display"
                                                            SelectedValueBinding="{Binding Duplex, UpdateSourceTrigger=PropertyChanged}"/>
                                    <DataGridComboBoxColumn Header="Kleur" Width="*" SelectedValuePath="Value" DisplayMemberPath="Display"
                                                            SelectedValueBinding="{Binding Color, UpdateSourceTrigger=PropertyChanged}"/>
                                    <DataGridComboBoxColumn Header="Papier" Width="*" SelectedValuePath="Value" DisplayMemberPath="Display"
                                                            SelectedValueBinding="{Binding PaperSize, UpdateSourceTrigger=PropertyChanged}"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </DockPanel>
                    </GroupBox>
                </Grid>
            </TabItem>

            <!-- ===== Tab 3 ===== -->
            <TabItem Header="  3. Package  ">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="10">
                        <GroupBox Header="Intune app">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="170"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                                </Grid.RowDefinitions>

                                <TextBlock Grid.Row="0" Text="Packagenaam" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtPackageName" ToolTip="Leeg = 'Printer - &lt;eerste printernaam&gt;'"/>

                                <TextBlock Grid.Row="1" Text="Versie" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtVersion" Text="1.0" Width="120" HorizontalAlignment="Left"/>

                                <TextBlock Grid.Row="2" Text="Publisher" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtPublisher" Text="Easy Office Online"/>

                                <TextBlock Grid.Row="3" Text="Logo (optioneel)" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="3" Grid.Column="1" x:Name="txtLogo"/>
                                <Button Grid.Row="3" Grid.Column="2" x:Name="btnBrowseLogo" Content="Bladeren..."/>

                                <TextBlock Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2" Style="{StaticResource Hint}"
                                           Text="Leeg = er wordt automatisch een printer-icoon gemaakt als &lt;Packagenaam&gt;_logo.png."/>

                                <TextBlock Grid.Row="5" Text="IntuneWinAppUtil.exe" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="5" Grid.Column="1" x:Name="txtIntuneWin"/>
                                <Button Grid.Row="5" Grid.Column="2" x:Name="btnBrowseIntuneWin" Content="Bladeren..."/>

                                <TextBlock Grid.Row="6" Grid.Column="1" Grid.ColumnSpan="2" Style="{StaticResource Hint}"
                                           Text="Wordt automatisch gezocht in 'Win32 Content Prep Tool'. Niet gevonden? Dan wordt aangeboden hem van GitHub (Microsoft) te downloaden."/>
                            </Grid>
                        </GroupBox>

                        <GroupBox Header="Wat wordt er gemaakt?">
                            <TextBlock Style="{StaticResource Hint}" Margin="0,4">
                                Na 'Package maken...' kies je de map waarin het pakket wordt opgeslagen. Daarin komt een map &lt;Packagenaam&gt; met:<LineBreak/>
                                - Source\install.ps1, uninstall.ps1 en Drivers\ (de inhoud van het .intunewin-bestand)<LineBreak/>
                                - Output\install.intunewin (uploaden in Intune)<LineBreak/>
                                - detect.ps1 (custom detectiescript), info.txt (alle Intune-instellingen), logo en PrinterPackage.json (om later weer te laden)
                            </TextBlock>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
    </DockPanel>
</Window>
'@

#endregion

#region Helpers

function Test-IsAdmin {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Update-UI {
    $script:Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Write-Log {
    param([string]$Message)
    $line = '[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message
    $ui.txtLog.AppendText($line + [Environment]::NewLine)
    $ui.txtLog.ScrollToEnd()
    Update-UI
}

function Show-Message {
    param(
        [string]$Text,
        [string]$Icon = 'Information',
        [string]$Buttons = 'OK'
    )
    return [System.Windows.MessageBox]::Show($script:Window, $Text, 'Printer Intune Packager', $Buttons, $Icon)
}

function Invoke-Safe {
    param([scriptblock]$Action)
    try {
        & $Action
    }
    catch {
        Write-Log "FOUT: $($_.Exception.Message)"
        Show-Message -Text $_.Exception.Message -Icon Error | Out-Null
    }
    finally {
        Set-Busy $false
    }
}

function Set-Busy {
    param([bool]$Busy)
    $script:Window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
    foreach ($name in 'btnBuild', 'btnReadOptions', 'btnLoad', 'btnSave', 'btnPing') { $ui[$name].IsEnabled = -not $Busy }
    Update-UI
}

function Invoke-ExternalProcess {
    param([string]$FilePath, [string]$Arguments)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $process = [System.Diagnostics.Process]::Start($psi)
    $stdOut = $process.StandardOutput.ReadToEndAsync()
    $stdErr = $process.StandardError.ReadToEndAsync()
    while (-not $process.HasExited) {
        Update-UI
        Start-Sleep -Milliseconds 150
    }
    $process.WaitForExit()
    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        Output   = $stdOut.Result
        Error    = $stdErr.Result
    }
}

function Get-SafeName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $clean = -join ($Name.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '-' } else { $_ } })
    return $clean.Trim().TrimEnd('.')
}

function Save-TextFile {
    param([string]$Path, [string]$Content)
    [IO.File]::WriteAllText($Path, $Content, $script:Utf8Bom)
}

function Confirm-GridEdits {
    foreach ($grid in $ui.dgOptions, $ui.dgQueues) {
        [void]$grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
    }
}

function Find-IntuneWinAppUtil {
    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add((Join-Path $script:ToolDir 'IntuneWinAppUtil.exe'))
    $dir = $script:ToolDir
    while ($dir) {
        $candidates.Add((Join-Path $dir 'Win32 Content Prep Tool\IntuneWinAppUtil.exe'))
        $parent = Split-Path $dir -Parent
        if ($parent -eq $dir) { break }
        $dir = $parent
    }
    return $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}

#endregion

#region INF parsing

function Remove-InfComment {
    param([string]$Line)
    $semicolon = $Line.IndexOf(';')
    if ($semicolon -lt 0) { return $Line }
    $quote = $Line.IndexOf('"')
    if ($quote -lt 0 -or $quote -gt $semicolon) { return $Line.Substring(0, $semicolon) }
    $inQuotes = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $c = $Line[$i]
        if ($c -eq '"') { $inQuotes = -not $inQuotes }
        elseif ($c -eq ';' -and -not $inQuotes) { return $Line.Substring(0, $i) }
    }
    return $Line
}

function Get-InfSections {
    param([string]$Path)
    $sections = @{}
    $current = $null
    foreach ($raw in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $line = (Remove-InfComment $raw).Trim()
        if (-not $line) { continue }
        if ($line -match '^\[(.+?)\]$') {
            $current = $Matches[1].Trim()
            if (-not $sections.ContainsKey($current)) { $sections[$current] = New-Object System.Collections.Generic.List[string] }
            continue
        }
        if ($current) { $sections[$current].Add($line) }
    }
    return $sections
}

function ConvertFrom-InfValue {
    param([string]$Value, [hashtable]$Strings)
    $v = $Value.Trim()
    if ($v -match '^%(.+)%$' -and $Strings.ContainsKey($Matches[1])) { $v = $Strings[$Matches[1]] }
    if ($v.Length -ge 2 -and $v.StartsWith('"') -and $v.EndsWith('"')) { $v = $v.Substring(1, $v.Length - 2) }
    return $v.Replace('""', '"').Trim()
}

function Get-InfDriverNames {
    param([string]$Path)

    $sections = Get-InfSections -Path $Path

    # [Strings] eerst, daarna taalvarianten (Strings.0409 ...) voor ontbrekende sleutels
    $strings = @{}
    $stringSections = @($sections.Keys | Where-Object { $_ -match '^Strings(\..+)?$' } | Sort-Object { if ($_ -ieq 'Strings') { 0 } else { 1 } })
    foreach ($s in $stringSections) {
        foreach ($line in $sections[$s]) {
            if ($line -match '^\s*([^=]+?)\s*=\s*(.*)$' -and -not $strings.ContainsKey($Matches[1])) {
                $key = $Matches[1]
                $strings[$key] = ConvertFrom-InfValue -Value $Matches[2] -Strings @{}
            }
        }
    }

    $modelSections = New-Object System.Collections.Generic.List[string]
    if ($sections.ContainsKey('Manufacturer')) {
        foreach ($line in $sections['Manufacturer']) {
            $right = if ($line -match '=(.*)$') { $Matches[1] } else { $line }
            $parts = @($right.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if (-not $parts) { continue }
            $base = $parts[0]
            $decorations = @($parts | Select-Object -Skip 1)
            $amd64 = @($decorations | Where-Object { $_ -match 'amd64' })
            if ($amd64) { $amd64 | ForEach-Object { $modelSections.Add("$base.$_") } }
            else {
                # Geen 64-bit sectie: alleen ongedecoreerde/neutrale secties (geen x86/ARM-only INF)
                $modelSections.Add($base)
                $decorations | Where-Object { $_ -notmatch 'x86|arm|ia64' } | ForEach-Object { $modelSections.Add("$base.$_") }
            }
        }
    }

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($section in $modelSections) {
        if (-not $sections.ContainsKey($section)) { continue }
        foreach ($line in $sections[$section]) {
            if ($line -match '^\s*(.+?)\s*=') {
                $name = ConvertFrom-InfValue -Value $Matches[1] -Strings $strings
                if ($name -and -not $names.Contains($name)) { $names.Add($name) }
            }
        }
    }
    return @($names | Sort-Object)
}

function Test-InfIsPrinter {
    param([string]$Path)
    try {
        $head = (Get-Content -LiteralPath $Path -TotalCount 80 -ErrorAction Stop) -join "`n"
        return $head -match '(?im)^\s*Class\s*=\s*"?Printer"?\s*(;.*)?$'
    }
    catch { return $false }
}

function Update-InfList {
    $folder = $ui.txtDriverFolder.Text.Trim().TrimEnd('\')
    $ui.cmbInf.Items.Clear()
    $ui.cmbDriverName.Items.Clear()
    $script:ModelIndex.Clear()
    Update-ModelList
    if (-not $folder -or -not (Test-Path -LiteralPath $folder -PathType Container)) {
        $ui.txtDriverInfo.Text = 'Kies de map met uitgepakte driverbestanden (.inf, .cat, .cab ...). Pak een .exe/.zip-driver eerst uit.'
        return
    }

    Set-Busy $true
    Write-Log "Drivermap scannen: $folder"
    $all = @(Get-ChildItem -LiteralPath $folder -Filter '*.inf' -Recurse -File -ErrorAction SilentlyContinue)
    $printerInfs = @($all | Where-Object { Test-InfIsPrinter $_.FullName })
    $list = if ($printerInfs.Count) { $printerInfs } else { $all }
    # 64-bit INF's (bijv. Kyocera ...\64bit\OEMSETUP.INF) bovenaan
    $list = @($list | Sort-Object @{ Expression = { $_.FullName -notmatch '(?<!arm)64' } }, FullName)
    foreach ($file in $list) { [void]$ui.cmbInf.Items.Add($file.FullName.Substring($folder.Length + 1)) }

    # Modellen uit alle INF's lezen voor de model-zoeker
    $script:ModelIndex.Clear()
    $seenModels = @{}
    foreach ($file in $list) {
        $rel = $file.FullName.Substring($folder.Length + 1)
        try {
            foreach ($name in (Get-InfDriverNames -Path $file.FullName)) {
                # Elk model een keer tonen; de 64-bit INF staat vooraan en wint
                if ($seenModels.ContainsKey($name)) { continue }
                $seenModels[$name] = $true
                $script:ModelIndex.Add((New-Object EooPrinterPkg.Choice("$name    [$rel]", "$rel`t$name")))
            }
        }
        catch { Write-Log "INF lezen mislukt ($rel): $($_.Exception.Message)" }
        Update-UI
    }
    Write-Log ('{0} printermodel(len) gevonden in {1} INF-bestand(en).' -f $script:ModelIndex.Count, $list.Count)
    Update-ModelList

    $ui.txtDriverInfo.Text = '{0} INF-bestand(en) gevonden, waarvan {1} printerdriver(s). Map: {2:N1} MB' -f $all.Count, $printerInfs.Count,
        ((Get-ChildItem -LiteralPath $folder -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1MB)
    if ($all.Count -eq 0) { Write-Log 'Geen INF-bestanden gevonden in deze map.' }
    if ($ui.cmbInf.Items.Count) { $ui.cmbInf.SelectedIndex = 0 }
    Set-Busy $false
}

function Update-ModelList {
    $words = @($ui.txtModelSearch.Text.Trim() -split '\s+' | Where-Object { $_ })
    $found = @($script:ModelIndex | Where-Object {
            $display = $_.Display
            -not ($words | Where-Object { $display -notlike "*$_*" })
        })
    $ui.lstModels.ItemsSource = $found
    if ($script:ModelIndex.Count -eq 0) {
        $ui.txtModelCount.Text = 'Nog geen modellen: kies eerst een drivermap met .inf-bestanden.'
    }
    else {
        $ui.txtModelCount.Text = '{0} van {1} modellen. Klik op een model om INF-bestand en drivernaam in te vullen.' -f $found.Count, $script:ModelIndex.Count
    }
    # Precies een treffer (bijv. geplakte volledige drivernaam): direct selecteren
    if ($words.Count -and $found.Count -eq 1) { $ui.lstModels.SelectedIndex = 0 }
}

function Select-Model {
    $choice = $ui.lstModels.SelectedItem
    if (-not $choice) { return }
    $rel, $name = $choice.Value -split "`t", 2
    if ([string]$ui.cmbInf.SelectedItem -ne $rel) { $ui.cmbInf.SelectedItem = $rel }
    $ui.cmbDriverName.Text = $name
    Write-Log "Model gekozen: $name ($rel)"
}

function Update-DriverNames {
    $previous = $ui.cmbDriverName.Text
    $ui.cmbDriverName.Items.Clear()
    $rel = [string]$ui.cmbInf.SelectedItem
    if (-not $rel) { return }
    $path = Join-Path $ui.txtDriverFolder.Text.Trim() $rel
    try {
        $names = Get-InfDriverNames -Path $path
        foreach ($n in $names) { [void]$ui.cmbDriverName.Items.Add($n) }
        Write-Log ("{0}: {1} drivernaam/-namen gevonden." -f $rel, $names.Count)
    }
    catch {
        Write-Log "INF lezen mislukt: $($_.Exception.Message)"
    }
    if ($previous -and $ui.cmbDriverName.Items.Contains($previous)) { $ui.cmbDriverName.Text = $previous }
    elseif ($ui.cmbDriverName.Items.Count) { $ui.cmbDriverName.SelectedIndex = 0 }
}

#endregion

#region Choices and state

function Reset-TrayChoices {
    $script:TrayChoices.Clear()
    $script:TrayChoices.Add((New-Object EooPrinterPkg.Choice('(driverstandaard)', '')))
    foreach ($name in 'AutoSelect', 'Manual', 'Cassette', 'AutoSheetFeeder', 'Tractor') {
        $script:TrayChoices.Add((New-Object EooPrinterPkg.Choice("$name  [psk]", "{$script:PskUri}$name")))
    }
}

function Add-TrayChoiceIfMissing {
    param([string]$Value, [string]$Display)
    if (-not $Value) { return }
    if (-not ($script:TrayChoices | Where-Object { $_.Value -eq $Value })) {
        if (-not $Display) { $Display = ($Value -replace '^\{[^}]*\}', '') }
        $script:TrayChoices.Add((New-Object EooPrinterPkg.Choice($Display, $Value)))
    }
}

function Get-TrayDisplay {
    param([string]$Value)
    if (-not $Value) { return '(driverstandaard)' }
    $choice = $script:TrayChoices | Where-Object { $_.Value -eq $Value } | Select-Object -First 1
    if ($choice) { return $choice.Display }
    return $Value
}

function Get-ChoiceDisplay {
    param($Choices, [string]$Value)
    $choice = $Choices | Where-Object { $_.Value -eq $Value } | Select-Object -First 1
    if ($choice) { return $choice.Display }
    return $Value
}

function Add-QueueRow {
    param([string]$Name = '')
    $row = New-Object EooPrinterPkg.QueueRow
    $row.Name = $Name
    $script:Queues.Add($row)
    return $row
}

function Get-GuiState {
    Confirm-GridEdits
    return [ordered]@{
        PrinterAddress          = $ui.txtAddress.Text.Trim()
        PortName                = $ui.txtPortName.Text.Trim()
        Protocol                = if ($ui.cmbProtocol.SelectedIndex -eq 1) { 'LPR' } else { 'RAW' }
        PortNumber              = $ui.txtPortNumber.Text.Trim()
        LprQueue                = $ui.txtLprQueue.Text.Trim()
        SnmpEnabled             = [bool]$ui.chkSnmp.IsChecked
        SnmpCommunity           = $ui.txtSnmpCommunity.Text.Trim()
        DriverFolder            = $ui.txtDriverFolder.Text.Trim().TrimEnd('\')
        InfPath                 = [string]$ui.cmbInf.SelectedItem
        DriverName              = $ui.cmbDriverName.Text.Trim()
        StageAllInfs            = [bool]$ui.chkStageAll.IsChecked
        RemoveDriverOnUninstall = [bool]$ui.chkRemoveDriver.IsChecked
        DeviceOptions           = @($script:DeviceOptions | Where-Object { "$($_.Property)".Trim() } | ForEach-Object {
                [ordered]@{ Property = "$($_.Property)".Trim(); Default = "$($_.Default)"; Value = "$($_.Value)".Trim() } })
        Queues                  = @($script:Queues | ForEach-Object {
                [ordered]@{
                    Name = "$($_.Name)".Trim(); Location = "$($_.Location)".Trim(); Comment = "$($_.Comment)".Trim()
                    Tray = "$($_.Tray)"; Duplex = "$($_.Duplex)"; Color = "$($_.Color)"; PaperSize = "$($_.PaperSize)"
                } })
        TrayChoices             = @($script:TrayChoices | Where-Object { $_.Value } | ForEach-Object { [ordered]@{ Display = $_.Display; Value = $_.Value } })
        PackageName             = $ui.txtPackageName.Text.Trim()
        Version                 = $ui.txtVersion.Text.Trim()
        Publisher               = $ui.txtPublisher.Text.Trim()
        LogoPath                = $ui.txtLogo.Text.Trim()
        IntuneWinAppUtil        = $ui.txtIntuneWin.Text.Trim()
    }
}

function Set-GuiState {
    param($State)
    $ui.txtAddress.Text       = "$($State.PrinterAddress)"
    $ui.txtPortName.Text      = "$($State.PortName)"
    $script:LastAutoPort      = ''
    $ui.cmbProtocol.SelectedIndex = if ($State.Protocol -eq 'LPR') { 1 } else { 0 }
    if ($State.PortNumber) { $ui.txtPortNumber.Text = "$($State.PortNumber)" }
    $ui.txtLprQueue.Text      = "$($State.LprQueue)"
    $ui.chkSnmp.IsChecked     = [bool]$State.SnmpEnabled
    if ($State.SnmpCommunity) { $ui.txtSnmpCommunity.Text = "$($State.SnmpCommunity)" }
    $ui.chkStageAll.IsChecked     = [bool]$State.StageAllInfs
    $ui.chkRemoveDriver.IsChecked = [bool]$State.RemoveDriverOnUninstall

    $ui.txtDriverFolder.Text = "$($State.DriverFolder)"
    $script:LastDriverFolder = $ui.txtDriverFolder.Text
    Update-InfList
    if ($State.InfPath -and $ui.cmbInf.Items.Contains("$($State.InfPath)")) { $ui.cmbInf.SelectedItem = "$($State.InfPath)" }
    if ($State.DriverName) { $ui.cmbDriverName.Text = "$($State.DriverName)" }

    Reset-TrayChoices
    foreach ($t in @($State.TrayChoices)) { if ($t) { Add-TrayChoiceIfMissing -Value $t.Value -Display $t.Display } }

    $script:DeviceOptions.Clear()
    foreach ($o in @($State.DeviceOptions)) {
        if (-not $o) { continue }
        $row = New-Object EooPrinterPkg.OptionRow
        $row.Property = "$($o.Property)"; $row.Default = "$($o.Default)"; $row.Value = "$($o.Value)"
        $script:DeviceOptions.Add($row)
    }

    $script:Queues.Clear()
    foreach ($q in @($State.Queues)) {
        if (-not $q) { continue }
        $row = Add-QueueRow -Name "$($q.Name)"
        $row.Location = "$($q.Location)"; $row.Comment = "$($q.Comment)"
        $row.Tray = "$($q.Tray)"; $row.Duplex = "$($q.Duplex)"; $row.Color = "$($q.Color)"; $row.PaperSize = "$($q.PaperSize)"
        Add-TrayChoiceIfMissing -Value $row.Tray
    }
    if ($script:Queues.Count -eq 0) { [void](Add-QueueRow) }
    $ui.dgQueues.Items.Refresh()
    $ui.dgOptions.Items.Refresh()

    if ($State.PackageName) { $ui.txtPackageName.Text = "$($State.PackageName)" }
    if ($State.Version)     { $ui.txtVersion.Text = "$($State.Version)" }
    if ($State.Publisher)   { $ui.txtPublisher.Text = "$($State.Publisher)" }
    $ui.txtLogo.Text = "$($State.LogoPath)"
    if ($State.IntuneWinAppUtil -and (Test-Path -LiteralPath "$($State.IntuneWinAppUtil)")) { $ui.txtIntuneWin.Text = "$($State.IntuneWinAppUtil)" }
}

function Import-StateFile {
    param([string]$Path)
    $json = [IO.File]::ReadAllText($Path)
    $state = $json | ConvertFrom-Json
    Set-GuiState -State $state
    Write-Log "Configuratie geladen: $Path"
}

#endregion

#region Actions

function Test-PrinterConnection {
    $address = $ui.txtAddress.Text.Trim()
    if (-not $address) { Show-Message 'Vul eerst een IP-adres of hostnaam in.' -Icon Warning | Out-Null; return }
    Set-Busy $true
    try {
        $reply = (New-Object System.Net.NetworkInformation.Ping).Send($address, 1500)
        Write-Log ("Ping {0}: {1}" -f $address, $reply.Status)
    }
    catch { Write-Log "Ping $address mislukt: $($_.Exception.InnerException.Message)" }

    $port = if ($ui.cmbProtocol.SelectedIndex -eq 1) { 515 } else { [int]("0$($ui.txtPortNumber.Text.Trim())") }
    if ($port -gt 0) {
        $client = New-Object System.Net.Sockets.TcpClient
        try {
            $task = $client.ConnectAsync($address, $port)
            $open = $task.Wait(2000) -and $client.Connected
            Write-Log ("TCP-poort {0} op {1}: {2}" -f $port, $address, $(if ($open) { 'open' } else { 'niet bereikbaar' }))
        }
        catch { Write-Log ("TCP-poort {0} op {1}: niet bereikbaar" -f $port, $address) }
        finally { $client.Dispose() }
    }
}

function Restart-Elevated {
    $tmp = Join-Path $env:TEMP ('PrinterPackager_{0}.json' -f [guid]::NewGuid())
    Save-TextFile -Path $tmp -Content (Get-GuiState | ConvertTo-Json -Depth 6)
    $arguments = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$script:ToolPath`" -LoadConfig `"$tmp`""
    try {
        Start-Process -FilePath (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -ArgumentList $arguments -Verb RunAs -ErrorAction Stop
        $script:Window.Close()
    }
    catch {
        Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
        Write-Log 'Opnieuw starten als administrator geannuleerd.'
    }
}

function Install-LocalDriver {
    param([string]$InfFullPath, [string]$DriverName)

    Write-Log "Driver toevoegen aan driver store (pnputil): $InfFullPath"
    $result = Invoke-ExternalProcess -FilePath (Join-Path $env:WINDIR 'System32\pnputil.exe') -Arguments "/add-driver `"$InfFullPath`""
    Write-Log "pnputil exitcode: $($result.ExitCode)"

    $packageVersion = Get-InfDriverVersion -Path $InfFullPath
    $existing = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
    if ($existing) {
        $installedVersion = ConvertTo-DriverVersion $existing.DriverVersion
        Write-Log "Driver al aanwezig op deze pc: versie $installedVersion (gekozen INF: $packageVersion)."
        if (-not $packageVersion -or ($installedVersion -and $installedVersion -ge $packageVersion)) { return }
        Write-Log 'Oudere versie; bijwerken zodat de lades/opties van de juiste versie worden uitgelezen...'
    }
    else {
        Write-Log "Printerdriver '$DriverName' installeren (kan even duren)..."
    }
    $storeInf = Find-DriverStoreInf -InfName (Split-Path $InfFullPath -Leaf) -Version $packageVersion
    if ($storeInf) { Add-PrinterDriver -Name $DriverName -InfPath $storeInf -ErrorAction Stop }
    else { Add-PrinterDriver -Name $DriverName -ErrorAction Stop }
    Write-Log "Driver geinstalleerd: versie $(ConvertTo-DriverVersion (Get-PrinterDriver -Name $DriverName).DriverVersion)"
}

function Get-TrayChoicesFromCapabilities {
    param([string]$CapabilitiesXml)

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($CapabilitiesXml)
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('psf', $script:PsfUri)

    $features = @($doc.DocumentElement.SelectNodes('psf:Feature', $nsm))
    $feature = $null
    foreach ($wanted in 'PageInputBin', 'JobInputBin', 'DocumentInputBin') {
        $feature = $features | Where-Object { ($_.GetAttribute('name') -split ':')[-1] -eq $wanted } | Select-Object -First 1
        if ($feature) { break }
    }
    if (-not $feature) { return @() }

    $result = @()
    foreach ($option in $feature.SelectNodes('psf:Option', $nsm)) {
        $qname = $option.GetAttribute('name')
        if (-not $qname) { continue }
        if ($qname.Contains(':')) {
            $prefix, $local = $qname -split ':', 2
            $uri = $option.GetNamespaceOfPrefix($prefix)
        }
        else { $local = $qname; $uri = '' }

        $display = $local
        $displayProp = @($option.SelectNodes('psf:Property', $nsm)) | Where-Object { $_.GetAttribute('name') -match '(^|:)DisplayName$' } | Select-Object -First 1
        if ($displayProp) {
            $valueNode = $displayProp.SelectSingleNode('psf:Value', $nsm)
            if ($valueNode -and $valueNode.InnerText) { $display = $valueNode.InnerText }
        }
        $result += New-Object EooPrinterPkg.Choice("$display  [$local]", "{$uri}$local")
    }
    return $result
}

function Get-CapabilityOptions {
    param([string]$CapabilitiesXml, [string[]]$FeatureNames)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($CapabilitiesXml)
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('psf', $script:PsfUri)
    $feature = @($doc.DocumentElement.SelectNodes('psf:Feature', $nsm)) |
        Where-Object { $FeatureNames -contains (($_.GetAttribute('name') -split ':')[-1]) } | Select-Object -First 1
    if (-not $feature) { return '(niet ondersteund door driver)' }
    return (@($feature.SelectNodes('psf:Option', $nsm)) | ForEach-Object { ($_.GetAttribute('name') -split ':')[-1] }) -join ', '
}

function Read-DriverOptions {
    if (-not (Test-IsAdmin)) {
        $answer = Show-Message -Text "Voor het uitlezen van lades en opties moet de driver lokaal worden geinstalleerd. Dat vereist administratorrechten.`n`nTool opnieuw starten als administrator? (Je ingevulde gegevens blijven behouden.)" -Icon Question -Buttons YesNo
        if ($answer -eq 'Yes') { Restart-Elevated }
        return
    }

    $state = Get-GuiState
    if (-not $state.InfPath -or -not $state.DriverName) { throw 'Kies eerst op tabblad 1 de drivermap, het INF-bestand en de drivernaam.' }
    $infFull = Join-Path $state.DriverFolder $state.InfPath
    if (-not (Test-Path -LiteralPath $infFull)) { throw "INF-bestand niet gevonden: $infFull" }

    Set-Busy $true
    $ui.txtReadStatus.Text = 'Bezig met uitlezen...'
    Install-LocalDriver -InfFullPath $infFull -DriverName $state.DriverName

    $tempName = 'EOO-TEMP-OptiesUitlezen'
    $tempPort = $null
    $portName = Get-PrinterPort | Where-Object { $_.Name -in 'PORTPROMPT:', 'nul:', 'FILE:' } | Select-Object -ExpandProperty Name -First 1
    if (-not $portName) {
        $tempPort = 'EOO_TEMP_PORT'
        if (-not (Get-PrinterPort -Name $tempPort -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $tempPort -PrinterHostAddress '127.0.0.1' }
        $portName = $tempPort
    }
    if (Get-Printer -Name $tempName -ErrorAction SilentlyContinue) { Remove-Printer -Name $tempName -Confirm:$false }
    Write-Log "Tijdelijke printer '$tempName' aanmaken..."
    Add-Printer -Name $tempName -DriverName $state.DriverName -PortName $portName -ErrorAction Stop

    try {
        # Eerst de gewenste opties toepassen, zodat bijv. extra lades zichtbaar worden in de capabilities
        foreach ($row in @($script:DeviceOptions | Where-Object { $_.Property -and $_.Value })) {
            try { Set-PrinterProperty -PrinterName $tempName -PropertyName $row.Property -Value $row.Value -ErrorAction Stop }
            catch { Write-Log "Optie $($row.Property) = $($row.Value) toepassen mislukt: $($_.Exception.Message)" }
        }

        $props = @(Get-PrinterProperty -PrinterName $tempName -ErrorAction SilentlyContinue | Where-Object { $_.PropertyName -like 'Config:*' })
        foreach ($p in $props) {
            $existing = $script:DeviceOptions | Where-Object { $_.Property -eq $p.PropertyName } | Select-Object -First 1
            if ($existing) {
                if (-not $existing.Default) { $existing.Default = "$($p.Value)" }
            }
            else {
                $row = New-Object EooPrinterPkg.OptionRow
                $row.Property = $p.PropertyName; $row.Default = "$($p.Value)"; $row.Value = "$($p.Value)"
                $script:DeviceOptions.Add($row)
            }
        }
        $ui.dgOptions.Items.Refresh()

        $printConfig = Get-PrintConfiguration -PrinterName $tempName -ErrorAction Stop
        $trays = @(Get-TrayChoicesFromCapabilities -CapabilitiesXml $printConfig.PrintCapabilitiesXML)

        $selected = @($script:Queues | ForEach-Object { $_.Tray } | Where-Object { $_ })
        $script:TrayChoices.Clear()
        $script:TrayChoices.Add((New-Object EooPrinterPkg.Choice('(driverstandaard)', '')))
        foreach ($t in $trays) { $script:TrayChoices.Add($t) }
        foreach ($s in $selected) { Add-TrayChoiceIfMissing -Value $s }
        $ui.dgQueues.Items.Refresh()

        $msg = '{0} lade(s) en {1} installeerbare optie(s) gevonden ({2:HH:mm}).' -f $trays.Count, $props.Count, (Get-Date)
        $ui.txtReadStatus.Text = $msg
        Write-Log $msg
        foreach ($t in $trays) { Write-Log "    Lade: $($t.Display)" }
        Write-Log "    Dubbelzijdig-opties: $(Get-CapabilityOptions -CapabilitiesXml $printConfig.PrintCapabilitiesXML -FeatureNames 'JobDuplexAllDocumentsContiguously', 'DocumentDuplex')"
        Write-Log "    Kleur-opties: $(Get-CapabilityOptions -CapabilitiesXml $printConfig.PrintCapabilitiesXML -FeatureNames 'PageOutputColor')"
        Write-Log "    Papier-opties (eerste 10): $(((Get-CapabilityOptions -CapabilitiesXml $printConfig.PrintCapabilitiesXML -FeatureNames 'PageMediaSize') -split ', ' | Select-Object -First 10) -join ', ')"
        if ($props.Count -eq 0) { Write-Log '    Deze driver heeft geen Config:-opties (vaak bij v4/class drivers).' }
    }
    finally {
        Remove-Printer -Name $tempName -Confirm:$false -ErrorAction SilentlyContinue
        if ($tempPort) { Remove-PrinterPort -Name $tempPort -ErrorAction SilentlyContinue }
        Write-Log 'Tijdelijke printer verwijderd.'
    }
}

function Test-State {
    param($State)
    $errors = New-Object System.Collections.Generic.List[string]
    if (-not $State.PrinterAddress) { $errors.Add('IP-adres / hostnaam is leeg.') }
    if (-not $State.PortName) { $errors.Add('Poortnaam is leeg.') }
    if ($State.Protocol -eq 'LPR' -and -not $State.LprQueue) { $errors.Add('LPR-wachtrijnaam is leeg.') }
    if ($State.Protocol -eq 'RAW') {
        $n = 0
        if (-not [int]::TryParse($State.PortNumber, [ref]$n) -or $n -lt 1 -or $n -gt 65535) { $errors.Add('Poortnummer is ongeldig.') }
    }
    if (-not $State.DriverFolder -or -not (Test-Path -LiteralPath $State.DriverFolder -PathType Container)) { $errors.Add('Drivermap bestaat niet.') }
    elseif (-not $State.InfPath -or -not (Test-Path -LiteralPath (Join-Path $State.DriverFolder $State.InfPath))) { $errors.Add('Geen geldig INF-bestand gekozen.') }
    if (-not $State.DriverName) { $errors.Add('Drivernaam is leeg.') }

    $queues = @($State.Queues)
    if ($queues.Count -eq 0) { $errors.Add('Voeg minimaal een printerwachtrij toe.') }
    $names = @{}
    foreach ($q in $queues) {
        if (-not $q.Name) { $errors.Add('Een printerwachtrij heeft geen naam.'); continue }
        if ($q.Name -match '[\\,!]') { $errors.Add("Printernaam '$($q.Name)' bevat een ongeldig teken (\ , !).") }
        if ($names.ContainsKey($q.Name)) { $errors.Add("Printernaam '$($q.Name)' komt dubbel voor.") }
        $names[$q.Name] = $true
    }
    if (-not $State.Version) { $errors.Add('Versie is leeg.') }
    if ($State.LogoPath -and -not (Test-Path -LiteralPath $State.LogoPath)) { $errors.Add('Logobestand bestaat niet.') }
    return $errors
}

function Get-PackageConfig {
    param($State, [string]$PackageName)
    return [ordered]@{
        PackageName             = $PackageName
        Version                 = $State.Version
        Publisher               = $State.Publisher
        PrinterAddress          = $State.PrinterAddress
        PortName                = $State.PortName
        Protocol                = $State.Protocol
        PortNumber              = [int]("0$($State.PortNumber)")
        LprQueue                = $State.LprQueue
        SnmpEnabled             = $State.SnmpEnabled
        SnmpCommunity           = $State.SnmpCommunity
        InfPath                 = $State.InfPath
        DriverName              = $State.DriverName
        StageAllInfs            = $State.StageAllInfs
        RemoveDriverOnUninstall = $State.RemoveDriverOnUninstall
        DeviceOptions           = @($State.DeviceOptions | Where-Object { $_.Value -and $_.Value -ne $_.Default } | ForEach-Object {
                [ordered]@{ Property = $_.Property; Value = $_.Value } })
        Queues                  = @($State.Queues | ForEach-Object {
                [ordered]@{
                    Name = $_.Name; Location = $_.Location; Comment = $_.Comment
                    Tray = $_.Tray; Duplex = $_.Duplex; Color = $_.Color; PaperSize = $_.PaperSize
                } })
    }
}

function New-ScriptFromTemplate {
    param([string]$Template, $PackageConfig)
    $json = $PackageConfig | ConvertTo-Json -Depth 6
    $configBlock = "`$Config = @'`r`n$json`r`n'@ | ConvertFrom-Json"
    $content = $Template.Replace('__CONFIG_BLOCK__', $configBlock)
    $content = $content.Replace('__PACKAGE_NAME__', $PackageConfig.PackageName)
    $content = $content.Replace('__VERSION__', $PackageConfig.Version)
    $content = $content.Replace('__DATE__', (Get-Date -Format 'yyyy-MM-dd HH:mm'))
    return $content.Replace('__REG_BASE__', $script:RegBase)
}

function New-PrinterLogo {
    param([string]$Path)
    $size = 256
    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $g.Clear([System.Drawing.Color]::Transparent)

        $radius = 48
        $pathShape = New-Object System.Drawing.Drawing2D.GraphicsPath
        $pathShape.AddArc(0, 0, $radius, $radius, 180, 90)
        $pathShape.AddArc($size - $radius - 1, 0, $radius, $radius, 270, 90)
        $pathShape.AddArc($size - $radius - 1, $size - $radius - 1, $radius, $radius, 0, 90)
        $pathShape.AddArc(0, $size - $radius - 1, $radius, $radius, 90, 90)
        $pathShape.CloseFigure()
        $g.FillPath((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(15, 108, 189))), $pathShape)

        $installedFonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }
        $fontName = @('Segoe Fluent Icons', 'Segoe MDL2 Assets') | Where-Object { $installedFonts -contains $_ } | Select-Object -First 1
        $format = New-Object System.Drawing.StringFormat
        $format.Alignment = [System.Drawing.StringAlignment]::Center
        $format.LineAlignment = [System.Drawing.StringAlignment]::Center
        $rect = New-Object System.Drawing.RectangleF(0, 8, $size, $size)
        if ($fontName) {
            $font = New-Object System.Drawing.Font($fontName, 140, [System.Drawing.GraphicsUnit]::Pixel)
            $g.DrawString([string][char]0xE749, $font, [System.Drawing.Brushes]::White, $rect, $format)
        }
        else {
            $font = New-Object System.Drawing.Font('Segoe UI', 90, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
            $g.DrawString('P', $font, [System.Drawing.Brushes]::White, $rect, $format)
        }
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $g.Dispose()
        $bmp.Dispose()
    }
}

function Save-Logo {
    param([string]$SourcePath, [string]$TargetPath)
    if (-not $SourcePath) {
        New-PrinterLogo -Path $TargetPath
        return 'automatisch gegenereerd printer-icoon'
    }
    if ([IO.Path]::GetExtension($SourcePath) -ieq '.png') {
        Copy-Item -LiteralPath $SourcePath -Destination $TargetPath -Force
    }
    else {
        $img = [System.Drawing.Image]::FromFile($SourcePath)
        try { $img.Save($TargetPath, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $img.Dispose() }
    }
    return "overgenomen van $SourcePath"
}

function Get-IntuneWinAppUtilPath {
    $exe = $ui.txtIntuneWin.Text.Trim()
    if ($exe -and (Test-Path -LiteralPath $exe)) { return $exe }

    $exe = Find-IntuneWinAppUtil
    if ($exe) { $ui.txtIntuneWin.Text = $exe; return $exe }

    $answer = Show-Message -Text "IntuneWinAppUtil.exe is niet gevonden.`n`nDownloaden van GitHub (microsoft/Microsoft-Win32-Content-Prep-Tool) naar de map van deze tool?" -Icon Question -Buttons YesNo
    if ($answer -ne 'Yes') { return $null }
    $target = Join-Path $script:ToolDir 'IntuneWinAppUtil.exe'
    Write-Log 'IntuneWinAppUtil.exe downloaden...'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe' -OutFile $target -UseBasicParsing
    $ui.txtIntuneWin.Text = $target
    return $target
}

function New-InfoText {
    param($State, $PackageConfig, [string]$PackageDir, [string]$LogoFile, [string]$LogoSource, [string]$IntuneWinFile)

    $duplexChoices = $script:DuplexChoices
    $colorChoices  = $script:ColorChoices
    $name    = $PackageConfig.PackageName
    $queues  = @($PackageConfig.Queues)
    $nl      = "`r`n"
    $sb      = New-Object System.Text.StringBuilder

    $description = 'Installeert printer(s) {0} op {1} ({2}) met driver "{3}".' -f (($queues | ForEach-Object { '"' + $_.Name + '"' }) -join ', '), $PackageConfig.PrinterAddress, $PackageConfig.PortName, $PackageConfig.DriverName

    [void]$sb.Append(@"
==============================================================
 PRINTER PACKAGE: $name
==============================================================
Gemaakt op : $(Get-Date -Format 'dd-MM-yyyy HH:mm') door $env:USERNAME
Versie     : $($PackageConfig.Version)
Map        : $PackageDir

--------------------------------------------------------------
 INTUNE - WIN32 APP INSTELLINGEN
--------------------------------------------------------------
App package file : $IntuneWinFile

[App information]
Name             : $name
Description      : $description
Publisher        : $($PackageConfig.Publisher)
App version      : $($PackageConfig.Version)
Category         : Printers
Logo             : $LogoFile ($LogoSource)

[Program]
Install command   : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
Uninstall command : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1
Installation time required (mins) : 30
Allow available uninstall         : Yes
Install behavior                  : System
Device restart behavior           : No specific action
Return codes                      : standaard (0 = Success, 1707 = Success, 3010 = Soft reboot, 1641 = Hard reboot, 1618 = Retry)

[Requirements]
Operating system architecture : 64-bit (x64)
Minimum operating system      : Windows 10 21H2 / Windows 11

[Detection rules]
Rules format : Use a custom detection script
Script file  : detect.ps1 (in deze map)
Run script as 32-bit process on 64-bit clients : No
Enforce script signature check and run script silently : No

Alternatief (handmatige regel):
  Rule type : Registry
  Key path  : HKEY_LOCAL_MACHINE\$($script:RegBase)\$name
  Value name: Version
  Detection method : String comparison - Equals - $($PackageConfig.Version)
  Associated with a 32-bit app on 64-bit clients : No

[Assignments]
Wijs bij voorkeur toe aan een apparaatgroep (installatie draait in SYSTEM-context).

--------------------------------------------------------------
 PRINTERGEGEVENS
--------------------------------------------------------------
IP-adres / hostnaam : $($PackageConfig.PrinterAddress)
Poortnaam           : $($PackageConfig.PortName)
Protocol            : $($PackageConfig.Protocol)$(if ($PackageConfig.Protocol -eq 'LPR') { " (wachtrij: $($PackageConfig.LprQueue))" } else { " (poort $($PackageConfig.PortNumber))" })
SNMP                : $(if ($PackageConfig.SnmpEnabled) { "Aan (community: $($PackageConfig.SnmpCommunity))" } else { 'Uit' })
Drivernaam          : $($PackageConfig.DriverName)
Driverversie (INF)  : $(Get-InfDriverVersion -Path (Join-Path $State.DriverFolder $State.InfPath))  (oudere versie op client wordt bijgewerkt, zelfde/nieuwere blijft)
INF-bestand         : Drivers\$($PackageConfig.InfPath)
Alle INF's stagen   : $(if ($PackageConfig.StageAllInfs) { 'Ja' } else { 'Nee' })
Driver bij uninstall: $(if ($PackageConfig.RemoveDriverOnUninstall) { 'Verwijderen (indien niet meer in gebruik)' } else { 'Laten staan' })
Bron drivermap      : $($State.DriverFolder)

"@)

    [void]$sb.Append('Installeerbare opties (Set-PrinterProperty):' + $nl)
    $opts = @($PackageConfig.DeviceOptions)
    if ($opts.Count) { foreach ($o in $opts) { [void]$sb.Append("  - $($o.Property) = $($o.Value)$nl") } }
    else { [void]$sb.Append("  (geen - driverstandaard)$nl") }

    [void]$sb.Append("$nl" + 'Printerwachtrijen:' + $nl)
    foreach ($q in $queues) {
        [void]$sb.Append("  * $($q.Name)$nl")
        if ($q.Location) { [void]$sb.Append("      Locatie       : $($q.Location)$nl") }
        if ($q.Comment)  { [void]$sb.Append("      Opmerking     : $($q.Comment)$nl") }
        [void]$sb.Append("      Standaardlade : $(Get-TrayDisplay $q.Tray)$(if ($q.Tray) { "  ($($q.Tray))" })$nl")
        [void]$sb.Append("      Dubbelzijdig  : $(Get-ChoiceDisplay $duplexChoices $q.Duplex)$nl")
        [void]$sb.Append("      Kleur         : $(Get-ChoiceDisplay $colorChoices $q.Color)$nl")
        [void]$sb.Append("      Papier        : $(if ($q.PaperSize) { $q.PaperSize } else { '(driverstandaard)' })$nl")
    }

    [void]$sb.Append(@"

--------------------------------------------------------------
 BESTANDEN
--------------------------------------------------------------
Source\install.ps1        Installeert driver, poort, wachtrijen, opties en lades
Source\uninstall.ps1      Verwijdert wachtrijen, poort$(if ($PackageConfig.RemoveDriverOnUninstall) { ' en driver' }) en registry-marker
Source\Drivers\           Driverbestanden
Output\$(Split-Path $IntuneWinFile -Leaf)   Uploaden in Intune
detect.ps1                Custom detectiescript
PrinterPackage.json       Configuratie; te laden in New-PrinterIntunePackage.ps1 voor een nieuwe versie

Afdrukvoorkeuren (dubbelzijdig, kleur, papier, lade):
  Worden als printerstandaard (alle gebruikers) ingesteld via een PrintTicket; per instelling
  wordt de optie gekozen die de driver aanbiedt (bijv. Monochrome of Grayscale) en daarna
  gecontroleerd. Oude per-gebruiker voorkeuren van aangemelde gebruikers worden gewist.
  In het install-log staat per printer een regel "Resultaat: ..." met de werkelijke waarden.

Logbestanden op de client:
  C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\$name-install.log
  C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\$name-uninstall.log

Nieuwe versie uitrollen (bijv. andere lade of IP):
  Laad PrinterPackage.json, verhoog de versie, maak het package opnieuw en
  vervang het .intunewin-bestand + detectiescript in de bestaande Intune-app.
"@)
    return $sb.ToString()
}

function Write-DriverSignatureInfo {
    param([string]$InfFullPath)
    $cats = @(Get-ChildItem -LiteralPath (Split-Path $InfFullPath -Parent) -Filter '*.cat' -File -ErrorAction SilentlyContinue)
    if ($cats.Count -eq 0) {
        Write-Log 'LET OP: geen .cat-bestand naast het INF-bestand. Windows zal om bevestiging vragen of de driver weigeren.'
        return
    }
    $signers = @{}
    foreach ($cat in $cats) {
        $signature = Get-AuthenticodeSignature -FilePath $cat.FullName -ErrorAction SilentlyContinue
        $subject = if ($signature.SignerCertificate) { $signature.SignerCertificate.Subject } else { '(niet ondertekend)' }
        $signers["$subject|$($signature.Status)"] = $true
    }
    foreach ($key in $signers.Keys) {
        $subject, $status = $key -split '\|', 2
        $kind = if ($subject -match 'Microsoft Windows Hardware Compatibility') { 'WHQL (Microsoft)' } else { 'uitgever wordt bij installatie vertrouwd (TrustedPublisher)' }
        Write-Log "Driverhandtekening: $status - $subject -> $kind"
        if ($status -ne 'Valid') { Write-Log 'LET OP: handtekening is niet geldig; een stille installatie via Intune kan mislukken.' }
    }
}

function New-Package {
    $state = Get-GuiState
    $errors = Test-State -State $state
    if ($errors.Count) {
        Show-Message -Text ("Controleer de volgende punten:`n`n- " + ($errors -join "`n- ")) -Icon Warning | Out-Null
        return
    }

    $packageName = if ($state.PackageName) { $state.PackageName } else { 'Printer - ' + @($state.Queues)[0].Name }
    $packageName = Get-SafeName $packageName
    $ui.txtPackageName.Text = $packageName
    $state.PackageName = $packageName

    $exe = Get-IntuneWinAppUtilPath
    if (-not $exe) { return }

    # Vraag de opslaglocatie
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Waar moet het printerpakket '$packageName' opgeslagen worden?`n(er wordt een submap met de packagenaam aangemaakt)"
    $dialog.ShowNewFolderButton = $true
    $dialog.SelectedPath = if ($script:LastOutput) { $script:LastOutput } else { Split-Path (Split-Path $script:ToolDir -Parent) -Parent }
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write-Log 'Package maken geannuleerd.'; return }
    $script:LastOutput = $dialog.SelectedPath

    $packageDir = Join-Path $dialog.SelectedPath $packageName
    $sourceDir  = Join-Path $packageDir 'Source'
    $outputDir  = Join-Path $packageDir 'Output'

    if ($state.DriverFolder.TrimEnd('\') -like "$($packageDir.TrimEnd('\'))*") {
        throw 'De drivermap ligt in de doelmap van het package. Kies een andere opslaglocatie.'
    }

    if (Test-Path -LiteralPath $packageDir) {
        $answer = Show-Message -Text "De map bestaat al:`n$packageDir`n`nSource en Output overschrijven?" -Icon Warning -Buttons YesNo
        if ($answer -ne 'Yes') { Write-Log 'Package maken geannuleerd.'; return }
        foreach ($dir in $sourceDir, $outputDir) {
            if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction Stop }
        }
    }

    Set-Busy $true
    Write-Log "=== Package maken: $packageName ==="
    New-Item -Path (Join-Path $sourceDir 'Drivers') -ItemType Directory -Force | Out-Null
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null

    Write-DriverSignatureInfo -InfFullPath (Join-Path $state.DriverFolder $state.InfPath)

    Write-Log 'Driverbestanden kopieren...'
    Copy-Item -Path (Join-Path $state.DriverFolder '*') -Destination (Join-Path $sourceDir 'Drivers') -Recurse -Force -ErrorAction Stop
    Update-UI

    $packageConfig = Get-PackageConfig -State $state -PackageName $packageName
    Save-TextFile -Path (Join-Path $sourceDir 'install.ps1')   -Content (New-ScriptFromTemplate -Template $script:InstallTemplate -PackageConfig $packageConfig)
    Save-TextFile -Path (Join-Path $sourceDir 'uninstall.ps1') -Content (New-ScriptFromTemplate -Template $script:UninstallTemplate -PackageConfig $packageConfig)
    Save-TextFile -Path (Join-Path $packageDir 'detect.ps1')   -Content (New-ScriptFromTemplate -Template $script:DetectTemplate -PackageConfig $packageConfig)
    Save-TextFile -Path (Join-Path $packageDir 'PrinterPackage.json') -Content ($state | ConvertTo-Json -Depth 6)
    Write-Log 'install.ps1, uninstall.ps1, detect.ps1 en PrinterPackage.json geschreven.'

    $logoFile = "$($packageName)_logo.png"
    $logoSource = Save-Logo -SourcePath $state.LogoPath -TargetPath (Join-Path $packageDir $logoFile)
    Write-Log "Logo: $logoFile ($logoSource)"

    Write-Log 'IntuneWinAppUtil uitvoeren (kan even duren bij grote drivers)...'
    $arguments = '-c "{0}" -s "{1}" -o "{2}" -q' -f $sourceDir, (Join-Path $sourceDir 'install.ps1'), $outputDir
    $result = Invoke-ExternalProcess -FilePath $exe -Arguments $arguments
    $intuneWin = Join-Path $outputDir 'install.intunewin'
    if ($result.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $intuneWin)) {
        Write-Log $result.Output
        Write-Log $result.Error
        throw "IntuneWinAppUtil is mislukt (exitcode $($result.ExitCode)). Zie het log hierboven."
    }
    Write-Log ('install.intunewin gemaakt ({0:N1} MB).' -f ((Get-Item -LiteralPath $intuneWin).Length / 1MB))

    $info = New-InfoText -State $state -PackageConfig $packageConfig -PackageDir $packageDir -LogoFile $logoFile -LogoSource $logoSource -IntuneWinFile 'Output\install.intunewin'
    Save-TextFile -Path (Join-Path $packageDir 'info.txt') -Content $info
    Write-Log 'info.txt geschreven.'
    Write-Log "=== Klaar: $packageDir ==="
    Set-Busy $false

    $answer = Show-Message -Text "Het package is gemaakt:`n$packageDir`n`nMap openen?" -Icon Information -Buttons YesNo
    if ($answer -eq 'Yes') { Start-Process explorer.exe -ArgumentList "`"$packageDir`"" }
}

#endregion

#region Build window

$reader = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)
$script:ui = @{}
foreach ($node in $xaml.SelectNodes("//*[@*[local-name()='Name']]")) {
    $name = $node.Attributes | Where-Object { $_.LocalName -eq 'Name' } | Select-Object -ExpandProperty Value -First 1
    $ui[$name] = $script:Window.FindName($name)
}

$script:Queues        = New-Object 'System.Collections.ObjectModel.ObservableCollection[EooPrinterPkg.QueueRow]'
$script:DeviceOptions = New-Object 'System.Collections.ObjectModel.ObservableCollection[EooPrinterPkg.OptionRow]'
$script:TrayChoices   = New-Object 'System.Collections.ObjectModel.ObservableCollection[EooPrinterPkg.Choice]'
$script:ModelIndex    = New-Object 'System.Collections.Generic.List[EooPrinterPkg.Choice]'
$script:DuplexChoices = @(
    (New-Object EooPrinterPkg.Choice('(driverstandaard)', '')),
    (New-Object EooPrinterPkg.Choice('Enkelzijdig', 'OneSided')),
    (New-Object EooPrinterPkg.Choice('Dubbelzijdig - lange zijde', 'TwoSidedLongEdge')),
    (New-Object EooPrinterPkg.Choice('Dubbelzijdig - korte zijde', 'TwoSidedShortEdge'))
)
$script:ColorChoices = @(
    (New-Object EooPrinterPkg.Choice('(driverstandaard)', '')),
    (New-Object EooPrinterPkg.Choice('Kleur', 'Color')),
    (New-Object EooPrinterPkg.Choice('Zwart-wit', 'Mono'))
)
$script:PaperChoices = @((New-Object EooPrinterPkg.Choice('(driverstandaard)', ''))) +
    @('A4', 'A3', 'A5', 'Letter', 'Legal' | ForEach-Object { New-Object EooPrinterPkg.Choice($_, $_) })

Reset-TrayChoices
$ui.dgQueues.ItemsSource  = $script:Queues
$ui.dgOptions.ItemsSource = $script:DeviceOptions
$ui.dgQueues.Columns[3].ItemsSource = $script:TrayChoices
$ui.dgQueues.Columns[4].ItemsSource = $script:DuplexChoices
$ui.dgQueues.Columns[5].ItemsSource = $script:ColorChoices
$ui.dgQueues.Columns[6].ItemsSource = $script:PaperChoices
[void](Add-QueueRow)

$ui.txtAdminStatus.Text = if (Test-IsAdmin) { 'Administrator: ja' } else { 'Administrator: nee (nodig voor lades uitlezen)' }
$foundExe = Find-IntuneWinAppUtil
if ($foundExe) { $ui.txtIntuneWin.Text = $foundExe }

# --- Events ---
$ui.txtAddress.Add_TextChanged({
        $address = $ui.txtAddress.Text.Trim()
        $current = $ui.txtPortName.Text
        if (-not $current -or $current -eq $script:LastAutoPort) {
            $script:LastAutoPort = if ($address) { "IP_$address" } else { '' }
            $ui.txtPortName.Text = $script:LastAutoPort
        }
    })

$ui.cmbProtocol.Add_SelectionChanged({
        $lpr = $ui.cmbProtocol.SelectedIndex -eq 1
        $ui.txtLprQueue.IsEnabled = $lpr
        $ui.txtPortNumber.IsEnabled = -not $lpr
    })

$ui.chkSnmp.Add_Click({ $ui.txtSnmpCommunity.IsEnabled = [bool]$ui.chkSnmp.IsChecked })

$ui.btnPing.Add_Click({ Invoke-Safe { Test-PrinterConnection } })

$ui.btnBrowseDriver.Add_Click({
        Invoke-Safe {
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = 'Kies de map met de uitgepakte printerdriver (.inf-bestanden)'
            if ($ui.txtDriverFolder.Text -and (Test-Path -LiteralPath $ui.txtDriverFolder.Text)) { $dialog.SelectedPath = $ui.txtDriverFolder.Text }
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $ui.txtDriverFolder.Text = $dialog.SelectedPath
                $script:LastDriverFolder = $dialog.SelectedPath
                Update-InfList
            }
        }
    })

$ui.txtDriverFolder.Add_LostFocus({
        if ($ui.txtDriverFolder.Text -ne $script:LastDriverFolder) {
            $script:LastDriverFolder = $ui.txtDriverFolder.Text
            Invoke-Safe { Update-InfList }
        }
    })

$ui.cmbInf.Add_SelectionChanged({ Invoke-Safe { Update-DriverNames } })

$ui.txtModelSearch.Add_TextChanged({ Invoke-Safe { Update-ModelList } })
$ui.btnClearSearch.Add_Click({ $ui.txtModelSearch.Text = '' })
$ui.lstModels.Add_SelectionChanged({ Invoke-Safe { Select-Model } })

$ui.btnReadOptions.Add_Click({ Invoke-Safe { Read-DriverOptions } })

$ui.btnAddOption.Add_Click({
        $row = New-Object EooPrinterPkg.OptionRow
        $row.Property = 'Config:'
        $script:DeviceOptions.Add($row)
    })

$ui.btnRemoveOption.Add_Click({
        $selected = $ui.dgOptions.SelectedItem
        if ($selected -is [EooPrinterPkg.OptionRow]) { [void]$script:DeviceOptions.Remove($selected) }
    })

$ui.btnAddQueue.Add_Click({
        Confirm-GridEdits
        $row = Add-QueueRow
        $ui.dgQueues.SelectedItem = $row
        $ui.dgQueues.ScrollIntoView($row)
    })

$ui.btnRemoveQueue.Add_Click({
        $selected = $ui.dgQueues.SelectedItem
        if ($selected -is [EooPrinterPkg.QueueRow]) {
            Confirm-GridEdits
            [void]$script:Queues.Remove($selected)
        }
    })

$ui.btnBrowseLogo.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'Afbeeldingen (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp'
        if ($dialog.ShowDialog() -eq $true) { $ui.txtLogo.Text = $dialog.FileName }
    })

$ui.btnBrowseIntuneWin.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'IntuneWinAppUtil.exe|IntuneWinAppUtil.exe|Programma (*.exe)|*.exe'
        if ($dialog.ShowDialog() -eq $true) { $ui.txtIntuneWin.Text = $dialog.FileName }
    })

$ui.btnSave.Add_Click({
        Invoke-Safe {
            $dialog = New-Object Microsoft.Win32.SaveFileDialog
            $dialog.Filter = 'Printer package configuratie (*.json)|*.json'
            $dialog.FileName = 'PrinterPackage.json'
            if ($dialog.ShowDialog() -eq $true) {
                Save-TextFile -Path $dialog.FileName -Content (Get-GuiState | ConvertTo-Json -Depth 6)
                Write-Log "Configuratie opgeslagen: $($dialog.FileName)"
            }
        }
    })

$ui.btnLoad.Add_Click({
        Invoke-Safe {
            $dialog = New-Object Microsoft.Win32.OpenFileDialog
            $dialog.Filter = 'Printer package configuratie (*.json)|*.json'
            if ($dialog.ShowDialog() -eq $true) { Import-StateFile -Path $dialog.FileName }
        }
    })

$ui.btnBuild.Add_Click({ Invoke-Safe { New-Package } })

$script:Window.Add_ContentRendered({
        Write-Log 'Klaar voor gebruik. Begin op tabblad 1 met het printeradres en de drivermap.'
        if ($LoadConfig -and (Test-Path -LiteralPath $LoadConfig)) {
            Invoke-Safe { Import-StateFile -Path $LoadConfig }
            if ($LoadConfig -like "$env:TEMP\PrinterPackager_*.json") {
                Remove-Item -LiteralPath $LoadConfig -ErrorAction SilentlyContinue
                $ui.tabs.SelectedIndex = 1
            }
        }
    })

[void]$script:Window.ShowDialog()

#endregion
