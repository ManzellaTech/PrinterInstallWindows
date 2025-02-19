<#
.SYNOPSIS
    Script installs a printer on a Windows computer.
.DESCRIPTION
    This script takes in four mandatory parameters.
    'PrinterName' is the name of the printer that will be displayed once it is successfully installed.
    'PrinterIP' is the IP address of the printer's port.
    'DriverPath' an absolute or relative path to an .inf file.  If a path to a folder or .zip file is provided, then the contents will be recursively searched for a .inf file and the first .inf file found will be used.
    'DriverName' name of the driver to use.  It needs to exactly match the driver name contained in the .inf file.
.PARAMETER PrinterName
    A mandatory parameter of type string.
.PARAMETER PrinterIP
    A mandatory parameter of type string.
.PARAMETER DriverPath
    A mandatory parameter of type string.
.PARAMETER DriverName
    A mandatory parameter of type string.
.EXAMPLE
    .\Install-Printer.ps1 -PrinterName "Front Desk" -DriverPath "C:\Drivers\Brandname m321\driver\x64\m321.inf" -PrinterIP "10.22.6.21" -DriverName "Brandname m321 PCL 6"
    .\Install-Printer.ps1 -PrinterName "Side Desk" -DriverPath "C:\Drivers\Brandname m322\" -PrinterIP "10.22.6.22" -DriverName "Brandname m322 PCL 6"
    .\Install-Printer.ps1 -PrinterName "Checkout" -DriverPath "C:\Drivers\Brandname m323 Driver.zip" -PrinterIP "10.22.6.23" -DriverName "Brandname m323 PCL 6"
#>
param (
    [string]$PrinterName,
    [string]$PrinterIP,
    [string]$DriverPath,
    [string]$DriverName
)

function Expand-ZipFile {
    param (
        [string]$ZipPath,
        [string]$DestinationFolder
    )

    if (Test-Path $DestinationFolder) {
        Remove-Item -Recurse -Force $DestinationFolder
    }

    Write-Host "Extracting driver files from: '$($ZipPath)'..."
    New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
    Expand-Archive -Path $ZipPath -DestinationPath $DestinationFolder -Force
    Write-Host "Extraction completed to: '$($DestinationFolder)'"
}

function Get-InfFileInChildFolder {
    param (
        [string]$FolderPath
    )
    $InfFilePath = (Get-ChildItem -Path $FolderPath -Filter "*.inf" -Recurse | Select-Object -ExpandProperty FullName -First 1).ToLower()
    if ($InfFilePath) {
        Write-Host "Found INF file: $InfFilePath"
        return $InfFilePath
    } else {
        Write-Host "No INF file found in extracted folder!" -ForegroundColor Red
        exit 1
    }
}

function Get-InfFilePath {
    param (
        [string]$DriverPath
    )
    $DriverPathItem = Get-Item -Path $DriverPath
    if ($DriverPathItem.PSIsContainer) {
        $InfFilePath = Get-InfFileInChildFolder -FolderPath $DriverPath
        return $InfFilePath
    }
    elseif ($DriverPathItem.Extension -eq ".zip") {
        $ExtractedDriverPath = "$env:TEMP\PrinterDriverExtract"
        Expand-ZipFile -ZipPath $DriverPath -DestinationFolder $ExtractedDriverPath
        $InfFilePath = Get-InfFileInChildFolder -FolderPath $ExtractedDriverPath
        return $InfFilePath
    }
    elseif ($DriverPathItem.Extension -eq ".inf") {
        $InfFilePath = $ExtractedDriverPath
        return $InfFilePath
    }
    else {
        Write-Error "The 'DriverPath' argument is not a folder, .zip, or .inf file.  The following file extension is not supported: $($DriverPathItem.Extension)"
        exit 1
    }
}

function Install-PrinterDriver {
    param (
        [string]$InfFilePath
    )
    Write-Host "Installing printer driver from: $InfFilePath"
    $result = pnputil /add-driver "$InfFilePath" /subdirs /install
    $SuccessResponse = "added successfully"
    if ($result -match $SuccessResponse) {
        Write-Host "Printer driver installed successfully. Response was: $($result)"
    } else {
        Write-Host "Printer driver installation may have failed. Response was: $($result)"
        exit 1
    }
}

function Add-PrinterPortByIp {
    param (
        [string]$PrinterIP
    )
    $IsExistingPrinterPort = Get-PrinterPort -Name $PrinterIP -ErrorAction SilentlyContinue
    if ($null -eq $IsExistingPrinterPort) {
        Write-Host "Creating TCP/IP Port for $PrinterIP..."
        Add-PrinterPort -Name $PrinterIP -PrinterHostAddress $PrinterIP
        Write-Host "TCP/IP Port $PrinterIP created successfully."
    } else {
        Write-Host "TCP/IP Port $PrinterIP already exists."
    }
}

function Add-PrinterByNameIpDrivername {
    param (
        [string]$PrinterName,
        [string]$PrinterIP,
        [string]$DriverName
    )

    if (!(Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
        Write-Host "Adding printer $PrinterName..."
        Add-Printer -Name $PrinterName -PortName $PrinterIP -DriverName $DriverName
        Write-Host "Printer $PrinterName added successfully."
    } else {
        Write-Host "Printer $PrinterName already exists."
    }
}

$InfFilePath = Get-InfFilePath -DriverPath $DriverPath
Install-PrinterDriver -InfFilePath $InfFilePath
Add-PrinterDriver -Name $DriverName
Add-PrinterPortByIp -PrinterIP $PrinterIP
Add-PrinterByNameIpDrivername -PrinterName $PrinterName -PrinterIP $PrinterIP -DriverName $DriverName

Write-Host "Printer installation process completed."