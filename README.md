# PrinterInstallWindows

PowerShell script to simplify printer installation on Windows computers.

## Download

Run the command below to download the script (avoiding PowerShell's Execution Policy warnings).  Review the script unless you're comfortable running arbitrary code written by strangers on the internet.
```powershell
irm https://raw.githubusercontent.com/ManzellaTech/PrinterInstallWindows/refs/heads/main/Install-Printer.ps1 | Set-Content Install-Printer.ps1
```

## Usage

Install a printer using a specific .inf driver file.
```powershell
.\Install-Printer.ps1 -PrinterName "Front Desk" -DriverPath "C:\Drivers\Brandname m321\driver\x64\m321.inf" -PrinterIP "10.22.6.21" -DriverName "Brandname m321 PCL 6"
```

Install a printer by passing a folder path that contains a .inf driver file (recursively searches the folder, first .inf file found will be used).
```powershell
.\Install-Printer.ps1 -PrinterName "Side Desk" -DriverPath "C:\Drivers\Brandname m322\" -PrinterIP "10.22.6.22" -DriverName "Brandname m322 PCL 6"
```

Install a printer by passing a .zip file path that contains a .inf driver file (recursively searches the archive, first .inf file found will be used).
```powershell
.\Install-Printer.ps1 -PrinterName "Checkout" -DriverPath "Drivers\Brandname m323 Driver.zip" -PrinterIP "10.22.6.23" -DriverName "Brandname m323 PCL 6"
```