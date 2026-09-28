<#
Create a virtual printer for verifying the print pipeline without using paper:
  Printer name: BDP Test PDF
  Driver:       Microsoft Print To PDF
  Port:         a local file port (every job is written to the same PDF file,
                so no "Save As" dialog is shown)

Run as administrator:
  powershell -ExecutionPolicy Bypass -File scripts\setup_test_pdf_printer.ps1 -OutputFile C:\bdp_test\out.pdf
Remove:
  Remove-Printer -Name "BDP Test PDF"; Remove-PrinterPort -Name C:\bdp_test\out.pdf

(This file is intentionally ASCII-only: Windows PowerShell 5.1 reads BOM-less
 scripts with the ANSI code page, which breaks non-ASCII text.)
#>
param(
    [string]$PrinterName = "BDP Test PDF",
    [string]$OutputFile = "C:\bdp_test\out.pdf"
)
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path (Split-Path $OutputFile) | Out-Null

$driver = Get-PrinterDriver -Name "Microsoft Print To PDF" -ErrorAction SilentlyContinue
if (-not $driver) {
    Write-Host "Driver 'Microsoft Print To PDF' not found, enabling the Windows feature..."
    Enable-WindowsOptionalFeature -Online -FeatureName "Printing-PrintToPDFServices-Features" -All -NoRestart | Out-Null
}
if (-not (Get-PrinterPort -Name $OutputFile -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $OutputFile
}
if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
    Remove-Printer -Name $PrinterName
}
Add-Printer -Name $PrinterName -DriverName "Microsoft Print To PDF" -PortName $OutputFile
Get-Printer -Name $PrinterName | Format-List Name, DriverName, PortName
