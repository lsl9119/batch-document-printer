<#
创建一个用于验证打印链路的虚拟打印机（不消耗纸张）：
  打印机名: BDP Test PDF
  驱动:     Microsoft Print To PDF
  端口:     本地文件端口（每个打印作业都写入同一个 PDF 文件，无需弹出"另存为"对话框）

需要管理员权限运行：
  powershell -ExecutionPolicy Bypass -File scripts\setup_test_pdf_printer.ps1 -OutputFile C:\bdp_test\out.pdf
删除：
  Remove-Printer -Name "BDP Test PDF"; Remove-PrinterPort -Name C:\bdp_test\out.pdf
#>
param(
    [string]$PrinterName = "BDP Test PDF",
    [string]$OutputFile = "C:\bdp_test\out.pdf"
)
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path (Split-Path $OutputFile) | Out-Null

$driver = Get-PrinterDriver -Name "Microsoft Print To PDF" -ErrorAction SilentlyContinue
if (-not $driver) {
    Write-Host "未找到 Microsoft Print To PDF 驱动，尝试启用 Windows 功能..."
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
