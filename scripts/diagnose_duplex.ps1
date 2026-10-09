<#
双面打印诊断脚本（PowerShell 版）

用法（PowerShell 里执行）：
  powershell -ExecutionPolicy Bypass -File .\scripts\diagnose_duplex.ps1
  powershell -ExecutionPolicy Bypass -File .\scripts\diagnose_duplex.ps1 "NPIDA8E61 (HP LaserJet MFP M227fdw)"

不传打印机名时，会自动用默认打印机。

它会打印：PowerShell 版本、是否管理员、打印机详细信息（驱动/端口/是否共享/是否网络打印机）、
当前双面设置、以及尝试改成双面后的完整异常信息。
#>
param(
    [string]$PrinterName = ""
)

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Show($t) { Write-Host $t }

Show "============================================================"
Show " Duplex diagnostic"
Show "============================================================"

Show ""
Show "[1] Environment"
Show ("  PSVersion : " + $PSVersionTable.PSVersion.ToString())
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Show ("  IsAdmin   : " + $isAdmin)
$mod = Get-Module -ListAvailable PrintManagement -ErrorAction SilentlyContinue
Show ("  PrintManagement module : " + [bool]$mod)

Show ""
Show "[2] Printers"
$printers = @()
try {
    $printers = Get-Printer -ErrorAction Stop | Sort-Object Name
} catch {
    Show ("  Get-Printer failed: " + $_.Exception.Message)
}
foreach ($p in $printers) {
    Show ("   - " + $p.Name)
}

if ($PrinterName -eq "") {
    $def = $printers | Where-Object { $_.Default -eq $true } | Select-Object -First 1
    if ($def) { $PrinterName = $def.Name } elseif ($printers.Count -gt 0) { $PrinterName = $printers[0].Name }
}
if ($PrinterName -eq "") {
    Show ""
    Show "No printer found. Exit."
    exit 1
}

Show ""
Show ("[3] Target printer : " + $PrinterName)
try {
    $p = Get-Printer -Name $PrinterName -ErrorAction Stop
    Show ("  DriverName   : " + $p.DriverName)
    Show ("  PortName     : " + $p.PortName)
    Show ("  Shared       : " + $p.Shared)
    Show ("  ShareName    : " + $p.ShareName)
    Show ("  PrinterStatus: " + $p.PrinterStatus)
    Show ("  Type         : " + $p.Type)
} catch {
    Show ("  Get-Printer failed: " + $_.Exception.Message)
}

try {
    $cfg = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop
    Show ("  DuplexingMode: " + $cfg.DuplexingMode)
    Show ("  PaperSize    : " + $cfg.PaperSize)
    Show ("  Color        : " + $cfg.Color)
} catch {
    Show ("  Get-PrintConfiguration failed: " + $_.Exception.Message)
}

Show ""
Show "[4] Try to set TwoSidedLongEdge"
try {
    Set-PrintConfiguration -PrinterName $PrinterName -DuplexingMode 'TwoSidedLongEdge' -Confirm:$false -ErrorAction Stop
    $after = (Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop).DuplexingMode
    Show ("  OK, readback = " + $after)
} catch {
    Show ("  FAILED: " + $_.Exception.Message)
    Show ("  FullRecord :")
    Show ($_.Exception.ToString())
}

try {
    $cfg2 = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop
    Show ""
    Show "[5] Driver capabilities"
    Show ("  DuplexingMode now : " + $cfg2.DuplexingMode)
    Show ("  Collate           : " + $cfg2.Collate)
} catch {
    Show ("  readback failed: " + $_.Exception.Message)
}

Show ""
Show "============================================================"
Show "Tips:"
Show "  - 'Access is denied' -> need admin, or the printer is a shared connection"
Show "  - readback stays 'OneSided' -> driver does not report duplex capability"
Show "  - for shared printers: set DUPLEX_STRATEGY = 'pdf' in config.py"
Show "============================================================"
