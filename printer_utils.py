# -*- coding: utf-8 -*-
"""打印机相关工具：枚举打印机、读取/设置双面（Duplex）模式。

所有 PowerShell 脚本都用 -EncodedCommand（base64 UTF-16LE）传递，
彻底绕开打印机名里的空格、括号、引号、中文造成的转义问题。
"""
import base64
import locale
import subprocess

PS_TIMEOUT = 40


def _ps_quote(s):
    """PowerShell 单引号字符串转义。"""
    return "'" + str(s).replace("'", "''") + "'"


def _ps(script, timeout=PS_TIMEOUT):
    """执行 PowerShell 脚本，返回 (returncode, stdout, stderr)。"""
    full = ("$ErrorActionPreference='Stop';\n"
            "$ProgressPreference='SilentlyContinue';\n"
            "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;\n"
            + script)
    enc = base64.b64encode(full.encode("utf-16-le")).decode("ascii")
    try:
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-EncodedCommand", enc],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
        return (p.returncode,
                p.stdout.decode("utf-8", errors="replace"),
                p.stderr.decode("utf-8", errors="replace"))
    except FileNotFoundError:
        return -1, "", "找不到 powershell.exe"
    except subprocess.TimeoutExpired:
        return -1, "", "PowerShell 执行超时"
    except Exception as e:
        return -1, "", str(e)


def list_printers():
    """返回本机可用打印机名称列表。"""
    rc, out, err = _ps("Get-Printer | Sort-Object Name | ForEach-Object { $_.Name }")
    names = [l.strip() for l in out.splitlines() if l.strip()]
    if names:
        return names

    # 回退：wmic（Win10 仍可用，但已被标记弃用）
    try:
        enc = locale.getpreferredencoding(False) or "gbk"
        p = subprocess.run(["wmic", "printer", "get", "name", "/value"],
                           stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=20)
        text = p.stdout.decode(enc, errors="replace")
        for line in text.splitlines():
            line = line.strip()
            if line.upper().startswith("NAME="):
                v = line.split("=", 1)[1].strip()
                if v:
                    names.append(v)
    except Exception:
        pass
    return names


def get_default_printer():
    rc, out, err = _ps("(Get-Printer | Where-Object { $_.Default -eq $true } | "
                       "Select-Object -First 1).Name")
    out = out.strip()
    if out:
        return out
    rc, out, err = _ps("(Get-WmiObject Win32_Printer -Filter \"Default=True\" | "
                       "Select-Object -First 1).Name")
    return out.strip()


def get_duplex(printer):
    """读取打印机当前双面模式：OneSided / TwoSidedLongEdge / TwoSidedShortEdge，失败返回 None"""
    rc, out, err = _ps("(Get-PrintConfiguration -PrinterName %s).DuplexingMode"
                       % _ps_quote(printer))
    out = out.strip()
    return out if out else None


# 依次尝试两种写法：按打印机名设置、按打印机对象设置；并回读校验是否真的生效
_PS_SET = """
$name = %(name)s
$mode = %(mode)s
$errs = @()
try {
    Set-PrintConfiguration -PrinterName $name -DuplexingMode $mode -Confirm:$false -ErrorAction Stop
    $after = (Get-PrintConfiguration -PrinterName $name -ErrorAction Stop).DuplexingMode
    Write-Output ("OK|printerName|" + $after)
} catch {
    $errs += "printerName: " + $_.Exception.Message
}
if ($errs.Count -gt 0) {
    try {
        $p = Get-Printer -Name $name -ErrorAction Stop
        Set-PrintConfiguration -PrinterObject $p -DuplexingMode $mode -Confirm:$false -ErrorAction Stop
        $after = (Get-PrintConfiguration -PrinterName $name -ErrorAction Stop).DuplexingMode
        Write-Output ("OK|printerObject|" + $after)
    } catch {
        $errs += "printerObject: " + $_.Exception.Message
    }
}
if ($errs.Count -gt 0) {
    Write-Output ("ERR|" + ($errs -join " || "))
}
"""


def set_duplex(printer, mode):
    """设置双面模式，返回 (是否生效, 详情文本)。

    注意：即使 PowerShell 没报错，也要回读一次确认值真的变了——
    有些驱动会"接受"设置但静默忽略（那就等于没生效）。
    """
    rc, out, err = _ps(_PS_SET % {"name": _ps_quote(printer), "mode": _ps_quote(mode)})
    out = (out or "").strip()

    if out.startswith("OK|"):
        parts = out.split("|")
        after = parts[2] if len(parts) > 2 else ""
        if after.strip().lower() == mode.strip().lower():
            return True, "已设置为 %s（回读确认）" % after
        return False, "PowerShell 返回成功，但回读仍是 %s（驱动未接受该设置）" % (after or "空")

    if out.startswith("ERR|"):
        return False, out[4:].strip() or "未知错误"
    if err.strip():
        return False, err.strip().splitlines()[0]
    return False, "PowerShell 无输出（退出码 %s）" % rc


def ps_available():
    """PrintManagement 模块是否可用。"""
    rc, out, err = _ps("(Get-Command Set-PrintConfiguration -ErrorAction SilentlyContinue) -ne $null")
    return "True" in out


def ps_module_report():
    """返回一段可读的诊断信息（给管理员排查用）。"""
    rc, out, err = _ps(
        "$v = $PSVersionTable.PSVersion.ToString(); "
        "Write-Output ('PSVersion=' + $v); "
        "$m = Get-Module -ListAvailable PrintManagement -ErrorAction SilentlyContinue; "
        "Write-Output ('PrintManagement=' + [bool]$m); "
        "Write-Output ('IsAdmin=' + ([Security.Principal.WindowsPrincipal]"
        "[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("
        "[Security.Principal.WindowsBuiltInRole]::Administrator))")
    return (out or err or "").strip() or "无法获取"
