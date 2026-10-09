# -*- coding: utf-8 -*-
"""全局配置：所有可调参数集中在这里。"""
import os
import locale

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 上传文件存放目录（自动创建）
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
# VBScript 脚本目录
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")

# ---------------- 服务监听 ----------------
HOST = "0.0.0.0"      # 局域网内可访问
PORT = 8080

# ---------------- 上传限制 ----------------
MAX_CONTENT_MB = 100
ALLOWED_EXT = {
    ".doc", ".docx", ".txt",
    ".xls", ".xlsx",
    ".pdf",
    ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot", ".potx",
}

# 按后端分组
WORD_EXT = {".doc", ".docx", ".txt"}
EXCEL_EXT = {".xls", ".xlsx"}
PPT_EXT = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot", ".potx"}

# ---------------- PowerPoint 讲义 ----------------
# 一张纸打印几张幻灯片的可选值（PowerPoint 只接受这几个）
SLIDES_PER_PAGE_CHOICES = (1, 2, 3, 4, 6, 9)
SLIDES_PER_PAGE_DEFAULT = 1

# ---------------- 任务 ----------------
JOB_TIMEOUT_SEC = 600        # 单个打印任务最长执行时间（秒）
JOB_RETENTION_HOURS = 6      # 任务记录 / 上传文件的保留时间
MAX_LOG_LINES = 500          # 单个任务最多保留的日志行数

# ---------------- 控制台编码 ----------------
# cscript.exe 使用系统 OEM/ANSI 代码页输出（简体中文 Windows 一般为 GBK）
CONSOLE_ENCODING = locale.getpreferredencoding(False) or "gbk"

# ---------------- cscript.exe 位置 ----------------
_cscript = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "cscript.exe")
CSCRIPT = _cscript if os.path.exists(_cscript) else "cscript"

# ---------------- SumatraPDF（用于打印 PDF） ----------------
# 留空则自动在程序目录 / SumatraPDF 子目录里找；也可以写绝对路径或相对本目录的路径
SUMATRA_PATH = ""

def find_sumatra():
    """查找 SumatraPDF：优先用 SUMATRA_PATH，其次在程序目录及子目录里找。"""
    if SUMATRA_PATH:
        p = SUMATRA_PATH if os.path.isabs(SUMATRA_PATH) \
            else os.path.join(BASE_DIR, SUMATRA_PATH)
        if os.path.exists(p):
            return p
    candidates = [
        os.path.join(BASE_DIR, "SumatraPDF.exe"),
        os.path.join(BASE_DIR, "SumatraPDF-64.exe"),
        os.path.join(BASE_DIR, "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(BASE_DIR, "SumatraPDF", "SumatraPDF-64.exe"),
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

SUMATRA = find_sumatra()

# 部分 SumatraPDF 版本在 -print-settings 中要求份数写作 "2x"，
# 若发现份数不生效，把这里改成 False（变成 "2"）。
SUMATRA_COPIES_X_SUFFIX = True

# SumatraPDF 的控制台输出是 UTF-8（cscript 是 GBK），两边必须分开解码，否则日志乱码
SUMATRA_ENCODING = "utf-8"

# ---------------- SumatraPDF 纸张 / 纸盒 ----------------
# SumatraPDF 默认按 PDF 页面自身的尺寸去选纸张。而 PPT 导出的 PDF 页面尺寸 = 幻灯片尺寸
# （16:9 时约 33.9 x 19.1 cm），驱动没有对应纸型就退化成“自定义”，打印机于是报
# “轻型纸，自定义”。显式指定纸张即可解决。
#   常用值：A4 / A3 / A5 / letter / legal
#   paper=auto  = 按每页自身尺寸（等于不加）
#   也可以写驱动报告的名字，用 SumatraPDF.exe -list-printers 查看
# 设为 "" 表示不指定（退回默认行为）
SUMATRA_PAPER = "A4"

# 纸盒选择："" = 不指定；"auto" = 让打印机按页面尺寸选纸盒；也可写编号或名字，如 "2"
SUMATRA_BIN = ""

# 缩放方式："" = 用 SumatraPDF 默认；可选 fit（适应可打印区域）/ shrink / noscale / stretch
SUMATRA_SCALE = ""

# ---------------- 双面打印策略 ----------------
# "auto"    : 先试「改打印机默认双面」，失败再试「Office 转 PDF → SumatraPDF 双面」
# "printer" : 只改打印机默认双面（改不动就按单面打，日志会说明）
# "pdf"     : 一律走「转 PDF → SumatraPDF 双面」（不需要改打印机全局设置，最稳）
# 说明：改打印机默认双面是全局生效的，且常因权限不足/打印机是共享连接而失败；
#      PDF 路线是按「单个任务」下发双面参数，不需要任何权限，优先推荐遇到问题的人使用。
# 例外：PowerPoint 讲义（一张纸 N 张幻灯片）无法导出成 PDF 保留版式，
#      所以讲义 + 双面只能走路线 1（改打印机默认双面）。
DUPLEX_STRATEGY = "auto"

# ---------------- 打印完成判定 ----------------
# 主路径是 COM 本身（Word/Excel 的同步 PrintOut、PPT 的 PrintInBackground=0），
# 下面的等待只是"确认"用的兜底：按 JobId 跟踪队列 + 读假脱机累计计数交叉验证。
# 设为 0 可关闭等待（只信 COM，速度最快）。

# Word/Excel：等待打印任务完成的最长秒数（超时放弃等待，不算失败）
OFFICE_QUEUE_WAIT_SEC = 300
# PowerPoint：同上
PPT_QUEUE_WAIT_SEC = 300
# 等待任务出现在打印队列里的宽限期（之后认为它已经打完，或被打印机拒收）
JOB_APPEAR_SEC = 60
PPT_JOB_APPEAR_SEC = JOB_APPEAR_SEC

# ---------------- PowerPoint 退出方式 ----------------
# 实测：自己启动的 PowerPoint 实例调用 Quit() 会阻塞 15 秒以上才返回（Word/Excel 不会）。
# True  = 导出/打印完成后直接 taskkill 掉自己启动的那个 PowerPoint（秒退，默认）
# False = 用 PowerPoint.Quit() 正常退出（更温和，但每个任务多花十几秒）
# 注意：只有当脚本确认「启动前没有任何 POWERPNT.EXE 进程」时才会杀进程，
#       因此不会误伤你自己已经打开的 PowerPoint。
PPT_FAST_QUIT = True

# ---------------- PowerPoint 指定打印机 ----------------
# PowerPoint 的 COM 接口没有 ActivePrinter，只能把目标打印机临时设为「Windows 默认打印机」，
# 打印完再恢复。若切换失败，脚本会直接报错中止（避免打到错误的打印机上）。
# 设为 True 则切换失败时仍然继续打（会打到当时的默认打印机上，不推荐）。
PPT_IGNORE_PRINTER_SWITCH_FAILURE = False

# ---------------- Office 进程 ----------------
# Word/Excel 以 COM 方式在后台运行；打印超时后是否强制结束残留进程。
# True 会杀掉所有 WINWORD/EXCEL 进程（可能误伤用户正在编辑的文档），默认关闭。
KILL_OFFICE_ON_TIMEOUT = False
