# 双面打印自检工具
#
# 用法（在打印服务器上、双击 run.bat 之外的另一个窗口里运行）：
#   python test_duplex.py "NPIDA8E61 (HP LaserJet MFP M227fdw)" long
#
# 第二个参数：long（长边翻转）/ short（短边翻转），省略则用 long
#
# 它会依次打印：
#   1. PowerShell / PrintManagement 是否可用、当前是否管理员
#   2. 打印机列表
#   3. 该打印机当前的双面设置
#   4. 尝试改成双面后的完整错误信息（这就是排查的关键）
#   5. 恢复原设置
#   6. SumatraPDF 是否找到

import sys

from config import SUMATRA, DUPLEX_STRATEGY
from printer_utils import (list_printers, get_duplex, set_duplex,
                           ps_available, ps_module_report)

MODES = {"long": "TwoSidedLongEdge", "short": "TwoSidedShortEdge"}


def main():
    print("=" * 60)
    print("双面打印自检")
    print("=" * 60)

    print("\n[1] PowerShell 环境")
    print("  PrintManagement 可用:", ps_available())
    for line in ps_module_report().splitlines():
        print("   ", line)
    print("  SumatraPDF:", SUMATRA or "未找到（Office 双面将只能走打印机默认设置路线）")
    print("  DUPLEX_STRATEGY:", DUPLEX_STRATEGY)

    print("\n[2] 打印机列表")
    printers = list_printers()
    if not printers:
        print("  （没有枚举到任何打印机）")
    for p in printers:
        print("   -", p)

    name = sys.argv[1] if len(sys.argv) > 1 else None
    if not name:
        if not printers:
            print("\n用法：python test_duplex.py \"打印机名\" [long|short]")
            return
        name = printers[0]
    mode = MODES.get((sys.argv[2] if len(sys.argv) > 2 else "long").lower(),
                     "TwoSidedLongEdge")

    print("\n[3] 目标打印机: %s" % name)
    if name not in printers:
        print("  警告：这个名字不在打印机列表里，请检查引号和空格是否完全一致")

    cur = get_duplex(name)
    print("  当前 DuplexingMode:", cur)

    print("\n[4] 尝试设置为 %s ..." % mode)
    ok, detail = set_duplex(name, mode)
    print("  结果:", "成功" if ok else "失败")
    print("  详情:", detail)

    if ok:
        print("\n[5] 回读确认:", get_duplex(name))
        print("  恢复为原设置 %s ..." % cur)
        if cur:
            ok2, d2 = set_duplex(name, cur)
            print("  恢复:", "成功" if ok2 else ("失败 -> " + d2))
    else:
        print("\n[5] 未做改动，无需恢复")
        print("""
  可能原因与处理：
  a) 提示 Access is denied / 权限不足
     - 用管理员身份重新运行本脚本确认；若管理员下成功，就把打印服务也用管理员身份启动
       （但 Office COM 在提权进程里偶尔会出问题，更推荐下面的 b）
  b) 打印机是「共享连接」（名字像 \\\\主机\\共享名）
     - 本地改不了服务器上的默认设置。请在 config.py 里把
       DUPLEX_STRATEGY 改为 "pdf"，即走「Office 转 PDF → SumatraPDF 双面」
  c) 提示驱动不支持 / 回读仍是 OneSided
     - 该打印机驱动没上报双面能力，检查驱动是否为完整版（而不是 Class Driver）
""")

    print("\n" + "=" * 60)


if __name__ == "__main__":
    main()
