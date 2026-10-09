# 局域网打印服务

办公室里的一台 Windows 10 电脑接上打印机、跑起这个服务，同一局域网的同事用浏览器
打开 `http://<这台电脑的IP>:8080`，上传文件 → 选打印机/方向/份数/页码（PPT 还能选一张纸几张）→ 提交，
就能在那台打印机上出纸，并可在状态页跟踪任务。

- **后端**：Python + Flask（任务队列、状态跟踪）
- **Office 交互**：VBScript（Word 打 doc/docx/txt，Excel 打 xls/xlsx，PowerPoint 打 ppt/pptx）
- **PDF**：程序目录下的 SumatraPDF（`-print-to` 静默打印）
- **前端**：原生 HTML + CSS + JS（不依赖 CDN，断网也能用）
- **无身份验证**：仅限局域网使用，请勿暴露到公网

## 1. 目录结构

```
printserver/
├─ app.py                # Flask 路由（页面 + API）
├─ config.py             # 全部可调参数（端口、超时、保留时间…）
├─ printer_utils.py      # 枚举打印机、读取/设置双面（PowerShell）
├─ job_manager.py        # 任务表 + 串行执行队列 + 调用打印
├─ scripts/
│   ├─ print_word.vbs    # Word COM 打印（doc/docx/txt）
│   ├─ print_excel.vbs   # Excel COM 打印（xls/xlsx）
│   ├─ print_ppt.vbs     # PowerPoint COM 打印（ppt/pptx，含讲义模式）
│   ├─ lib_printq.vbs    # 打印队列校验库（按 JobId 跟踪 + 累计计数），被上面三个引入
│   └─ diagnose_duplex.ps1  # 双面问题诊断
├─ templates/            # index.html / status.html
├─ static/               # style.css / index.js / status.js
├─ uploads/              # 上传文件（运行时自动创建、自动清理）
├─ SumatraPDF.exe        # ← 需要你自己放到这里（PDF 打印用）
├─ requirements.txt
└─ run.bat               # 双击启动
```

## 2. 安装与启动

```bat
:: 1) 安装依赖（只需一次）
pip install -r requirements.txt

:: 2) 把 SumatraPDF.exe 放到 printserver 目录（或 SumatraPDF\ 子目录）

:: 3) 启动
run.bat            :: 或  python app.py
```

首次运行建议先手动执行一次 `python app.py`，确认控制台里
`SumatraPDF: ...\SumatraPDF.exe` 已被识别到。

**放通防火墙**（管理员 PowerShell 执行一次即可）：

```powershell
netsh advfirewall firewall add rule name="LAN Print" dir=in action=allow protocol=TCP localport=8080
```

同事访问：`http://<服务器IP>:8080`（服务器 IP 用 `ipconfig` 查）。

## 3. 关键实现说明

| 需求 | 实现方式 |
| --- | --- |
| doc / docx | `cscript print_word.vbs`，Word COM `Documents.Open` + `PrintOut Background:=False, Range, Copies, Pages` |
| txt | 同上，但 `Open` 时显式指定 `Format:=wdOpenFormatEncodedText, Encoding:=65001`，避免弹出"文件转换"对话框 |
| xls / xlsx | `cscript print_excel.vbs`，Excel COM 遍历工作表，设置 `PageSetup.Orientation` 后 `PrintOut From, To, Copies, ActivePrinter` |
| pdf | `SumatraPDF.exe -silent -print-to "打印机" -print-settings "页码,份数,duplex" 文件.pdf` |
| ppt / pptx | `cscript print_ppt.vbs`，PowerPoint COM：设置 `PrintOptions.PrintType/OutputType/HandoutOrder`、`PageSetup.SlideOrientation` 后 `PrintOut From, To, "", Copies, Collate` |
| 方向 | Word：`PageSetup.Orientation=0/1`；Excel：`=1/2`；PDF：借 `pypdf` 把纵向页旋转 90° 生成临时文件（横向不是必需时可不装 pypdf） |
| 份数 | Word/Excel：`Copies`；SumatraPDF：`-print-settings` 的 `2x` |
| 页码范围 | Word：`Pages:="1-3,5"`；Excel：拆成连续区间多次 `PrintOut`；PPT：按幻灯片编号拆区间多次 `PrintOut`；SumatraPDF：`-print-settings` |
| 每张纸几张幻灯片 | PowerPoint `PrintOptions.OutputType`（1/2/3/4/6/9 张）+ `HandoutOrder`（横向/纵向优先）。可选 1、2、3、4、6、9 |
| 双面 | Word/Excel 没有可靠的无人值守双面参数，改为打印前用 PowerShell `Set-PrintConfiguration -DuplexingMode` 临时改打印机默认双面，打印完恢复；PDF 额外传 `duplex` / `duplexshort` 给 SumatraPDF |
| 状态跟踪 | 提交返回 `job_id` → 跳转 `/status/<id>`，前端每 1.5 秒轮询 `/api/jobs/<id>`，展示状态徽章与实时日志 |

任务**串行执行**：Office 的 COM 自动化不适合并发，多个任务同时打开文件很容易卡死或弹窗，
所以后台只跑一个 worker，逐个处理（状态页会提示"前面还有 N 个任务"）。

## 4. 部署注意事项（重要）

1. **必须以已登录的桌面用户身份运行**（双击 `run.bat` 即可），**不要**用 Windows 服务 / 计划任务
   的"不管用户是否登录"模式运行——Office COM 在非交互会话里无法启动。
   开机自启可以放一个 `run.bat` 的快捷方式到
   `shell:startup`（`C:\Users\<用户>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup`）。
2. **Office 需已完成首次启动向导**（打开过一次 Word / Excel 并接受协议），否则 COM 会卡在激活界面。
3. 若 Excel COM 报"无法创建对象"或卡住，常见原因是缺少 Desktop 目录，可手工创建：
   `C:\Windows\SysWOW64\config\systemprofile\Desktop`（32 位 Office 在 64 位系统上）。
4. 打印任务超时（`config.py` 的 `JOB_TIMEOUT_SEC`，默认 600 秒）会被强制结束；
   如果频繁残留 WINWORD.EXE / EXCEL.EXE，可把 `KILL_OFFICE_ON_TIMEOUT` 改成 `True`
   （注意：会杀掉所有 Word/Excel 进程，包括你正在编辑的文档）。
5. 打印期间**不要在服务器上手动打开 Word/Excel 做别的事**，COM 是单实例共享的。
6. SumatraPDF 的份数参数在少数版本里写作 `2` 而非 `2x`；若发现份数不对，
   把 `config.py` 里的 `SUMATRA_COPIES_X_SUFFIX` 改成 `False`。

## 5. VBScript 编写约定（改脚本前必读）

`scripts/*.vbs` 由 `cscript.exe` 执行，它跑的是 **VBScript，不是 VBA**，三条硬规则：

1. **不能用命名参数**：`doc.PrintOut Background:=False, Copies:=2` 会直接编译失败（"缺少 ')'"）。
   必须按位置传参：
   ```vbs
   ' PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages)
   doc.PrintOut False, False, 0, "", "", "", 0, nCopies          ' 全部页
   doc.PrintOut False, False, 4, "", "", "", 0, nCopies, sPages  ' 指定页 (Range=4)
   ' PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate)
   ws.PrintOut 1, 9999, nCopies, False, sUsedPrinter, False, True
   ```
2. **不要用行尾 `_` 续行**，一条语句写在一行里。
3. **文件保存为 ANSI/ASCII + CRLF**（用记事本改完记得编码选 ANSI）。

两个脚本都内置了**降级回退**：带参数的 `PrintOut` 一旦报错，会自动退到无参数的
`PrintOut`（份数用循环实现，页码范围失效），并在日志里打印 `WARN: falling back ...`，
这样即使参数在某些 Office 版本上不兼容，也至少能把文件打出来。

## 6. 双面打印是怎么实现的

Word / Excel 的 COM 接口**没有**无人值守的双面参数，所以只有两条路可走：

| 路线 | 做法 | 优点 | 缺点 |
| --- | --- | --- | --- |
| 1. `printer-default` | 打印前用 PowerShell `Set-PrintConfiguration -DuplexingMode` 临时改打印机默认双面，打印后恢复 | 直接打 Office，不转格式 | **全局生效**（同一瞬间别人打印会受影响）；常因权限不足、打印机是共享连接、驱动不上报双面能力而失败 |
| 2. `pdf` | Office 先把 doc/docx/txt/xls/xlsx 导出成 PDF，再由 SumatraPDF 用 `-print-settings "...duplex"` 按**单个任务**下发双面参数 | 不需要任何权限，共享打印机也能用 | 多一步格式转换（排版有极小概率与 Office 直接打印不同） |

`config.py` 里 `DUPLEX_STRATEGY` 控制选哪条：

```python
DUPLEX_STRATEGY = "auto"     # 先试 1，失败自动退到 2（推荐）
DUPLEX_STRATEGY = "printer"  # 只用 1
DUPLEX_STRATEGY = "pdf"      # 只用 2（改不动打印机默认设置时直接选这个）
```

任务日志会明确打出用了哪条路线，例如 `双面路线 2：Office 转 PDF → SumatraPDF 按任务双面打印`。

**PowerPoint 的特殊之处**：

- 每条路线 PPT 都能用：路线 1（改打印机默认双面）配合 COM 直接打印；路线 2（转 PDF）配合 SumatraPDF。
- **但讲义模式（一张纸 2/3/4/6/9 张）例外**：PowerPoint 的 `ExportAsFixedFormat` 只能导出
  "每页一张幻灯片"，**无法把讲义版式导出成 PDF**。所以「讲义 + 双面」只能走路线 1；
  此时即使 `DUPLEX_STRATEGY="pdf"`，程序也会自动改试路线 1，并在日志里说明。
- 另外 PowerPoint 的 COM **没有 `ActivePrinter`**，无法像 Word/Excel 那样直接指定打印机，
  只能把目标打印机临时设为 Windows 默认打印机（`WshNetwork.SetDefaultPrinter`），打完恢复。
  若切换失败，脚本会**直接中止**并报错，而不是偷偷打到别的打印机上；
  真要接受"打到当时默认的那台"，把 `config.py` 里的 `PPT_IGNORE_PRINTER_SWITCH_FAILURE` 改成 `True`。
- PowerPoint 直接打印失败时（非讲义模式）会自动兜底转 PDF 再由 SumatraPDF 打印。

## 7. 排错

| 现象 | 原因 / 处理 |
| --- | --- |
| 日志出现 `Microsoft VBScript 编译器错误` | VBScript **不是 VBA**：不支持 `Name:=value` 命名参数（只能按位置传参）、不支持行尾 `_` 续行、文件须为 ANSI/ASCII 编码。启动时控制台会做一次语法自检（显示 `[OK]` / `[FAIL]`），并会直接指出第几行用了 `:=` 或 `_`。 |
| 拖文件到框里页面显示了，提交却说"请先选择文件" | 已在 `static/index.js` 修复：拖放时把 `dataTransfer.files` 写回 `<input type=file>`，并用 `chosenFile` 变量兜底 |
| 打印机名带空格/括号 | 无需处理：PowerShell 脚本用 `-EncodedCommand` 传递，Python 用 argv 列表传参，都会自动处理好 |
| 选了双面却打出单面，日志有 `WARN: 双面路线 1 失败` | 见下一节「双面打印排查步骤」 |
| PPT 报 `cannot switch the Windows default printer` | PowerPoint 无法指定打印机，只能临时改 Windows 默认打印机；失败就中止以免打错打印机。确认前台里选的名字和「控制面板 → 设备和打印机」里的**完全一致**（含空格、括号）。共享打印机若始终切不动，把 `PPT_IGNORE_PRINTER_SWITCH_FAILURE` 设为 `True`（会打到当时默认的那台） |
| 日志里中文文件名变成乱码（如 `鍓鏈`） | 子进程输出编码不同：cscript 用系统 ANSI（GBK），SumatraPDF 用 UTF-8。已分开解码，且上传文件落盘改名成纯 ASCII 的 `source<扩展名>`（原中文名只用于页面展示），所以日志里不会再出现中文路径 |
| 打印机提示"轻型纸，自定义"而不是 A4（Word 正常、PDF/PPT 异常） | SumatraPDF 默认按 PDF 页面自身尺寸选纸。PPT 导出的 PDF 页面尺寸 = 幻灯片尺寸（16:9 约 33.9×19.1cm），驱动没有对应纸型就退化成"自定义"。已默认加 `-print-settings paper=A4`，见下一节 |
| 导出 PDF 成功到真正打印之间卡了十几秒 | 自己启动的 PowerPoint 调 `Quit()` 要阻塞 15s+。已改为直接结束该进程（仅在"启动前没有任何 POWERPNT.EXE"时才做，不会误伤你已打开的 PowerPoint），见下一节 |
| PPT 打印任务一直卡住 | PowerPoint COM 要求演示文稿能正常打开；先手工打开一次确认没弹出修复/密码框。必要时把 `KILL_OFFICE_ON_TIMEOUT` 设为 `True` 清理残留 POWERPNT.EXE |
| PPT 双面提示 `ERROR: export pdf failed: 类型不匹配` | PowerPoint 的 `ExportAsFixedFormat` 在某些版本上会报这个错。脚本已内置 5 种导出调用 × 3 种打开方式的自动回退；仍失败请按「PPT 转 PDF 调试」跑诊断，把结果贴出来 |
| PPT 提示成功但没出纸 | 见「PPT 打印调试」。最常见是后台打印被提前退出（已修：强制前台打印 + 打印后不再立刻 Quit） |
| 已打开的 PowerPoint 窗口被关掉 | 已修：PowerPoint 是单实例 COM，脚本会判断是不是自己启动的实例，只有自己启动的才 `Quit`；不是自己启动的只关闭本次打开的文档（Word / Excel 脚本同样处理） |

### 打印完成是怎么判定的

**主路径永远是 COM**，队列扫描只是兜底/交叉验证：

| 后端 | COM 主路径（本来就有的保证） | 兜底确认 |
| --- | --- | --- |
| Word | `PrintOut Background:=False`（按位置传参的第 1 个参数）会阻塞到任务入队 | 轮询 `Application.BackgroundPrintingStatus` 直到 0；拿不到该属性时才扫队列 |
| Excel | 根本没有后台打印开关，`PrintOut` 本身就是同步的 | 直接扫队列确认 |
| PowerPoint | `PrintOptions.PrintInBackground = 0` 让 `PrintOut` 阻塞 | 扫队列确认（PPT 没有类似的状态属性） |

**队列扫描的兜底逻辑**（三种格式共用 `scripts/lib_printq.vbs`）：打印前对
`Win32_PrintJob` 快照记下 JobId 集合 → 打印后只盯**新出现的 JobId** →
它们从队列消失即确认完成；状态含 `Error/Deleted/Offline/Paperout/需要干预` 即判失败；
两次轮询之间任务出现又消失的情况，用 `Win32_PerfRawData_Spooler_PrintQueue.TotalJobsPrinted`
累计计数交叉验证。**从不只靠队列长度判断**，且每次等待都有硬上限。

相关配置（`config.py`）：

```python
OFFICE_QUEUE_WAIT_SEC = 300   # Word/Excel 最长等待秒数（超时放弃等待，不算失败）
PPT_QUEUE_WAIT_SEC    = 300   # PPT 同上
JOB_APPEAR_SEC        = 60    # 等待任务出现在队列里的宽限期
```

设为 `0` 表示完全不等待（只信 COM，速度最快）。Python 侧进程超时自动取
`max(JOB_TIMEOUT_SEC, 等待上限 + 180)`，不会在脚本等待期间把它误杀；
启动时若配置冲突会打提示。

任务日志里能直接看出走了哪条：

```
INFO: print confirmed by COM - Word background print queue is empty after 0s
INFO: print confirmed by queue - tracked job(s) left the queue after 12s
WARN: UNCONFIRMED: no print job appeared in the queue within 60s
```

出现 `UNCONFIRMED` 时任务仍算成功，但状态页会写成"已提交，但未能确认打印机已完成"。

> 注：`lib_printq.vbs` 由各脚本用 `ExecuteGlobal` 引入。万一引入失败，脚本**照常打印**
> （COM 保证不受影响），只是没有队列确认，日志会提示 `queue helper library not loaded`。

只"等一会儿再看队列长度"是不可靠的：短文档可能在两次轮询之间就已经打完，
队列长度没变化就被误判成失败。所以脚本用了三层判断：

1. **COM 侧先保证同步**：`PrintOptions.PrintInBackground = 0`（前台打印），
   `PrintOut` 会阻塞到任务真正提交进假脱机，等价于 Word 的 `Background:=False`。
2. **按 JobId 逐个跟踪**（不是按数量）：打印前对 `Win32_PrintJob` 做一次快照记下
   JobId 集合，打印后只关注**新出现的 JobId**，轮询它们的状态直到：
   - 这些 JobId **从队列中消失** → 确认完成（返回 0）；
   - 出现 `Error / Deleted / Offline / Paperout / 需要用户干预` 等状态 → 判失败（返回 2）；
   - 超过 `PPT_QUEUE_WAIT_SEC` 还在队列里 → 放弃等待，记为"已提交"并告警（返回 1）。
3. **补一次计数器交叉验证**：万一任务在两次轮询之间就出现又消失了，
   兜底读 `Win32_PerfRawData_Spooler_PrintQueue.TotalJobsPrinted`（累计已打印任务数），
   只要它变大就确认打印过（返回 0）。
   只有在"计数器不可用 + 从未看到任务"时才落到"无法确认"的告警（返回 3），
   **不会**因此把任务判成失败。

相关配置（`config.py`）：

```python
PPT_QUEUE_WAIT_SEC = 300   # 单个任务最长等待时间，超时放弃等待（不算失败）
PPT_JOB_APPEAR_SEC = 60    # 等待任务出现在队列里的宽限期
```

Python 侧的进程超时会自动取 `max(JOB_TIMEOUT_SEC, PPT_QUEUE_WAIT_SEC + 180)`，
不会在脚本等待期间把它误杀。

### 日志编码 / 纸张 / 延迟：三个常见坑及处理

**1）日志乱码**

控制台输出的编码，两个子进程是不一样的：

| 子进程 | 输出编码 | 处理方式 |
| --- | --- | --- |
| `cscript.exe`（Word/Excel/PPT 脚本） | 系统 ANSI 代码页（简体中文 = GBK） | 用 `CONSOLE_ENCODING` 解码 |
| `SumatraPDF.exe` | UTF-8 | 用 `SUMATRA_ENCODING` 解码 |

之前统一按 GBK 解码，于是 SumatraPDF 那几行里的中文变成 `鍓鏈`（典型的"UTF-8 当 GBK 读"）。
现在按命令分别解码。

更根本的一点：**上传的文件落盘时改名为 `source<扩展名>`**（纯 ASCII），
原始中文文件名只保存在任务记录里用于页面展示，不参与任何命令行。
这样日志里根本不会出现中文路径 —— 顺带也避免了 Office COM 对中文路径偶发的"类型不匹配"。

**2）纸张变成"自定义 / 轻型纸"**

SumatraPDF 默认按 **PDF 页面自身的尺寸**去选纸型。而 PowerPoint 导出的 PDF，
页面尺寸等于幻灯片尺寸（16:9 时约 33.9 × 19.1 cm），驱动里没有这个纸型，
就退化成"自定义"，打印机于是报"轻型纸，自定义"。Word 走的是 Office COM，
直接读驱动默认 DEVMODE（A4），所以 Word 正常。

解决办法是显式指定纸张（`config.py`）：

```python
SUMATRA_PAPER = "A4"      # A2/A3/A4/A5/A6/letter/legal，或驱动报告的纸型名
SUMATRA_BIN   = ""        # "" = 不指定；"auto" = 让打印机按页面尺寸选纸盒
SUMATRA_SCALE = ""        # "" = 用默认；内容没缩放到纸上时改成 "fit"
```

想看驱动实际报告哪些纸型名，跑：

```bat
SumatraPDF.exe -list-printers
```

如果这台机器上的 SumatraPDF 版本不认 `paper=`（打印直接失败），
程序会**自动去掉这些参数重试一次**并在日志里提示，届时把 `SUMATRA_PAPER` 改成 `""` 即可。

**3）导出 PDF 到实际打印之间的十几秒延迟**

实测：**自己启动的** PowerPoint 实例，调用 `Quit()` 会阻塞 15 秒以上才返回
（Word / Excel 没有这个问题）。这段耗时正好落在"导出成功 → 任务返回"之间。

已改为：确认导出/打印完成、文档已关闭后，直接结束那个 PowerPoint 进程，秒退。

```vbs
Sub PptShutdown(pptApp, bWeStarted, bFast)
    If Not bWeStarted Then Exit Sub      ' 只有自己启动的才处理
    ...
    oSh.Run "cmd /c taskkill /F /IM POWERPNT.EXE", 0, False   ' 异步，不等它
End Sub
```

安全边界：只有当脚本启动前**没有任何** `POWERPNT.EXE` 进程时才会这么做，
所以不会误伤你已经打开的 PowerPoint；如果你本来就开着 PowerPoint，脚本只关闭本次打开的文档。
想改回温和退出，把 `config.py` 里的 `PPT_FAST_QUIT` 改成 `False`。

另外每步耗时超过 5 秒时，日志会多打一行 `（该步骤耗时 X.X 秒）`，方便继续定位瓶颈。

### PPT 转 PDF 调试

双面路线 2（转 PDF）依赖 PowerPoint 的导出接口，而这个接口在不同版本上表现不一致。
`scripts/print_ppt.vbs` 已内置**自动回退**，按下面的顺序依次尝试（**SaveAs 优先**）：

| 顺序 | 导出调用 | |
| --- | --- | --- |
| 1 | mode 4 | `SaveAs(Path, 32)` — ppSaveAsPDF ← **实测可用，默认首选** |
| 2 | mode 5 | `SaveAs(Path, 32, 0)` — ppSaveAsPDF + 不嵌字体 |
| 3 | mode 1 | `ExportAsFixedFormat(Path, 2, 2)` — PDF + 打印意图 |
| 4 | mode 2 | `ExportAsFixedFormat(Path, 2, 1)` — PDF + 屏幕意图 |
| 5 | mode 3 | `ExportAsFixedFormat(Path, 2, 2, 0, 1, 1, 0)` — 显式补齐可选参数 |

如果无窗口只读打开下全部失败，会**带窗口**重开再试一轮，最后**读写方式**重开再试 SaveAs
（共 3 种打开方式 × 5 种调用）。每次失败都打一行 `WARN: mode N raised 0x800A000D ...`，
成功则打 `OK: exported pdf [...] via mode 4 (windowless)`，能看出赢在哪条。

**实测结论（Office 16.0 简体中文）**：这台机器上 `ExportAsFixedFormat` 的三种参数组合
**全部**报 `0x800A000D 类型不匹配`，且与打开方式（无窗口/带窗口、只读/读写）无关 ——
说明不是打开方式的问题，而是这个 COM 方法在后绑定（VBScript）调用时的参数封送有问题。
而 `SaveAs(Path, ppSaveAsPDF)` 走另一条代码路径，工作正常、内容正确。
所以默认已改为 **SaveAs 优先**，`ExportAsFixedFormat` 降级为兜底（别的机器上它可能才是唯一可用的）。

另外导出文件名已固定为 `uploads/<任务ID>/export.pdf`（纯 ASCII），
因为含中文/空格的路径本身就是触发"类型不匹配"的常见原因之一。
导出成功后关闭前会先置 `pres.Saved = True`，避免 PowerPoint 弹"是否保存更改"把脚本挂住
（PowerPoint 没有 `DisplayAlerts` 可以关）。

想直接看这台机器上哪种调用能用，跑诊断脚本（不打印，只试导出）：

```bat
cscript //nologo scripts\test_ppt_pdf.vbs "D:\tmp\demo.pptx" "D:\tmp\out.pdf"
```

它会输出一张表，例如：

```
--- open mode 1: windowless, read-only ---
  opened, slides = 3
  [1] ExportAsFixedFormat(Path,2,2) -> ERROR 0x800A000D Type mismatch
  [2] ExportAsFixedFormat(Path,2,1) -> ERROR 0x800A000D Type mismatch
  [3] ExportAsFixedFormat(Path,2,2,0,1,1,0) -> ERROR 0x800A000D Type mismatch
  [4] SaveAs(Path,32) -> OK
RESULT: SUCCESS
```

把这段输出贴出来，就能精确锁定要保留哪种调用。

### 手工调试（Word / Excel / PPT 都适用）

三个脚本都能脱离服务单独跑，参数顺序见各自文件头部的注释：

```bat
cscript //nologo scripts\print_word.vbs   "D:\tmp\a.docx" "打印机名" 0 1 "" "" 300 60
cscript //nologo scripts\print_excel.vbs  "D:\tmp\b.xlsx" "打印机名" 1 1 ""  0 "" 300 60
cscript //nologo scripts\print_ppt.vbs    "D:\tmp\c.pptx" "打印机名" 0 1 ""  4 0
```

末尾两个数字是「最长等待秒数」和「任务出现宽限秒数」，都可省略。
想确认完成判定是否工作，看日志里是否有 `print confirmed by COM` 或
`print confirmed by queue`；只有 `UNCONFIRMED` 时才需要排查。

### PPT 打印调试

在服务器上手工跑一次脚本，把参数直接给它（顺序：文件 打印机 方向 份数 页码 每页张数 排列）：

```bat
cscript //nologo scripts\print_ppt.vbs "D:\tmp\test.pptx" "NPIDA8E61 (HP LaserJet MFP M227fdw)" 0 1 "" 4 0
```

（末尾可再加两个数字：最长等待秒数、任务出现宽限秒数，默认 300 / 60）

正常输出里应当出现这几行，**逐行核对**：

```
INFO: PowerPoint <版本>, <页数> slides, perPage=4
INFO: OutputType readback = 8 (expected 8)      <- 讲义版式是否真的生效
INFO: queue before: N job(s), spooler printed counter = M
OK: slides 1-N sent to printer
INFO: new print job 7 -> 新建 Microsoft PowerPoint 演示文稿 [Spooling]
INFO: print job confirmed finished - tracked job(s) left the queue after Xs
```

- 出现 `print job confirmed finished - spooler printed counter went M -> M+1`
  → 任务太快、在两次轮询之间就打完了，靠计数器确认，同样算成功。
- 出现 `WARN: no print job appeared in the queue within ...s` → 可能是打完了也可能是被拒收：
  检查打印机是否脱机/暂停/缺纸，或用 `Get-PrintJob -PrinterName "..."` 看队列。
- 出现 `ERROR: print job reported a problem - print job N reported [Error ...]` → 打印机侧报错，
  任务会被标记为失败。
- 出现 `ERROR: handout layout did NOT apply` → 讲义版式没设置成功，已中止（不会打到错误版式）。
- 只有 1 张/每页能打、讲义打不出 → PowerPoint 的讲义母版异常，用 PowerPoint 手动打开该文档，在「视图 → 讲义母版」里确认版式正常。

**关于 PowerPoint 的两个坑（已修，别改回去）**：

1. `PrintOptions` **没有** `PrintType` 属性 —— 给它赋值会报"对象不支持此属性或方法"，讲义版式不生效。
   正确做法是只设 `OutputType`（1/2/3/4/6/9 张分别对应值 1/2/3/8/4/9）。
2. `PrintInBackground` 默认为「后台打印」，`PrintOut` 立刻返回，此时若退出 PowerPoint
   会把排队中的任务一起杀掉 → **日志显示成功、实际不出纸**。现已强制前台打印。

### 双面打印排查步骤

在打印服务器上（另开一个命令行窗口）运行：

```bat
python test_duplex.py "NPIDA8E61 (HP LaserJet MFP M227fdw)" long
```

它会直接打出**失败的真实原因**（以前日志里只写了"PrintManagement 不可用？"，太笼统）。
也可以用 PowerShell 版看更细的信息（驱动名、端口、是否共享）：

```bat
powershell -ExecutionPolicy Bypass -File .\scripts\diagnose_duplex.ps1 "NPIDA8E61 (HP LaserJet MFP M227fdw)"
```

按打印出的原因对症处理：

| 报错 | 处理 |
| --- | --- |
| `Access is denied` / 权限不足 | 用管理员身份再跑一次 `test_duplex.py` 验证；若管理员下成功，就把打印服务改为管理员启动。**更推荐**：直接把 `DUPLEX_STRATEGY` 改成 `"pdf"`，绕开权限问题 |
| `回读仍是 OneSided`（设置成功但没生效） | 驱动没上报双面能力，装完整版驱动（别用 Class Driver / 通用驱动）。临时方案同上：改用 `pdf` 路线 |
| 打印机是共享连接（名字像 `\\主机\共享名`，或 `diagnose_duplex.ps1` 显示 Shared=True） | 本地无权改服务器上的默认设置，**只能用 `pdf` 路线** |
| 两条路线都失败 | `pdf` 路线还需要 `SumatraPDF.exe` 在程序目录；确认日志里没有 `转 PDF 失败` |

## 8. 已知限制

- 双面：路线 1 依赖 PowerShell `PrintManagement` 模块 + 打印机驱动上报双面能力；
  两者任一不满足就只能用路线 2（转 PDF）。
- 路线 2（转 PDF）对 Excel 会导出整个工作簿；勾选「仅打印第一张工作表」时脚本会
  只导出第 1 张表，但其余情况下与"逐张表打印"的语义不完全一致。
- PDF 强制横向需要 `pypdf`；未安装时按 PDF 原始页面方向打印。
- Excel 的页码范围是"每个工作表分别应用同一范围"，且只能是连续区间。
- 双面路线 1 修改的是打印机**全局**默认设置，脚本会在打印后立刻恢复原值，
  但同一瞬间其他人的打印任务可能受影响（概率极低）；用路线 2（`pdf`）则完全无此问题。
- 无身份验证，任何能访问该端口的人都能打印；请只在可信局域网内使用。
