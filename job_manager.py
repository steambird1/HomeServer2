# -*- coding: utf-8 -*-
"""打印任务管理：内存任务表 + 单线程串行执行队列。

之所以串行：Office COM（Word / Excel）本身不适合并发调用，
多个实例同时打印极易卡死或弹出对话框，因此这里用一个后台 worker 挨个处理。
"""
import os
import re
import uuid
import time
import queue
import shutil
import threading
import subprocess
from datetime import datetime, timedelta

from config import (UPLOAD_DIR, SCRIPTS_DIR, SUMATRA, CSCRIPT, CONSOLE_ENCODING,
                    JOB_TIMEOUT_SEC, JOB_RETENTION_HOURS, MAX_LOG_LINES,
                    SUMATRA_COPIES_X_SUFFIX, KILL_OFFICE_ON_TIMEOUT, DUPLEX_STRATEGY,
                    WORD_EXT, EXCEL_EXT, PPT_EXT, SLIDES_PER_PAGE_CHOICES,
                    SLIDES_PER_PAGE_DEFAULT, PPT_IGNORE_PRINTER_SWITCH_FAILURE,
                    PPT_QUEUE_WAIT_SEC, PPT_JOB_APPEAR_SEC,
                    OFFICE_QUEUE_WAIT_SEC, JOB_APPEAR_SEC,
                    SUMATRA_ENCODING, SUMATRA_PAPER, SUMATRA_BIN, SUMATRA_SCALE,
                    PPT_FAST_QUIT)
from printer_utils import get_duplex, set_duplex

JOBS = {}
_LOCK = threading.Lock()
_QUEUE = queue.Queue()

DUPLEX_MAP = {
    "one": "OneSided",
    "long": "TwoSidedLongEdge",
    "short": "TwoSidedShortEdge",
}

STATUS_TEXT = {
    "queued": "排队中",
    "printing": "正在打印",
    "success": "已发送到打印机",
    "error": "失败",
    "cancelled": "已取消",
}


# ---------------------------------------------------------------- 工具函数
def _log(job, msg):
    line = "[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg)
    if len(job["log"]) < MAX_LOG_LINES:
        job["log"].append(line)


def parse_pages(text):
    """把 "1-3,5,8-" 解析为 [(1,3),(5,5),(8,9999)]；空返回 []。"""
    text = (text or "").strip()
    if not text:
        return []
    out = []
    for part in re.split(r"[,\s]+", text):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.append((int(a) if a else 1, int(b) if b else 99999))
        else:
            n = int(part)
            out.append((n, n))
    return out


def _run(cmd, job, timeout=JOB_TIMEOUT_SEC, cwd=None, encoding=None):
    """执行命令并实时把输出写进任务日志，返回 (returncode, 全部输出)

    encoding 用于解码子进程输出：cscript 用系统 ANSI 代码页（CONSOLE_ENCODING），
    SumatraPDF 用 UTF-8。两者混用会直接导致日志乱码。
    """
    enc = encoding or CONSOLE_ENCODING
    _log(job, "$ " + " ".join(cmd))
    t0 = time.time()
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                cwd=cwd or SCRIPTS_DIR)
    except Exception as e:
        _log(job, "ERROR: 无法启动进程: %s" % e)
        return -1, ""

    lines = []

    def reader():
        try:
            for raw in iter(proc.stdout.readline, b""):
                s = raw.decode(enc, errors="replace").rstrip()
                if s:
                    lines.append(s)
                    _log(job, s)
        except Exception:
            pass

    t = threading.Thread(target=reader, daemon=True)
    t.start()
    try:
        proc.wait(timeout=timeout)
        rc = proc.returncode
    except subprocess.TimeoutExpired:
        try:
            proc.kill()
        except Exception:
            pass
        _log(job, "ERROR: 执行超时（%d 秒），已强制结束" % timeout)
        rc = -1
    t.join(5)

    elapsed = time.time() - t0
    if elapsed > 5:
        _log(job, "（该步骤耗时 %.1f 秒）" % elapsed)

    out = "\n".join(lines)
    if "编译器错误" in out or "编译错误" in out:
        _log(job, "HINT: VBScript 语法错误 —— 常见原因：用了 Name:=value 命名参数"
                  "（VBS 不支持，只能按位置传参）、脚本不是 ANSI/ASCII 编码、或行尾续行符“_”")
    return rc, out


# ---------------------------------------------------------------- 打印实现
def _print_word(job, path, printer, orient, copies, pages_raw, wait_sec=None):
    vbs = os.path.join(SCRIPTS_DIR, "print_word.vbs")
    wait_sec = OFFICE_QUEUE_WAIT_SEC if wait_sec is None else wait_sec
    # 参数顺序：文件 打印机 方向 份数 页码 [导出PDF] [最长等待] [出现宽限]
    cmd = [CSCRIPT, "//nologo", vbs, path, printer,
           "1" if orient == "landscape" else "0", str(copies), pages_raw, "",
           str(wait_sec), str(JOB_APPEAR_SEC)]
    return _run(cmd, job, timeout=_timeout_for(wait_sec))


def _timeout_for(wait_sec):
    """进程超时必须大于脚本内部的等待上限，否则会在等待期间被误杀。"""
    return max(JOB_TIMEOUT_SEC, int(wait_sec) + 180)


def _print_excel(job, path, printer, orient, copies, ranges, first_only, wait_sec=None):
    vbs = os.path.join(SCRIPTS_DIR, "print_excel.vbs")
    wait_sec = OFFICE_QUEUE_WAIT_SEC if wait_sec is None else wait_sec
    # 参数顺序：文件 打印机 方向 份数 页码 仅首表 [导出PDF] [最长等待] [出现宽限]
    cmd = [CSCRIPT, "//nologo", vbs, path, printer,
           "2" if orient == "landscape" else "1", str(copies), ranges,
           "1" if first_only else "0", "",
           str(wait_sec), str(JOB_APPEAR_SEC)]
    return _run(cmd, job, timeout=_timeout_for(wait_sec))


def _prepare_pdf(path, orient, job):
    """PDF 横向：若页面本身不是横向，则旋转 90 度生成临时文件（需要 pypdf）。"""
    if orient != "landscape":
        return path
    Reader = Writer = None
    try:
        from pypdf import PdfReader as Reader, PdfWriter as Writer
    except Exception:
        try:
            from PyPDF2 import PdfReader as Reader, PdfWriter as Writer
        except Exception:
            Reader = None
    if Reader is None:
        _log(job, "WARN: 未安装 pypdf，无法强制横向；请 `pip install pypdf`，"
                  "或改用打印机驱动里的方向设置")
        return path

    try:
        reader = Reader(path)
        writer = Writer()
        rotated = 0
        for page in reader.pages:
            box = page.mediabox
            if box.width < box.height:
                page.rotate(90)
                rotated += 1
            writer.add_page(page)
        if rotated == 0:
            return path
        out = os.path.splitext(path)[0] + "_landscape.pdf"
        with open(out, "wb") as f:
            writer.write(f)
        _log(job, "已生成横向版本（旋转 %d 页）" % rotated)
        return out
    except Exception as e:
        _log(job, "WARN: PDF 旋转失败，按原样打印: %s" % e)
        return path


def _print_ppt(job, path, printer, orient, copies, ranges, per_page, handout_order):
    vbs = os.path.join(SCRIPTS_DIR, "print_ppt.vbs")
    # 参数顺序：文件 打印机 方向 份数 页码 每页张数 排列 [导出PDF] [最长等待] [出现宽限] [快速退出]
    cmd = [CSCRIPT, "//nologo", vbs, path, printer,
           "1" if orient == "landscape" else "0", str(copies), ranges,
           str(per_page), "1" if handout_order == "vertical" else "0", "",
           str(PPT_QUEUE_WAIT_SEC), str(PPT_JOB_APPEAR_SEC),
           "1" if PPT_FAST_QUIT else "0"]
    # 脚本内部最长会等 PPT_QUEUE_WAIT_SEC 秒，进程超时必须比它更长，否则会被误杀
    return _run(cmd, job, timeout=_timeout_for(PPT_QUEUE_WAIT_SEC))


def _export_pdf(job, path, ext, orient, first_only, per_page=1):
    """用 Office COM 把 doc/docx/txt/xls/xlsx/ppt/pptx 导出为 PDF，成功返回路径，失败返回 None。

    这是双面打印的“备用路线”：SumatraPDF 可以对单个任务下发 duplex 参数，
    而 Word/Excel/PowerPoint 的 COM 接口没有无人值守的双面参数，只能改打印机全局默认设置。

    注意：PowerPoint 无法把“讲义版式（一张纸 N 张幻灯片）”导出成 PDF，
    所以 per_page > 1 时不要走这条路线。
    """
    # 输出文件名固定为 export.pdf：Office（尤其 PowerPoint 的 ExportAsFixedFormat）
    # 对含中文/空格的路径偶尔会报“类型不匹配”，用纯 ASCII 名字可完全绕开这类问题
    pdf = os.path.join(os.path.dirname(path), "export.pdf")
    try:
        if os.path.exists(pdf):
            os.remove(pdf)
    except OSError:
        pass

    if ext in EXCEL_EXT:
        vbs = os.path.join(SCRIPTS_DIR, "print_excel.vbs")
        cmd = [CSCRIPT, "//nologo", vbs, path, "",
               "2" if orient == "landscape" else "1", "1", "",
               "1" if first_only else "0", pdf]
    elif ext in PPT_EXT:
        vbs = os.path.join(SCRIPTS_DIR, "print_ppt.vbs")
        cmd = [CSCRIPT, "//nologo", vbs, path, "",
               "1" if orient == "landscape" else "0", "1", "",
               str(per_page), "0", pdf]
    else:
        vbs = os.path.join(SCRIPTS_DIR, "print_word.vbs")
        cmd = [CSCRIPT, "//nologo", vbs, path, "",
               "1" if orient == "landscape" else "0", "1", "", pdf]

    rc, _out = _run(cmd, job)
    if rc == 0 and os.path.exists(pdf):
        _log(job, "已转 PDF：%s" % os.path.basename(pdf))
        return pdf
    _log(job, "WARN: 转 PDF 失败（退出码 %s）" % rc)
    return None


def _sumatra_print(job, path, printer, settings):
    cmd = [SUMATRA, "-silent", "-print-to", printer]
    if settings:
        cmd += ["-print-settings", ",".join(settings)]
    cmd.append(path)
    return _run(cmd, job, timeout=JOB_TIMEOUT_SEC, cwd=os.path.dirname(SUMATRA),
                encoding=SUMATRA_ENCODING)


def _print_pdf(job, path, printer, orient, copies, pages_raw, duplex_mode):
    if SUMATRA is None:
        return -1, "ERROR: 未找到 SumatraPDF.exe，请放到程序目录"

    path = _prepare_pdf(path, orient, job)

    settings = []
    if pages_raw:
        settings.append(pages_raw)
    if copies > 1:
        settings.append(("%dx" % copies) if SUMATRA_COPIES_X_SUFFIX else str(copies))
    if duplex_mode == "long":
        settings.append("duplex")
    elif duplex_mode == "short":
        settings.append("duplexshort")

    # 纸张 / 纸盒 / 缩放：强制纸张是解决“自定义纸张”的关键，
    # 但老版本 SumatraPDF 可能不认这几个参数，所以失败时自动去掉再试一次。
    extra = []
    if SUMATRA_PAPER:
        extra.append("paper=" + SUMATRA_PAPER)
    if SUMATRA_BIN:
        extra.append("bin=" + SUMATRA_BIN)
    if SUMATRA_SCALE:
        extra.append(SUMATRA_SCALE)

    rc, out = _sumatra_print(job, path, printer, settings + extra)
    if rc != 0 and extra:
        _log(job, "WARN: 带纸张/纸盒参数打印失败，去掉这些参数重试一次")
        rc, out = _sumatra_print(job, path, printer, settings)
        if rc == 0:
            _log(job, "HINT: 这台机器上的 SumatraPDF 可能不支持 %s 参数，"
                      "请把 config.py 里的对应项改为空字符串" % ",".join(extra))
    return rc, out


# ---------------------------------------------------------------- 任务执行
def _execute(job):
    printer = job["printer"]
    path = job["saved_path"]
    ext = job["ext"]
    orient = job["orientation"]
    copies = job["copies"]
    pages_raw = (job["pages"] or "").strip()
    ranges = ";".join("%d-%d" % r for r in parse_pages(pages_raw))
    duplex_mode = job["duplex"]
    want_duplex = DUPLEX_MAP.get(duplex_mode, "OneSided")

    per_page = job.get("slides_per_page", SLIDES_PER_PAGE_DEFAULT) or 1
    handout_order = job.get("handout_order", "horizontal")
    is_ppt = ext in PPT_EXT
    # 讲义（一张纸 N 张幻灯片）无法导出成 PDF 保留版式，所以只能走路线 1
    handout = is_ppt and per_page > 1

    saved = None       # 打印机原来的双面设置（用于事后恢复）
    pdf_path = None    # PDF 路线生成的中间文件
    duplex_by = None   # "printer-default" / "pdf" / "pdf-native" / None

    # ---------------- 双面处理 ----------------
    if want_duplex != "OneSided":
        if handout:
            _log(job, "讲义模式（每张 %d 张）：PowerPoint 无法把讲义版式导出为 PDF，"
                      "双面只能走路线 1（改打印机默认双面）" % per_page)

        # 路线 1：改打印机默认双面设置（全局生效，打印后恢复）
        # 讲义模式下即使 DUPLEX_STRATEGY="pdf" 也必须试一次，否则完全没有双面可用
        if DUPLEX_STRATEGY in ("auto", "printer") or handout:
            saved = get_duplex(printer)
            _log(job, "双面路线 1：改打印机默认设置（当前 %s → 目标 %s）" % (saved, want_duplex))
            ok, detail = set_duplex(printer, want_duplex)
            if ok:
                duplex_by = "printer-default"
                _log(job, "双面路线 1 生效：%s" % detail)
            else:
                _log(job, "WARN: 双面路线 1 失败：%s" % detail)

        # 路线 2：先用 Office 转 PDF，再由 SumatraPDF 对“单个任务”下发双面参数
        pdf_possible = bool(SUMATRA) and ext != ".pdf" and not handout
        if duplex_by is None and pdf_possible and DUPLEX_STRATEGY in ("auto", "pdf"):
            _log(job, "双面路线 2：Office 转 PDF → SumatraPDF 按任务双面打印")
            pdf_path = _export_pdf(job, path, ext, orient, job["first_sheet_only"],
                                   per_page=per_page)
            if pdf_path:
                duplex_by = "pdf"
            else:
                _log(job, "WARN: 双面路线 2 不可用（转 PDF 失败）")
                if is_ppt:
                    _log(job, "HINT: 排查 PPT 转 PDF，请在服务器上运行："
                              "cscript //nologo scripts\\test_ppt_pdf.vbs \"<这个 pptx 文件>\"")

        if duplex_by is None:
            if ext == ".pdf" and SUMATRA:
                duplex_by = "pdf-native"
                _log(job, "PDF 文件，双面参数直接由 SumatraPDF 下发")
            elif handout:
                _log(job, "WARN: 讲义 + 双面未能生效（改不动打印机默认双面），"
                          "本次按打印机当前设置打印")
            else:
                _log(job, "WARN: 双面未能生效，本次将按打印机当前设置打印（很可能是单面）")

    try:
        _log(job, "开始打印：%s -> %s（%s，%d 份，双面=%s%s）"
             % (job["filename"], printer, orient, copies, duplex_by or "无",
                ("，每张 %d 张幻灯片" % per_page) if handout else ""))

        if ext == ".pdf":
            rc, out = _print_pdf(job, path, printer, orient, copies, pages_raw, duplex_mode)
        elif pdf_path:
            rc, out = _print_pdf(job, pdf_path, printer, orient, copies, pages_raw, duplex_mode)
        elif is_ppt:
            rc, out = _print_ppt(job, path, printer, orient, copies, ranges,
                                 per_page, handout_order)
            # 兜底 1：切不动默认打印机时，按配置改用「当前默认打印机」重试一次
            if rc != 0 and PPT_IGNORE_PRINTER_SWITCH_FAILURE:
                _log(job, "WARN: 已按 PPT_IGNORE_PRINTER_SWITCH_FAILURE 配置，"
                          "改用当前默认打印机重试（可能不是你选的那台）")
                rc, out = _print_ppt(job, path, "", orient, copies, ranges,
                                     per_page, handout_order)
            # 兜底 2：直接打印失败（例如默认打印机切换不了）时，非讲义模式可改走 PDF；
            #         讲义模式下 PDF 无法保留版式，只能提示改用每页 1 张或换 PDF 路线
            if rc != 0 and handout:
                _log(job, "HINT: 讲义模式的双面只能靠打印机默认设置；若想改用「转 PDF → SumatraPDF」"
                          "路线，请把每张纸改回 1 张幻灯片")
            if rc != 0 and SUMATRA and not handout:
                _log(job, "WARN: PowerPoint 直接打印失败，尝试转 PDF 后由 SumatraPDF 打印")
                alt = _export_pdf(job, path, ext, orient, False, per_page=1)
                if alt:
                    rc, out = _print_pdf(job, alt, printer, orient, copies,
                                         pages_raw, duplex_mode)
        elif ext in EXCEL_EXT:
            rc, out = _print_excel(job, path, printer, orient, copies, ranges,
                                   job["first_sheet_only"])
        else:
            rc, out = _print_word(job, path, printer, orient, copies, pages_raw)

        if rc == 0:
            job["status"] = "success"
            # 脚本未能确认打印任务真的完成时，会打一行 "WARN: UNCONFIRMED: ..."
            unconfirmed = any("UNCONFIRMED" in line for line in job["log"])
            job["message"] = ("已提交，但未能确认打印机已完成（请查看日志并到打印机旁确认）"
                              if unconfirmed else "已发送到打印机，请到打印机旁取件")
            _log(job, "完成")
        else:
            last_err = ""
            for line in reversed(job["log"]):
                if "ERROR:" in line:
                    last_err = line.split("ERROR:", 1)[1].strip()
                    break
            job["status"] = "error"
            job["message"] = ("打印失败：%s" % last_err) if last_err \
                else ("打印失败（退出码 %s），请查看日志" % rc)
            _log(job, "失败，退出码 %s" % rc)
            if KILL_OFFICE_ON_TIMEOUT and ext not in (".pdf",):
                _kill_office(ext)
    except Exception as e:
        job["status"] = "error"
        job["message"] = "打印出错：%s" % e
        _log(job, "异常：%s" % e)
    finally:
        if duplex_by == "printer-default" and saved:
            ok, detail = set_duplex(printer, saved)
            _log(job, "已恢复打印机双面设置为 %s（%s）" % (saved, "成功" if ok else "失败：" + detail))
        job["duplex_by"] = duplex_by
        job["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _kill_office(ext):
    """残留 Office 进程清理（默认只在开启 KILL_OFFICE_ON_TIMEOUT 时调用）。"""
    if ext in EXCEL_EXT:
        img = "EXCEL.EXE"
    elif ext in PPT_EXT:
        img = "POWERPNT.EXE"
    elif ext in WORD_EXT:
        img = "WINWORD.EXE"
    else:
        return
    try:
        subprocess.run(["taskkill", "/F", "/IM", img],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15)
    except Exception:
        pass


def _worker():
    while True:
        job_id = _QUEUE.get()
        try:
            job = JOBS.get(job_id)
            if job is None or job["status"] == "cancelled":
                continue
            job["status"] = "printing"
            job["message"] = "正在调用 Office / SumatraPDF 打印…"
            _execute(job)
        finally:
            _QUEUE.task_done()


threading.Thread(target=_worker, daemon=True).start()


def _cleaner():
    while True:
        time.sleep(600)
        try:
            deadline = datetime.now() - timedelta(hours=JOB_RETENTION_HOURS)
            with _LOCK:
                for job_id, job in list(JOBS.items()):
                    if job.get("finished_at") and \
                       datetime.strptime(job["finished_at"], "%Y-%m-%d %H:%M:%S") < deadline:
                        try:
                            shutil.rmtree(os.path.dirname(job["saved_path"]), ignore_errors=True)
                        except Exception:
                            pass
                        JOBS.pop(job_id, None)
        except Exception:
            pass


threading.Thread(target=_cleaner, daemon=True).start()


# ---------------------------------------------------------------- 对外接口
def create_job(form):
    """form 为已校验的参数 dict，返回 job_id"""
    return create_job_with_id(uuid.uuid4().hex[:12], form)


def create_job_with_id(job_id, form):
    """用外部指定的 job_id 建任务（保存文件需要先知道 id），返回 job_id"""
    job = {
        "id": job_id,
        "filename": form["filename"],
        "ext": form["ext"],
        "saved_path": form["saved_path"],
        "printer": form["printer"],
        "orientation": form["orientation"],
        "copies": form["copies"],
        "pages": form["pages"],
        "duplex": form["duplex"],
        "first_sheet_only": form["first_sheet_only"],
        "slides_per_page": form.get("slides_per_page", SLIDES_PER_PAGE_DEFAULT),
        "handout_order": form.get("handout_order", "horizontal"),
        "status": "queued",
        "message": "排队中…",
        "log": [],
        "duplex_by": None,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "finished_at": None,
    }
    with _LOCK:
        JOBS[job_id] = job
    _QUEUE.put(job_id)
    return job_id


def get_job(job_id):
    with _LOCK:
        job = JOBS.get(job_id)
        if not job:
            return None
        pos = 0
        if job["status"] == "queued":
            pos = max(0, _QUEUE.qsize())
        return {
            "id": job["id"],
            "filename": job["filename"],
            "printer": job["printer"],
            "orientation": "横向" if job["orientation"] == "landscape" else "纵向",
            "copies": job["copies"],
            "pages": job["pages"] or "全部",
            "duplex": {"one": "单面", "long": "双面（长边）", "short": "双面（短边）"}.get(job["duplex"], "单面"),
            "duplex_by": job.get("duplex_by"),
            "first_sheet_only": job["first_sheet_only"],
            "slides_per_page": job.get("slides_per_page", SLIDES_PER_PAGE_DEFAULT),
            "handout_order": ("纵向优先" if job.get("handout_order") == "vertical"
                              else "横向优先"),
            "is_ppt": job["ext"] in PPT_EXT,
            "status": job["status"],
            "status_text": STATUS_TEXT.get(job["status"], job["status"]),
            "message": job["message"],
            "queue_pos": pos,
            "log": list(job["log"]),
            "created_at": job["created_at"],
            "finished_at": job["finished_at"],
        }


def cancel_job(job_id):
    with _LOCK:
        job = JOBS.get(job_id)
        if job and job["status"] == "queued":
            job["status"] = "cancelled"
            job["message"] = "任务已取消"
            job["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return True
    return False


def ensure_upload_dir(job_id):
    d = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(d, exist_ok=True)
    return d
