# -*- coding: utf-8 -*-
"""局域网打印服务（Flask）

- 前端提交文件与打印参数
- 后端把任务放进串行队列，由 worker 调用 VBScript（Word / Excel）或 SumatraPDF（PDF）打印
- 状态页通过 /api/jobs/<id> 轮询任务进度

运行：
    python app.py
访问：
    http://<服务器IP>:8080
"""
import os
import re
import uuid
import threading

from flask import Flask, request, jsonify, render_template

from config import (BASE_DIR, HOST, PORT, MAX_CONTENT_MB, ALLOWED_EXT, SUMATRA,
                    CONSOLE_ENCODING, PPT_EXT, SLIDES_PER_PAGE_CHOICES,
                    SLIDES_PER_PAGE_DEFAULT, SUMATRA_ENCODING, SUMATRA_PAPER)
from printer_utils import list_printers, get_default_printer, ps_available
import job_manager as jm

app = Flask(__name__, static_folder="static", template_folder="templates")
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_MB * 1024 * 1024
try:
    app.json.ensure_ascii = False        # Flask >= 2.3
except AttributeError:
    app.config["JSON_AS_ASCII"] = False  # Flask 旧版本

INVALID_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')

_PS_OK = None


def ps_available_cached():
    """PrintManagement 模块是否可用（决定双面能否生效），只检查一次。"""
    global _PS_OK
    if _PS_OK is None:
        _PS_OK = bool(ps_available())
    return _PS_OK


def safe_filename(name):
    """仅用于展示的安全文件名（保留中文），不用于落盘。"""
    name = os.path.basename(name or "")
    name = INVALID_CHARS.sub("_", name).strip().strip(". ")
    return name[:120] or "file"


# ------------------------------------------------------------------ 页面
@app.route("/")
def index():
    return render_template("index.html", max_mb=MAX_CONTENT_MB,
                           sumatra=bool(SUMATRA), ps=ps_available_cached())


@app.route("/status/<job_id>")
def status_page(job_id):
    return render_template("status.html", job_id=job_id)


# ------------------------------------------------------------------ API
@app.route("/api/printers")
def api_printers():
    return jsonify({
        "printers": list_printers(),
        "default": get_default_printer(),
        "sumatra": bool(SUMATRA),
    })


@app.route("/api/jobs", methods=["POST"])
def api_create_job():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "请选择要打印的文件"}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"error": "不支持的文件类型：%s" % ext}), 400

    printer = (request.form.get("printer") or "").strip()
    if not printer:
        return jsonify({"error": "请选择打印机"}), 400

    orientation = request.form.get("orientation", "portrait")
    if orientation not in ("portrait", "landscape"):
        return jsonify({"error": "方向参数错误"}), 400

    try:
        copies = int(request.form.get("copies", 1))
        copies = max(1, min(99, copies))
    except ValueError:
        return jsonify({"error": "份数必须是数字"}), 400

    pages = (request.form.get("pages") or "").strip()
    if pages and not re.match(r'^[\d\s,\-]+$', pages):
        return jsonify({"error": "页码范围格式不对，例如 1-3,5,8-"}), 400

    duplex = request.form.get("duplex", "one")
    if duplex not in ("one", "long", "short"):
        return jsonify({"error": "双面参数错误"}), 400

    first_only = request.form.get("first_sheet_only") in ("1", "true", "on")

    # PowerPoint 专有参数：一张纸打印几张幻灯片 + 讲义排列顺序
    try:
        slides_per_page = int(request.form.get("slides_per_page", SLIDES_PER_PAGE_DEFAULT))
    except ValueError:
        return jsonify({"error": "每张幻灯片数必须是数字"}), 400
    if slides_per_page not in SLIDES_PER_PAGE_CHOICES:
        slides_per_page = SLIDES_PER_PAGE_DEFAULT

    handout_order = request.form.get("handout_order", "horizontal")
    if handout_order not in ("horizontal", "vertical"):
        handout_order = "horizontal"

    if ext == ".pdf" and not SUMATRA:
        return jsonify({"error": "服务器上没有找到 SumatraPDF.exe，无法打印 PDF"}), 400

    jid = uuid.uuid4().hex[:12]          # 先定 job_id，再用它命名上传目录
    form = {
        "filename": safe_filename(f.filename),
        "ext": ext,
        "saved_path": "",
        "printer": printer,
        "orientation": orientation,
        "copies": copies,
        "pages": pages,
        "duplex": duplex,
        "first_sheet_only": first_only,
        "slides_per_page": slides_per_page,
        "handout_order": handout_order,
    }

    folder = jm.ensure_upload_dir(jid)
    # 落盘文件名固定为 source<扩展名>（纯 ASCII）：
    #   1) cscript / SumatraPDF 的控制台输出编码不一致，中文路径必然导致日志乱码；
    #   2) Office COM 对含中文或空格的路径偶发“类型不匹配”。
    # 原始文件名只用于页面展示，不参与任何命令行。
    saved_path = os.path.join(folder, "source%s" % ext)
    f.save(saved_path)
    form["saved_path"] = saved_path

    job_id = jm.create_job_with_id(jid, form)
    return jsonify({"job_id": job_id})


@app.route("/api/jobs/<job_id>")
def api_job(job_id):
    job = jm.get_job(job_id)
    if not job:
        return jsonify({"error": "任务不存在或已过期"}), 404
    return jsonify(job)


@app.route("/api/jobs/<job_id>/cancel", methods=["POST"])
def api_cancel_job(job_id):
    ok = jm.cancel_job(job_id)
    return jsonify({"ok": ok})


# ------------------------------------------------------------------ 启动
SCRIPT_NAMES = ("print_word.vbs", "print_excel.vbs", "print_ppt.vbs")


def lint_scripts():
    """纯文本检查：VBScript 不是 VBA，不能用 Name:=value 命名参数，也不能用行尾 '_' 续行，
    且脚本最好是纯 ASCII（cscript 按系统 ANSI 代码页读文件，中文注释常变乱码）。"""
    import os as _os
    from config import SCRIPTS_DIR
    issues = []
    for name in SCRIPT_NAMES:
        path = _os.path.join(SCRIPTS_DIR, name)
        if not _os.path.exists(path):
            issues.append((name, "脚本文件缺失"))
            continue
        with open(path, "rb") as fh:
            raw = fh.read()
        if any(b > 127 for b in raw):
            issues.append((name, "含非 ASCII 字符，请另存为 ANSI 编码"))
        text = raw.decode("ascii", errors="replace")
        for no, line in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
            stripped = line.strip()
            if stripped.startswith("'") or stripped == "":
                continue
            if ":=" in stripped:
                issues.append((name, "第 %d 行用了命名参数 ':='（VBS 不支持，改按位置传参）" % no))
            if stripped.endswith("_"):
                issues.append((name, "第 %d 行以续行符 '_' 结尾（VBS 里容易出问题，请合并成一行）" % no))
    return issues


def check_scripts():
    """启动时做一次 VBS 语法自检：不带参数运行，正常应打印 bad arguments（退出码 9）；
    若脚本有语法错误，cscript 会直接报“编译器错误”，这里提前暴露出来。"""
    import subprocess
    from config import SCRIPTS_DIR, CSCRIPT
    lint = dict((n, m) for n, m in lint_scripts())
    report = []
    for name in SCRIPT_NAMES:
        path = os.path.join(SCRIPTS_DIR, name)
        if not os.path.exists(path):
            report.append((name, False, "脚本文件缺失"))
            continue
        if name in lint:
            report.append((name, False, lint[name]))
            continue
        try:
            p = subprocess.run([CSCRIPT, "//nologo", path], stdout=subprocess.PIPE,
                               stderr=subprocess.STDOUT, timeout=30)
            out = p.stdout.decode(CONSOLE_ENCODING, errors="replace").strip()
        except Exception as e:
            report.append((name, False, str(e)))
            continue
        report.append((name, "bad arguments" in out, out))
    return report


def main():
    os.makedirs(os.path.join(BASE_DIR, "uploads"), exist_ok=True)
    print("=" * 60)
    print(" 局域网打印服务")
    print(" 监听地址: http://0.0.0.0:%d" % PORT)
    print(" SumatraPDF: %s" % (SUMATRA or "未找到（PDF 打印不可用）"))
    print(" 控制台编码: %s（cscript）/ %s（SumatraPDF）" % (CONSOLE_ENCODING, SUMATRA_ENCODING))
    print(" 纸张设置: %s" % (("paper=" + SUMATRA_PAPER) if SUMATRA_PAPER else "不指定（用 PDF 自身尺寸）"))
    print("-" * 60)
    for name, ok, msg in check_scripts():
        print(" %s %s %s" % ("[OK]  " if ok else "[FAIL]", name, "" if ok else ("-> " + msg)))
    from config import (PPT_QUEUE_WAIT_SEC, OFFICE_QUEUE_WAIT_SEC, JOB_TIMEOUT_SEC)
    if max(PPT_QUEUE_WAIT_SEC, OFFICE_QUEUE_WAIT_SEC) + 180 > JOB_TIMEOUT_SEC:
        print(" [WARN] JOB_TIMEOUT_SEC(%d) 小于等待上限 + 180，建议调大"
              % JOB_TIMEOUT_SEC)
        print("        （PPT=%d, Office=%d；等待期间进程超时会把任务误杀）"
              % (PPT_QUEUE_WAIT_SEC, OFFICE_QUEUE_WAIT_SEC))
    print("=" * 60)
    try:
        from waitress import serve
        serve(app, host=HOST, port=PORT, threads=8)
    except ImportError:
        app.run(host=HOST, port=PORT, threaded=True, debug=False)


if __name__ == "__main__":
    main()
