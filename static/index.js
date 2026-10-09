/* 首页：上传 + 参数 + 提交 */
(function () {
  var form = document.getElementById('printForm');
  var drop = document.getElementById('drop');
  var fileInput = document.getElementById('file');
  var dropHint = document.getElementById('dropHint');
  var fileInfo = document.getElementById('fileInfo');
  var fileName = document.getElementById('fileName');
  var fileSize = document.getElementById('fileSize');
  var clearFile = document.getElementById('clearFile');
  var printerSel = document.getElementById('printer');
  var pagesInput = document.getElementById('pages');
  var copiesInput = document.getElementById('copies');
  var duplexSel = document.getElementById('duplex');
  var sheetRow = document.getElementById('sheetRow');
  var pptRow = document.getElementById('pptRow');
  var pptTip = document.getElementById('pptTip');
  var slidesSel = document.getElementById('slidesPerPage');
  var orderSel = document.getElementById('handoutOrder');
  var alertBox = document.getElementById('alert');
  var submitBtn = document.getElementById('submitBtn');

  // 拖拽进来的 File 对象不会自动进入 <input type=file>，单独存一份
  var chosenFile = null;

  var OK_EXT = ['doc', 'docx', 'txt', 'xls', 'xlsx',
                'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'ppsm', 'pot', 'potx', 'pdf'];
  var PPT_EXT = ['ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'ppsm', 'pot', 'potx'];
  var XLS_EXT = ['xls', 'xlsx'];

  function showError(msg) {
    alertBox.textContent = msg;
    alertBox.classList.remove('hidden');
  }
  function clearError() { alertBox.classList.add('hidden'); }

  function fmtSize(n) {
    if (n < 1024) return n + ' B';
    if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
    return (n / 1048576).toFixed(1) + ' MB';
  }

  function currentFile() {
    if (chosenFile) return chosenFile;
    if (fileInput.files && fileInput.files.length) return fileInput.files[0];
    return null;
  }

  function pickFile(f) {
    if (!f) return;
    if (f.size === 0) { showError('这个文件是空的，请重新选择'); return; }
    var ext = (f.name.split('.').pop() || '').toLowerCase();
    if (OK_EXT.indexOf(ext) < 0) { showError('不支持的文件类型：.' + ext); return; }

    chosenFile = f;                       // 拖拽 / 选择两条路都记下来
    clearError();
    fileName.textContent = f.name;
    fileSize.textContent = fmtSize(f.size);
    fileInfo.classList.remove('hidden');
    dropHint.classList.add('hidden');

    var isXls = XLS_EXT.indexOf(ext) >= 0;
    var isPpt = PPT_EXT.indexOf(ext) >= 0;
    sheetRow.classList.toggle('hidden', !isXls);
    pptRow.classList.toggle('hidden', !isPpt);
    pptTip.classList.toggle('hidden', !isPpt);
  }

  // 讲义模式下提示：无法转 PDF，双面只能靠打印机默认设置
  function refreshPptTip() {
    pptTip.classList.toggle('hidden',
      pptRow.classList.contains('hidden') || slidesSel.value === '1');
  }
  slidesSel.addEventListener('change', refreshPptTip);

  function resetFile() {
    chosenFile = null;
    try { fileInput.value = ''; } catch (e) {}
    fileInfo.classList.add('hidden');
    dropHint.classList.remove('hidden');
    sheetRow.classList.add('hidden');
    pptRow.classList.add('hidden');
    pptTip.classList.add('hidden');
  }

  drop.addEventListener('click', function (e) {
    if (e.target === clearFile) return;
    fileInput.click();
  });

  fileInput.addEventListener('change', function () {
    pickFile(this.files && this.files[0]);
  });

  // 拖放到虚线框
  ['dragenter', 'dragover'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) {
      e.preventDefault();
      e.stopPropagation();
      drop.classList.add('drag');
    });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) {
      e.preventDefault();
      e.stopPropagation();
      drop.classList.remove('drag');
    });
  });

  drop.addEventListener('drop', function (e) {
    var dt = e.dataTransfer;
    if (!dt || !dt.files || !dt.files.length) return;
    // 关键：把拖进来的文件写回 input，这样后续统一从 fileInput.files 读取
    try {
      fileInput.files = dt.files;
    } catch (err) {
      /* 少数浏览器不允许写入 input.files，此时靠 chosenFile 兜底 */
    }
    pickFile(dt.files[0]);
  });

  // 拖到页面其它位置时别让浏览器直接打开文件
  ['dragover', 'drop'].forEach(function (ev) {
    document.addEventListener(ev, function (e) {
      if (!drop.contains(e.target)) e.preventDefault();
    });
  });

  clearFile.addEventListener('click', function (e) {
    e.stopPropagation();
    resetFile();
  });

  // 打印机列表
  fetch('/api/printers')
    .then(function (r) { return r.json(); })
    .then(function (data) {
      printerSel.innerHTML = '';
      if (!data.printers || !data.printers.length) {
        var o = document.createElement('option');
        o.value = '';
        o.textContent = '（服务器上没有找到打印机）';
        printerSel.appendChild(o);
        return;
      }
      data.printers.forEach(function (name) {
        var opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name + (name === data.default ? '（默认）' : '');
        printerSel.appendChild(opt);
      });
      if (data.default) printerSel.value = data.default;
    })
    .catch(function () {
      printerSel.innerHTML = '<option value="">读取打印机失败，请刷新重试</option>';
    });

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    clearError();

    var f = currentFile();
    if (!f) { showError('请先选择文件（点击选择或把文件拖到虚线框里）'); return; }
    if (!printerSel.value) { showError('请选择打印机'); return; }

    var pages = pagesInput.value.trim();
    if (pages && !/^[\d\s,\-]+$/.test(pages)) {
      showError('页码范围格式不对，例子：1-3,5,8-');
      return;
    }
    var copies = parseInt(copiesInput.value, 10);
    if (!(copies >= 1 && copies <= 99)) { showError('份数需要在 1 ~ 99 之间'); return; }

    var fd = new FormData();
    fd.append('file', f, f.name);          // 显式给文件名，避免个别浏览器丢失
    fd.append('printer', printerSel.value);
    fd.append('orientation', document.querySelector('input[name="orientation"]:checked').value);
    fd.append('copies', copies);
    fd.append('pages', pages);
    fd.append('duplex', duplexSel.value);
    fd.append('first_sheet_only', document.getElementById('firstSheetOnly').checked ? '1' : '0');
    fd.append('slides_per_page', slidesSel.value);
    fd.append('handout_order', orderSel.value);

    submitBtn.disabled = true;
    submitBtn.textContent = '正在上传…';

    fetch('/api/jobs', { method: 'POST', body: fd })
      .then(function (r) {
        return r.json().then(function (data) {
          if (!r.ok) throw new Error(data.error || '提交失败');
          return data;
        });
      })
      .then(function (data) { location.href = '/status/' + data.job_id; })
      .catch(function (err) {
        showError(err.message);
        submitBtn.disabled = false;
        submitBtn.textContent = '提交打印';
      });
  });
})();
