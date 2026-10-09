/* 状态页：轮询任务状态并展示日志 */
(function () {
  var jobId = window.JOB_ID;
  var badge = document.getElementById('badge');
  var message = document.getElementById('message');
  var logBox = document.getElementById('log');
  var alertBox = document.getElementById('alert');
  var cancelBtn = document.getElementById('cancelBtn');
  var timer = null;
  var lastCount = 0;

  var slidesRow = document.getElementById('mSlidesRow');
  var slidesCell = document.getElementById('mSlides');

  function setMeta(j) {
    document.getElementById('mFile').textContent = j.filename || '-';
    document.getElementById('mPrinter').textContent = j.printer || '-';
    document.getElementById('mOrient').textContent = j.orientation + ' / ' + j.copies + ' 份';
    document.getElementById('mPages').textContent = j.pages || '全部';
    document.getElementById('mDuplex').textContent = j.duplex || '单面';
    document.getElementById('mTime').textContent = j.created_at || '-';

    if (j.is_ppt) {
      slidesRow.classList.remove('hidden');
      slidesCell.textContent = j.slides_per_page + ' 张幻灯片（' + j.handout_order + '）';
    } else {
      slidesRow.classList.add('hidden');
    }
  }

  function render(j) {
    setMeta(j);
    badge.textContent = j.status_text;
    badge.className = 'badge ' + j.status;
    message.textContent = j.status === 'queued' && j.queue_pos > 1
      ? '排队中，前面还有 ' + (j.queue_pos - 1) + ' 个任务'
      : (j.message || '');

    if (j.log.length !== lastCount) {
      lastCount = j.log.length;
      logBox.textContent = j.log.join('\n') || '等待服务器处理…';
      logBox.scrollTop = logBox.scrollHeight;
    }

    cancelBtn.classList.toggle('hidden', j.status !== 'queued');

    if (j.status === 'error') {
      alertBox.textContent = j.message + '。常见原因：Office 弹窗卡住、打印机离线、纸张/页码范围超出文档页数。';
      alertBox.classList.remove('hidden');
    } else {
      alertBox.classList.add('hidden');
    }

    if (j.status === 'success' || j.status === 'error' || j.status === 'cancelled') {
      clearInterval(timer);
    }
  }

  function poll() {
    fetch('/api/jobs/' + jobId)
      .then(function (r) {
        if (!r.ok) throw new Error('任务不存在或已过期');
        return r.json();
      })
      .then(render)
      .catch(function (err) {
        clearInterval(timer);
        badge.textContent = '未知';
        badge.className = 'badge error';
        message.textContent = err.message;
      });
  }

  cancelBtn.addEventListener('click', function () {
    fetch('/api/jobs/' + jobId + '/cancel', { method: 'POST' })
      .then(poll);
  });

  poll();
  timer = setInterval(poll, 1500);
})();
