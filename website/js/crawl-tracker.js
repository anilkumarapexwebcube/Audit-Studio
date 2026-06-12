/* GSC Crawl Date Tracker - web logic
 * Talks to the same Google Apps Script webhook the original tool used.
 */
(function () {
  const WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbydUPonVySIRG_Icz2ygWRmDJ1_qEGjhS9vMP9INZF5_f_sD1ZX2d7kRnOqN8YbWeN_/exec';

  const $ = (id) => document.getElementById(id);

  let workflow = { items: [], results: [], currentIndex: 0, state: 'idle', startTime: null, batchId: '' };

  function updateCount() {
    const lines = $('urlsInput').value.split('\n').map(l => l.trim()).filter(Boolean);
    const counter = $('urlCount');
    const summary = $('domainSummary');
    const unique = new Set(lines);
    const dupCount = lines.length - unique.size;

    counter.firstElementChild.textContent = lines.length + ' / 100 URLs' +
      (dupCount > 0 ? ' (' + dupCount + ' duplicate' + (dupCount === 1 ? '' : 's') + ')' : '');
    counter.classList.toggle('over', lines.length > 100);

    const domains = new Set();
    unique.forEach(u => {
      const m = u.match(/^https?:\/\/([^/]+)/i);
      if (m) domains.add(m[1].toLowerCase().replace(/^www\./, ''));
    });
    summary.textContent = domains.size > 0 ? domains.size + ' unique domain' + (domains.size === 1 ? '' : 's') : '';
  }

  async function startWorkflow() {
    const rawLines = $('urlsInput').value.split('\n').map(l => l.trim()).filter(Boolean);
    if (rawLines.length === 0) { GSCUI.toast('Please paste at least one URL.', 'warn'); return; }

    const seen = new Set();
    const urls = [];
    let duplicatesRemoved = 0;
    rawLines.forEach(url => {
      if (seen.has(url)) duplicatesRemoved++;
      else { seen.add(url); urls.push(url); }
    });

    if (urls.length > 100) {
      await GSCUI.alert('Maximum 100 unique URLs allowed per batch. You have ' + urls.length + ' after removing duplicates.', { title: 'Too many URLs', type: 'warn' });
      return;
    }
    if (duplicatesRemoved > 0) { $('urlsInput').value = urls.join('\n'); updateCount(); }

    const introMsg = duplicatesRemoved > 0
      ? 'Validating ' + urls.length + ' unique URLs (' + duplicatesRemoved + ' duplicate' + (duplicatesRemoved === 1 ? '' : 's') + ' removed)...'
      : 'Validating ' + urls.length + ' URLs...';
    $('progressBox').classList.add('visible');
    setProgressState('running', introMsg);
    setButtons({ run: false, pause: false, resume: false, stop: false });

    try {
      const resp = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'validate_batch', urls })
      });
      const data = await resp.json();

      if (!data.success) {
        setProgressState('stopped', data.error);
        setButtons({ run: true, pause: false, resume: false, stop: false });
        return;
      }

      workflow.items = data.items;
      workflow.batchId = data.batchId || '';
      workflow.results = data.items.map(i => ({
        url: i.url,
        domain: i.domain || '-',
        accountKey: i.accountKey || '',
        lastCrawlDate: i.ready ? 'Pending' : 'Skipped',
        indexStatus: i.ready ? 'Pending' : (i.error || 'Unknown'),
        status: i.ready ? 'pending' : 'error'
      }));
      workflow.currentIndex = 0;
      workflow.state = 'running';
      workflow.startTime = Date.now();

      $('progressBox').classList.add('visible');
      $('resultsCard').style.display = 'block';
      renderTable();
      setButtons({ run: false, pause: true, resume: false, stop: true });
      processNext();
    } catch (err) {
      setProgressState('stopped', 'Connection failed: ' + err.message);
      setButtons({ run: true, pause: false, resume: false, stop: false });
    }
  }

  async function processNext() {
    if (workflow.state !== 'running') return;

    while (workflow.currentIndex < workflow.items.length && !workflow.items[workflow.currentIndex].ready) {
      workflow.currentIndex++;
    }
    if (workflow.currentIndex >= workflow.items.length) {
      workflow.state = 'completed';
      setProgressState('completed', 'All URLs processed.');
      setButtons({ run: true, pause: false, resume: false, stop: false });
      updateProgress();
      return;
    }

    const item = workflow.items[workflow.currentIndex];
    renderTable();
    updateProgress();

    try {
      const resp = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'inspect_single',
          url: item.url,
          domain: item.domain,
          accountKey: item.accountKey,
          batchId: workflow.batchId
        })
      });
      const data = await resp.json();
      if (data.success) {
        workflow.results[workflow.currentIndex] = {
          url: data.url, domain: data.domain, accountKey: data.accountKey,
          lastCrawlDate: data.lastCrawlDate, indexStatus: data.indexStatus, status: 'ok'
        };
      } else {
        workflow.results[workflow.currentIndex].lastCrawlDate = 'Error';
        workflow.results[workflow.currentIndex].indexStatus = data.error || 'Unknown error';
        workflow.results[workflow.currentIndex].status = 'error';
      }
    } catch (err) {
      workflow.results[workflow.currentIndex].lastCrawlDate = 'Error';
      workflow.results[workflow.currentIndex].indexStatus = 'Network error';
      workflow.results[workflow.currentIndex].status = 'error';
    }

    renderTable();
    workflow.currentIndex++;
    await sleep(300);
    if (workflow.state === 'running') processNext();
  }

  function pauseWorkflow() {
    if (workflow.state !== 'running') return;
    workflow.state = 'paused';
    setProgressState('paused', 'Paused. Click Resume to continue.');
    setButtons({ run: false, pause: false, resume: true, stop: true });
  }
  function resumeWorkflow() {
    if (workflow.state !== 'paused') return;
    workflow.state = 'running';
    setProgressState('running', 'Resuming...');
    setButtons({ run: false, pause: true, resume: false, stop: true });
    processNext();
  }
  async function stopWorkflow() {
    const ok = await GSCUI.confirm('Stop the workflow? Already-fetched results will be kept and can be downloaded.', { title: 'Stop run', okLabel: 'Stop', danger: true });
    if (!ok) return;
    workflow.state = 'stopped';
    setProgressState('stopped', 'Stopped by user. ' + countCompleted() + ' URLs fetched.');
    setButtons({ run: true, pause: false, resume: false, stop: false });
  }
  function countCompleted() {
    return workflow.results.filter(r => r.status === 'ok' || r.status === 'error').length;
  }

  function setButtons({ run, pause, resume, stop }) {
    $('runBtn').disabled = !run;
    $('pauseBtn').disabled = !pause;
    $('resumeBtn').disabled = !resume;
    $('stopBtn').disabled = !stop;
  }

  function setProgressState(state, detail) {
    const stateEl = $('progressState');
    const fill = $('progressFill');
    stateEl.className = 'progress-state ' + state;
    stateEl.textContent = state;
    fill.className = 'progress-bar-fill' + (state === 'paused' ? ' paused' : state === 'stopped' ? ' stopped' : '');
    $('progressDetail').textContent = detail;
    updateProgress();
  }

  function updateProgress() {
    const done = workflow.currentIndex;
    const total = workflow.items.length;
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    $('progressCount').textContent = done + ' / ' + total + ' (' + pct + '%)';
    $('progressFill').style.width = pct + '%';
    if (workflow.state === 'running' && done > 0) {
      const elapsed = (Date.now() - workflow.startTime) / 1000;
      const perItem = elapsed / done;
      const remaining = Math.round((total - done) * perItem);
      $('progressEta').textContent = 'ETA: ~' + remaining + 's';
    } else {
      $('progressEta').textContent = '';
    }
  }

  function renderTable() {
    $('resultsBody').innerHTML = workflow.results.map((r, i) => {
      let rowClass = '';
      if (r.status === 'pending') rowClass = 'pending';
      else if (i === workflow.currentIndex && workflow.state === 'running') rowClass = 'processing';

      let badgeClass = 'badge-pending';
      let badgeText = r.lastCrawlDate;
      if (r.status === 'ok') badgeClass = r.lastCrawlDate === 'Never Crawled' ? 'badge-never' : 'badge-ok';
      else if (r.status === 'error') badgeClass = 'badge-err';
      else if (i === workflow.currentIndex && workflow.state === 'running') { badgeClass = 'badge-running'; badgeText = 'Checking...'; }

      return '<tr class="' + rowClass + '"><td>' + (i + 1) + '</td><td>' + escapeHtml(r.url) + '</td><td>' +
        escapeHtml(r.domain) + '</td><td><span class="' + badgeClass + '">' + escapeHtml(badgeText) +
        '</span></td><td>' + escapeHtml(r.indexStatus) + '</td></tr>';
    }).join('');
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function downloadCSV() {
    if (workflow.results.length === 0) { GSCUI.toast('No results to download yet.', 'warn'); return; }
    const header = ['URL', 'Domain', 'Last Crawl Date', 'Index Status'];
    const rows = workflow.results.map(r => [r.url, r.domain, r.lastCrawlDate, r.indexStatus]);
    const csv = [header, ...rows].map(row => row.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'gsc_crawl_dates_' + new Date().toISOString().split('T')[0] + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  // ===== Wire up =====
  $('urlsInput').addEventListener('input', updateCount);
  $('runBtn').addEventListener('click', startWorkflow);
  $('pauseBtn').addEventListener('click', pauseWorkflow);
  $('resumeBtn').addEventListener('click', resumeWorkflow);
  $('stopBtn').addEventListener('click', stopWorkflow);
  $('downloadCsvBtn').addEventListener('click', downloadCSV);
  updateCount();
})();
