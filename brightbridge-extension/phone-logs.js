// phone-logs.js — BrightBridge phone logs (Teacher Tools parity, extension storage)
(function (global) {
  const STORAGE_KEY = 'bb_phone_logs_v1';

  let BB = null;
  let phoneCache = {};
  let inited = false;

  function setStatus(text, type = 'idle') {
    const bar = document.getElementById('phoneLogsStatus');
    const el = document.getElementById('phoneLogsStatusText');
    if (el) el.textContent = text;
    if (bar) bar.className = `status-bar status-${type}`;
  }

  async function loadCache() {
    try {
      const data = await chrome.storage.local.get(STORAGE_KEY);
      phoneCache = data[STORAGE_KEY] && typeof data[STORAGE_KEY] === 'object' ? data[STORAGE_KEY] : {};
    } catch (_) {
      phoneCache = {};
    }
  }

  async function saveCache() {
    try {
      await chrome.storage.local.set({ [STORAGE_KEY]: phoneCache });
    } catch (_) {}
  }

  function classKey(name) {
    return String(name || '').toLowerCase().trim().replace(/\s+/g, ' ');
  }

  function studentStorageKey(className, studentName) {
    return `${classKey(className)}|${classKey(studentName)}`;
  }

  function hasPhoneLogEntry(entry) {
    if (!entry) return false;
    return !!(entry.date || entry.time || (entry.log && String(entry.log).trim()));
  }

  function formatPhoneLogTime(t) {
    if (!t) return '';
    const [h, m] = String(t).split(':').map(Number);
    if (isNaN(h) || isNaN(m)) return t;
    const suffix = h >= 12 ? 'PM' : 'AM';
    const h12 = h % 12 || 12;
    return `${h12}:${String(m).padStart(2, '0')} ${suffix}`;
  }

  function escapeHtml(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function render() {
    const root = document.getElementById('phoneLogsRoot');
    if (!root || !BB) return;

    const students = BB.getStudents();
    const className = BB.getClassName() || '';
    root.innerHTML = '';

    if (!students.length) {
      setStatus('Load grades from the Brightspace Grades page (↻), or open Report Card first.', 'error');
      root.innerHTML = '<p class="phone-logs-empty">No students loaded. Use ↻ after opening your gradebook.</p>';
      return;
    }

    const wrap = document.createElement('div');
    wrap.className = 'phone-logs-layout';

    const upPanel = document.createElement('div');
    upPanel.className = 'phone-logs-underperforming';
    const upTitle = document.createElement('div');
    upTitle.className = 'phone-logs-up-title';
    upTitle.textContent = 'At or below 60%';
    const upCount = document.createElement('span');
    upCount.className = 'phone-logs-up-count';
    const upList = document.createElement('div');
    upList.className = 'phone-logs-up-list';
    upPanel.appendChild(upTitle);
    upTitle.appendChild(upCount);
    upPanel.appendChild(upList);

    const search = document.createElement('input');
    search.type = 'search';
    search.className = 'phone-logs-search';
    search.placeholder = 'Search students…';

    const list = document.createElement('div');
    list.className = 'phone-logs-list';

    const sorted = [...students].sort((a, b) =>
      a.name.localeCompare(b.name, undefined, { sensitivity: 'base' })
    );

    const underperforming = sorted.filter(s =>
      s.gradeNum != null && s.gradeNum <= BB.UNDERPERFORM_THRESHOLD
    );
    upCount.textContent = underperforming.length ? String(underperforming.length) : '';

    if (!underperforming.length) {
      upList.innerHTML = '<p class="muted-tiny">No students at or below 60%.</p>';
    } else {
      underperforming.forEach(s => {
        const sk = studentStorageKey(className, s.name);
        const logged = hasPhoneLogEntry(phoneCache[sk]);
        const row = document.createElement('div');
        row.className = 'phone-logs-up-item';
        row.title = 'Scroll to ' + s.name;
        row.innerHTML = `<span>${escapeHtml(s.name)}</span>` +
          (logged ? '<span class="up-logged-badge">✓</span>' : '') +
          `<span class="up-mark">${s.gradeNum}%</span>`;
        row.addEventListener('click', () => {
          const card = Array.from(list.querySelectorAll('.phone-log-entry')).find(el => el.dataset.sk === sk);
          if (card) {
            card.classList.remove('phone-log-highlight');
            void card.offsetWidth;
            card.classList.add('phone-log-highlight');
            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
            setTimeout(() => card.classList.remove('phone-log-highlight'), 2000);
          }
        });
        upList.appendChild(row);
      });
    }

    function buildCards(filterQ) {
      list.innerHTML = '';
      const q = (filterQ || '').trim().toLowerCase();
      sorted.forEach(s => {
        if (q && !s.name.toLowerCase().includes(q)) return;
        const sk = studentStorageKey(className, s.name);
        const entry = phoneCache[sk] || {};

        const card = document.createElement('div');
        card.className = 'phone-log-entry';
        card.dataset.sk = sk;

        const nameEl = document.createElement('div');
        nameEl.className = 'phone-log-name' +
          (s.gradeNum != null && s.gradeNum <= BB.UNDERPERFORM_THRESHOLD ? ' phone-log-under' : '');
        nameEl.textContent = s.name;
        card.appendChild(nameEl);

        const fieldsRow = document.createElement('div');
        fieldsRow.className = 'phone-log-fields';

        const dateLabel = document.createElement('label');
        const dateSpan = document.createElement('span');
        dateSpan.textContent = 'Date';
        const dateInput = document.createElement('input');
        dateInput.type = 'date';
        dateInput.value = entry.date || '';
        dateInput.addEventListener('change', () => {
          if (!phoneCache[sk]) phoneCache[sk] = {};
          phoneCache[sk].date = dateInput.value;
          saveCache();
        });
        dateLabel.appendChild(dateSpan);
        dateLabel.appendChild(dateInput);

        const timeLabel = document.createElement('label');
        const timeSpan = document.createElement('span');
        timeSpan.textContent = 'Time';
        const timeInput = document.createElement('input');
        timeInput.type = 'time';
        timeInput.value = entry.time || '';
        timeInput.addEventListener('change', () => {
          if (!phoneCache[sk]) phoneCache[sk] = {};
          phoneCache[sk].time = timeInput.value;
          saveCache();
        });
        timeLabel.appendChild(timeSpan);
        timeLabel.appendChild(timeInput);

        fieldsRow.appendChild(dateLabel);
        fieldsRow.appendChild(timeLabel);
        card.appendChild(fieldsRow);

        const textarea = document.createElement('textarea');
        textarea.className = 'phone-log-textarea';
        textarea.placeholder = 'Enter phone log notes…';
        textarea.value = entry.log || '';
        textarea.addEventListener('input', () => {
          if (!phoneCache[sk]) phoneCache[sk] = {};
          phoneCache[sk].log = textarea.value;
          saveCache();
        });
        card.appendChild(textarea);

        const saveRow = document.createElement('div');
        saveRow.className = 'phone-log-save-row';
        const saveBtn = document.createElement('button');
        saveBtn.type = 'button';
        saveBtn.className = 'btn phone-log-save-btn' + (hasPhoneLogEntry(phoneCache[sk]) ? ' saved' : '');
        saveBtn.textContent = hasPhoneLogEntry(phoneCache[sk]) ? 'Saved ✓' : 'Save Log';
          saveBtn.addEventListener('click', () => {
          if (!phoneCache[sk]) phoneCache[sk] = {};
          phoneCache[sk].date = dateInput.value;
          phoneCache[sk].time = timeInput.value;
          phoneCache[sk].log = textarea.value;
          saveCache();
          saveBtn.textContent = 'Saved ✓';
          saveBtn.classList.add('saved');
          setTimeout(() => {
            saveBtn.textContent = hasPhoneLogEntry(phoneCache[sk]) ? 'Saved ✓' : 'Save Log';
            if (!hasPhoneLogEntry(phoneCache[sk])) saveBtn.classList.remove('saved');
          }, 1500);
        });
        saveRow.appendChild(saveBtn);
        card.appendChild(saveRow);

        list.appendChild(card);
      });
    }

    search.addEventListener('input', () => buildCards(search.value));
    buildCards('');

    const actions = document.createElement('div');
    actions.className = 'phone-logs-actions';
    const printBtn = document.createElement('button');
    printBtn.type = 'button';
    printBtn.className = 'btn primary';
    printBtn.textContent = 'Print phone logs';
    printBtn.addEventListener('click', () => printPhoneLogsSheet(className, sorted));
    actions.appendChild(printBtn);

    wrap.appendChild(upPanel);
    wrap.appendChild(search);
    wrap.appendChild(list);
    wrap.appendChild(actions);
    root.appendChild(wrap);

    setStatus(`✓ ${students.length} students — phone logs`, 'success');
  }

  function printPhoneLogsSheet(className, studentsSorted) {
    const teacher = '';
    const classInfo = className || '';
    const termLabel = '';
    const today = new Date().toLocaleDateString('en-CA');
    const ROWS_PER_PAGE = 20;
    const pages = Math.ceil(studentsSorted.length / ROWS_PER_PAGE) || 1;
    let html = '';

    for (let p = 0; p < pages; p++) {
      const slice = studentsSorted.slice(p * ROWS_PER_PAGE, (p + 1) * ROWS_PER_PAGE);
      const metaLine = [
        teacher && `Teacher: ${teacher}`,
        classInfo && `Class: ${classInfo}`,
        termLabel && `Term: ${termLabel}`,
        `Printed: ${today}`
      ].filter(Boolean).join(' &nbsp;·&nbsp; ');

      html += `<div class="phone-logs-print-page">
  <div class="phone-logs-print-page-header">
    <h2>Phone Log</h2>
    <div class="phone-logs-print-meta">${metaLine}</div>
  </div>
  <table class="phone-logs-print-table">
    <thead><tr>
      <th style="width:22%">Student</th>
      <th style="width:12%">Date</th>
      <th style="width:9%">Time</th>
      <th>Notes</th>
    </tr></thead><tbody>`;

      slice.forEach(s => {
        const sk = studentStorageKey(className, s.name);
        const entry = phoneCache[sk] || {};
        const date = entry.date || '';
        const time = formatPhoneLogTime(entry.time || '');
        const log = escapeHtml(entry.log || '').replace(/\n/g, '<br>');
        html += `<tr>
          <td>${escapeHtml(s.name)}</td>
          <td>${date}</td>
          <td>${time}</td>
          <td class="log-notes">${log}</td>
        </tr>`;
      });

      html += '</tbody></table></div>';
    }

    const fullDoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Phone Log</title>
    <style>
      body{font-family:system-ui,sans-serif;font-size:12px;margin:0;padding:12px;}
      .phone-logs-print-page{page-break-after:always;}
      .phone-logs-print-page-header h2{margin:0 0 6px;font-size:18px;}
      .phone-logs-print-meta{font-size:11px;color:#444;margin-bottom:8px;}
      table{border-collapse:collapse;width:100%;}
      th,td{border:1px solid #ccc;padding:4px 6px;text-align:left;vertical-align:top;}
      th{background:#f3f4f6;}
      @media print{@page{margin:12mm;}body{padding:0;}}
    </style></head><body>${html}</body></html>`;

    if (typeof window.bbPrintHtml === 'function') {
      window.bbPrintHtml(fullDoc);
    }
  }

  function init(api) {
    BB = api;
    if (inited) return;
    inited = true;
    loadCache();
    const refresh = document.getElementById('phoneLogsRefreshBtn');
    if (refresh) {
      refresh.addEventListener('click', async () => {
        setStatus('Loading grades from page…', 'loading');
        await BB.loadGrades();
        await loadCache();
        render();
      });
    }
  }

  async function renderWithCache() {
    await loadCache();
    render();
  }

  global.BBPhoneLogs = {
    init,
    render: renderWithCache,
    setStatus
  };
})(window);
