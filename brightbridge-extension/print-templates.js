// print-templates.js — BrightBridge print templates + accumulated roster (pagination)
(function (global) {
  const PRINT_TEMPLATE_DEFS = {
    attendance: { id: 'attendance', label: 'Attendance Sheet' },
    marking: { id: 'marking', label: 'Marking Sheet' },
    drill: { id: 'drill', label: 'Drill Sheet' },
    reportCard: { id: 'reportCard', label: 'Report Card' }
  };

  let BB = null;
  let inited = false;

  /** @type {{ headers: string[], rows: object[], className: string }} */
  let accumulated = { headers: [], rows: [], className: '' };

  let selectedTemplate = 'attendance';
  let selectedMarkingCols = new Set();
  let drillCount = 12;
  let printMeta = { teacher: '', classInfo: '', term: 'T1' };

  function setStatus(text, type = 'idle') {
    const bar = document.getElementById('printTemplatesStatus');
    const el = document.getElementById('printTemplatesStatusText');
    if (el) el.textContent = text;
    if (bar) bar.className = `status-bar status-${type}`;
  }

  function getTermWeekRange(term) {
    if (term === 'T2') return { start: 14, end: 26 };
    if (term === 'T3') return { start: 27, end: 39 };
    return { start: 1, end: 13 };
  }

  function getTermLabel(term) {
    if (term === 'T2') return 'Term 2';
    if (term === 'T3') return 'Term 3';
    return 'Term 1';
  }

  function rowStudentKey(row, headers) {
    const c = BB.detectCols(headers);
    if (c.fullCol) return BB.cleanNameText(BB.rowGet(row, c.fullCol)).toLowerCase().trim();
    if (c.firstCol && c.lastCol) {
      return `${BB.cleanNameText(BB.rowGet(row, c.firstCol))} ${BB.cleanNameText(BB.rowGet(row, c.lastCol))}`.toLowerCase().trim();
    }
    return String(BB.rowGet(row, headers[0] || '') || '').toLowerCase().trim();
  }

  /** After merging pages, order must not depend on which page was added first — sort A→Z by student key. */
  function sortAccumulatedRowsInPlace() {
    const h = accumulated.headers;
    if (!h.length || !accumulated.rows.length) return;
    accumulated.rows.sort((a, b) => {
      const ka = rowStudentKey(a, h) || '\uffff';
      const kb = rowStudentKey(b, h) || '\uffff';
      return ka.localeCompare(kb, undefined, { sensitivity: 'base', numeric: true });
    });
  }

  function mergeIntoAccumulated(newHeaders, newRows, className) {
    if (!newHeaders || !newRows) return;
    if (!accumulated.headers.length) {
      accumulated.headers = [...newHeaders];
      accumulated.rows = newRows.map(r => expandRow(r, accumulated.headers));
      accumulated.className = className || '';
      sortAccumulatedRowsInPlace();
      return;
    }
    const union = [...new Set([...accumulated.headers, ...newHeaders])];
    const map = new Map();
    accumulated.rows.forEach(r => {
      const k = rowStudentKey(r, union);
      if (k) map.set(k, expandRow(r, union));
    });
    newRows.forEach(r => {
      const k = rowStudentKey(r, union);
      if (!k) return;
      if (!map.has(k)) map.set(k, expandRow(r, union));
    });
    accumulated.headers = union;
    accumulated.rows = [...map.values()];
    if (className) accumulated.className = className;
    sortAccumulatedRowsInPlace();
  }

  function expandRow(row, headers) {
    const o = {};
    headers.forEach(h => { o[h] = row[h] != null ? String(row[h]) : ''; });
    return o;
  }

  function clearAccumulated() {
    accumulated = { headers: [], rows: [], className: '' };
    selectedMarkingCols = new Set();
  }

  function buildContext() {
    const headers = accumulated.headers;
    const cols = BB.detectCols(headers);
    const nameCol = cols.fullCol || cols.firstCol || headers[0];
    return {
      name: accumulated.className || 'Class',
      rows: accumulated.rows,
      columnOrder: headers,
      allColumns: headers,
      studentNameColumn: nameCol,
      firstNameKey: cols.firstCol,
      lastNameKey: cols.lastCol
    };
  }

  function deriveContextStudentName(ctx, row) {
    if (ctx.studentNameColumn) {
      const t = String(row[ctx.studentNameColumn] ?? '').trim();
      if (t) return BB.cleanNameText(t);
    }
    if (ctx.firstNameKey && ctx.lastNameKey) {
      const first = String(row[ctx.firstNameKey] ?? '').trim();
      const last = String(row[ctx.lastNameKey] ?? '').trim();
      const c = `${first} ${last}`.trim();
      if (c) return BB.cleanNameText(c);
    }
    return 'Student';
  }

  function getMarkingSourceColumns(ctx) {
    const cols = BB.detectCols(ctx.columnOrder);
    return ctx.columnOrder.filter(h => h && !cols.skipCols.has(h));
  }

  function buildTemplateColumns(ctx, templateId) {
    const base = [
      { id: 'index', label: '#', className: 'index-col' },
      { id: 'name', label: 'Student Name', className: 'student-name-col' }
    ];
    if (templateId === 'attendance') {
      const range = getTermWeekRange(printMeta.term || 'T1');
      let maxWeek = range.end;
      const weekCols = ctx.columnOrder.filter(c => /week\s*\d+/i.test(c));
      if (weekCols.length) {
        const nums = weekCols.map(c => parseInt((c.match(/\d+/) || ['0'])[0], 10))
          .filter(n => n >= range.start && n <= range.end);
        if (nums.length) maxWeek = Math.max(...nums);
      }
      for (let week = range.start; week <= maxWeek; week++) {
        const wk = ctx.columnOrder.find(c => new RegExp(`week\\s*${week}\\b`, 'i').test(String(c)));
        base.push({
          id: wk || `__week${week}`,
          label: wk || `Week ${week}`,
          weekNum: week
        });
      }
    } else if (templateId === 'drill') {
      const n = Math.max(1, Math.min(30, drillCount || 12));
      for (let i = 1; i <= n; i++) {
        base.push({ id: `drill${i}`, label: `Drill ${i}` });
      }
    } else if (templateId === 'marking') {
      const source = getMarkingSourceColumns(ctx);
      const chosen = source.filter(c => selectedMarkingCols.has(c));
      (chosen.length ? chosen : source.slice(0, 12)).forEach((label, idx) => {
        base.push({ id: `marking-col-${idx}`, label });
      });
    }
    return base;
  }

  function renderTemplatePage(ctx, templateId) {
    if (templateId === 'reportCard') {
      const d = document.createElement('div');
      d.className = 'print-preview-empty';
      d.innerHTML = '<p><strong>Report Card</strong> template is available in Teacher Tools with full comment integration.</p><p class="muted-tiny">Use the web app for report card print, or use Attendance / Marking / Drill here.</p>';
      return d;
    }

    const columns = buildTemplateColumns(ctx, templateId);
    if (columns.length < 2) return null;

    const teacherName = (printMeta.teacher || '').trim();
    const classDisplay = (printMeta.classInfo || '').trim() || ctx.name || '';
    const termLabel = getTermLabel(printMeta.term);
    const templateLabel = PRINT_TEMPLATE_DEFS[templateId]?.label || 'Template';

    const students = ctx.rows.map((row, i) => ({
      rowIdx: i,
      name: deriveContextStudentName(ctx, row),
      row
    }));

    const page = document.createElement('div');
    page.className = 'class-template-page';

    const header = document.createElement('div');
    header.className = 'page-header';
    const title = document.createElement('div');
    title.className = 'page-title';
    title.textContent = templateLabel;
    header.appendChild(title);
    const meta = document.createElement('div');
    meta.className = 'page-meta';
    meta.textContent = [`Class: ${classDisplay}`, `Teacher: ${teacherName}`].join('   ');
    header.appendChild(meta);
    page.appendChild(header);

    if (templateId === 'drill') {
      const drillNote = document.createElement('div');
      drillNote.className = 'drill-note';
      drillNote.innerHTML = '<strong>Drill Type(s):</strong> _______________________________';
      page.appendChild(drillNote);
    }

    const table = document.createElement('table');
    table.className = 'template-table' + (templateId === 'marking' ? ' template-table-marking' : '');
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    columns.forEach(col => {
      const th = document.createElement('th');
      th.textContent = col.label;
      if (col.className) th.classList.add(col.className);
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    if (!students.length) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = columns.length;
      td.textContent = 'No students in template.';
      tr.appendChild(td);
      tbody.appendChild(tr);
    } else {
      students.forEach((entry, idx) => {
        const tr = document.createElement('tr');
        const row = entry.row;
        columns.forEach(col => {
          const td = document.createElement('td');
          if (col.id === 'index') td.textContent = String(idx + 1);
          else if (col.id === 'name') {
            td.textContent = entry.name;
            td.classList.add('student-name-cell');
          } else if (templateId === 'attendance') {
            let key = col.id;
            if (String(key).startsWith('__week')) {
              const wn = col.weekNum;
              key = ctx.columnOrder.find(c => new RegExp(`week\\s*${wn}\\b`, 'i').test(String(c))) || '';
            }
            td.textContent = key && row[key] != null ? String(row[key]) : '';
          } else if (templateId === 'drill') {
            td.innerHTML = '&nbsp;';
          } else if (templateId === 'marking') {
            const val = row[col.label] != null ? String(row[col.label]) : '';
            td.textContent = val;
          }
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
    }
    table.appendChild(tbody);

    const wrap = document.createElement('div');
    wrap.className = 'template-table-scroll-wrap';
    wrap.appendChild(table);
    page.appendChild(wrap);

    const footer = document.createElement('div');
    footer.className = 'print-footer';
    const ftLeft = document.createElement('div');
    ftLeft.className = 'print-footer-left';
    ftLeft.textContent = classDisplay || '[class info]';
    const ftCenter = document.createElement('div');
    ftCenter.className = 'print-footer-center';
    ftCenter.textContent = termLabel || '';
    const ftRight = document.createElement('div');
    ftRight.className = 'print-footer-right';
    ftRight.textContent = teacherName || '[teacher name]';
    footer.appendChild(ftLeft);
    footer.appendChild(ftCenter);
    footer.appendChild(ftRight);
    page.appendChild(footer);

    return page;
  }

  function renderPreview(root) {
    const preview = root.querySelector('#ptPreviewMount');
    if (!preview) return;
    preview.innerHTML = '';
    if (!accumulated.rows.length) {
      preview.appendChild(Object.assign(document.createElement('div'), {
        className: 'print-preview-empty',
        textContent: 'Click “Add students to template” on each Brightspace page (merge builds your class list).'
      }));
      return;
    }
    const ctx = buildContext();
    const page = renderTemplatePage(ctx, selectedTemplate);
    if (page) preview.appendChild(page);
  }

  async function printPreview() {
    const preview = document.getElementById('printTemplatesRoot');
    const panel = preview && preview.querySelector('#ptBuilt');
    const mount = preview && preview.querySelector('#ptPreviewMount');
    if (!mount || !mount.innerHTML.trim()) {
      setStatus('Nothing to print — add students and preview first.', 'error');
      return;
    }
    if (!(printMeta.teacher || '').trim() && BB.fetchTeacherNameFromPage) {
      try {
        const t = await BB.fetchTeacherNameFromPage();
        if (t) {
          printMeta.teacher = t;
          const ptIn = document.getElementById('ptTeacher');
          if (ptIn) ptIn.value = t;
          if (panel) renderPreview(panel);
        }
      } catch (_) {}
    }
    const htmlAfter = preview && preview.querySelector('#ptPreviewMount')?.innerHTML;
    if (!htmlAfter || !htmlAfter.trim()) {
      setStatus('Nothing to print — add students and preview first.', 'error');
      return;
    }
    const fullDoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Print</title>
      <style>${PRINT_CSS}</style>
    </head><body><div class="print-root">${htmlAfter}</div></body></html>`;

    if (typeof window.bbPrintHtml === 'function') {
      window.bbPrintHtml(fullDoc);
    } else {
      setStatus('Print not ready — reload the side panel.', 'error');
    }
  }

  const PRINT_CSS = `
    :root{ --yellow:#ffdd00; --magenta:#af006f; --ink:#1b1b1b; }
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;font-size:12px;margin:0;padding:12px;color:var(--ink);}
    .class-template-page{
      background:#fff;border:1px solid #d1d5db;border-radius:12px;padding:14px 16px 34px;margin-bottom:12px;
      box-shadow:0 6px 16px rgba(15,23,42,.08);position:relative;break-inside:avoid;
    }
    .page-header{
      display:flex;justify-content:space-between;align-items:baseline;gap:12px;
      border-bottom:2px solid var(--yellow);padding-bottom:6px;margin-bottom:8px;
    }
    .page-title{
      font-weight:600;font-size:1rem;text-transform:uppercase;letter-spacing:.06em;color:var(--magenta);
    }
    .page-meta{font-size:0.85rem;color:var(--ink);}
    .template-table{width:100%;border-collapse:collapse;font-size:0.85rem;margin-top:4px;}
    .template-table th,.template-table td{border:1px solid #d1d5db;padding:3px 6px;vertical-align:middle;}
    .template-table th{background:#f9fafb;font-weight:600;}
    .template-table-marking{font-size:10px;}
    .template-table-scroll-wrap{overflow-x:auto;max-width:100%;}
    .print-footer{
      position:absolute;left:16px;right:16px;bottom:12px;margin-top:12px;
      display:grid;grid-template-columns:1fr auto 1fr;align-items:end;gap:10px;
    }
    .print-footer-left,.print-footer-right{
      font-size:22px;font-weight:800;line-height:1.15;color:#111827;
      white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
    }
    .print-footer-left{justify-self:start;text-align:left;}
    .print-footer-right{justify-self:end;text-align:right;}
    .print-footer-center{
      justify-self:center;text-align:center;font-size:12px;font-weight:600;color:#6b7280;line-height:1.2;white-space:nowrap;
    }
    .print-preview-empty{padding:16px;text-align:center;color:#6b7280;border:1px dashed #ccc;border-radius:8px;}
    .drill-note{font-size:0.85rem;margin:8px 0;}
    @media print{
      .class-template-page{break-inside:avoid;box-shadow:none;}
    }
  `;

  function buildUI() {
    const root = document.getElementById('printTemplatesRoot');
    if (!root || root.querySelector('#ptBuilt')) return;
    root.innerHTML = '';

    const panel = document.createElement('div');
    panel.className = 'print-templates-panel';
    panel.id = 'ptBuilt';

    const rowCount = document.createElement('div');
    rowCount.className = 'pt-count';
    rowCount.id = 'ptRowCount';

    const actions = document.createElement('div');
    actions.className = 'pt-actions';
    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'btn primary';
    addBtn.textContent = 'Add students to template';
    addBtn.addEventListener('click', async () => {
      try {
        setStatus('Reading gradebook from page…', 'loading');
        const data = await BB.scrapeGradebookRaw();
        mergeIntoAccumulated(data.headers, data.rows, data.className);
        document.getElementById('printTemplateClassHint').textContent =
          (accumulated.className || 'Gradebook').slice(0, 36);
        rowCount.textContent = `${accumulated.rows.length} student(s) in template`;
        setStatus(`Merged — ${accumulated.rows.length} students total`, 'success');
        rebuildMarkingCheckboxes(panel);
        renderPreview(panel);
      } catch (e) {
        setStatus(`⚠ ${e.message}`, 'error');
      }
    });

    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'btn';
    clearBtn.textContent = 'Clear roster';
    clearBtn.addEventListener('click', () => {
      clearAccumulated();
      rowCount.textContent = '0 students in template';
      document.getElementById('printTemplateClassHint').textContent = 'Gradebook';
      setStatus('Roster cleared.', 'idle');
      rebuildMarkingCheckboxes(panel);
      renderPreview(panel);
    });

    actions.appendChild(addBtn);
    actions.appendChild(clearBtn);
    actions.appendChild(rowCount);

    const metaRow = document.createElement('div');
    metaRow.className = 'pt-meta';
    metaRow.innerHTML = `
      <label>Teacher <input type="text" id="ptTeacher" class="pt-input" placeholder="Name"/></label>
      <label>Class <input type="text" id="ptClass" class="pt-input" placeholder="Class info"/></label>
      <label>Term <select id="ptTerm" class="pt-input">
        <option value="T1">Term 1</option><option value="T2">Term 2</option><option value="T3">Term 3</option>
      </select></label>`;
    metaRow.querySelector('#ptTeacher').addEventListener('input', e => { printMeta.teacher = e.target.value; renderPreview(panel); });
    metaRow.querySelector('#ptClass').addEventListener('input', e => { printMeta.classInfo = e.target.value; renderPreview(panel); });
    metaRow.querySelector('#ptTerm').addEventListener('change', e => { printMeta.term = e.target.value; renderPreview(panel); });

    const tmpl = document.createElement('div');
    tmpl.className = 'pt-templates';
    tmpl.innerHTML = '<span class="pt-label">Template</span>';
    Object.entries(PRINT_TEMPLATE_DEFS).forEach(([id, def]) => {
      const lab = document.createElement('label');
      lab.className = 'pt-radio';
      const rb = document.createElement('input');
      rb.type = 'radio';
      rb.name = 'ptTpl';
      rb.value = id;
      if (id === selectedTemplate) rb.checked = true;
      rb.addEventListener('change', () => {
        selectedTemplate = id;
        renderPreview(panel);
      });
      lab.appendChild(rb);
      lab.appendChild(document.createTextNode(' ' + def.label));
      tmpl.appendChild(lab);
    });

    const drillRow = document.createElement('div');
    drillRow.className = 'pt-drill';
    drillRow.innerHTML = '<label>Drill count <input type="number" id="ptDrill" min="1" max="30" value="12"/></label>';
    drillRow.querySelector('#ptDrill').addEventListener('input', e => {
      drillCount = parseInt(e.target.value, 10) || 12;
      renderPreview(panel);
    });

    const markWrap = document.createElement('div');
    markWrap.className = 'pt-marking-list';
    markWrap.id = 'ptMarkingCols';

    const previewMount = document.createElement('div');
    previewMount.className = 'pt-preview';
    previewMount.id = 'ptPreviewMount';

    const printBtn = document.createElement('button');
    printBtn.type = 'button';
    printBtn.className = 'btn primary';
    printBtn.textContent = 'Print';
    printBtn.addEventListener('click', printPreview);

    panel.appendChild(actions);
    panel.appendChild(metaRow);
    panel.appendChild(tmpl);
    panel.appendChild(drillRow);
    panel.appendChild(markWrap);
    panel.appendChild(previewMount);
    panel.appendChild(printBtn);
    root.appendChild(panel);

    rebuildMarkingCheckboxes(panel);
    renderPreview(panel);

    if (BB.fetchTeacherNameFromPage) {
      queueMicrotask(async () => {
        try {
          const t = await BB.fetchTeacherNameFromPage();
          if (t && !(printMeta.teacher || '').trim()) {
            printMeta.teacher = t;
            const ptIn = document.getElementById('ptTeacher');
            if (ptIn) ptIn.value = t;
            renderPreview(panel);
          }
        } catch (_) {}
      });
    }
  }

  function rebuildMarkingCheckboxes(panel) {
    const wrap = panel.querySelector('#ptMarkingCols');
    if (!wrap) return;
    wrap.innerHTML = '<div class="pt-label">Marking columns (optional — defaults to first 12 assignment cols)</div>';
    if (!accumulated.headers.length) {
      wrap.innerHTML += '<p class="muted-tiny">Add students first.</p>';
      return;
    }
    const ctx = buildContext();
    const source = getMarkingSourceColumns(ctx);
    source.forEach(col => {
      const lab = document.createElement('label');
      lab.className = 'pt-cb';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = col;
      cb.checked = selectedMarkingCols.has(col);
      cb.addEventListener('change', () => {
        if (cb.checked) selectedMarkingCols.add(col);
        else selectedMarkingCols.delete(col);
        renderPreview(panel);
      });
      lab.appendChild(cb);
      lab.appendChild(document.createTextNode(' ' + col.slice(0, 48)));
      wrap.appendChild(lab);
    });
  }

  function init(api) {
    BB = api;
    if (inited) return;
    inited = true;
    try {
      const s = JSON.parse(localStorage.getItem('bb_print_meta') || '{}');
      if (s.teacher) printMeta.teacher = s.teacher;
      if (s.classInfo) printMeta.classInfo = s.classInfo;
      if (s.term) printMeta.term = s.term;
    } catch (_) {}
    buildUI();
    setStatus('Open Brightspace gradebook. Click “Add students to template” on each page.', 'idle');
    ['ptTeacher', 'ptClass', 'ptTerm'].forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (id === 'ptTeacher') el.value = printMeta.teacher;
      if (id === 'ptClass') el.value = printMeta.classInfo;
      if (id === 'ptTerm') el.value = printMeta.term;
    });
    setInterval(() => {
      try {
        localStorage.setItem('bb_print_meta', JSON.stringify({
          teacher: printMeta.teacher,
          classInfo: printMeta.classInfo,
          term: printMeta.term
        }));
      } catch (_) {}
    }, 2000);
  }

  global.BBPrintTemplates = { init, setStatus };
})(window);
