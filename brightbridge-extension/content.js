// content.js — BrightBridge
// Runs silently on every Brightspace page.
// Finds the grade table and responds to messages from the side panel.

// ── Text helpers ───────────────────────────────────────────────────────────────

function cleanHeaderText(text) {
  // Strip sort-arrow characters and extra whitespace from header cells
  return (text || '')
    .replace(/[▼▲►◄▸▾⌄⌃⚑⚐]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanCellText(text) {
  return (text || '').replace(/\s+/g, ' ').trim();
}

// ── Table finding ──────────────────────────────────────────────────────────────

function findGradeTable() {
  // Helper: search for grade table inside a DOM root (regular or shadow)
  function searchRoot(root) {
    const allTables = Array.from(root.querySelectorAll('table'));

    // Priority 1: table with a <th> that says "Learner" or "Student"
    const byLearner = allTables.find(tbl =>
      Array.from(tbl.querySelectorAll('th')).some(th =>
        /^(learner|student name?|student)$/i.test(cleanHeaderText(th.innerText))
      )
    );
    if (byLearner) return byLearner;

    // Priority 2: D2L-specific class selectors
    for (const sel of ['table.d2l-table', '.d2l-table-wrapper table', '[data-grade-type] table']) {
      try {
        const el = root.querySelector(sel);
        if (el) {
          const tbl = el.tagName === 'TABLE' ? el : el.querySelector('table');
          if (tbl && tbl.rows.length > 1) return tbl;
        }
      } catch (_) {}
    }

    // Priority 3: largest table with 3+ rows
    return allTables
      .filter(t => t.rows.length > 2)
      .sort((a, b) =>
        (b.rows.length * (b.rows[0] ? b.rows[0].cells.length : 1)) -
        (a.rows.length * (a.rows[0] ? a.rows[0].cells.length : 1))
      )[0] || null;
  }

  // Search regular DOM first (Brightspace typically renders in regular DOM)
  const fromDOM = searchRoot(document);
  if (fromDOM) return fromDOM;

  // Fallback: search open shadow roots (some Brightspace instances use web components)
  for (const el of Array.from(document.querySelectorAll('*'))) {
    try {
      if (el.shadowRoot) {
        const fromShadow = searchRoot(el.shadowRoot);
        if (fromShadow) return fromShadow;
      }
    } catch (_) {}
  }

  return null;
}

// ── Header building ────────────────────────────────────────────────────────────
//
// Brightspace uses a 2-row <thead>:
//
//   Row 0 (group headers):
//     "" rs=2 | "Learner" rs=2 | "Final Grades" cs=1 | "Enrolled…" rs=2 |
//     "Marked Assignments" cs=19 | "Homework Completion" cs=2 | "Tests and Drills" cs=18 |
//     ... more groups ... | "Final Exam Version" rs=2 | "Bonus Marks" rs=2 | ...
//
//   Row 1 (individual column names):
//     "Final Calculated Grade" | "L3 - Review 1" | "L3 - Review 2" | ...
//     (Only 103 cells — 8 positions are already claimed by rowspan=2 cells from row 0)
//
// KEY BUG (now fixed): The old code started col=0 for row 1 without skipping the
// positions claimed by rowspan=2 cells from row 0. So "Final Calculated Grade"
// landed at col 0 (overwriting ""), "L3 - Review 1" landed at col 1 (overwriting
// "Learner"), etc. — ALL row-1 headers were shifted left, causing every mark to
// be stored under the wrong column key.
//
// FIX: Use a 2D grid (sparse object) to track which [row][col] positions are
// already occupied by a rowspan from above. For each cell, advance past occupied
// positions before placing the cell's text.

function buildHeaders(tbl) {
  const thead = tbl.querySelector('thead');
  const headerRows = thead
    ? Array.from(thead.querySelectorAll('tr'))
    : (() => {
        const r = Array.from(tbl.rows).find(row => row.querySelector('th'));
        return r ? [r] : [tbl.rows[0]];
      })();

  if (!headerRows.length) return [];

  // 2D sparse grid: grid[rowIndex][colIndex] = headerText
  // Cells with rowspan>1 fill multiple row positions in the grid, so later rows
  // correctly skip those positions when advancing the column counter.
  const grid = {}; // { [r]: { [c]: string } }
  let maxCols = 0;

  for (let r = 0; r < headerRows.length; r++) {
    if (!grid[r]) grid[r] = {};
    let c = 0;

    for (const cell of Array.from(headerRows[r].cells)) {
      // Skip past columns already claimed by a rowspan from a previous row
      while (grid[r][c] !== undefined) c++;

      const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
      const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
      const text    = cleanHeaderText(cell.innerText);

      // Fill the grid for every position this cell spans
      for (let rs = 0; rs < rowspan; rs++) {
        if (!grid[r + rs]) grid[r + rs] = {};
        for (let cs = 0; cs < colspan; cs++) {
          grid[r + rs][c + cs] = text;
        }
      }

      c += colspan;
      if (c > maxCols) maxCols = c;
    }
  }

  // Build the flat headers array.
  // For each logical column, the value from the LAST row that has text wins
  // (individual column names in row 1 override group header names from row 0).
  // Rowspan=2 cells set the same text in both rows, so they stay correct too.
  const headers = [];
  for (let c = 0; c < maxCols; c++) {
    let value = '';
    for (let r = 0; r < headerRows.length; r++) {
      const v = (grid[r] && grid[r][c] !== undefined) ? grid[r][c] : '';
      if (v) value = v; // last non-empty wins = most specific row
    }
    headers.push(value);
  }

  return headers;
}

// ── Table parsing ──────────────────────────────────────────────────────────────

function parseTable(tbl) {
  if (!tbl) return null;

  const headers = buildHeaders(tbl);
  if (!headers.length) return null;

  // Use <tbody> rows as data rows; if no <tbody>, use all non-<th> rows
  const tbody = tbl.querySelector('tbody');
  const dataRows = tbody
    ? Array.from(tbody.querySelectorAll('tr'))
    : Array.from(tbl.rows).filter(row => !row.querySelector('th'));

  if (!dataRows.length) return null;

  // Each data row has one cell per logical column (matching headers[]).
  // Direct index mapping: cells[i] → headers[i].
  const rows = dataRows.map(row => {
    const obj = {};
    const cells = Array.from(row.cells);
    headers.forEach((h, i) => {
      const key = h || `_col${i}`;
      obj[key] = cells[i] ? cleanCellText(cells[i].innerText) : '';
    });
    return obj;
  }).filter(row =>
    // Remove rows where every cell is empty or just dashes
    Object.values(row).some(v => v && v !== '-' && v !== '- -' && v !== '--')
  );

  return { headers, rows };
}

// ── Pagination helpers ─────────────────────────────────────────────────────────

// Finds the "Results Per Page" <select> and maximises it so all students are
// visible in one scrape. Returns true if the value was changed (meaning the
// page will reload — the teacher should click ↻ again after it settles).
function tryMaximiseRowsPerPage() {
  // Brightspace labels the select with an off-screen <label> containing
  // "Results Per Page" linked via a `for` attribute.
  const label = Array.from(document.querySelectorAll('label'))
    .find(l => /results per page/i.test(l.textContent));
  if (!label) return false;

  let sel = document.getElementById(label.getAttribute('for'));
  if (!sel) {
    // Fall back: find a <select> inside the element following the label
    const next = label.nextElementSibling;
    sel = next && next.querySelector('select');
  }
  if (!sel || !sel.options.length) return false;

  // If Brightspace doesn't offer a 200-option, add one so we can use it
  if (!Array.from(sel.options).some(o => o.value === '200')) {
    sel.appendChild(new Option('200 per page', '200'));
  }

  if (sel.value === '200') return false; // already showing all

  sel.value = '200';
  sel.dispatchEvent(new Event('change', { bubbles: true }));
  console.log('[BrightBridge] Results Per Page changed to 200 — table will reload');
  return true;
}

// Returns { current, total } page numbers if Brightspace is paginating,
// or null if only one page (or pagination control not found).
function detectPagination() {
  // The "Page Number" label is linked to a <select> whose options are "1", "2", ...
  const label = Array.from(document.querySelectorAll('label'))
    .find(l => /^page number$/i.test(l.textContent.trim()));
  if (!label) return null;

  let sel = document.getElementById(label.getAttribute('for'));
  if (!sel) {
    const next = label.nextElementSibling;
    sel = next && next.querySelector('select');
  }
  if (!sel || sel.options.length <= 1) return null; // only 1 page

  const current = parseInt(sel.value) || 1;
  const total   = sel.options.length; // each option = one page
  return { current, total };
}

// ── Class name detection ───────────────────────────────────────────────────────

function getClassName() {
  const breadSelectors = [
    '.d2l-navigation-s-breadcrumbs li:last-child',
    '.d2l-breadcrumbs li:last-child',
    'nav[aria-label] li:last-child',
    'ol.d2l-link li:last-child'
  ];
  for (const sel of breadSelectors) {
    const el = document.querySelector(sel);
    if (el && el.innerText.trim()) return el.innerText.trim();
  }
  return document.title.replace(/\s*[-|–]\s*Brightspace.*/i, '').trim();
}

// ── Main scrape ────────────────────────────────────────────────────────────────

function scrapeGrades() {
  const allTables = Array.from(document.querySelectorAll('table'));
  console.log('[BrightBridge] total tables on page:', allTables.length);

  const tbl = findGradeTable();
  if (!tbl) {
    return {
      success: false,
      error: `No grade table found. ${allTables.length} table(s) on page. Make sure you're on the Brightspace Grades page.`
    };
  }

  const thead = tbl.querySelector('thead');
  const tbody = tbl.querySelector('tbody');
  console.log('[BrightBridge] table found — rows:', tbl.rows.length,
    '| thead rows:', thead ? thead.querySelectorAll('tr').length : 0,
    '| tbody rows:', tbody ? tbody.querySelectorAll('tr').length : 0);

  // Check for pagination — Brightspace may be showing only one page of students.
  // We cannot reliably trigger a reload via JS (Brightspace ignores programmatic
  // change events on its selects), so we scrape what is visible and return a
  // warning so the side panel can prompt the teacher to fix it manually.
  const pagination = detectPagination();
  const paginationWarning = (pagination && pagination.total > 1)
    ? `⚠️ Brightspace is showing page ${pagination.current} of ${pagination.total} — not all students are loaded. In Brightspace, set "Results Per Page" to the highest value, then click ↻.`
    : null;

  if (paginationWarning) {
    console.log(`[BrightBridge] Pagination detected: page ${pagination.current} of ${pagination.total}`);
  }

  const data = parseTable(tbl);
  console.log('[BrightBridge] parseTable result — headers:', JSON.stringify(data?.headers),
    '| data rows:', data?.rows?.length ?? 0);
  if (data?.rows?.length) console.log('[BrightBridge] first data row:', JSON.stringify(data.rows[0]));

  if (!data || !data.rows.length) {
    return {
      success: false,
      error: `Grade table found (${tbl.rows.length} total rows) but no data rows extracted. Try switching to Spreadsheet View in Brightspace.`
    };
  }

  return {
    success: true,
    data,
    className: getClassName(),
    paginationWarning
  };
}

// ── Single-student page scraper ────────────────────────────────────────────────

function scrapeSingleStudent() {
  const h1 = document.querySelector('h1');
  if (!h1 || !/final calculated grade/i.test(document.body.innerText)) {
    return {
      success: false,
      error: "Not on a single-student grade page. In Brightspace, open the Grades tab and click on a student's name first."
    };
  }

  // 1. Student name — strip D2L dropdown arrow glyphs from h1 text
  const studentName = h1.innerText
    .replace(/[▼▲›∨⌄↓⌃\uFE0F]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  // 2. Final grade % — find "Scheme: XX.XX %" within 400 chars after "Final Calculated Grade"
  let finalPercent = null;
  const bodyText = document.body.innerText;
  const fcgMatch = bodyText.match(/final calculated grade[\s\S]{0,400}?scheme:\s*(\d+\.?\d*)\s*%/i);
  if (fcgMatch) finalPercent = parseFloat(fcgMatch[1]);

  // 3. Grade items — find elements whose text contains a lesson number (L1–L99)
  //    then walk up the DOM to find their enclosing container's "Scheme: XX %" value.
  //    Items WITH a scheme % = graded assignments.
  //    Items WITHOUT a scheme % = upcoming (no mark entered yet).
  const assignments   = [];
  const upcomingItems = [];
  const seen = new Set();
  document.querySelectorAll('h3, h4, strong, [class*="name"], [class*="title"]').forEach(el => {
    const text = el.innerText?.trim();
    if (!text || seen.has(text) || text.length > 120) return;
    if (!/\bL\d{1,2}\b/i.test(text)) return;
    seen.add(text);
    let percent = null;
    let node = el.parentElement;
    for (let i = 0; i < 6; i++) {
      if (!node) break;
      const m = node.innerText?.match(/scheme:\s*(\d+\.?\d*)\s*%/i);
      if (m) { percent = parseFloat(m[1]); break; }
      node = node.parentElement;
    }
    if (percent !== null) {
      assignments.push({ name: text, percent });
    } else {
      upcomingItems.push({ name: text });
    }
  });

  return { success: true, isSingleStudent: true, studentName, finalPercent, assignments, upcomingItems };
}

// ── Message listener ───────────────────────────────────────────────────────────
// Guard: only register once, even if content.js is injected multiple times.

// ── Fill grades (Marking Assistant push) ──────────────────────────────────────

function waitForInput(container, timeoutMs) {
  return new Promise(resolve => {
    const check = () => {
      const input = container.querySelector('input[type="text"], input:not([type])')
        || container.closest('tr')?.querySelector('input[type="text"], input:not([type])');
      if (input) { obs.disconnect(); resolve(input); }
    };
    const obs = new MutationObserver(check);
    obs.observe(document.body, { childList: true, subtree: true });
    check();
    setTimeout(() => { obs.disconnect(); resolve(null); }, timeoutMs);
  });
}

async function fillGrades(rows, columnName) {
  const tbl = findGradeTable();
  if (!tbl) return { success: false, error: 'Grade table not found. Navigate to the Brightspace gradebook page first.' };

  const parsed = parseTable(tbl);
  if (!parsed) return { success: false, error: 'Could not parse grade table.' };

  const { headers } = parsed;
  const colIdx = headers.findIndex(h => h === columnName);
  if (colIdx === -1) return { success: false, error: `Column "${columnName}" not found in the table.` };

  const studentColIdx = headers.findIndex(h => /^(learner|student)/i.test(h));
  const tbody = tbl.querySelector('tbody');
  const dataRows = tbody ? [...tbody.querySelectorAll('tr')] : [];

  const normalize = s => s.toLowerCase().replace(/[,.]/g, '').replace(/\s+/g, ' ').trim();

  const results = [];
  for (const row of rows) {
    const name  = row['Student Name'] || '';
    const score = row['Score'] || '';
    if (!name || name === '?') { results.push({ name, status: 'skipped' }); continue; }

    const normName = normalize(name);
    const parts    = normName.split(' ');
    const swapped  = parts.length >= 2 ? `${parts.slice(1).join(' ')} ${parts[0]}` : normName;

    const matchRow = dataRows.find(tr => {
      const cell = tr.cells[studentColIdx];
      if (!cell) return false;
      const cellText = normalize(cell.innerText);
      return cellText.includes(normName) || cellText.includes(swapped);
    });

    if (!matchRow) { results.push({ name, status: 'not_found' }); continue; }

    const targetCell = matchRow.cells[colIdx];
    if (!targetCell) { results.push({ name, status: 'no_cell' }); continue; }

    targetCell.dispatchEvent(new MouseEvent('click',     { bubbles: true }));
    targetCell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));

    const input = await waitForInput(targetCell, 1500);
    if (!input) { results.push({ name, status: 'input_timeout' }); continue; }

    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
    if (setter) setter.call(input, score);
    else input.value = score;

    input.dispatchEvent(new Event('input',  { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.dispatchEvent(new FocusEvent('blur', { bubbles: true }));

    await new Promise(r => setTimeout(r, 350));
    results.push({ name, status: 'ok' });
  }

  // Try to click Brightspace's save button if one exists
  let saved = false;
  const saveBtn = document.querySelector('[data-key="save"], button[title*="Save"], .d2l-button-primary[type="submit"]');
  if (saveBtn) { saveBtn.click(); saved = true; }

  return { success: true, results, saved };
}

// ── Message listener ──────────────────────────────────────────────────────────

if (!window.__bbListenerRegistered) {
  window.__bbListenerRegistered = true;

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message.action === 'scrapeGrades') {
      sendResponse(scrapeGrades());
      return false;
    }
    if (message.action === 'scrapeSingleStudent') {
      sendResponse(scrapeSingleStudent());
      return false;
    }
    if (message.action === 'fillGrades') {
      fillGrades(message.rows, message.columnName).then(sendResponse);
      return true; // async
    }
    return false;
  });
}

// ── MutationObserver: notify panel when grade table appears ───────────────────

let bbObserver = null;
function startObserver() {
  if (bbObserver) bbObserver.disconnect();
  bbObserver = new MutationObserver(() => {
    if (findGradeTable()) {
      chrome.runtime.sendMessage({ action: 'gradeTableDetected' }).catch(() => {});
      bbObserver.disconnect();
    }
  });
  bbObserver.observe(document.body, { childList: true, subtree: true });
}

if (document.body) {
  startObserver();
} else {
  document.addEventListener('DOMContentLoaded', startObserver);
}
