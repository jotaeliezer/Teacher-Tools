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
  const allTables = Array.from(document.querySelectorAll('table'));

  // Priority 1: a table that has a <th> cell whose text is exactly "Learner" or "Student"
  const byLearner = allTables.find(tbl => {
    return Array.from(tbl.querySelectorAll('th')).some(th =>
      /^(learner|student name?|student)$/i.test(cleanHeaderText(th.innerText))
    );
  });
  if (byLearner) return byLearner;

  // Priority 2: a D2L-specific selector
  for (const sel of ['table.d2l-table', '.d2l-table-wrapper table', '[data-grade-type] table']) {
    try {
      const el = document.querySelector(sel);
      if (el) {
        const tbl = el.tagName === 'TABLE' ? el : el.querySelector('table');
        if (tbl && tbl.rows.length > 1) return tbl;
      }
    } catch (_) {}
  }

  // Priority 3: largest table with at least 3 rows
  return allTables
    .filter(t => t.rows.length > 2)
    .sort((a, b) => (b.rows.length * (b.rows[0] ? b.rows[0].cells.length : 1))
                  - (a.rows.length * (a.rows[0] ? a.rows[0].cells.length : 1)))[0] || null;
}

// ── Header building ────────────────────────────────────────────────────────────
// Brightspace has a 2-row <thead>:
//   Row 1: "Learner" | "Final Grades" | "Enrolled before Week 26?" | ...
//   Row 2: ""        | "Final Calculated Grade ▼" | ""             | "L4..." | "L5..."
//
// We merge both rows per-column so we get the most specific header name.
// colspan attributes are handled so cell positions stay aligned.

function buildHeaders(tbl) {
  const thead = tbl.querySelector('thead');
  const headerRows = thead
    ? Array.from(thead.querySelectorAll('tr'))
    : (() => {
        const r = Array.from(tbl.rows).find(row => row.querySelector('th'));
        return r ? [r] : [tbl.rows[0]];
      })();

  if (!headerRows.length) return [];

  // Find total number of logical columns (accounting for colspan)
  let maxCols = 0;
  for (const row of headerRows) {
    let count = 0;
    for (const cell of Array.from(row.cells)) {
      count += parseInt(cell.getAttribute('colspan') || '1', 10);
    }
    maxCols = Math.max(maxCols, count);
  }

  // Fill headers: more specific (lower) header rows override group header rows
  const headers = new Array(maxCols).fill('');
  for (const row of headerRows) {
    let col = 0;
    for (const cell of Array.from(row.cells)) {
      const span = parseInt(cell.getAttribute('colspan') || '1', 10);
      const text = cleanHeaderText(cell.innerText);
      if (text) headers[col] = text; // last non-empty wins = most specific
      col += span;
    }
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

  const rows = dataRows.map(row => {
    const obj = {};
    const cells = Array.from(row.cells);
    headers.forEach((h, i) => {
      const key = h || `_col${i}`; // placeholder keeps cell alignment correct
      obj[key] = cells[i] ? cleanCellText(cells[i].innerText) : '';
    });
    return obj;
  }).filter(row =>
    // Remove rows where every cell is empty or just dashes
    Object.values(row).some(v => v && v !== '-' && v !== '- -' && v !== '--')
  );

  return { headers, rows };
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
    className: getClassName()
  };
}

// ── Message listener ───────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.action === 'scrapeGrades') {
    sendResponse(scrapeGrades());
  }
  return true; // keep channel open for async
});

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
