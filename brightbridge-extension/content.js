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

// ── Visibility helper ──────────────────────────────────────────────────────────
// Returns true if an element is visible in the layout (not hidden/collapsed).

function isCellVisible(cell) {
  if (!cell) return false;
  if (cell.hidden) return false;
  // getComputedStyle is reliable for detecting display:none on collapsed columns
  try {
    const style = window.getComputedStyle(cell);
    return style.display !== 'none' && style.visibility !== 'collapse';
  } catch (_) {
    return true; // if we can't check, assume visible
  }
}

// ── Header building ────────────────────────────────────────────────────────────
// Brightspace has a 2-row <thead>:
//   Row 1: "Learner" | "Final Grades" | "Tests & Quizzes" (colspan=6) | ...
//   Row 2: ""        | "Final Calculated Grade ▼" | "L6..." | "L7..." | "L35..."
//
// We merge both rows per-column so we get the most specific header name.
// colspan attributes are handled so cell positions stay aligned.
//
// IMPORTANT: Brightspace allows column groups to be collapsed. When collapsed,
// individual column <th> cells get display:none and the corresponding <td> cells
// are removed from / hidden in the tbody rows — causing an index offset if we
// naively map cells[i] → headers[i].
//
// We fix this by also tracking visibleIndices: the logical column positions that
// map to actually-visible <th> cells in the last header row. parseTable() uses
// this to match visible body cells to their correct header.

function buildHeaders(tbl) {
  const thead = tbl.querySelector('thead');
  const headerRows = thead
    ? Array.from(thead.querySelectorAll('tr'))
    : (() => {
        const r = Array.from(tbl.rows).find(row => row.querySelector('th'));
        return r ? [r] : [tbl.rows[0]];
      })();

  if (!headerRows.length) return { headers: [], visibleIndices: [] };

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

  // Build visibleIndices: logical column positions whose <th> in the last header
  // row is currently visible (not hidden by Brightspace column-group collapse).
  const lastRow = headerRows[headerRows.length - 1];
  const visibleIndices = [];
  let vi = 0;
  for (const cell of Array.from(lastRow.cells)) {
    const span = parseInt(cell.getAttribute('colspan') || '1', 10);
    if (isCellVisible(cell)) {
      for (let s = 0; s < span; s++) visibleIndices.push(vi + s);
    }
    vi += span;
  }

  // Fallback: if all cells appear hidden (e.g. getComputedStyle unavailable),
  // treat every column as visible so behaviour matches the original.
  const effectiveIndices = visibleIndices.length
    ? visibleIndices
    : headers.map((_, i) => i);

  return { headers, visibleIndices: effectiveIndices };
}

// ── Table parsing ──────────────────────────────────────────────────────────────

function parseTable(tbl) {
  if (!tbl) return null;

  const { headers, visibleIndices } = buildHeaders(tbl);
  if (!headers.length) return null;

  // Use <tbody> rows as data rows; if no <tbody>, use all non-<th> rows
  const tbody = tbl.querySelector('tbody');
  const dataRows = tbody
    ? Array.from(tbody.querySelectorAll('tr'))
    : Array.from(tbl.rows).filter(row => !row.querySelector('th'));

  if (!dataRows.length) return null;

  const rows = dataRows.map(row => {
    const obj = {};

    // Only read visible cells — hidden cells from collapsed column groups are
    // skipped so that cells[j] correctly maps to visibleIndices[j].
    const visibleCells = Array.from(row.cells).filter(cell => isCellVisible(cell));

    // If the visible-cell count doesn't match visibleIndices (unexpected DOM state),
    // fall back to using all cells with a direct index mapping.
    const useFallback = visibleCells.length === 0 ||
      Math.abs(visibleCells.length - visibleIndices.length) > 2;

    if (useFallback) {
      // Original behaviour: map cells[i] → headers[i]
      Array.from(row.cells).forEach((cell, i) => {
        const key = headers[i] || `_col${i}`;
        obj[key] = cleanCellText(cell.innerText);
      });
    } else {
      visibleCells.forEach((cell, j) => {
        const logicalIdx = visibleIndices[j] !== undefined ? visibleIndices[j] : j;
        const key = headers[logicalIdx] || `_col${logicalIdx}`;
        obj[key] = cleanCellText(cell.innerText);
      });
    }

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
