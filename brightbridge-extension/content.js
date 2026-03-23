// content.js — BrightBridge
// Runs silently on every Brightspace page.
// Finds the grade table and responds to messages from the side panel.

const BB_SELECTORS = [
  'd2l-grades-table table',
  'table.d2l-table',
  '.d2l-table-wrapper table',
  '[data-grade-type] table',
  '#z_grades table',
  'table[class*="grade"]',
  'table' // generic fallback
];

function findGradeTable() {
  for (const sel of BB_SELECTORS) {
    try {
      const el = document.querySelector(sel);
      if (!el) continue;
      const tbl = el.tagName === 'TABLE' ? el : el.querySelector('table');
      if (tbl && tbl.rows && tbl.rows.length > 1) return tbl;

      // Check shadow DOM (D2L uses web components)
      if (el.shadowRoot) {
        const shadowTbl = el.shadowRoot.querySelector('table');
        if (shadowTbl && shadowTbl.rows && shadowTbl.rows.length > 1) return shadowTbl;
      }
    } catch (_) { /* ignore bad selectors */ }
  }
  return null;
}

function parseTable(tbl) {
  if (!tbl) return null;

  const allRows = Array.from(tbl.rows);
  if (allRows.length < 2) return null;

  // Find header row — first row with <th> elements, or first row
  const headerRow = allRows.find(r => r.querySelector('th')) || allRows[0];
  const headers = Array.from(headerRow.cells)
    .map(c => c.innerText.trim())
    .filter(Boolean);

  if (!headers.length) return null;

  // All rows after the header
  const headerIdx = allRows.indexOf(headerRow);
  const dataRows = allRows
    .slice(headerIdx + 1)
    .map(row => {
      const obj = {};
      const cells = Array.from(row.cells);
      headers.forEach((h, i) => {
        obj[h] = cells[i] ? cells[i].innerText.trim() : '';
      });
      return obj;
    })
    .filter(row => Object.values(row).some(v => v && v !== '- -')); // skip empty rows

  return { headers, rows: dataRows };
}

function getClassName() {
  // Try breadcrumb → page title → fallback
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

function scrapeGrades() {
  const tbl = findGradeTable();
  if (!tbl) {
    return {
      success: false,
      error: 'No grade table found on this page. Make sure you\'re on the Brightspace Grades page.'
    };
  }

  const data = parseTable(tbl);
  if (!data || !data.rows.length) {
    return {
      success: false,
      error: 'Grade table found but no student rows could be extracted.'
    };
  }

  return {
    success: true,
    data,
    className: getClassName()
  };
}

// Listen for messages from the side panel
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.action === 'scrapeGrades') {
    sendResponse(scrapeGrades());
  }
  return true; // keep channel open for async
});

// Watch for the grade table appearing after dynamic load
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
