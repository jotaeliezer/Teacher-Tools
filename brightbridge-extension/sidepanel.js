// sidepanel.js — BrightBridge
// Handles all side-panel UI: loading grades, rendering students, generating comments.

// ── Constants ──────────────────────────────────────────────────────────────────

const API_ENDPOINT = 'https://jotaeliezer-teacher-tools-api-fj1k.vercel.app/api/generate-comment';
const CONCURRENCY  = 2;       // max parallel API calls (mirrors Teacher Tools)
const UNDERPERFORM_THRESHOLD = 60; // % at or below = underperforming

// ── State ──────────────────────────────────────────────────────────────────────

let students            = [];   // processed student objects
let generatedComments   = {};   // { studentIdx: string }
let showOnlyUnderperf   = false;
let isGeneratingAll     = false;

// ── DOM ────────────────────────────────────────────────────────────────────────

const $  = id => document.getElementById(id);
const refreshBtn             = $('refreshBtn');
const statusBar              = $('statusBar');
const statusText             = $('statusText');
const classLabel             = $('classLabel');
const controls               = $('controls');
const termSelect             = $('termSelect');
const gradeGroupSelect       = $('gradeGroupSelect');
const structureSelect        = $('structureSelect');
const filterUnderperformingBtn = $('filterUnderperformingBtn');
const studentCount           = $('studentCount');
const generateAllBtn         = $('generateAllBtn');
const studentList            = $('studentList');
const emptyState             = $('emptyState');

// ── Status helpers ─────────────────────────────────────────────────────────────

function setStatus(text, type = 'idle') {
  statusText.textContent = text;
  statusBar.className = `status-bar status-${type}`;
}

// ── Column detection ───────────────────────────────────────────────────────────

const GRADE_HINTS = [
  'final calculated grade',
  'calculated final grade',
  'final grade',
  'overall grade',
  'course grade',
  'grade'
];

function norm(h) {
  return (h || '')
    .replace(/[▼▲►◄▸▾⌄⌃]/g, '') // strip sort arrows Brightspace puts in headers
    .toLowerCase()
    .replace(/\s*#\s*$/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function detectCols(headers) {
  const norms = headers.map(norm);

  const lastCol  = headers.find((_, i) => ['last name', 'lastname', 'surname'].includes(norms[i]));
  const firstCol = headers.find((_, i) => ['first name', 'firstname', 'given name'].includes(norms[i]));
  // Include 'learner' and 'student' — used by Spirit of Math / D2L Brightspace
  const fullCol  = headers.find((_, i) => ['student name', 'name', 'full name', 'learner', 'student'].includes(norms[i]));

  let gradeCol = null;
  for (const hint of GRADE_HINTS) {
    gradeCol = headers.find((_, i) => norms[i] === hint || norms[i].startsWith(hint));
    if (gradeCol) break;
  }
  // fallback: any col with grade/mark/percent/score
  if (!gradeCol) {
    gradeCol = headers.find(h => /grade|mark|percent|score/i.test(h));
  }

  // Skip cols: ID, email, org cols
  const skipPatterns = /email|username|orgid|org\s*defined|id\s*#|userid/i;
  const skipCols = new Set(
    [lastCol, firstCol, fullCol, gradeCol, ...headers.filter(h => skipPatterns.test(h))].filter(Boolean)
  );

  return { lastCol, firstCol, fullCol, gradeCol, skipCols };
}

// Strip non-name characters (icons, arrows, checkboxes) that Brightspace injects into cells
function cleanNameText(raw) {
  return (raw || '')
    .replace(/[▼▲►◄▸▾⌄⌃⌘⚑⚐]/g, '')       // arrows and icons
    .replace(/[^\p{L}\p{M}\s'\-\.]/gu, ' ')  // keep letters (incl. accented), spaces, hyphens, apostrophes
    .replace(/\s+/g, ' ')
    .trim();
}

// ── Student processing ─────────────────────────────────────────────────────────

function processStudents({ headers, rows }) {
  const cols = detectCols(headers);

  return rows.map((row, idx) => {
    // Name — clean up Brightspace cell noise (icons, arrows, checkboxes)
    let name = '';
    if (cols.fullCol) {
      name = cleanNameText(row[cols.fullCol] || '');
    } else if (cols.firstCol && cols.lastCol) {
      const fn = cleanNameText(row[cols.firstCol] || '');
      const ln = cleanNameText(row[cols.lastCol] || '');
      name = `${fn} ${ln}`.trim();
    } else {
      // Broader fallback: matches "Learner", "Student Name", any name-like column
      const fallback = headers.find(h =>
        /name|learner|student/i.test(h) && !/grade|mark|score|percent/i.test(h)
      );
      name = fallback ? cleanNameText(row[fallback] || '') : '';
    }
    if (!name || name.length < 2) return null;

    const firstName = cols.firstCol
      ? cleanNameText(row[cols.firstCol] || '')
      : name.split(' ')[0];
    const lastName = cols.lastCol ? cleanNameText(row[cols.lastCol] || '') : '';

    // Grade
    const gradeRaw = cols.gradeCol ? (row[cols.gradeCol] || '') : '';
    const gradeNum = parseGradeNum(gradeRaw);

    // Assignment columns (everything except name/id/grade/email cols)
    const assignments = headers
      .filter(h => !cols.skipCols.has(h))
      .map(h => ({ label: h, value: (row[h] || '').trim() }))
      .filter(a => a.value && a.value !== '- -' && a.value !== '--' && a.value !== 'N/A');

    return { idx, name, firstName, lastName, gradeRaw, gradeNum, assignments, row };
  }).filter(Boolean);
}

function parseGradeNum(raw) {
  if (!raw) return null;
  // Extract the first number in the string (handles "81.34 % 👁", "0 %", "- -", etc.)
  const match = String(raw).match(/(\d+[.,]?\d*)/);
  if (!match) return null;
  const n = parseFloat(match[1].replace(',', '.'));
  return isNaN(n) ? null : n;
}

function formatGrade(num) {
  if (num === null) return '—';
  return num.toFixed(1).replace(/\.0$/, '') + '%';
}

function isUnderperforming(s) {
  return s.gradeNum !== null && s.gradeNum <= UNDERPERFORM_THRESHOLD;
}

function perfCode(gradeNum) {
  if (gradeNum === null) return 'satisfactory';
  if (gradeNum >= 90)   return 'good';
  if (gradeNum >= 75)   return 'satisfactory';
  if (gradeNum >= 60)   return 'average';
  return 'needs_support';
}

// ── Render ─────────────────────────────────────────────────────────────────────

function renderStudents() {
  studentList.innerHTML = '';

  if (!students.length) {
    emptyState.style.display = '';
    controls.style.display = 'none';
    return;
  }

  emptyState.style.display = 'none';
  controls.style.display = '';

  const visible = showOnlyUnderperf ? students.filter(isUnderperforming) : students;
  const sorted  = [...visible].sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

  const underCount = students.filter(isUnderperforming).length;
  studentCount.textContent = showOnlyUnderperf
    ? `${visible.length} underperforming`
    : `${students.length} students · ${underCount} ⚠️`;

  sorted.forEach(s => studentList.appendChild(buildCard(s)));
}

function buildCard(student) {
  const under   = isUnderperforming(student);
  const comment = generatedComments[student.idx] || '';

  const card = document.createElement('div');
  card.className = 'student-card' + (under ? ' underperforming' : '');
  card.dataset.idx = student.idx;

  // ── Card header ──
  const header = document.createElement('div');
  header.className = 'card-header';

  if (under) {
    const badge = document.createElement('span');
    badge.className = 'warning-badge';
    badge.textContent = '⚠️';
    header.appendChild(badge);
  }

  const nameEl = document.createElement('span');
  nameEl.className = 'student-name';
  nameEl.textContent = student.name;

  const gradeEl = document.createElement('span');
  gradeEl.className = 'student-grade' + (under ? ' grade-low' : '');
  gradeEl.textContent = formatGrade(student.gradeNum);

  header.appendChild(nameEl);
  header.appendChild(gradeEl);
  card.appendChild(header);

  // ── Comment textarea ──
  const textarea = document.createElement('textarea');
  textarea.className = 'comment-textarea';
  textarea.placeholder = 'Click ↺ Generate to create a comment…';
  textarea.value = comment;
  textarea.addEventListener('input', () => {
    generatedComments[student.idx] = textarea.value;
    updateCopyBtn(copyBtn, textarea.value);
  });
  card.appendChild(textarea);

  // ── Action row ──
  const actions = document.createElement('div');
  actions.className = 'card-actions';

  const genBtn = document.createElement('button');
  genBtn.type = 'button';
  genBtn.className = 'btn gen-btn';
  genBtn.textContent = '↺ Generate';

  const copyBtn = document.createElement('button');
  copyBtn.type = 'button';
  copyBtn.className = 'btn copy-btn' + (comment ? ' has-content' : '');
  copyBtn.textContent = comment ? 'Copy ✓' : 'Copy';

  genBtn.addEventListener('click', () => generateOne(student, textarea, genBtn, copyBtn));
  copyBtn.addEventListener('click', () => {
    const text = textarea.value.trim();
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      copyBtn.textContent = '✓ Copied!';
      setTimeout(() => {
        copyBtn.textContent = 'Copy ✓';
      }, 1600);
    });
  });

  actions.appendChild(genBtn);
  actions.appendChild(copyBtn);
  card.appendChild(actions);

  return card;
}

function updateCopyBtn(btn, text) {
  const has = !!(text && text.trim());
  btn.classList.toggle('has-content', has);
  btn.textContent = has ? 'Copy ✓' : 'Copy';
}

// ── API & comment generation ───────────────────────────────────────────────────

function buildPayload(student) {
  const term        = termSelect.value;
  const gradeGroup  = gradeGroupSelect.value;
  const structure   = structureSelect.value;
  const perf        = perfCode(student.gradeNum);

  // Up to 2 assignment facts for the prompt
  const assignmentFacts = student.assignments.slice(0, 2).map(a => ({
    label: a.label,
    value: a.value
  }));

  const classFirstNames = students.map(s => s.firstName).filter(Boolean).slice(0, 120);

  return {
    mode:                    'basic_bulk',
    reviseMode:              'basic_bulk',
    targetStructure:         structure,
    basicStructure:          structure,
    termLabel:               term,
    studentName:             student.name,
    studentFirstName:        student.firstName,
    pronounGuess:            'unknown',
    finalMark:               student.gradeRaw || '',
    performanceLevel:        perf,
    performanceLabel:        perf.replace('_', ' '),
    needsSupport:            perf === 'needs_support',
    gradeGroup,
    assignmentFacts,
    upcomingTests:           [],
    allowedAssignmentLabels: student.assignments.map(a => a.label),
    classFirstNames
  };
}

function parseApiResponse(data) {
  if (!data || typeof data !== 'object') return '';
  for (const key of ['comment', 'text', 'output', 'result']) {
    const t = String(data[key] || '').trim();
    if (t) return t;
  }
  return '';
}

// Typewriter animation — types the comment character by character into the textarea
function typewriterAnimate(textarea, text, onDone) {
  textarea.value = '';
  let i = 0;
  const CHAR_DELAY = 12; // ms per character — adjust for speed

  function typeNext() {
    if (i < text.length) {
      textarea.value += text[i];
      // Auto-grow height as text fills in
      textarea.style.height = 'auto';
      textarea.style.height = textarea.scrollHeight + 'px';
      i++;
      setTimeout(typeNext, CHAR_DELAY);
    } else {
      if (onDone) onDone();
    }
  }

  typeNext();
}

async function generateOne(student, textarea, genBtn, copyBtn) {
  genBtn.disabled = true;
  genBtn.textContent = '…';
  textarea.classList.add('loading');
  textarea.value = '';

  try {
    const res = await fetch(API_ENDPOINT, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(buildPayload(student))
    });

    if (!res.ok) throw new Error(`Server error ${res.status}`);
    const data    = await res.json();
    const comment = parseApiResponse(data);
    if (!comment) throw new Error('Empty response from server');

    generatedComments[student.idx] = comment;
    textarea.classList.remove('loading');

    // Typewriter animation — re-enable buttons only after typing finishes
    typewriterAnimate(textarea, comment, () => {
      updateCopyBtn(copyBtn, comment);
      genBtn.disabled    = false;
      genBtn.textContent = '↺ Generate';
    });
    return; // buttons re-enabled inside onDone above

  } catch (err) {
    textarea.value = `⚠ Error: ${err.message}. Click ↺ to retry.`;
  }

  textarea.classList.remove('loading');
  genBtn.disabled    = false;
  genBtn.textContent = '↺ Generate';
}

async function generateAll() {
  if (isGeneratingAll || !students.length) return;
  isGeneratingAll = true;
  generateAllBtn.disabled  = true;
  generateAllBtn.textContent = '⏳ Generating…';

  const visible = showOnlyUnderperf ? students.filter(isUnderperforming) : students;
  const queue   = [...visible]; // copy so mutations don't affect it
  let done = 0;

  setStatus(`Generating comments for ${visible.length} students…`, 'loading');

  async function worker() {
    while (queue.length) {
      const student = queue.shift();
      const card    = studentList.querySelector(`.student-card[data-idx="${student.idx}"]`);
      if (!card) continue;

      const textarea = card.querySelector('.comment-textarea');
      const genBtn   = card.querySelector('.gen-btn');
      const copyBtn  = card.querySelector('.copy-btn');
      if (textarea && genBtn && copyBtn) {
        await generateOne(student, textarea, genBtn, copyBtn);
      }
      done++;
      setStatus(`Generated ${done} / ${visible.length}…`, 'loading');
    }
  }

  // Run CONCURRENCY workers in parallel
  await Promise.all(Array.from({ length: CONCURRENCY }, worker));

  setStatus(`✓ ${done} comments ready!`, 'success');
  generateAllBtn.disabled  = false;
  generateAllBtn.textContent = '▶ Generate All Comments';
  isGeneratingAll = false;
}

// ── Data loading ───────────────────────────────────────────────────────────────

async function loadGrades() {
  setStatus('Loading grades from page…', 'loading');
  refreshBtn.disabled = true;
  emptyState.style.display = 'none';

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab) throw new Error('No active tab found.');

    // Inject content script if it hasn't loaded yet (e.g. page was already open)
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files:  ['content.js']
    }).catch(() => {}); // ignore if already injected

    const result = await chrome.tabs.sendMessage(tab.id, { action: 'scrapeGrades' });
    if (!result) throw new Error('No response from page. Try refreshing Brightspace.');
    if (!result.success) throw new Error(result.error);

    students = processStudents(result.data);
    if (!students.length) throw new Error('No student rows found in the grade table.');

    // Class label
    classLabel.textContent = result.className
      ? result.className.slice(0, 40)
      : `${students.length} students`;

    // Restore settings
    restoreSettings();

    generatedComments = {}; // reset comments on fresh load
    renderStudents();
    setStatus(`✓ ${students.length} students loaded`, 'success');

  } catch (err) {
    setStatus(`⚠ ${err.message}`, 'error');
    emptyState.style.display = '';
    controls.style.display   = 'none';
    studentList.innerHTML    = '';
  }

  refreshBtn.disabled = false;
}

// ── Settings persistence ───────────────────────────────────────────────────────

function saveSettings() {
  try {
    localStorage.setItem('bb_settings', JSON.stringify({
      term:       termSelect.value,
      gradeGroup: gradeGroupSelect.value,
      structure:  structureSelect.value
    }));
  } catch (_) {}
}

function restoreSettings() {
  try {
    const s = JSON.parse(localStorage.getItem('bb_settings') || '{}');
    if (s.term       && termSelect.querySelector(`option[value="${s.term}"]`))       termSelect.value = s.term;
    if (s.gradeGroup && gradeGroupSelect.querySelector(`option[value="${s.gradeGroup}"]`)) gradeGroupSelect.value = s.gradeGroup;
    if (s.structure  && structureSelect.querySelector(`option[value="${s.structure}"]`))  structureSelect.value = s.structure;
  } catch (_) {}
}

// ── Event wiring ───────────────────────────────────────────────────────────────

refreshBtn.addEventListener('click', loadGrades);

generateAllBtn.addEventListener('click', generateAll);

filterUnderperformingBtn.addEventListener('click', () => {
  showOnlyUnderperf = !showOnlyUnderperf;
  filterUnderperformingBtn.classList.toggle('active', showOnlyUnderperf);
  filterUnderperformingBtn.textContent = showOnlyUnderperf
    ? '✕ Show All Students'
    : '⚠️ Show Underperforming';
  renderStudents();
});

[termSelect, gradeGroupSelect, structureSelect].forEach(el => {
  el.addEventListener('change', saveSettings);
});

// Listen for content script notification (table appeared dynamically)
chrome.runtime.onMessage.addListener(msg => {
  if (msg.action === 'gradeTableDetected') {
    setStatus('Grade table detected — click ↻ to load', 'idle');
  }
});

// Auto-load on panel open
loadGrades();
