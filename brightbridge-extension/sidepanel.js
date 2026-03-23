// sidepanel.js — BrightBridge

// ── Constants ──────────────────────────────────────────────────────────────────

const API_ENDPOINT = 'https://jotaeliezer-teacher-tools-api-fj1k.vercel.app/api/generate-comment';
const CONCURRENCY  = 2;
const UNDERPERFORM_THRESHOLD = 60;

const COMMENT_BANK_MINI = [
  { category: 'Participation', items: [
    'actively participates in class discussions',
    'consistently contributes thoughtful answers'
  ]},
  { category: 'Homework', items: [
    'completes all assigned work on time',
    'demonstrates strong effort on homework'
  ]},
  { category: 'Seeking Help', items: [
    'proactively asks questions when unsure',
    'makes great use of extra help sessions'
  ]},
  { category: 'Personal Qualities', items: [
    'shows excellent perseverance and resilience',
    'demonstrates a positive and cooperative attitude'
  ]},
  { category: 'Looking Ahead', items: [
    'is encouraged to review key concepts regularly',
    'has strong potential to excel next term'
  ]}
];

// ── Name-based pronoun detection ───────────────────────────────────────────────

const MALE_NAMES = new Set([
  'aaron','adam','adil','adnan','adrian','ahmed','aidan','aiden','alan','albert',
  'alejandro','alex','alexander','alexei','ali','amir','andre','andrew','andy',
  'anthony','antonio','arjun','arthur','ashraf','austin','ayaan','ayden',
  'beau','ben','benjamin','bob','brady','brandon','brian','bruce','bryan',
  'caleb','carlos','carter','charles','charlie','christian','christopher',
  'colton','connor','cory','curtis','cyrus',
  'daniel','darius','david','dean','derek','devraj','diego','dominic','dylan',
  'eli','elijah','elliot','ethan','evan',
  'fares','farhan','farhaan','fariz','felix','finn','francisco',
  'gabriel','gavin','george','giovanni','grant','griffin',
  'hamza','hassan','henry','hudson','hugo','hunter',
  'ian','ibrahim','isaac','isaiah','ivan',
  'jack','jackson','jacob','jake','james','jason','javier','jayden','jeremy',
  'jesse','joe','joel','john','jonathan','jose','joseph','joshua',
  'juan','julian','justin',
  'karim','kevin','khalid','kyle',
  'landon','leo','levi','liam','logan','lorenzo','lucas','luke',
  'marcus','mark','mason','matthew','max','michael','miguel','miles',
  'mohammed','muhammad','mustafa',
  'nate','nathan','nicholas','noah',
  'oliver','omar','owen',
  'patrick','paulo','peter','philip',
  'rafael','rajan','rami','richard','robert','ryan',
  'samir','samuel','sean','sebastian','seth','shaan','shane','simon',
  'stephen','steven',
  'thomas','tim','tobias','tyler',
  'umar','usman',
  'victor','vincent',
  'wesley','william',
  'xavier',
  'yousuf','yusef',
  'zachary','zane','ziad','zubair'
]);

const FEMALE_NAMES = new Set([
  'aaliyah','abby','abigail','ada','aisha','alana','alannah','alexa','alexandra',
  'alexia','alice','alicia','alisha','alina','aliya','aliyah','alondra','alyssa',
  'amanda','amber','amelia','amira','amy','ana','anastasia','aneesa','aneesha',
  'angela','angie','annika','annie','ariana','asha','ashley','astrid','audrey',
  'ava','ayesha',
  'beatrice','bella','beth','bianca','brianna','brooke','brooklyn',
  'caitlin','camila','cara','carolina','caroline','cassandra','charlotte',
  'chloe','christina','claire','clara','claudia',
  'daisy','dana','daniela','diana',
  'elena','eliza','elizabeth','ella','emilia','emily','emma','erica','eva','evelyn',
  'faith','fatima','fiona','freya',
  'gabriela','gemma','gianna','grace',
  'hana','hanna','hannah','harper','hayley','heather','helen','holly',
  'isabella','isla',
  'jade','jasmine','jennifer','jessica','jillian','julia','julianna','julie',
  'kaitlyn','karen','katelyn','katherine','kathryn','katie','kayla','kelly',
  'kendra','kylie',
  'laila','layla','laura','lauren','leah','leila','lillian','lily','lindsay',
  'lisa','lola','lucy','luna','lydia',
  'madison','maeve','maria','mariam','marie','maya','megan','mia','michelle',
  'miranda','molly',
  'nadia','natalie','nicole','noor','nora',
  'olivia',
  'paige','patricia','penelope','priya',
  'rachel','rebecca','rosa','rose','ruby',
  'samantha','sandra','sara','sarah','savannah','scarlett','serena','sienna',
  'simone','sofia','sophia','stephanie','suman',
  'tiffany',
  'valentina','vanessa','victoria','violet','vivera',
  'whitney',
  'yara',
  'zainab','zara','zoe','zoey'
]);

function guessPronounFromName(firstName) {
  if (!firstName) return 'unknown';
  const lower = firstName.toLowerCase().trim();
  if (MALE_NAMES.has(lower))   return 'he';
  if (FEMALE_NAMES.has(lower)) return 'she';
  return 'unknown';
}

// ── State ──────────────────────────────────────────────────────────────────────

let students          = [];
let generatedComments = {};
let showOnlyUnderperf = false;
let isGeneratingAll   = false;

// Raw scrape data — kept for assignment overlay stats
let rawRows       = [];
let allAssignCols = [];

// Selections from the Generate All overlay
let overlayAssignCols   = [];
let overlayUpcomingCols = [];

const cardStates = {};

function getCardState(idx, student) {
  if (!cardStates[idx]) {
    cardStates[idx] = {
      collapsed:    false,
      pronoun:      student ? guessPronounFromName(student.firstName) : 'unknown',
      customNote:   '',
      selectedBank: new Set(),
      advancedOpen: false,
      perfOverride: null
    };
  } else if (cardStates[idx].pronoun === 'unknown' && student) {
    // Re-detect in case this is a stale state from a previous load
    const guess = guessPronounFromName(student.firstName);
    if (guess !== 'unknown') cardStates[idx].pronoun = guess;
  }
  return cardStates[idx];
}

// Normalize "Term 1" / "TERM 1" / "T1" → "T1" for TERM_LESSON_RANGES lookup
function normalizeTermCode(val) {
  const v = String(val || '').trim().toUpperCase().replace(/\s+/g, '');
  if (v === 'T1' || v === 'TERM1') return 'T1';
  if (v === 'T2' || v === 'TERM2') return 'T2';
  if (v === 'T3' || v === 'TERM3') return 'T3';
  return '';
}

// ── DOM ────────────────────────────────────────────────────────────────────────

const $  = id => document.getElementById(id);
const refreshBtn               = $('refreshBtn');
const darkToggleBtn            = $('darkToggleBtn');
const statusBar                = $('statusBar');
const statusText               = $('statusText');
const classLabel               = $('classLabel');
const controls                 = $('controls');
const termSelect               = $('termSelect');
const gradeGroupSelect         = $('gradeGroupSelect');
const structureSelect          = $('structureSelect');
const filterUnderperformingBtn = $('filterUnderperformingBtn');
const studentCount             = $('studentCount');
const generateAllBtn           = $('generateAllBtn');
const studentList              = $('studentList');
const emptyState               = $('emptyState');

// ── Dark mode toggle ────────────────────────────────────────────────────────────

function applyDarkMode(dark) {
  document.body.classList.toggle('dark', dark);
  darkToggleBtn.textContent = dark ? '☀️' : '🌙';
  darkToggleBtn.title = dark ? 'Switch to light mode' : 'Switch to dark mode';
}

// Restore preference on load (default: light)
applyDarkMode(localStorage.getItem('bb-dark') === '1');

darkToggleBtn.addEventListener('click', () => {
  const isDark = document.body.classList.contains('dark');
  localStorage.setItem('bb-dark', isDark ? '0' : '1');
  applyDarkMode(!isDark);
});

// ── Status helpers ─────────────────────────────────────────────────────────────

function setStatus(text, type = 'idle') {
  statusText.textContent = text;
  statusBar.className = `status-bar status-${type}`;
}

// ── Column detection ───────────────────────────────────────────────────────────

const GRADE_HINTS = [
  'final calculated grade','calculated final grade','final grade',
  'overall grade','course grade','grade'
];

function norm(h) {
  return (h || '')
    .replace(/[▼▲►◄▸▾⌄⌃]/g, '')
    .toLowerCase().replace(/\s*#\s*$/, '').replace(/\s+/g, ' ').trim();
}

function detectCols(headers) {
  const norms = headers.map(norm);
  const lastCol  = headers.find((_, i) => ['last name','lastname','surname'].includes(norms[i]));
  const firstCol = headers.find((_, i) => ['first name','firstname','given name'].includes(norms[i]));
  const fullCol  = headers.find((_, i) => ['student name','name','full name','learner','student'].includes(norms[i]));

  let gradeCol = null;
  for (const hint of GRADE_HINTS) {
    gradeCol = headers.find((_, i) => norms[i] === hint || norms[i].startsWith(hint));
    if (gradeCol) break;
  }
  if (!gradeCol) gradeCol = headers.find(h => /grade|mark|percent|score/i.test(h));

  const skipPatterns = /email|username|orgid|org\s*defined|id\s*#|userid/i;
  const skipCols = new Set(
    [lastCol, firstCol, fullCol, gradeCol, ...headers.filter(h => skipPatterns.test(h))].filter(Boolean)
  );
  return { lastCol, firstCol, fullCol, gradeCol, skipCols };
}

function cleanNameText(raw) {
  return (raw || '')
    .replace(/[▼▲►◄▸▾⌄⌃⌘⚑⚐]/g, '')
    .replace(/[^\p{L}\p{M}\s'\-\.]/gu, ' ')
    .replace(/\s+/g, ' ').trim();
}

// ── Student processing ─────────────────────────────────────────────────────────

function rowGet(row, key) {
  if (!key) return '';
  if (row[key] !== undefined) return row[key];
  const lk = key.toLowerCase().trim();
  const found = Object.keys(row).find(k => k.toLowerCase().trim() === lk);
  return found ? row[found] : '';
}

function processStudents({ headers, rows }) {
  console.log('[BrightBridge] headers:', JSON.stringify(headers));
  console.log('[BrightBridge] row count:', rows.length);
  if (rows.length) console.log('[BrightBridge] first row sample:', JSON.stringify(rows[0]));

  const cols = detectCols(headers);
  console.log('[BrightBridge] detected cols — fullCol:', cols.fullCol, '| gradeCol:', cols.gradeCol);

  // Save for overlay use
  rawRows = rows;
  allAssignCols = headers.filter(h => h && !cols.skipCols.has(h));

  return rows.map((row, idx) => {
    let name = '';
    if (cols.fullCol) {
      name = cleanNameText(rowGet(row, cols.fullCol));
    } else if (cols.firstCol && cols.lastCol) {
      name = `${cleanNameText(rowGet(row, cols.firstCol))} ${cleanNameText(rowGet(row, cols.lastCol))}`.trim();
    }
    if (!name || name.length < 2) {
      const fallback = headers.find(h => /name|learner|student/i.test(h) && !/grade|mark|score|percent|enrolled/i.test(h));
      if (fallback) name = cleanNameText(rowGet(row, fallback));
    }
    if (!name || name.length < 2) {
      for (const h of headers) {
        const val = cleanNameText(rowGet(row, h));
        if (val.length >= 2 && /[a-zA-Z]{2,}/.test(val) && !/grade|calculated|percent|enrolled|week|exercise|mastermind/i.test(val)) {
          name = val; break;
        }
      }
    }
    if (!name || name.length < 2) return null;

    const firstName = cols.firstCol ? cleanNameText(rowGet(row, cols.firstCol)) : name.split(' ')[0];
    const lastName  = cols.lastCol  ? cleanNameText(rowGet(row, cols.lastCol))  : '';
    const gradeRaw  = cols.gradeCol ? rowGet(row, cols.gradeCol) : '';
    const gradeNum  = parseGradeNum(gradeRaw);

    const assignments = headers
      .filter(h => !cols.skipCols.has(h))
      .map(h => ({ label: h, value: (rowGet(row, h) || '').trim() }))
      .filter(a => a.value && /\d/.test(a.value));

    return { idx, name, firstName, lastName, gradeRaw, gradeNum, assignments, row };
  }).filter(Boolean);
}

function parseGradeNum(raw) {
  if (!raw) return null;
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

// ── Assignment overlay helpers ─────────────────────────────────────────────────

function getLessonNum(col) {
  const m = String(col || '').toUpperCase().match(/L\s*(\d+)/);
  return m ? (Number(m[1]) || null) : null;
}

const TERM_LESSON_RANGES = { T1: [1, 13], T2: [14, 26], T3: [27, 39] };

function getTermAssignCols(termCode) {
  const tc = normalizeTermCode(termCode);
  if (!tc) return allAssignCols.filter(c => colHasMarks(c));
  return allAssignCols.filter(col => {
    const n = getLessonNum(col);
    if (n == null) return false;
    const [min, max] = TERM_LESSON_RANGES[tc] || [];
    return min != null && n >= min && n <= max;
  });
}

// A cell value counts as "has a mark" only when it contains at least one digit
// Handles Brightspace's various no-mark formats: "- -", "--", "N/A", "-%", "- %", empty
function hasMark(v) {
  const s = String(v || '').trim();
  return s.length > 0 && /\d/.test(s);
}

function colHasMarks(col) {
  return rawRows.some(row => hasMark(row[col]));
}

function getColClassAvg(col) {
  const nums = rawRows.map(row => parseGradeNum(row[col])).filter(n => n !== null);
  if (!nums.length) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

function getColSubmitRate(col) {
  if (!rawRows.length) return null;
  const submitted = rawRows.filter(row => hasMark(row[col])).length;
  return (submitted / rawRows.length) * 100;
}

function cleanAssignLabel(label) {
  let s = String(label || '').trim();
  s = s.replace(/\s+(Scheme Symbol|Points Grade|Points Points|Grade Points|Percentage|Letter Grade|GPA Scale|Pass\/Fail|Complete\/Incomplete)\s*$/i, '').trim();
  const m = s.match(/^L\d+\s*-\s*(.*)$/i);
  return m ? (m[1].trim() || s) : s;
}

function avgColor(avg) {
  if (!Number.isFinite(avg)) return '#bbb';
  if (avg >= 80) return '#1a7f4b';
  if (avg >= 70) return '#b07800';
  return '#c0186a';
}

// ── Assignment selection overlay (2-step) ──────────────────────────────────────

function showAssignOverlay(onConfirm) {
  const overlay = document.createElement('div');
  overlay.className = 'assign-overlay';

  const modal = document.createElement('div');
  modal.className = 'assign-modal';

  // ── Step 1: Assignments ──
  const step1 = document.createElement('div');
  step1.className = 'assign-step';

  const s1Title = document.createElement('div');
  s1Title.className = 'assign-title';
  s1Title.innerHTML = '<strong>Step 1 of 2</strong> — Select assignments to highlight';

  const s1Hint = document.createElement('div');
  s1Hint.className = 'assign-hint';
  s1Hint.textContent = 'Choose up to 3. Avg = class average · Sub = submission rate.';

  const s1Search = document.createElement('input');
  s1Search.className = 'assign-search';
  s1Search.placeholder = 'Search assignments…';
  s1Search.type = 'text';

  const s1Warning = document.createElement('div');
  s1Warning.className = 'assign-warning';
  s1Warning.textContent = 'Maximum 3 assignments can be selected.';
  s1Warning.style.display = 'none';

  const s1List = document.createElement('div');
  s1List.className = 'assign-list';

  const termCode = normalizeTermCode(termSelect.value);
  let termCols = getTermAssignCols(termCode);
  if (!termCols.length) termCols = allAssignCols.filter(c => colHasMarks(c));

  if (!termCols.length) {
    s1List.innerHTML = '<p class="assign-empty">No assignment columns found. Make sure your Brightspace grade table has assignment columns with marks.</p>';
  } else {
    termCols.forEach(col => {
      const label = cleanAssignLabel(col);
      const avg   = getColClassAvg(col);
      const sub   = getColSubmitRate(col);
      const safeId = `bka-${col.replace(/[^a-z0-9]/gi, '_')}`;

      const row = document.createElement('div');
      row.className = 'assign-row';
      row.dataset.label = label.toLowerCase();

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.id = safeId;
      cb.dataset.col = col;
      if (overlayAssignCols.includes(col)) cb.checked = true;

      cb.addEventListener('change', () => {
        const checked = s1List.querySelectorAll('input[type="checkbox"]:checked');
        if (checked.length > 3) { cb.checked = false; s1Warning.style.display = ''; return; }
        s1Warning.style.display = 'none';
      });
      row.addEventListener('click', e => {
        if (e.target === cb) return;
        cb.checked = !cb.checked;
        cb.dispatchEvent(new Event('change'));
      });

      const lbl = document.createElement('label');
      lbl.htmlFor = safeId;
      lbl.className = 'assign-label';
      lbl.textContent = label;

      const avgBadge = document.createElement('span');
      avgBadge.className = 'assign-avg';
      avgBadge.style.color = avgColor(avg);
      avgBadge.title = 'Class average';
      avgBadge.textContent = avg != null ? `${Math.round(avg)}%` : '—';

      const subBadge = document.createElement('span');
      subBadge.className = 'assign-sub';
      subBadge.style.color = avgColor(sub);
      subBadge.title = 'Submission rate';
      subBadge.textContent = sub != null ? `${Math.round(sub)}%` : '—';

      row.appendChild(cb);
      row.appendChild(lbl);
      row.appendChild(avgBadge);
      row.appendChild(subBadge);
      s1List.appendChild(row);
    });
  }

  s1Search.addEventListener('input', () => {
    const q = s1Search.value.toLowerCase();
    s1List.querySelectorAll('.assign-row').forEach(r => {
      r.style.display = r.dataset.label.includes(q) ? '' : 'none';
    });
  });

  const s1Footer = document.createElement('div');
  s1Footer.className = 'assign-footer';

  const cancelBtn = document.createElement('button');
  cancelBtn.className = 'btn';
  cancelBtn.textContent = 'Cancel';
  cancelBtn.addEventListener('click', () => overlay.remove());

  const nextBtn = document.createElement('button');
  nextBtn.className = 'btn assign-primary-btn';
  nextBtn.textContent = 'Continue →';

  s1Footer.appendChild(cancelBtn);
  s1Footer.appendChild(nextBtn);

  step1.appendChild(s1Title);
  step1.appendChild(s1Hint);
  step1.appendChild(s1Search);
  step1.appendChild(s1Warning);
  step1.appendChild(s1List);
  step1.appendChild(s1Footer);

  // ── Step 2: Upcoming tests ──
  const step2 = document.createElement('div');
  step2.className = 'assign-step';
  step2.style.display = 'none';

  const s2Title = document.createElement('div');
  s2Title.className = 'assign-title';
  s2Title.innerHTML = '<strong>Step 2 of 2</strong> — Upcoming test <span style="font-weight:400;color:var(--muted)">(optional)</span>';

  const s2Hint = document.createElement('div');
  s2Hint.className = 'assign-hint';
  s2Hint.textContent = 'Optionally pick one upcoming assignment/test to mention encouragingly. Only columns with no marks yet (for this term) appear here.';

  const s2Search = document.createElement('input');
  s2Search.className = 'assign-search';
  s2Search.placeholder = 'Search tests…';
  s2Search.type = 'text';

  const s2List = document.createElement('div');
  s2List.className = 'assign-list';

  // Show all term columns that have NO marks yet — these are upcoming/ungraded work.
  // We don't filter by the word "test" because Spirit of Math columns are named things
  // like "L12-Challenge", "L12-Review", etc. — any ungraded column counts as upcoming.
  const termColsForUpcoming = getTermAssignCols(termCode);
  const upcomingCols = (termColsForUpcoming.length ? termColsForUpcoming : allAssignCols)
    .filter(col => !colHasMarks(col));

  if (!upcomingCols.length) {
    s2List.innerHTML = '<p class="assign-empty">No upcoming (ungraded) columns found for this term. Once a column in Brightspace has no marks yet it will appear here.</p>';
  } else {
    upcomingCols.forEach(col => {
      const label = cleanAssignLabel(col);
      const safeId = `bku-${col.replace(/[^a-z0-9]/gi, '_')}`;

      const row = document.createElement('div');
      row.className = 'assign-row';
      row.dataset.label = label.toLowerCase();

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.id = safeId;
      cb.dataset.col = col;
      if (overlayUpcomingCols.includes(col)) cb.checked = true;

      cb.addEventListener('change', () => {
        if (cb.checked) {
          s2List.querySelectorAll('input[type="checkbox"]').forEach(o => { if (o !== cb) o.checked = false; });
        }
      });
      row.addEventListener('click', e => {
        if (e.target === cb) return;
        cb.checked = !cb.checked;
        cb.dispatchEvent(new Event('change'));
      });

      const lbl = document.createElement('label');
      lbl.htmlFor = safeId;
      lbl.className = 'assign-label';
      lbl.textContent = label;

      row.appendChild(cb);
      row.appendChild(lbl);
      s2List.appendChild(row);
    });
  }

  s2Search.addEventListener('input', () => {
    const q = s2Search.value.toLowerCase();
    s2List.querySelectorAll('.assign-row').forEach(r => {
      r.style.display = r.dataset.label.includes(q) ? '' : 'none';
    });
  });

  const s2Footer = document.createElement('div');
  s2Footer.className = 'assign-footer';

  const backBtn = document.createElement('button');
  backBtn.className = 'btn';
  backBtn.textContent = '← Back';
  backBtn.addEventListener('click', () => {
    step1.style.display = '';
    step2.style.display = 'none';
  });

  const generateBtn = document.createElement('button');
  generateBtn.className = 'btn assign-primary-btn';
  generateBtn.textContent = '▶ Generate All';

  s2Footer.appendChild(backBtn);
  s2Footer.appendChild(generateBtn);

  step2.appendChild(s2Title);
  step2.appendChild(s2Hint);
  step2.appendChild(s2Search);
  step2.appendChild(s2List);
  step2.appendChild(s2Footer);

  // Advance to step 2
  nextBtn.addEventListener('click', () => {
    overlayAssignCols = Array.from(s1List.querySelectorAll('input[type="checkbox"]:checked'))
      .map(cb => cb.dataset.col).filter(Boolean);
    step1.style.display = 'none';
    step2.style.display = '';
  });

  // Confirm and generate
  generateBtn.addEventListener('click', () => {
    overlayUpcomingCols = Array.from(s2List.querySelectorAll('input[type="checkbox"]:checked'))
      .map(cb => cb.dataset.col).filter(Boolean);
    overlay.remove();
    onConfirm();
  });

  modal.appendChild(step1);
  modal.appendChild(step2);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);
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

// ── Card builder ───────────────────────────────────────────────────────────────

function buildCard(student) {
  const under   = isUnderperforming(student);
  const perf    = perfCode(student.gradeNum);
  const comment = generatedComments[student.idx] || '';
  const state   = getCardState(student.idx, student);

  const card = document.createElement('div');
  card.className = `student-card perf-${perf}` + (under ? ' underperforming' : '');
  card.dataset.idx = student.idx;

  // Card header
  const header = document.createElement('div');
  header.className = 'card-header';

  const collapseBtn = document.createElement('button');
  collapseBtn.type = 'button';
  collapseBtn.className = 'btn collapse-btn';
  collapseBtn.title = 'Collapse / expand';
  collapseBtn.textContent = state.collapsed ? '▶' : '▼';
  header.appendChild(collapseBtn);

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
  gradeEl.className = `student-grade grade-${perf}`;
  gradeEl.textContent = formatGrade(student.gradeNum);

  const editBtn = document.createElement('button');
  editBtn.type = 'button';
  editBtn.className = 'btn edit-btn';
  editBtn.textContent = '✏️';
  editBtn.title = 'Advanced options';

  header.appendChild(nameEl);
  header.appendChild(gradeEl);
  header.appendChild(editBtn);
  card.appendChild(header);

  // Collapsible body
  const body = document.createElement('div');
  body.className = 'card-body' + (state.collapsed ? ' collapsed' : '');

  // Pronoun toast — only shown when pronoun is unknown
  const pronounRow = document.createElement('div');
  pronounRow.className = 'pronoun-row';
  pronounRow.style.display = (state.pronoun === 'unknown') ? '' : 'none';

  const toastBtn = document.createElement('button');
  toastBtn.type = 'button';
  toastBtn.className = 'pronoun-toast';
  toastBtn.title = 'Pronoun unclear — tap to set';
  toastBtn.textContent = '🔁 Pronoun unclear — tap to set';

  const chooser = document.createElement('div');
  chooser.className = 'pronoun-chooser';
  chooser.style.display = 'none';

  const heBtn = document.createElement('button');
  heBtn.type = 'button';
  heBtn.className = 'pronoun-choice';
  heBtn.dataset.pronoun = 'he';
  heBtn.textContent = '🚹 He/Him';

  const sheBtn = document.createElement('button');
  sheBtn.type = 'button';
  sheBtn.className = 'pronoun-choice';
  sheBtn.dataset.pronoun = 'she';
  sheBtn.textContent = '🚺 She/Her';

  chooser.appendChild(heBtn);
  chooser.appendChild(sheBtn);
  pronounRow.appendChild(toastBtn);
  pronounRow.appendChild(chooser);
  body.appendChild(pronounRow);

  // Comment textarea
  const textarea = document.createElement('textarea');
  textarea.className = 'comment-textarea';
  textarea.placeholder = 'Click ↺ Generate to create a comment…';
  textarea.value = comment;
  textarea.addEventListener('input', () => {
    generatedComments[student.idx] = textarea.value;
    updateCopyBtn(copyBtn, textarea.value);
  });
  body.appendChild(textarea);

  // Action row
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

  genBtn.addEventListener('click', () =>
    generateOne(student, textarea, genBtn, copyBtn, pronounRow, state)
  );
  copyBtn.addEventListener('click', () => {
    const text = textarea.value.trim();
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      copyBtn.textContent = '✓ Copied!';
      setTimeout(() => { copyBtn.textContent = 'Copy ✓'; }, 1600);
    });
  });

  actions.appendChild(genBtn);
  actions.appendChild(copyBtn);
  body.appendChild(actions);

  // Advanced panel
  const advPanel = buildAdvancedPanel(student, state, textarea, genBtn, copyBtn, pronounRow);
  advPanel.style.display = state.advancedOpen ? '' : 'none';
  body.appendChild(advPanel);

  card.appendChild(body);

  // Wire collapse
  collapseBtn.addEventListener('click', () => {
    state.collapsed = !state.collapsed;
    collapseBtn.textContent = state.collapsed ? '▶' : '▼';
    body.classList.toggle('collapsed', state.collapsed);
  });

  // Wire edit
  editBtn.addEventListener('click', () => {
    state.advancedOpen = !state.advancedOpen;
    advPanel.style.display = state.advancedOpen ? '' : 'none';
    editBtn.classList.toggle('active', state.advancedOpen);
    if (state.collapsed && state.advancedOpen) {
      state.collapsed = false;
      collapseBtn.textContent = '▼';
      body.classList.remove('collapsed');
    }
  });

  // Wire toast + chooser
  toastBtn.addEventListener('click', () => {
    chooser.style.display = chooser.style.display === 'none' ? '' : 'none';
  });

  [heBtn, sheBtn].forEach(btn => {
    btn.addEventListener('click', () => {
      state.pronoun = btn.dataset.pronoun;
      [heBtn, sheBtn].forEach(b => b.classList.toggle('active', b === btn));
      chooser.style.display = 'none';
      pronounRow.style.display = 'none'; // hide toast once set
    });
  });

  return card;
}

// ── Advanced panel ─────────────────────────────────────────────────────────────

function buildAdvancedPanel(student, state, textarea, genBtn, copyBtn, pronounRow) {
  const panel = document.createElement('div');
  panel.className = 'advanced-panel';

  // Pronoun selector
  const pronounSection = document.createElement('div');
  pronounSection.className = 'adv-section';
  const pronounLabel = document.createElement('div');
  pronounLabel.className = 'adv-label';
  pronounLabel.textContent = 'Pronouns';
  const pronounGroup = document.createElement('div');
  pronounGroup.className = 'pronoun-radio-group';

  [
    { val: 'unknown', label: 'Auto-detect' },
    { val: 'he',      label: '🚹 He/Him'   },
    { val: 'she',     label: '🚺 She/Her'  }
  ].forEach(opt => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'pronoun-radio-btn' + (state.pronoun === opt.val ? ' active' : '');
    btn.dataset.val = opt.val;
    btn.textContent = opt.label;
    btn.addEventListener('click', () => {
      state.pronoun = opt.val;
      pronounGroup.querySelectorAll('.pronoun-radio-btn').forEach(b =>
        b.classList.toggle('active', b.dataset.val === opt.val)
      );
      pronounRow.style.display = (state.pronoun === 'unknown') ? '' : 'none';
    });
    pronounGroup.appendChild(btn);
  });

  pronounSection.appendChild(pronounLabel);
  pronounSection.appendChild(pronounGroup);
  panel.appendChild(pronounSection);

  // Performance override
  const perfSection = document.createElement('div');
  perfSection.className = 'adv-section';
  const perfLabel = document.createElement('div');
  perfLabel.className = 'adv-label';
  perfLabel.textContent = 'Performance';
  const perfSelect = document.createElement('select');
  perfSelect.className = 'adv-select';
  [
    { val: '',              label: `Auto (${perfCode(student.gradeNum).replace('_', ' ')})` },
    { val: 'good',          label: 'Good' },
    { val: 'satisfactory',  label: 'Satisfactory' },
    { val: 'average',       label: 'Average' },
    { val: 'needs_support', label: 'Needs Support' }
  ].forEach(opt => {
    const o = document.createElement('option');
    o.value = opt.val;
    o.textContent = opt.label;
    if (opt.val === (state.perfOverride || '')) o.selected = true;
    perfSelect.appendChild(o);
  });
  perfSelect.addEventListener('change', () => { state.perfOverride = perfSelect.value || null; });
  perfSection.appendChild(perfLabel);
  perfSection.appendChild(perfSelect);
  panel.appendChild(perfSection);

  // Custom note
  const noteSection = document.createElement('div');
  noteSection.className = 'adv-section';
  const noteLabel = document.createElement('div');
  noteLabel.className = 'adv-label';
  noteLabel.textContent = 'Custom Note';
  const noteInput = document.createElement('textarea');
  noteInput.className = 'adv-note';
  noteInput.placeholder = 'e.g. "Recently joined from another school." — added to the prompt';
  noteInput.value = state.customNote;
  noteInput.rows = 2;
  noteInput.addEventListener('input', () => { state.customNote = noteInput.value; });
  noteSection.appendChild(noteLabel);
  noteSection.appendChild(noteInput);
  panel.appendChild(noteSection);

  // Comment bank
  const bankSection = document.createElement('div');
  bankSection.className = 'adv-section';
  const bankLabel = document.createElement('div');
  bankLabel.className = 'adv-label';
  bankLabel.textContent = 'Comment Bank';
  bankSection.appendChild(bankLabel);

  COMMENT_BANK_MINI.forEach(cat => {
    const catEl = document.createElement('div');
    catEl.className = 'bank-category';
    catEl.textContent = cat.category;
    bankSection.appendChild(catEl);
    cat.items.forEach(item => {
      const itemRow = document.createElement('label');
      itemRow.className = 'bank-item';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = state.selectedBank.has(item);
      cb.addEventListener('change', () => {
        if (cb.checked) state.selectedBank.add(item);
        else            state.selectedBank.delete(item);
      });
      itemRow.appendChild(cb);
      itemRow.appendChild(document.createTextNode(' ' + item));
      bankSection.appendChild(itemRow);
    });
  });
  panel.appendChild(bankSection);

  // Regenerate
  const regenBtn = document.createElement('button');
  regenBtn.type = 'button';
  regenBtn.className = 'btn gen-btn adv-regen-btn';
  regenBtn.textContent = '↺ Regenerate with options';
  regenBtn.addEventListener('click', () =>
    generateOne(student, textarea, genBtn, copyBtn, pronounRow, state)
  );
  panel.appendChild(regenBtn);

  return panel;
}

// ── Copy button helper ─────────────────────────────────────────────────────────

function updateCopyBtn(btn, text) {
  const has = !!(text && text.trim());
  btn.classList.toggle('has-content', has);
  btn.textContent = has ? 'Copy ✓' : 'Copy';
}

// ── API & comment generation ───────────────────────────────────────────────────

function buildPayload(student, state) {
  const term       = termSelect.value;
  const gradeGroup = gradeGroupSelect.value;
  const structure  = structureSelect.value;

  const resolvedPerf   = (state && state.perfOverride) ? state.perfOverride : perfCode(student.gradeNum);
  const resolvedPronoun = state ? state.pronoun : guessPronounFromName(student.firstName);

  // Assignment facts: prefer overlay selections, fall back to top student assignments
  let assignmentFacts = [];
  if (overlayAssignCols.length) {
    assignmentFacts = overlayAssignCols.map(col => {
      const a = student.assignments.find(x => x.label === col);
      return a ? { label: cleanAssignLabel(a.label), value: a.value } : null;
    }).filter(Boolean).slice(0, 3);
  }
  if (!assignmentFacts.length) {
    assignmentFacts = student.assignments.slice(0, 2).map(a => ({
      label: cleanAssignLabel(a.label), value: a.value
    }));
  }

  const upcomingTests = overlayUpcomingCols.map(col => cleanAssignLabel(col)).filter(Boolean);

  let additionalContext = '';
  if (state && state.selectedBank.size > 0) {
    additionalContext += 'Incorporate these notes: ' + [...state.selectedBank].join('; ') + '.';
  }
  if (state && state.customNote.trim()) {
    additionalContext += (additionalContext ? ' ' : '') + state.customNote.trim();
  }

  const classFirstNames = students.map(s => s.firstName).filter(Boolean).slice(0, 120);

  return {
    mode:                    'basic_bulk',
    reviseMode:              'basic_bulk',
    targetStructure:         structure,
    basicStructure:          structure,
    termLabel:               term,
    studentName:             student.name,
    studentFirstName:        student.firstName,
    pronounGuess:            resolvedPronoun,
    finalMark:               student.gradeRaw || '',
    performanceLevel:        resolvedPerf,
    performanceLabel:        resolvedPerf.replace('_', ' '),
    needsSupport:            resolvedPerf === 'needs_support',
    gradeGroup,
    assignmentFacts,
    upcomingTests,
    allowedAssignmentLabels: student.assignments.map(a => cleanAssignLabel(a.label)),
    classFirstNames,
    additionalContext:       additionalContext || undefined
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

function typewriterAnimate(textarea, text, onDone) {
  textarea.value = '';
  textarea.style.overflow = 'hidden';
  let i = 0;
  const CHAR_DELAY = 12;
  function typeNext() {
    if (i < text.length) {
      textarea.value += text[i];
      textarea.style.height = 'auto';
      textarea.style.height = textarea.scrollHeight + 'px';
      i++;
      setTimeout(typeNext, CHAR_DELAY);
    } else {
      textarea.style.overflow = '';
      if (onDone) onDone();
    }
  }
  typeNext();
}

async function generateOne(student, textarea, genBtn, copyBtn, pronounRow, state) {
  genBtn.disabled = true;
  genBtn.textContent = '…';
  textarea.classList.add('loading');
  textarea.value = '';

  try {
    const res = await fetch(API_ENDPOINT, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(buildPayload(student, state || getCardState(student.idx, student)))
    });

    if (!res.ok) throw new Error(`Server error ${res.status}`);
    const data    = await res.json();
    const comment = parseApiResponse(data);
    if (!comment) throw new Error('Empty response from server');

    generatedComments[student.idx] = comment;
    textarea.classList.remove('loading');

    typewriterAnimate(textarea, comment, () => {
      updateCopyBtn(copyBtn, comment);
      genBtn.disabled    = false;
      genBtn.textContent = '↺ Generate';
    });
    return;

  } catch (err) {
    textarea.value = `⚠ Error: ${err.message}. Click ↺ to retry.`;
  }

  textarea.classList.remove('loading');
  genBtn.disabled    = false;
  genBtn.textContent = '↺ Generate';
}

async function generateAll() {
  if (isGeneratingAll || !students.length) return;

  const term = String(termSelect.value || '').trim();
  if (!term) {
    setStatus('⚠ Select a term before generating all comments.', 'error');
    return;
  }

  showAssignOverlay(async () => {
    isGeneratingAll = true;
    generateAllBtn.disabled    = true;
    generateAllBtn.textContent = '⏳ Generating…';

    const visible = showOnlyUnderperf ? students.filter(isUnderperforming) : students;
    const queue   = [...visible];
    let done = 0;

    setStatus(`Generating comments for ${visible.length} students…`, 'loading');

    async function worker() {
      while (queue.length) {
        const student    = queue.shift();
        const card       = studentList.querySelector(`.student-card[data-idx="${student.idx}"]`);
        if (!card) continue;

        const textarea   = card.querySelector('.comment-textarea');
        const genBtn     = card.querySelector('.gen-btn');
        const copyBtn    = card.querySelector('.copy-btn');
        const pronounRow = card.querySelector('.pronoun-row');
        const state      = getCardState(student.idx, student);

        if (textarea && genBtn && copyBtn) {
          await generateOne(student, textarea, genBtn, copyBtn, pronounRow, state);
        }
        done++;
        setStatus(`Generated ${done} / ${visible.length}…`, 'loading');
      }
    }

    await Promise.all(Array.from({ length: CONCURRENCY }, worker));

    setStatus(`✓ ${done} comments ready!`, 'success');
    generateAllBtn.disabled    = false;
    generateAllBtn.textContent = '▶ Generate All Comments';
    isGeneratingAll = false;
  });
}

// ── Data loading ───────────────────────────────────────────────────────────────

async function loadGrades() {
  setStatus('Loading grades from page…', 'loading');
  refreshBtn.disabled = true;
  emptyState.style.display = 'none';

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab) throw new Error('No active tab found.');

    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files:  ['content.js']
    }).catch(() => {});

    const result = await chrome.tabs.sendMessage(tab.id, { action: 'scrapeGrades' });
    if (!result)         throw new Error('No response from page. Try refreshing Brightspace.');
    if (!result.success) throw new Error(result.error);

    students = processStudents(result.data);
    if (!students.length) throw new Error('No student rows found in the grade table.');

    classLabel.textContent = result.className
      ? result.className.slice(0, 40)
      : `${students.length} students`;

    restoreSettings();
    generatedComments   = {};
    overlayAssignCols   = [];
    overlayUpcomingCols = [];
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
    if (s.term       && termSelect.querySelector(`option[value="${s.term}"]`))                 termSelect.value = s.term;
    if (s.gradeGroup && gradeGroupSelect.querySelector(`option[value="${s.gradeGroup}"]`)) gradeGroupSelect.value = s.gradeGroup;
    if (s.structure  && structureSelect.querySelector(`option[value="${s.structure}"]`))     structureSelect.value = s.structure;
  } catch (_) {}
}

// ── Event wiring ───────────────────────────────────────────────────────────────

refreshBtn.addEventListener('click', loadGrades);
generateAllBtn.addEventListener('click', generateAll);

filterUnderperformingBtn.addEventListener('click', () => {
  showOnlyUnderperf = !showOnlyUnderperf;
  filterUnderperformingBtn.classList.toggle('active', showOnlyUnderperf);
  filterUnderperformingBtn.textContent = showOnlyUnderperf ? '✕ Show All Students' : '⚠️ Show Underperforming';
  renderStudents();
});

[termSelect, gradeGroupSelect, structureSelect].forEach(el => {
  el.addEventListener('change', saveSettings);
});

chrome.runtime.onMessage.addListener(msg => {
  if (msg.action === 'gradeTableDetected') {
    setStatus('Grade table detected — click ↻ to load', 'idle');
  }
});

loadGrades();
