// sidepanel.js — BrightBridge
// Handles all side-panel UI: loading grades, rendering students, generating comments.

// ── Constants ──────────────────────────────────────────────────────────────────

const API_ENDPOINT = 'https://jotaeliezer-teacher-tools-api-fj1k.vercel.app/api/generate-comment';
const CONCURRENCY  = 2;       // max parallel API calls (mirrors Teacher Tools)
const UNDERPERFORM_THRESHOLD = 60; // % at or below = underperforming

const COMMENT_BANK_MINI = [
  {
    category: 'Participation',
    items: [
      'actively participates in class discussions',
      'consistently contributes thoughtful answers'
    ]
  },
  {
    category: 'Homework',
    items: [
      'completes all assigned work on time',
      'demonstrates strong effort on homework'
    ]
  },
  {
    category: 'Seeking Help',
    items: [
      'proactively asks questions when unsure',
      'makes great use of extra help sessions'
    ]
  },
  {
    category: 'Personal Qualities',
    items: [
      'shows excellent perseverance and resilience',
      'demonstrates a positive and cooperative attitude'
    ]
  },
  {
    category: 'Looking Ahead',
    items: [
      'is encouraged to review key concepts regularly',
      'has strong potential to excel next term'
    ]
  }
];

// ── Pronoun guess from first name (mirrors Teacher Tools main.js) ──────────────

function guessPronounFromFirstName(firstName) {
  const name = String(firstName || '').trim().toLowerCase();
  if (!name) return 'they';

  const female = new Set([
    // English
    'emma','olivia','ava','sophia','isabella','mia','charlotte','amelia','evelyn','abigail',
    'emily','elizabeth','ella','avery','sofia','camila','aria','scarlett','victoria','madison',
    'luna','grace','chloe','penelope','layla','riley','zoey','nora','lily','eleanor','hannah',
    'lillian','aubrey','ellie','stella','natalia','zoe','leah','hazel','violet','aurora',
    'savannah','audrey','brooklyn','bella','claire','skylar','lucy','anna','caroline','nova',
    'emilia','kennedy','samantha','maya','willow','naomi','aaliyah','elena','sarah','ariana',
    'allison','gabriella','alice','ruby','eva','serenity','autumn','hailey','gianna',
    'valentina','isla','eliana','quinn','ivy','sadie','piper','lydia','alexa','josephine',
    'julia','delilah','arianna','vivian','kaylee','sophie','madeline','peyton','rylee','clara',
    'hadley','amber','anaya','anita','ana','irene','annie','aisha','sara','sehar','jessica',
    'jennifer','ashley','amanda','megan','rachel','laura','natalie','kayla','brianna','taylor',
    'jasmine','vanessa','caitlin','shannon','melissa','tiffany','brittany','danielle','crystal',
    'holly','katie','kelly','lindsey','miranda','nikki','paige','shelby','stacy','wendy',
    'sandra','patricia','linda','barbara','margaret','carol','donna','diane',
    // Chinese female (pinyin)
    'mei','fang','jing','ling','xiu','hui','yan','ying','xin','yue','zhen','lan','rong',
    'juan','min','fen','xia','na','qing','hua','shan','dan','rou','xue','ting','lian',
    'qian','shu','tian','ping','yanyan','tingting','linlin','xiaoyan','xiaoling','xiaomei',
    'yinyin','jingjing','huihui','shanshan','xinxin','yueyue','fangfang','rongrong',
    // Indian / South-Asian female
    'priya','ananya','divya','pooja','shreya','riya','neha','aditi','swati','deepa',
    'kavya','nisha','sunita','meera','anjali','sita','lakshmi','puja','isha','diya',
    'sonal','tanya','vanya','mansi','sakshi','smita','nita','rita','gita','rashmi',
    'shweta','preeti','seema','reena','rina','reema','rekha','renu','usha','uma',
    'vidya','vinita','varsha','vandana','veena','yamini','yashika','yashna','anika',
    'aastha','anushka','avni','charu','devika','esha','gargi','harsha','hema','indu',
    'jaya','jyoti','kajal','kamla','karuna','komal','kriti','lavanya','madhuri','manisha',
    'megha','minal','mohini','namita','nandita','natasha','nidhi','nilam','paro','parvati',
    'poonam','prachi','pragya','rachna','radhika','radha','rani','riddhi','ritu','ruhi',
    'rupal','rupali','sandhya','sanjana','saraswati','savita','shanti','shilpa','shobha',
    'shruti','simran','sonali','sonam','sudha','supriya','surbhi','sushma','swapna',
    'tanvi','taruna','trisha','trishna','urvashi','vasudha','vibha','vimala','vrinda',
    'anam','fatima','zainab','maryam','nadia','hana','layla','yasmin','sana',
    'bushra','farah','hira','iram','noor','rabia','rima','sahar','samia','sobia','sofia',
  ]);

  const male = new Set([
    // English
    'liam','noah','william','james','oliver','benjamin','elijah','lucas','mason','logan',
    'alexander','ethan','jacob','michael','daniel','henry','jackson','sebastian','aiden','matthew',
    'samuel','david','joseph','carter','owen','wyatt','john','jack','luke','jayden',
    'dylan','grayson','levi','isaac','gabriel','julian','mateo','anthony','jaxon','lincoln',
    'joshua','christopher','andrew','theodore','caleb','ryan','asher','nathaniel','thomas','leo',
    'christian','jonathan','ezra','charles','colton','cameron','eli','hudson','aaron','landon',
    'adam','dominic','austin','evan','parker','tyler','blake','chase','garrett','grant',
    'ian','kyle','nolan','seth','tanner','trevor','troy','victor','wade','zach','zachary',
    'mikhail','zaydan','wesley','kanav','kabir','aadit','hunter','hayden','cole','jordan',
    'kevin','brian','jason','justin','brandon','sean','derek','eric','greg',
    'mark','paul','scott','steven','robert','richard','edward','george',
    // Chinese male (pinyin)
    'wei','jun','hao','ming','jie','tao','bin','rui','long','peng',
    'gang','feng','bo','yi','dong','kang','qiang','chao','jian','kai',
    'liang','xiang','yang','yu','zheng','jiahao','jiaming','junhao','junming',
    'mingzhe','ruihao','tianhao','yuhao','zihao','ziyang','ziyuan',
    // Indian / South-Asian male
    'arjun','rahul','rohan','raj','vikram','arun','suresh','rajesh','kiran','amit',
    'nikhil','vivek','sanjay','gaurav','vishal','akash','ravi','kunal','aditya','saurabh',
    'pranav','varun','ishaan','aarav','advait','dhruv','harsh','krishna','manav','nakul',
    'omkar','parth','prateek','rishabh','rohit','samarth','siddharth','tanmay','uday',
    'vaibhav','vedant','vikas','vinay','vineet','yash','yogesh','ankit','apoorv',
    'aryan','ashish','ayush','bharat','chirag','deepak','devesh','dheeraj',
    'dinesh','dipak','girish','gopal','govind','harish','hitesh','jatin','jayesh',
    'karan','kartik','mahesh','manish','mayank','mohit','naresh','naveen','nilesh','nishant',
    'pankaj','paras','piyush','prakash','prasad','pratik','praveen','pushkar',
    'raghav','rakesh','ramesh','ritesh','sachin','sahil','santosh','satish',
    'shyam','sumit','sunil','suraj','tushar','umesh','vipul',
    // Muslim / Sikh male
    'muhammad','omar','ali','hassan','hussain','ibrahim','ismail','yusuf','tariq','khalid',
    'adnan','bilal','faisal','hamza','imran','jawad','kamran','naveed','sajid','shahid',
    'sufyan','usman','waleed','zain','zubair','gurpreet','harpreet','jaswinder','kulwant',
    'mandeep','navneet','paramjit','rajinder','surjit','amrit','balwinder','davinder',
    'gurmeet','inderjit','jaskaran','jaspal','lakhvir','manjit','narinder','parminder',
    'ranjit','satinder','tarsem','tejinder',
  ]);

  if (female.has(name)) return 'she';
  if (male.has(name))   return 'he';

  // ── Heuristic patterns (same as Teacher Tools) ─────────────────────────────
  if (/i[ck]a$|ita$|ina$|n[iy]a$|l[iy]a$|s[iy]a$|m[iy]a$|[iy][ay]$|shi$|thi$|dhi$|nee$|dee$|jot[i]?$|vani$|wati$|kumari$/.test(name)) return 'she';
  if (/raj$|han$|nav$|esh$|jit$|pal$|bir$|vin$|kar$|jot$|vir$|deep$|nand$|preet$|meet$|inder$|winder$/.test(name)) return 'he';
  if (/(ling|mei|xin|yan|ying|yue|fang|xiu|rong|juan|shan|rou|lian|qian|shu)$/.test(name)) return 'she';
  if (/(jun|hao|wei|jie|long|peng|yang|ming|feng|gang|dong|kang|bin|rui|tao|liang)$/.test(name)) return 'he';
  if (/a$/.test(name) && !/sha$|cha$|ssa$|kha$|pha$/.test(name)) return 'she';

  return 'they'; // genuinely ambiguous — teacher will be prompted
}

// ── State ──────────────────────────────────────────────────────────────────────

let students            = [];   // processed student objects
let generatedComments   = {};   // { studentIdx: string }
let showOnlyUnderperf   = false;
let isGeneratingAll     = false;

// Per-card state: collapse/expand, pronoun override, advanced panel data
const cardStates = {};   // { [idx]: CardState }

function getCardState(idx, firstName) {
  if (!cardStates[idx]) {
    cardStates[idx] = {
      collapsed:    false,
      pronoun:      guessPronounFromFirstName(firstName),  // 'he' | 'she' | 'they'
      customNote:   '',
      selectedBank: new Set(),
      advancedOpen: false,
      perfOverride: null         // null = auto | 'good' | 'satisfactory' | 'average' | 'needs_support'
    };
  }
  return cardStates[idx];
}

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

// Extract the first (given) name from a full name string.
// Handles both "First Last" and Brightspace's "Last, First" comma-separated format.
function extractFirstName(fullName) {
  if (!fullName) return '';
  if (fullName.includes(',')) {
    // "Ghindea, Darius" → "Darius"
    const afterComma = fullName.split(',')[1] || '';
    return afterComma.trim().split(/\s+/)[0];
  }
  // "Darius Ghindea" → "Darius"
  return fullName.split(/\s+/)[0];
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

// Case-insensitive key lookup on a row object — handles minor whitespace/case mismatches
function rowGet(row, key) {
  if (!key) return '';
  if (row[key] !== undefined) return row[key]; // exact match first
  const lk = key.toLowerCase().trim();
  const found = Object.keys(row).find(k => k.toLowerCase().trim() === lk);
  return found ? row[found] : '';
}

function processStudents({ headers, rows }) {
  // ── Debug logging — open Side Panel DevTools (right-click → Inspect) to view ──
  console.log('[BrightBridge] headers:', JSON.stringify(headers));
  console.log('[BrightBridge] row count:', rows.length);
  if (rows.length) console.log('[BrightBridge] first row keys:', JSON.stringify(Object.keys(rows[0])));
  if (rows.length) console.log('[BrightBridge] first row sample:', JSON.stringify(rows[0]));

  const cols = detectCols(headers);
  console.log('[BrightBridge] detected cols — fullCol:', cols.fullCol, '| gradeCol:', cols.gradeCol);

  return rows.map((row, idx) => {
    // Name — clean up Brightspace cell noise (icons, arrows, checkboxes)
    let name = '';

    if (cols.fullCol) {
      const raw = cleanNameText(rowGet(row, cols.fullCol));
      // Normalise "Last, First" → "First Last" for display
      if (raw.includes(',')) {
        const parts = raw.split(',');
        name = (parts[1].trim() + ' ' + parts[0].trim()).trim();
      } else {
        name = raw;
      }
    } else if (cols.firstCol && cols.lastCol) {
      const fn = cleanNameText(rowGet(row, cols.firstCol));
      const ln = cleanNameText(rowGet(row, cols.lastCol));
      name = `${fn} ${ln}`.trim();
    }

    // Broader fallback 1: any header containing name/learner/student
    if (!name || name.length < 2) {
      const fallback = headers.find(h =>
        /name|learner|student/i.test(h) && !/grade|mark|score|percent|enrolled/i.test(h)
      );
      if (fallback) name = cleanNameText(rowGet(row, fallback));
    }

    // Broader fallback 2: first column that contains at least 2 letters and looks like a name
    if (!name || name.length < 2) {
      for (const h of headers) {
        const val = cleanNameText(rowGet(row, h));
        if (
          val.length >= 2 &&
          /[a-zA-Z]{2,}/.test(val) &&             // at least 2 consecutive letters
          !/grade|calculated|percent|enrolled|week|exercise|mastermind/i.test(val)
        ) {
          name = val;
          break;
        }
      }
    }

    if (!name || name.length < 2) return null;

    const firstName = cols.firstCol
      ? cleanNameText(rowGet(row, cols.firstCol))
      : extractFirstName(name);  // handles both "First Last" and "Last, First" (Brightspace Learner col)
    const lastName = cols.lastCol ? cleanNameText(rowGet(row, cols.lastCol)) : '';

    // Grade
    const gradeRaw = cols.gradeCol ? rowGet(row, cols.gradeCol) : '';
    const gradeNum = parseGradeNum(gradeRaw);

    // Assignment columns (everything except name/id/grade/email cols)
    const assignments = headers
      .filter(h => !cols.skipCols.has(h))
      .map(h => ({ label: h, value: (rowGet(row, h) || '').trim() }))
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

// ── Pronoun detection ──────────────────────────────────────────────────────────

function detectTheyThem(text) {
  return /\b(they|them|their|themself|themselves)\b/i.test(text || '');
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
  const comment = generatedComments[student.idx] || '';
  const state   = getCardState(student.idx, student.firstName);

  const card = document.createElement('div');
  card.className = 'student-card' + (under ? ' underperforming' : '');
  card.dataset.idx = student.idx;

  // ── Card header ──
  const header = document.createElement('div');
  header.className = 'card-header';

  // Collapse / expand button
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
  gradeEl.className = 'student-grade' + (under ? ' grade-low' : '');
  gradeEl.textContent = formatGrade(student.gradeNum);

  // Edit button
  const editBtn = document.createElement('button');
  editBtn.type = 'button';
  editBtn.className = 'btn edit-btn';
  editBtn.textContent = '✏️';
  editBtn.title = 'Edit advanced options';

  header.appendChild(nameEl);
  header.appendChild(gradeEl);
  header.appendChild(editBtn);
  card.appendChild(header);

  // ── Collapsible body ──
  const body = document.createElement('div');
  body.className = 'card-body' + (state.collapsed ? ' collapsed' : '');

  // Pronoun toast row — shows when pronoun is ambiguous (they) or already confirmed (he/she)
  const pronounRow = document.createElement('div');
  pronounRow.className = 'pronoun-row';
  // Show toast when pronoun is genuinely unclear (they) so teacher can assign he/she,
  // or when it's been confirmed (he/she) so teacher can see/change the assignment.
  pronounRow.style.display = '';

  const toastBtn = document.createElement('button');
  toastBtn.type = 'button';
  toastBtn.className = 'pronoun-toast';
  toastBtn.title = 'Click to set pronouns';

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

  updatePronounToast(state, toastBtn);

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

  // Advanced panel (initially hidden unless state.advancedOpen)
  const advPanel = buildAdvancedPanel(student, state, textarea, genBtn, copyBtn, pronounRow, toastBtn);
  advPanel.style.display = state.advancedOpen ? '' : 'none';
  body.appendChild(advPanel);

  card.appendChild(body);

  // ── Wire collapse ──
  collapseBtn.addEventListener('click', () => {
    state.collapsed = !state.collapsed;
    collapseBtn.textContent = state.collapsed ? '▶' : '▼';
    body.classList.toggle('collapsed', state.collapsed);
  });

  // ── Wire edit button ──
  editBtn.addEventListener('click', () => {
    state.advancedOpen = !state.advancedOpen;
    advPanel.style.display = state.advancedOpen ? '' : 'none';
    editBtn.classList.toggle('active', state.advancedOpen);
    if (state.collapsed && state.advancedOpen) {
      // Auto-expand card if collapsed when opening edit panel
      state.collapsed = false;
      collapseBtn.textContent = '▼';
      body.classList.remove('collapsed');
    }
  });

  // ── Wire pronoun toast toggle ──
  toastBtn.addEventListener('click', () => {
    chooser.style.display = chooser.style.display === 'none' ? '' : 'none';
  });

  // ── Wire pronoun choice buttons ──
  [heBtn, sheBtn].forEach(btn => {
    btn.addEventListener('click', () => {
      state.pronoun = btn.dataset.pronoun;
      updatePronounToast(state, toastBtn);
      [heBtn, sheBtn].forEach(b => b.classList.toggle('active', b === btn));
      chooser.style.display = 'none';
    });
  });

  // Set active pronoun button on initial render
  [heBtn, sheBtn].forEach(b =>
    b.classList.toggle('active', b.dataset.pronoun === state.pronoun)
  );

  return card;
}

// ── Advanced panel ─────────────────────────────────────────────────────────────

function buildAdvancedPanel(student, state, textarea, genBtn, copyBtn, pronounRow, toastBtn) {
  const panel = document.createElement('div');
  panel.className = 'advanced-panel';

  // ── Pronoun selector ──
  const pronounSection = document.createElement('div');
  pronounSection.className = 'adv-section';

  const pronounLabel = document.createElement('div');
  pronounLabel.className = 'adv-label';
  pronounLabel.textContent = 'Pronouns';

  const pronounGroup = document.createElement('div');
  pronounGroup.className = 'pronoun-radio-group';

  const pronounOpts = [
    { val: 'he',   label: '🚹 He/Him' },
    { val: 'she',  label: '🚺 She/Her' }
  ];

  pronounOpts.forEach(opt => {
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
      // Sync the toast in the card header
      if (toastBtn) updatePronounToast(state, toastBtn);
    });
    pronounGroup.appendChild(btn);
  });

  pronounSection.appendChild(pronounLabel);
  pronounSection.appendChild(pronounGroup);
  panel.appendChild(pronounSection);

  // ── Performance override ──
  const perfSection = document.createElement('div');
  perfSection.className = 'adv-section';

  const perfLabel = document.createElement('div');
  perfLabel.className = 'adv-label';
  perfLabel.textContent = 'Performance';

  const perfSelect = document.createElement('select');
  perfSelect.className = 'adv-select';
  [
    { val: '',             label: `Auto (${perfCode(student.gradeNum).replace('_', ' ')})` },
    { val: 'good',         label: 'Good' },
    { val: 'satisfactory', label: 'Satisfactory' },
    { val: 'average',      label: 'Average' },
    { val: 'needs_support',label: 'Needs Support' }
  ].forEach(opt => {
    const o = document.createElement('option');
    o.value = opt.val;
    o.textContent = opt.label;
    if (opt.val === (state.perfOverride || '')) o.selected = true;
    perfSelect.appendChild(o);
  });
  perfSelect.addEventListener('change', () => {
    state.perfOverride = perfSelect.value || null;
  });

  perfSection.appendChild(perfLabel);
  perfSection.appendChild(perfSelect);
  panel.appendChild(perfSection);

  // ── Custom note ──
  const noteSection = document.createElement('div');
  noteSection.className = 'adv-section';

  const noteLabel = document.createElement('div');
  noteLabel.className = 'adv-label';
  noteLabel.textContent = 'Custom Note';

  const noteInput = document.createElement('textarea');
  noteInput.className = 'adv-note';
  noteInput.placeholder = 'e.g. "She recently moved schools." — added to the prompt';
  noteInput.value = state.customNote;
  noteInput.rows = 2;
  noteInput.addEventListener('input', () => { state.customNote = noteInput.value; });

  noteSection.appendChild(noteLabel);
  noteSection.appendChild(noteInput);
  panel.appendChild(noteSection);

  // ── Comment bank ──
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

  // ── Regenerate button ──
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

// ── Pronoun toast helper ───────────────────────────────────────────────────────

function updatePronounToast(state, toastBtn) {
  const labels = {
    they: '🔁 Pronoun unclear — tap to set',
    he:   '🚹 He/Him',
    she:  '🚺 She/Her'
  };
  toastBtn.textContent = labels[state.pronoun] || labels.they;
  toastBtn.dataset.pronoun = state.pronoun;
}

// ── Copy button helper ─────────────────────────────────────────────────────────

function updateCopyBtn(btn, text) {
  const has = !!(text && text.trim());
  btn.classList.toggle('has-content', has);
  btn.textContent = has ? 'Copy ✓' : 'Copy';
}

// ── API & comment generation ───────────────────────────────────────────────────

function buildPayload(student, state) {
  const term        = termSelect.value;
  const gradeGroup  = gradeGroupSelect.value;
  const structure   = structureSelect.value;

  const resolvedPerf = state && state.perfOverride
    ? state.perfOverride
    : perfCode(student.gradeNum);

  const resolvedPronoun = state ? state.pronoun : 'unknown';

  // Up to 4 assignment facts for the prompt
  const assignmentFacts = student.assignments.slice(0, 4).map(a => ({
    label: a.label,
    value: a.value
  }));

  // Append selected comment bank items + custom note to additional context
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
    upcomingTests:           [],
    allowedAssignmentLabels: student.assignments.map(a => a.label),
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

// Typewriter animation — types the comment character by character into the textarea
function typewriterAnimate(textarea, text, onDone) {
  textarea.value = '';
  textarea.style.overflow = 'hidden';
  let i = 0;
  const CHAR_DELAY = 12; // ms per character

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
      body:    JSON.stringify(buildPayload(student, state || getCardState(student.idx)))
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
  isGeneratingAll = true;
  generateAllBtn.disabled  = true;
  generateAllBtn.textContent = '⏳ Generating…';

  const visible = showOnlyUnderperf ? students.filter(isUnderperforming) : students;
  const queue   = [...visible];
  let done = 0;

  setStatus(`Generating comments for ${visible.length} students…`, 'loading');

  async function worker() {
    while (queue.length) {
      const student = queue.shift();
      const card    = studentList.querySelector(`.student-card[data-idx="${student.idx}"]`);
      if (!card) continue;

      const textarea  = card.querySelector('.comment-textarea');
      const genBtn    = card.querySelector('.gen-btn');
      const copyBtn   = card.querySelector('.copy-btn');
      const pronounRow = card.querySelector('.pronoun-row');
      const state     = getCardState(student.idx);

      if (textarea && genBtn && copyBtn) {
        await generateOne(student, textarea, genBtn, copyBtn, pronounRow, state);
      }
      done++;
      setStatus(`Generated ${done} / ${visible.length}…`, 'loading');
    }
  }

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
