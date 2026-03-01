  // ==== State ====
  const STUDENT_NAME_LABEL = "Student Name";
  const END_OF_LINE_COL = "End-of-line indicator";
  const END_OF_LINE_VALUE = "#";
  const FINAL_GRADE_COLUMN = "Calculated Final Grade";
  const MIN_ZOOM = 0.8;
  const MAX_ZOOM = 1.4;
  const ZOOM_STEP = 0.1;
  const COMMENT_DEFAULTS = {
    gradeColumn: "",
    highThreshold: 90,
    midThreshold: 75,
    satisfactoryThreshold: 70,
    highTemplate: "{name} consistently exceeds expectations and shows strong reasoning. With {mark}, [he/she] explains solutions clearly, checks work carefully, and extends tasks with enrichment. Next term we will push multi-step investigations and advanced challenges.",
    midTemplate: "{name} is making steady progress and meets most expectations. With {mark}, [he/she] benefits from showing full reasoning and checking units.",
    lowTemplate: "{name} is building core skills and currently holds {mark}. With guided practice and clear routines, [he/she] can improve accuracy and confidence. Next term we will target fundamentals and consistent homework habits.",
    extraColumns: []
  };
  const GRADE_GROUP_TEMPLATES = {
    elem: {
      high: "{name} is doing very well and shows growing confidence. [He/She] explains thinking clearly and works carefully. Next term we will keep building problem-solving skills.",
      mid: "{name} is making steady progress. [He/She] benefits from clear steps, checklists, and regular practice. Next term we will build fluency and confidence.",
      low: "{name} is building core skills. With guided practice and routines, [he/she] can improve accuracy and confidence."
    },
    middle: {
      high: "{name} is performing strongly and shows solid reasoning. With {mark}, [he/she] works carefully and applies strategies independently.",
      mid: "{name} is progressing well and meets most expectations. With {mark}, [he/she] benefits from showing full reasoning and checking work.",
      low: "{name} is building foundational skills and currently holds {mark}. With structured practice and feedback, [he/she] can improve consistency. Next term we will reinforce core skills."
    },
    high: {
      high: "{name} is performing at a high level and communicates solutions clearly. With {mark}, [he/she] demonstrates strong mastery and disciplined study habits. Next term we will push higher-order applications and exam readiness.",
      mid: "{name} is meeting expectations and shows steady growth. With {mark}, [he/she] benefits from concise reasoning and careful checking. Next term we will emphasize multi-step reasoning and assessment readiness.",
      low: "{name} is developing essential skills and currently holds {mark}. With targeted practice and support, [he/she] can improve accuracy and confidence. Next term we will focus on core concepts and consistency."
    }
  };
  const TEMPLATE_VARIANTS = {
    high: [
      [
        "{name} approaches every lesson with curiosity and initiative, frequently extending ideas beyond the core expectation.",
        "Their current mark of {mark} showcases precise reasoning, fluent computation, and articulate communication.",
        "{name} routinely finishes core tasks early and then designs extension questions to deepen understanding.",
        "During collaborative tasks they summarize peer strategies, highlight mathematical connections, and keep teams focused.",
        "Feedback is implemented immediately; {name} keeps a reflection log to track goals and next steps.",
        "Next term will emphasize multi-step investigations and contest-style problems, and {name} is ready to mentor classmates through those challenges.",
        "Keep nurturing this blend of creativity and disciplined practice."
      ],
      [
        "{name} demonstrates sophisticated problem solving and eagerly applies concepts to authentic contexts.",
        "Earning {mark} reflects a balance of accuracy, fluency, and elegant explanation.",
        "They lead group discussions by asking probing questions and pushing peers to justify reasoning with evidence.",
        "{name} voluntarily revisits enrichment problems to test alternate strategies and records insights in a math journal.",
        "When challenges arise, they persevere, consult resources, and articulate what was learned from the attempt.",
        "Upcoming investigations will involve open-ended projects where {name} can showcase creativity and leadership.",
        "This exemplary work ethic continues to inspire the class community."
      ],
      [
        "{name} consistently applies higher-order thinking to connect new ideas with past learning.",
        "The {mark} standing highlights precision in calculations and clarity in written reasoning.",
        "They offer insightful reflections after each assessment, naming strategies that led to success.",
        "{name} regularly volunteers to model solutions at the board, guiding peers through alternate pathways.",
        "They thrive when challenged with open-ended tasks and seek feedback to polish each presentation.",
        "Our next unit’s inquiry projects will provide another platform for {name} to innovate and lead.",
        "This combination of humility and excellence makes {name} a role model in the classroom."
      ],
      [
        "{name} immerses themselves in complex scenarios, often extending assignments into independent research.",
        "Their {mark} average represents both speed and accuracy while maintaining elegant work habits.",
        "{name} articulates mathematical thinking using precise vocabulary and visual models.",
        "During collaborative work they facilitate equitable participation and encourage classmates to justify claims.",
        "They synthesize teacher feedback into concrete goals and track progress through self-designed rubrics.",
        "The upcoming contests and investigations will allow {name} to stretch creativity even further.",
        "Please continue providing enrichment opportunities that match this remarkable drive."
      ],
      [
        "{name} excels at transferring skills across novel contexts and frequently mentors peers.",
        "Holding {mark} demonstrates mastery of both procedural fluency and conceptual depth.",
        "They avidly explore optional resources, from math circles to challenge problems, and bring insights back to class.",
        "{name} communicates respectfully, posing questions that push the entire group to think critically.",
        "Each reflection they submit includes personal goals, evidence of growth, and next steps.",
        "Next term we will co-design capstone tasks so {name} can pursue individualized investigations.",
        "Their joyful perseverance continues to elevate our learning community."
      ]
    ],
    mid: [
      [
        "{name} demonstrates steady progress and thoughtful engagement with new material.",
        "With a current mark of {mark} they meet expectations on most outcomes and are beginning to take more academic risks.",
        "{name} listens attentively, asks clarifying questions, and revisits notes to strengthen recall.",
        "When feedback is provided they revise solutions, highlight learning points, and try similar problems independently.",
        "Continued focus on showing complete reasoning and checking units will unlock even higher accuracy.",
        "Next term we will target multi-step problems and timed practice to build automaticity.",
        "I appreciate the consistent effort and positive attitude {name} brings to every session."
      ],
      [
        "{name} contributes to discussions and is learning to explain ideas with greater precision.",
        "The mark of {mark} shows that most foundational outcomes are secure while a few concepts still need polishing.",
        "{name} benefits from pausing to plan before solving and from circling final answers for clarity.",
        "Independent study habits are improving; they now reference worked examples and annotate steps in the margin.",
        "Regular review of math vocabulary and mental math warmups will support faster recall.",
        "Upcoming tasks will emphasize reasoning through word problems, and {name} is ready for that stretch.",
        "Keep encouraging them to ask why and to celebrate the steady growth they are achieving."
      ],
      [
        "{name} arrives prepared and engages respectfully in class conversations, often paraphrasing instructions for peers.",
        "Their {mark} result reflects reliable achievement with occasional slips when rushing.",
        "{name} benefits from colour-coding steps and checking each computation with a different method.",
        "They respond well to feedback, adjusting strategies and noting reminders in the margin.",
        "Daily warmups plus end-of-lesson summaries will sharpen accuracy and long-term retention.",
        "In the coming unit we will emphasize logical proofs, and {name} is poised to participate fully.",
        "Celebrating each milestone continues to boost their motivation."
      ],
      [
        "{name} is strengthening stamina with extended tasks and currently maintains {mark}.",
        "They clarify misunderstandings quickly by referencing anchor charts or mini-videos.",
        "{name} benefits from organizing work in a table before solving each part of a problem.",
        "They now use peer feedback more thoughtfully, acknowledging next steps and trying again.",
        "Consistent practice with estimation and mental math will raise confidence on quizzes.",
        "Next term we will incorporate more hands-on explorations, which suit {name}'s learning style.",
        "Their cooperative spirit and steady effort are appreciated."
      ],
      [
        "{name} is beginning to generalize patterns and explain reasoning aloud.",
        "A {mark} average indicates that essential concepts are solid while advanced applications still need refinement.",
        "{name} thrives when given structured graphic organizers to plan multi-step solutions.",
        "They have started keeping a checklist of common errors, which has reduced avoidable mistakes.",
        "Targeted practice on fraction/decimal conversions will further support success.",
        "Upcoming projects will allow {name} to demonstrate learning through visuals and presentations.",
        "With continued encouragement, their confidence will keep rising."
      ]
    ],
    low: [
      [
        "{name} is developing foundational skills and currently holds {mark}.",
        "They contribute ideas willingly yet benefit from additional guided practice to solidify routines.",
        "{name} responds well to checklists that break complex tasks into smaller checkpoints.",
        "During conferences we co-create goals, model sample solutions, and celebrate each successful attempt.",
        "Home review of vocabulary and worked examples will help transfer strategies across contexts.",
        "We will continue providing targeted small-group lessons, manipulatives, and sentence frames for explanations.",
        "With this supportive plan {name} will keep growing in confidence and accuracy."
      ],
      [
        "{name} is building stamina with multi-step questions and presently has {mark}.",
        "They are most successful when given sentence starters and graphic organizers to capture thinking.",
        "Short, daily practice sessions help {name} remember procedures before tackling longer assignments.",
        "In class we model strategies, rehearse them aloud, and then let {name} teach the process back to a peer.",
        "Families can reinforce progress by celebrating each checkpoint mastered rather than focusing solely on the final score.",
        "Next term we will provide more manipulatives and visuals so abstract ideas feel concrete.",
        "With patience and consistent routines, {name}'s confidence and accuracy will continue to climb."
      ],
      [
        "{name} is learning to manage productive struggle and currently sits at {mark}.",
        "They participate enthusiastically but sometimes rush through steps, so we slow down and highlight key words.",
        "{name} benefits from using math sentence frames to describe operations before computing.",
        "We often use think-alouds so they can mimic the language of reasoning.",
        "Daily review of flashcards and quick drills at home will help facts stick.",
        "Small wins are carefully celebrated so momentum feels tangible.",
        "Together we will keep building routines that lead to confident, independent problem solving."
      ],
      [
        "{name} continues to develop number sense, and their current {mark} reflects gradual improvements.",
        "They feel most confident when manipulatives or visuals are introduced alongside symbolic notation.",
        "{name} benefits from frequent check-ins during independent practice to maintain focus.",
        "We use personalized goal sheets that outline the exact skills being targeted each week.",
        "Families can support by asking {name} to explain the 'why' behind each step rather than just the answer.",
        "Next unit we will embed more game-based reviews to keep engagement high.",
        "With ongoing collaboration, {name}'s perseverance and accuracy will steadily grow."
      ],
      [
        "{name} is developing perseverance with challenging questions and currently holds {mark}.",
        "They appreciate when instructions are chunked and when model solutions are close at hand.",
        "{name} benefits from rehearsing vocabulary aloud before tackling written tasks.",
        "We encourage them to highlight clues in word problems to maintain direction.",
        "Extra practice on math facts and estimation at home will strengthen automaticity.",
        "More scaffolded group work next term will allow {name} to observe peer strategies in real time.",
        "Every incremental success is acknowledged so motivation remains high."
      ]
    ]
  };

  function getDefaultCommentConfig(){
    return {
      ...COMMENT_DEFAULTS,
      extraColumns: [...(COMMENT_DEFAULTS.extraColumns || [])]
    };
  }

  let rows = [];           // [{col:val,...}]
  let originalRows = [];   // deep clone for Reset
  let allColumns = [];     // all column names
  let visibleColumns = []; // chosen columns
  let columnOrder = [];    // current order for all columns
  let filteredIdx = [];    // current row indices after filter
  let sortState = {key:null, dir:null}; // dir: 'asc'|'desc'|null
  let currentFileName = null;
  let isTransposed = false;
  let firstNameKey = null;
  let lastNameKey = null;
  let studentNameColumn = null;
let wrapHeadersEnabled = false;
let visibleRowSet = null;
let studentNameWarning = '';
let darkModeEnabled = false;
let commentConfig = getDefaultCommentConfig();
let generatedComments = [];
let zoomLevel = 1;
const templateIndices = { high:0, mid:0, low:0 };
let columnFilterText = '';
let dataColumnsOnly = false;
const MARK_COLUMN_HINTS = ['grade','mark','score','percent','percentage','assessment','average','result'];
let markColorsEnabled = true;
let pendingMarkWarningAction = null;
let performanceFilterLevel = null;
let printMeta = { teacher:'', classInfo:'', title:'', term:'T1' };
let reportStudentIndex = null;
let reportCardCommentsMap = new Map();
let reportCardPronoun = 'male';
let builderSelectedRowIndex = null;
let builderCommentCheckboxes = [];
let builderBankRendered = false;
let commentTermFilter = 'all';
let commentOrderMode = 'sandwich'; // 'sandwich' | 'selection' | 'bullet'
let savedReports = [];
const SAVED_REPORT_LIMIT = 200;
const PRINT_TEMPLATE_DEFS = {
  attendance: { id:'attendance', label:'Attendance Sheet' },
  marking: { id:'marking', label:'Marking Sheet' },
  drill: { id:'drill', label:'Drill Sheet' },
  reportCard: { id:'reportCard', label:'Report Card' }
};
const PRINT_EXTRA_BLANK_ROWS = 4;
let selectedTemplateId = 'attendance';
let selectedClassIds = new Set();
let printAllClasses = false;
let printReportAllStudents = false;
let builderAiEndpoint = '';
const textMeasureCanvas = document.createElement('canvas');
const textMeasureCtx = textMeasureCanvas.getContext('2d');
function getTemplateFontVariants(){
  const bodyStyle = window.getComputedStyle(document.body || document.documentElement);
  const baseSize = parseFloat(bodyStyle.fontSize) || 16;
  const tableFontSize = `${(baseSize * 0.85).toFixed(2)}px`;
  const fontFamily = bodyStyle.fontFamily || 'system-ui, -apple-system, Segoe UI, sans-serif';
  return {
    regular: `400 ${tableFontSize} ${fontFamily}`,
    bold: `600 ${tableFontSize} ${fontFamily}`
  };
}
function measureTemplateTextWidth(text, weight = 'regular'){
  if (!textMeasureCtx) return (text || '').length * 9;
  const fonts = getTemplateFontVariants();
  textMeasureCtx.font = weight === 'bold' ? fonts.bold : fonts.regular;
  const metrics = textMeasureCtx.measureText(text || '');
  return metrics?.width || 0;
}
function computePrintColumnWidths(students){
  const list = Array.isArray(students) ? students : [];
  const sampleName = list.reduce((longest, entry) => {
    const name = (entry?.name || '').trim();
    return name.length > longest.length ? name : longest;
  }, 'Student Name');
  const totalRows = Math.max(1, list.length);
  const nameWidth = Math.ceil(measureTemplateTextWidth(sampleName || 'Student Name', 'bold')) + 32;
  const indexWidth = Math.ceil(measureTemplateTextWidth(String(totalRows), 'regular')) + 16;
  return {
    name: Math.min(520, Math.max(120, nameWidth)),
    index: Math.min(80, Math.max(32, indexWidth))
  };
}

  // Undo/Redo stacks
  let undoStack = [];
  let redoStack = [];
  const UNDO_LIMIT = 50;

  let draggingCol = null;

  // Persisted UI prefs
  const LS_KEY = "teacher_tools_settings";

  // Multi-file state
  const fileContexts = [];
  let activeContext = null;
  let nextFileId = 1;

  // ==== Elements ====
  const dropzone = document.getElementById('dropzone');
  const appRoot = document.getElementById('appRoot');
  const fileInput = document.getElementById('file');
  const loadDirBtn = document.getElementById('loadDirBtn');
  const lastFolderLabel = document.getElementById('lastFolderLabel');
  const folderPrompt = document.getElementById('folderPrompt');
  const folderPromptText = document.getElementById('folderPromptText');
  const folderPromptCancel = document.getElementById('folderPromptCancel');
  const folderPromptConfirm = document.getElementById('folderPromptConfirm');
  const colsDiv   = document.getElementById('cols');
  const table     = document.getElementById('grid');
  const tableWrapper = document.querySelector('.table-wrapper');
  const thead     = table.querySelector('thead');
  const tbody     = table.querySelector('tbody');
  const statusEl  = document.getElementById('status');
  const metaEl    = document.getElementById('meta');
  const countsEl  = document.getElementById('counts');
  const colCountEl= document.getElementById('colCount');
  const rowCountEl= document.getElementById('rowCount');
  const selectAllColsBtn = document.getElementById('selectAll');
  const clearAllColsBtn = document.getElementById('clearAll');
  const termBtn1 = document.getElementById('termBtn1');
  const termBtn2 = document.getElementById('termBtn2');
  const termBtn3 = document.getElementById('termBtn3');
  const exportCsv = document.getElementById('exportCsvBtn');
  const exportXlsx= document.getElementById('exportXlsxBtn');
  const resetBtn  = document.getElementById('resetBtn');
  const searchEl  = document.getElementById('search');
  const editHeadersBtn = document.getElementById('editHeadersBtn');
  const transposeBtn = document.getElementById('transposeBtn');
  const transposeBanner = document.getElementById('transposeBanner');
  const filesList = document.getElementById('filesList');
  const fileCountEl = document.getElementById('fileCount');
  const rowsDiv   = document.getElementById('rows');
  const selectAllRowsBtn = document.getElementById('selectAllRows');
  const clearRowsBtn = document.getElementById('clearRows');
  const rowFilterGoodBtn = document.getElementById('rowFilterGood');
  const rowFilterMidBtn = document.getElementById('rowFilterMid');
  const rowFilterNeedsBtn = document.getElementById('rowFilterNeeds');
  const rowFilterClearBtn = document.getElementById('rowFilterClear');
  const rowFilterStatusEl = document.getElementById('rowFilterStatus');
  const wrapHeadersToggle = document.getElementById('wrapHeadersToggle');
  const markColorsToggle = document.getElementById('markColorsToggle');
  const filesWarningEl = document.getElementById('filesWarning');
  const filesWarningTextEl = document.getElementById('filesWarningText');
  const colSearchInput = document.getElementById('colSearch');
  const columnDataFilterBtn = document.getElementById('columnDataFilterBtn');
  const darkModeBtn = document.getElementById('darkModeBtn');
  const zoomInBtn = document.getElementById('zoomInBtn');
  const zoomOutBtn = document.getElementById('zoomOutBtn');
  const filesDrawer = document.getElementById('filesDrawer');
  const filesDrawerToggle = document.getElementById('filesDrawerToggle');
  const filesDrawerMiniList = document.getElementById('filesDrawerMiniList');
  const tabDataBtn = document.getElementById('tabDataBtn');
  const tabPrintBtn = document.getElementById('tabPrintBtn');
  const tabPhoneLogsBtn = document.getElementById('tabPhoneLogsBtn');
  const tabCommentsBtn = document.getElementById('tabCommentsBtn');
  const dataTabSection = document.getElementById('dataTabSection');
  const commentsTabSection = document.getElementById('commentsModal');
  const printTabSection = document.getElementById('printPreviewModal');
  const phoneLogsSection = document.getElementById('phoneLogsSection');
  const phoneLogsList = document.getElementById('phoneLogsList');
  const printPhoneLogsBtn = document.getElementById('printPhoneLogsBtn');
  const phoneLogsPrintContainer = document.getElementById('phoneLogsPrintContainer');
  const phoneLogsSearch = document.getElementById('phoneLogsSearch');
  const underperformingList = document.getElementById('underperformingList');
  const underperformingCount = document.getElementById('underperformingCount');
  const commentsModal = document.getElementById('commentsModal');
  const commentsCloseBtn = document.getElementById('commentsCloseBtn');
  const commentsCopyAllBtn = document.getElementById('commentsCopyAll');
  const commentGradeSelect = document.getElementById('commentGradeSelect');
  const commentHighThresholdInput = document.getElementById('commentHighThreshold');
  const commentMidThresholdInput = document.getElementById('commentMidThreshold');
  const commentHighTemplateInput = document.getElementById('commentHighTemplate');
  const commentMidTemplateInput = document.getElementById('commentMidTemplate');
  const commentLowTemplateInput = document.getElementById('commentLowTemplate');
  const commentHighRestructureBtn = document.getElementById('commentHighRestructure');
  const commentMidRestructureBtn = document.getElementById('commentMidRestructure');
  const commentLowRestructureBtn = document.getElementById('commentLowRestructure');
  const commentExtrasList = document.getElementById('commentExtrasList');
  const commentsPreviewEl = document.getElementById('commentsPreview');
  const commentsStatusEl = document.getElementById('commentsStatus');
  const commentsCountEl = document.getElementById('commentsCount');
  const commentTermAllBtn = document.getElementById('commentTermAllBtn');
  const commentTermT1Btn = document.getElementById('commentTermT1Btn');
  const commentTermT2Btn = document.getElementById('commentTermT2Btn');
  const commentTermT3Btn = document.getElementById('commentTermT3Btn');
  const builderStudentSelect = document.getElementById('builderStudentSelect');
  const builderStudentNameInput = document.getElementById('builderStudentName');
  const builderGradeGroupSelect = document.getElementById('builderGradeGroup');
  const builderIncludeFinalGradeInput = document.getElementById('builderIncludeFinalGrade');
  const builderPronounMaleInput = document.getElementById('builderPronounMale');
  const builderPronounFemaleInput = document.getElementById('builderPronounFemale');
  const builderPrefillBtn = document.getElementById('builderPrefillBtn');
  const builderCorePerformanceSelect = document.getElementById('builderCorePerformance');
  const builderTermSelector = document.getElementById('builderTermSelector');
  const builderCCPerformanceSelect = document.getElementById('builderCCPerformance');
  const builderTrigTest1Input = document.getElementById('builderTrigTest1');
  const builderTrigTest2Input = document.getElementById('builderTrigTest2');
  const builderRetestInput = document.getElementById('builderRetestScore');
  const builderOriginalInput = document.getElementById('builderOriginalScore');
  const builderTermAverageInput = document.getElementById('builderTermAverage');
  const builderCustomCommentInput = document.getElementById('builderCustomComment');
  const builderCommentOrderToggle = document.getElementById('builderCommentOrderToggle');
  const builderCommentBankEl = document.getElementById('builderCommentBank');
  const builderSelectedCommentsEl = document.getElementById('builderSelectedComments');
  const builderSelectedTagsEl = document.getElementById('builderSelectedTags');
  const builderSaveBtn = document.getElementById('builderSaveBtn');
  const savedReportsOpenBtn = document.getElementById('savedReportsOpenBtn');
  const savedReportsCloseBtn = document.getElementById('savedReportsCloseBtn');
  const savedReportsBackdrop = document.getElementById('savedReportsBackdrop');
  const savedReportsPanel = document.getElementById('savedReportsPanel');
  const builderGenerateBtn = document.getElementById('builderGenerateBtn');
  const builderGenerateAiBtn = document.getElementById('builderGenerateAiBtn');
  const builderCreateAiBtn = document.getElementById('builderCreateAiBtn');
  const builderAiCreatedOutput = document.getElementById('builderAiCreatedOutput');
  const createConfirmModal = document.getElementById('createConfirmModal');
  const createConfirmYes = document.getElementById('createConfirmYes');
  const createConfirmNo = document.getElementById('createConfirmNo');
  const builderCopyBtn = document.getElementById('builderCopyBtn');
  const builderClearBtn = document.getElementById('builderClearBtn');
  const builderOutputWrap = document.getElementById('builderOutputWrap');
  const builderReportOutput = document.getElementById('builderReportOutput');
  const builderReportOverlay = document.getElementById('builderReportOverlay');
  const builderRevisedOutput = document.getElementById('builderRevisedOutput');
  const savedReportsListEl = document.getElementById('savedReportsList');
  const savedReportsClearBtn = document.getElementById('savedReportsClearBtn');
  const printContainer = document.getElementById('printContainer');
  const printTeacherInput = document.getElementById('printTeacherInput');
  const printClassInput = document.getElementById('printClassInput');
  const printTitleInput = document.getElementById('printTitleInput');
  const printPreviewModal = document.getElementById('printPreviewModal');
  const printPreviewContent = document.getElementById('printPreviewContent');
  const printPreviewCloseBtn = document.getElementById('printPreviewClose');
  const printPreviewCancelBtn = document.getElementById('printPreviewCancel');
  const printPreviewPrintTopBtn = document.getElementById('printPreviewPrintTop');
  const printPreviewPrintBtn = document.getElementById('printPreviewPrint');
const printMarkingPanel = document.getElementById('printMarkingPanel');
const printMarkingColumnsList = document.getElementById('printMarkingColumnsList');
const printMarkingSelectAllBtn = document.getElementById('printMarkingSelectAll');
const printMarkingClearBtn = document.getElementById('printMarkingClear');
const printClassList = document.getElementById('printClassList');
const printAllClassesToggle = document.getElementById('printAllClassesToggle');
const printDrillCountInput = document.getElementById('printDrillCount');
const printReportStudentSelect = document.getElementById('printReportStudentSelect');
const printReportAllStudentsToggle = document.getElementById('printReportAllStudents');
const printCommentsBtn = document.getElementById('printCommentsBtn');
const printCommentsDrawer = document.getElementById('printCommentsDrawer');
const printCommentsClose = document.getElementById('printCommentsClose');
const printCommentsList = document.getElementById('printCommentsList');
const printCommentsPronounMale = document.getElementById('printCommentsPronounMale');
const printCommentsPronounFemale = document.getElementById('printCommentsPronounFemale');
const printTemplateInputs = Array.from(document.querySelectorAll('input[name="printTemplate"]'));
const printTermInputs = Array.from(document.querySelectorAll('input[name="printTerm"]'));
  const printMarkingSelections = {};
  const renamePrompt = document.getElementById('renamePrompt');
  const renameInput = document.getElementById('renameInput');
  const renameCancel = document.getElementById('renameCancel');
  const renameSave = document.getElementById('renameSave');
  let renameTargetId = null;
  const markWarningModal = document.getElementById('markWarningModal');
  const markWarningContinue = document.getElementById('markWarningContinue');
  const markWarningCancel = document.getElementById('markWarningCancel');
  const markWarningMessage = document.getElementById('markWarningMessage');
  const teacherNamePrompt = document.getElementById('teacherNamePrompt');
  const teacherNamePromptInput = document.getElementById('teacherNamePromptInput');
  const teacherNamePromptCancel = document.getElementById('teacherNamePromptCancel');
  const teacherNamePromptSave = document.getElementById('teacherNamePromptSave');

  // Modal elements
  const modalBackdrop = document.getElementById('modalBackdrop');
  const hdrFind = document.getElementById('hdrFind');
  const hdrReplace = document.getElementById('hdrReplace');
  const hdrRmStart = document.getElementById('hdrRmStart');
  const hdrRmEnd = document.getElementById('hdrRmEnd');
  const hdrBetweenPre = document.getElementById('hdrBetweenPre');
  const hdrBetweenSuf = document.getElementById('hdrBetweenSuf');
  const hdrAddPre = document.getElementById('hdrAddPre');
  const hdrAddSuf = document.getElementById('hdrAddSuf');
  const hdrCase = document.getElementById('hdrCase');
  const hdrTrim = document.getElementById('hdrTrim');
  const hdrCollapse = document.getElementById('hdrCollapse');
  const hdrLower = document.getElementById('hdrLower');
  const hdrInclusive = document.getElementById('hdrInclusive');
  const hdrNormalize = document.getElementById('hdrNormalize');
  const hdrTargetFilter = document.getElementById('hdrTargetFilter');
  const hdrPreview = document.getElementById('hdrPreview');
  const hdrApply = document.getElementById('hdrApply');
  const hdrCancel = document.getElementById('hdrCancel');
  const headerCountChip = document.getElementById('headerCountChip');
