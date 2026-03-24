// sidepanel.js — BrightBridge

// ── Constants ──────────────────────────────────────────────────────────────────

const API_ENDPOINT = 'https://jotaeliezer-teacher-tools-api-fj1k.vercel.app/api/generate-comment';
const CONCURRENCY  = 2;
const UNDERPERFORM_THRESHOLD = 60;

const COMMENT_BANK_MINI = [
  {
    id: 'participation', title: '🙋 Class Participation',
    options: [
      {id:'cp1', label:'Consistently participates, valuable insights', text:"[Student] consistently raises [his/her] hand and contributes valuable insights to class discussions.", type:'positive'},
      {id:'cp2', label:'Eagerly contributes, actively engages', text:"[Student] eagerly contributes to class discussions and actively engages with the material.", type:'positive'},
      {id:'cp3', label:'Active participant, positive environment', text:"[Student] is an active participant who helps create a positive learning environment for peers.", type:'positive'},
      {id:'cp4', label:'Needs to participate more, be more outgoing', text:"[Student] is encouraged to be more outgoing in class and volunteer answers more frequently.", type:'constructive'},
      {id:'cp5', label:'Adequate participation, could improve', text:"[Student] participates adequately but would benefit from raising [his/her] hand more often.", type:'constructive'},
      {id:'cp6', label:'Rarely participates, needs engagement', text:"[Student] rarely volunteers answers or participates in class discussions and should make a stronger effort to engage.", type:'constructive'},
      {id:'cp7', label:'Asks insightful questions, curious', text:"[Student] often asks insightful questions about alternative solutions and demonstrates genuine curiosity about the material.", type:'positive'},
      {id:'cp8', label:'Good group work and collaboration', text:"[Student] works well in group discussions and contributes constructively to collaborative work.", type:'positive'},
      {id:'cp9', label:'Makes most of class time, enthusiastic', text:"[Student] makes the most of class time, engaging enthusiastically and dedicating [himself/herself] to discussions.", type:'positive'},
      {id:'cp10', label:'Be more present, ask clarifying questions', text:"[Student] is encouraged to be more present during class by asking questions to clarify understanding.", type:'constructive'}
    ]
  },
  {
    id: 'peer', title: '🤝 Peer Teaching & Helping',
    options: [
      {id:'peer1', label:'Serves as teacher among peers', text:"[Student] serves as a teacher among [his/her] peers and is always willing to help classmates succeed.", type:'positive'},
      {id:'peer2', label:'Willing to help peers succeed', text:"[Student] is always willing to help [his/her] peers succeed and contributes to a supportive learning environment.", type:'positive'},
      {id:'peer3', label:'Helps classmates, works constructively', text:"[Student] helps out [his/her] classmates and works constructively with peers.", type:'positive'},
      {id:'peer4', label:'Cheerful demeanor enhances classroom morale', text:"[His/Her] cheerful demeanor has enhanced the morale of our classroom.", type:'positive'},
      {id:'peer5', label:'Presence helps all students', text:"[His/Her] presence in the class helps all students and contributes to a positive atmosphere.", type:'positive'}
    ]
  },
  {
    id: 'homeworkQuality', title: '📚 Homework Quality',
    options: [
      {id:'hw1', label:'Evident effort, quality and detail', text:"[Student] consistently puts evident effort into [his/her] assignments, completing them with quality and attention to detail.", type:'positive'},
      {id:'hw2', label:'Assignments reflect hard work and detail', text:"[His/Her] assignments consistently reflect [his/her] hard work and attention to detail.", type:'positive'},
      {id:'hw3', label:'Care and detail evident in high-quality work', text:"[Student]'s care and attention to detail are evident in [his/her] high-quality assignments.", type:'positive'},
      {id:'hw4', label:'Diligence shines through in quality', text:"[Student]'s diligence shines through in the high quality of [his/her] assignments.", type:'positive'},
      {id:'hw5', label:'High quality, detailed problem-solving', text:"[Student]'s problem-solving assignments consistently demonstrate high quality, detailed solutions.", type:'positive'},
      {id:'hw6', label:'Needs better presentation/neatness', text:"[Student] needs to improve the presentation and neatness of [his/her] assignments.", type:'constructive'},
      {id:'hw7', label:'Be more detailed for bonus marks', text:"[Student] is encouraged to be more detailed with [his/her] solutions on problem-solving assignments to pick up bonus marks.", type:'constructive'}
    ]
  },
  {
    id: 'homeworkCompletion', title: '⏰ Homework Completion',
    options: [
      {id:'hwc1', label:'Start homework earlier in week', text:"[Student] should start homework earlier in the week to allow more time for thoughtful completion.", type:'constructive'},
      {id:'hwc2', label:'Excellent time management, starts early', text:"[Student] demonstrates excellent time management by starting [his/her] homework early in the week.", type:'positive'},
      {id:'hwc3', label:'Submits assignments on time', text:"[Student] demonstrates strong time management by consistently submitting assignments on time.", type:'positive'},
      {id:'hwc4', label:'Needs to complete homework fully, not rush', text:"[Student] should focus on completing homework assignments completely rather than rushing through them.", type:'constructive'},
      {id:'hwc5', label:'Complete fully AND start earlier', text:"[Student] needs to focus on completing homework assignments completely and starting them earlier in the week.", type:'constructive'},
      {id:'hwc6', label:'Prioritize homework for full benefit of class', text:"[Student] is encouraged to prioritize homework completion so that [he/she] can receive the full benefit of in-class instruction.", type:'constructive'},
      {id:'hwc7', label:'Submit homework on time', text:"[Student] should focus on submitting homework on time to stay on track with the curriculum.", type:'constructive'}
    ]
  },
  {
    id: 'studyHabits', title: '📖 Study Habits',
    options: [
      {id:'sh1', label:'Redo past assignments before tests', text:"[Student] should focus on test preparation by redoing previous assignments and reviewing errors carefully.", type:'constructive'},
      {id:'sh2', label:'Strong test preparation', text:"[Student] demonstrates strong test preparation skills and consistently performs well on assessments.", type:'positive'},
      {id:'sh3', label:'Lacks adequate test preparation', text:"[Student]'s test scores indicate a lack of adequate preparation and [he/she] needs to dedicate more study time.", type:'constructive'},
      {id:'sh4', label:'Make corrections, learn from mistakes', text:"[Student] should focus on making corrections to tests and learning from [his/her] mistakes.", type:'constructive'},
      {id:'sh5', label:'Review to recognize patterns', text:"[Student] is encouraged to review previous assignments to recognize recurring problem types and patterns.", type:'constructive'},
      {id:'sh6', label:'Avoid careless mistakes', text:"[Student] should be more attentive to avoid careless mistakes on tests and assignments.", type:'constructive'},
      {id:'sh7', label:'Impressive retest improvement (positive)', text:"[Student] showed impressive improvement on [his/her] retest, demonstrating strong ability to learn from mistakes and apply corrections.", type:'positive'},
      {id:'sh8', label:'Set consistent SOM study time', text:"[Student] is encouraged to prioritize a consistent time for Spirit of Math work to engage more deeply with the material.", type:'constructive'}
    ]
  },
  {
    id: 'assignmentTestGap', title: '📊 Assignment vs Test Gap',
    options: [
      {id:'gap1', label:'Does well on HW, struggles on tests', text:"Patterns indicate that while [Student] does very well on assignments, [he/she] struggles to translate this knowledge on the corresponding tests.", type:'constructive'},
      {id:'gap2', label:'Good on assignments, review needed for tests', text:"[Student] continues to do very well on assignments but sometimes struggles with tests, suggesting [he/she] needs to focus more on review before assessments.", type:'constructive'},
      {id:'gap3', label:'High quality work, but more test review needed', text:"[His/Her] work continues to be high quality, but test performance suggests more thorough review is needed when preparing for assessments.", type:'constructive'},
      {id:'gap4', label:'Strong in HW, needs better test translation', text:"[Student] demonstrates strong understanding in assignments but needs to translate this more effectively to test situations.", type:'constructive'}
    ]
  },
  {
    id: 'brightspace', title: '💻 Brightspace & Resources',
    options: [
      {id:'bs1', label:'Watch Brightspace videos after lessons', text:"[Student] is encouraged to watch the Brightspace videos after each lesson to reinforce [his/her] understanding of new concepts.", type:'constructive'},
      {id:'bs2', label:'Watch videos before starting homework', text:"[Student] should develop the habit of watching Brightspace videos to review lesson content before starting homework.", type:'constructive'},
      {id:'bs3', label:'Use videos to start homework', text:"[Student] should use Brightspace videos as a catalyst to begin [his/her] homework questions.", type:'constructive'},
      {id:'bs4', label:'Reattempt class examples independently', text:"[Student] is encouraged to reattempt examples taken up in class by [himself/herself] to reinforce understanding of new concepts.", type:'constructive'},
      {id:'bs5', label:'Utilize Brightspace resources', text:"[Student] should utilize the resources available on Brightspace to supplement [his/her] learning.", type:'constructive'},
      {id:'bs6', label:'Reinforce concepts with videos (formal)', text:"Moving forward, [Student] should reinforce [his/her] understanding of core concepts by watching the Brightspace videos after each lesson.", type:'constructive'},
      {id:'bs7', label:'Videos help with HW completion', text:"These videos will help [him/her] review the information necessary to complete homework successfully.", type:'constructive'}
    ]
  },
  {
    id: 'helpSeeking', title: '🆘 Seeking Help',
    options: [
      {id:'help1', label:'Use Microsoft Teams for help', text:"[Student] is encouraged to reach out for help on Microsoft Teams when needed.", type:'constructive'},
      {id:'help2', label:'Takes initiative to ask questions', text:"[Student] takes initiative to ask questions when [he/she] doesn't understand something.", type:'positive'},
      {id:'help3', label:'Needs to seek help more proactively', text:"[Student] needs to be more proactive about seeking help and should not hesitate to ask questions.", type:'constructive'},
      {id:'help4', label:'Use class time to ask questions', text:"[Student] should actively use class time to ask questions when needed, so that any difficulties can be addressed promptly.", type:'constructive'},
      {id:'help5', label:'Rarely seeks help, should ask more', text:"[Student] rarely seeks help and would benefit greatly from asking questions more frequently.", type:'constructive'},
      {id:'help6', label:'Reach out for extra help', text:"[Student] is encouraged to reach out for extra help as needed and continue practicing diligently.", type:'constructive'},
      {id:'help7', label:'Ready to learn, asks clarifying questions', text:"[Student] consistently comes to class ready to learn, asking insightful questions to clarify any uncertainties.", type:'positive'},
      {id:'help8', label:'Seek help BEFORE class', text:"[Student] is encouraged to seek help from [his/her] teacher prior to class to address difficulties early.", type:'constructive'}
    ]
  },
  {
    id: 'attendance', title: '📅 Attendance & Behavior',
    options: [
      {id:'att1', label:'Too many absences/make-ups', text:"[Student] would benefit from attending class more regularly and reducing reliance on make-up sessions.", type:'constructive'},
      {id:'att2', label:'Excellent attendance', text:"[Student] has excellent attendance and is consistently engaged during class time.", type:'positive'},
      {id:'att3', label:'Positive attitude, enthusiastic', text:"[Student] demonstrates a positive attitude and enthusiasm towards learning.", type:'positive'},
      {id:'att4', label:'Listens well, follows directions', text:"[Student] listens well to instructions and follows directions carefully.", type:'positive'},
      {id:'att5', label:'Pleasure to have in class', text:"[Student] is a pleasure to have in class and contributes to a positive classroom atmosphere.", type:'positive'},
      {id:'att6', label:'Pleasure to have in class (formal)', text:"It has been a pleasure having [Student] in the class this term.", type:'positive'},
      {id:'att7', label:'Ready-to-learn attitude', text:"[Student] comes to class each week with a ready-to-learn attitude.", type:'positive'},
      {id:'att8', label:'Regular attendance needed for curriculum', text:"Attending class regularly and reducing reliance on make-ups would help [Student] stay on track with the curriculum.", type:'constructive'},
      {id:'att9', label:'Attendance affecting assessment', text:"[Student]'s attendance in class has made it difficult to accurately assess [his/her] participation and progress.", type:'constructive'}
    ]
  },
  {
    id: 'personalQualities', title: '✨ Personal Qualities',
    options: [
      {id:'pq1',  label:'Warm and friendly', text:"[Student] is a warm and friendly student with a positive demeanor.", type:'positive'},
      {id:'pq2',  label:'Kind-hearted, offers valuable insights', text:"[Student] is a kind-hearted student who consistently offers valuable insights during class discussions.", type:'positive'},
      {id:'pq3',  label:'Honest and hardworking', text:"[Student] is an honest and hardworking student whose diligence has yielded remarkable results this term.", type:'positive'},
      {id:'pq4',  label:'Demonstrates humility and willingness to learn', text:"[Student] demonstrates humility, effort, and willingness to learn.", type:'positive'},
      {id:'pq5',  label:'Pleasant, remains positive through challenges', text:"[Student] is a pleasant student who remains positive while working through challenging problems.", type:'positive'},
      {id:'pq6',  label:'Smart and capable, great potential', text:"[Student] is a smart and capable student with great potential.", type:'positive'},
      {id:'pq7',  label:'Sincere, committed to overcome challenges', text:"[Student] demonstrates sincerity and a commitment to overcome challenges.", type:'positive'},
      {id:'pq8',  label:'Dedicated, diligent in all work', text:"[Student] is a dedicated student whose diligence is evident in all aspects of [his/her] work.", type:'positive'},
      {id:'pq9',  label:'Fortitude and commitment to excellence', text:"[Student] demonstrates fortitude and commitment to excellence that are admirable.", type:'positive'},
      {id:'pq10', label:'Takes work seriously, strives for excellence', text:"[Student] takes [his/her] work seriously and strives for excellence in all that [he/she] does.", type:'positive'},
      {id:'pq11', label:'Cheerful and kind', text:"[Student] is cheerful, kind, and contributes to a positive classroom atmosphere.", type:'positive'},
      {id:'pq12', label:'Polite and well-behaved', text:"[Student] is polite, well-behaved, and shows respect to teachers and peers alike.", type:'positive'},
      {id:'pq13', label:'Shows effort despite struggles', text:"Even when struggling with some homework submissions, [Student] remains present, attentive, and gives the impression of doing [his/her] best.", type:'positive'},
      {id:'pq14', label:'Reliable classroom presence', text:"[Student] consistently shows up ready to participate, bringing steady effort and a positive attitude to class.", type:'positive'}
    ]
  },
  {
    id: 'understanding', title: '🎓 Understanding & Mastery',
    options: [
      {id:'und1',  label:'Demonstrates mastery', text:"[Student] demonstrates mastery of core concepts and applies them skillfully.", type:'positive'},
      {id:'und2',  label:'Natural aptitude, impressive problem-solving', text:"[Student] demonstrates a natural aptitude for mathematics and impressive problem-solving skills.", type:'positive'},
      {id:'und3',  label:'Propensity and love for mathematics', text:"[Student] demonstrates a propensity and love for mathematics through active engagement with the material.", type:'positive'},
      {id:'und4',  label:'Evident skill, asset to class', text:"[His/Her] evident skill in mathematics has made [him/her] a highly engaged student and an asset to the class.", type:'positive'},
      {id:'und5',  label:'Advanced understanding', text:"[Student] demonstrates an advanced understanding of the material.", type:'positive'},
      {id:'und6',  label:'Understands basics, struggles with difficult', text:"[Student] shows understanding of many concepts but sometimes struggles with more challenging applications.", type:'constructive'},
      {id:'und7',  label:'Inconsistent understanding', text:"[Student] demonstrates inconsistent understanding and would benefit from more thorough review of fundamental concepts.", type:'constructive'},
      {id:'und8',  label:'Capable, needs more confidence', text:"[Student] is capable of excellent work but needs to develop more confidence in [his/her] abilities.", type:'constructive'},
      {id:'und9',  label:'Expertly applies rules, sophisticated techniques', text:"[Student] expertly applies mathematical rules and demonstrates sophisticated problem-solving techniques.", type:'positive'},
      {id:'und10', label:'Creative solutions to higher-order problems', text:"[Student] consistently impresses with [his/her] ability to think of creative solutions to higher-order problems.", type:'positive'}
    ]
  },
  {
    id: 'newStudent', title: '🆕 New Student Context',
    options: [
      {id:'new1', label:'Great commitment as new student', text:"[Student] has shown great commitment, perseverance, and dedication as a new student in the rigorous SOM program.", type:'positive'},
      {id:'new2', label:'Impressive adaptability, first year', text:"[Student] has demonstrated impressive adaptability in [his/her] first year in the program.", type:'positive'},
      {id:'new3', label:'Good progress for first-time student', text:"[Student] is showing a lot of progress despite being a first-time SOM student.", type:'positive'},
      {id:'new4', label:'Enthusiastically attends as new student', text:"[Student] continues to enthusiastically attend the Spirit of Math program as a new student.", type:'positive'},
      {id:'new5', label:'Fast-paced curriculum needs HW completion', text:"As our curriculum is sophisticated and fast-paced, [Student] is encouraged to stay on top of homework to receive the full benefit of instruction.", type:'constructive'},
      {id:'new6', label:'Adapting to pace, willing to learn', text:"[Student] is still adapting to the pace of the class but shows genuine willingness to keep learning and applying each concept.", type:'constructive'},
      {id:'new7', label:'Positive attitude despite lower marks', text:"Even though current marks are below target, [Student] stays positive, asks questions, and is beginning to adjust to the program's expectations.", type:'constructive'},
      {id:'new8', label:'Progressing steadily with support', text:"As a new enrollee, [Student] is progressing steadily; continued practice and feedback will help [him/her] apply concepts more confidently.", type:'constructive'},
      {id:'new9', label:'Building stamina for fast lessons', text:"[Student] is building stamina for our faster lessons; with ongoing practice, [he/she] will translate this effort into stronger results.", type:'constructive'}
    ]
  },
  {
    id: 'future', title: '🔮 Looking Ahead',
    options: [
      {id:'future1',  label:'On track for successful year', text:"[Student] is well-positioned for a successful year ahead.", type:'positive'},
      {id:'future2',  label:'Hope to see continued success', text:"I hope to see [Student] continue [his/her] strong performance in Term 2.", type:'positive'},
      {id:'future3',  label:'Confident of better results with effort', text:"With more consistent effort, I am confident [Student] is capable of achieving stronger results in the coming term.", type:'constructive'},
      {id:'future4',  label:'Hope to see leadership', text:"I hope to see [Student] being a leader in the class for the rest of the year.", type:'positive'},
      {id:'future5',  label:'Next term = opportunity', text:"The next term offers an opportunity to build on [his/her] strengths and address areas for improvement.", type:'constructive'},
      {id:'future6',  label:'Confident feedback will bring success', text:"By implementing this feedback, I am confident [Student] will see greater success in Term 2.", type:'constructive'},
      {id:'future7',  label:'Certain of success with feedback', text:"As [Student] implements this feedback, I am certain [he/she] will have a successful term 2.", type:'constructive'},
      {id:'future8',  label:'Look forward to working next term', text:"I look forward to continuing to work with [Student] next term.", type:'positive'},
      {id:'future9',  label:'Will excel with ongoing effort', text:"With ongoing effort, [he/she] will certainly continue to excel.", type:'positive'},
      {id:'future10', label:'Successful year with adjustments', text:"[Student] is on [his/her] way to a successful year if [he/she] makes a few adjustments.", type:'constructive'}
    ]
  },
  {
    id: 'familyCommunication', title: '📞 Family Communication',
    options: [
      {id:'fam1', label:'Kept parents informed about progress', text:"I have been in regular contact with [Student]'s family to keep them informed about progress and next steps.", type:'positive'},
      {id:'fam2', label:'Discussed missing assignments with parent', text:"I have spoken with [Student]'s parent about missing assignments and shared a plan to get caught up.", type:'positive'},
      {id:'fam3', label:'Answered homework questions via email', text:"I have responded to family emails with clarifications on homework questions to support [Student]'s learning at home.", type:'positive'},
      {id:'fam4', label:'Collaborating with family on support plan', text:"I am collaborating with [Student]'s family to provide consistent support between class and home.", type:'positive'},
      {id:'fam5', label:'Parent updated on progress checkpoints', text:"[Student]'s parent has been updated on progress checkpoints and understands the upcoming expectations.", type:'positive'},
      {id:'fam6', label:'Strong rapport with parents', text:"There is a strong rapport with [Student]'s parents; our ongoing communication helps keep goals aligned.", type:'positive'},
      {id:'fam7', label:'Parents responsive and supportive', text:"[Student]'s parents have been responsive and supportive in our conversations about progress and next steps.", type:'positive'}
    ]
  }
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

let currentMode       = 'bulk'; // 'bulk' | 'single'
let students          = [];
let generatedComments = {};
let showOnlyUnderperf = false;
let isGeneratingAll   = false;

// Single student state
let singleStudentData = null;
let singleState = {
  pronoun:          'unknown',
  perfOverride:     null,
  selectedBank:     new Set(),   // kept for bulk card compat; single student uses selectedBankItems
  selectedBankItems: new Map(),  // Map<id, {id, categoryId, text, type}>
  customNote:       '',
  selectedAssigns:  [],
  selectedUpcoming: null,        // { name } or null
};

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
const bulkPanel                = $('bulkPanel');
const singlePanel              = $('singlePanel');
const singleContent            = $('singleContent');
const modeBulkBtn              = $('modeBulkBtn');
const modeSingleBtn            = $('modeSingleBtn');

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

  // Columns to exclude from the assignment list:
  // - System/metadata cols (email, id, enrollment, etc.)
  // - Aggregate/summary cols (subtotal, average, total, term summary)
  // - Non-assignment graded items (drills, contests, bonus, participation, exam version, etc.)
  // Assignment columns we WANT have an "Lxx - Name" pattern — those are kept.
  const skipPatterns = /^(email|username|orgid|org\s*defined|id\s*#|userid|subtotal|average|total|bonus|participation|cooperation|contest|competition|olympiad|cemc|cnml|caribou|gauss|euclid|cayley|fermat|fryer|galois|hypatia?|pascal|mathematica|kangaroo|drill|probation|withdrawn|year.end|assessment|somc|term\s*\d|final\s*exam|exam\s*version|enrolled|homework\s*comp)/i;
  const skipCols = new Set(
    [lastCol, firstCol, fullCol, gradeCol, ...headers.filter(h => skipPatterns.test(norm(h)))].filter(Boolean)
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
  const termCols = getTermAssignCols(termCode);
  // NOTE: No cross-term fallback here. If the selected term has no lesson-numbered
  // columns, we show a message rather than silently showing another term's data.

  if (!termCols.length) {
    s1List.innerHTML = `<p class="assign-empty">No assignment columns found for ${termSelect.value}. Make sure Brightspace assignment columns include a lesson number in their name (e.g. "L14 – Assignment Name").</p>`;
  } else {
    termCols.forEach(col => {
      const label  = cleanAssignLabel(col);
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

      row.appendChild(cb);
      row.appendChild(lbl);
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
  s2Hint.textContent = 'Optionally pick one upcoming test to mention encouragingly. Shows columns with "Test" in the name and no marks entered yet for this term.';

  const s2Search = document.createElement('input');
  s2Search.className = 'assign-search';
  s2Search.placeholder = 'Search tests…';
  s2Search.type = 'text';

  const s2List = document.createElement('div');
  s2List.className = 'assign-list';

  // Show test columns for the selected term that have ZERO marks entered.
  // A column with no marks = the test hasn't happened yet = it's upcoming.
  // Spirit of Math names all test columns with the word "Test"
  // e.g. "L6 - Relocation Property Test", "L7 - Mastermind Test".
  // Now that the rowspan data bug is fixed, colHasMarks() is accurate so
  // we can use the strict rule: any mark at all = test already happened.
  const termColsForUpcoming = getTermAssignCols(termCode);
  const upcomingCols = (termColsForUpcoming.length ? termColsForUpcoming : allAssignCols)
    .filter(col => /\btest\b/i.test(cleanAssignLabel(col)) && !colHasMarks(col));

  if (!upcomingCols.length) {
    s2List.innerHTML = '<p class="assign-empty">No upcoming tests found for this term. Columns with "Test" in the name and no marks entered yet will appear here.</p>';
  } else {
    upcomingCols.forEach(col => {
      const label  = cleanAssignLabel(col);
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

  // Performance override (versioned)
  const perfSection = document.createElement('div');
  perfSection.className = 'adv-section';
  const perfLabel = document.createElement('div');
  perfLabel.className = 'adv-label';
  perfLabel.textContent = 'Performance';
  const perfSelect = document.createElement('select');
  perfSelect.className = 'adv-select';

  const autoOpt = document.createElement('option');
  autoOpt.value = '';
  autoOpt.textContent = `Auto (${perfCode(student.gradeNum).replace('_', ' ')})`;
  if (!state.perfOverride) autoOpt.selected = true;
  perfSelect.appendChild(autoOpt);

  [
    { group: 'Good',          base: 'good' },
    { group: 'Satisfactory',  base: 'satisfactory' },
    { group: 'Average',       base: 'average' },
    { group: 'New Student',   base: 'newstu' },
    { group: 'Needs Support', base: 'poor' }
  ].forEach(({ group, base }) => {
    const og = document.createElement('optgroup');
    og.label = group;
    [1, 2, 3].forEach(v => {
      const o = document.createElement('option');
      o.value = `${base}${v}`;
      o.textContent = `${group} – Version ${v}`;
      if (o.value === state.perfOverride) o.selected = true;
      og.appendChild(o);
    });
    perfSelect.appendChild(og);
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
    const catHeader = document.createElement('div');
    catHeader.className = 'bank-category bank-category-header';
    const catTitle = document.createElement('span');
    catTitle.textContent = cat.title;
    const catArrow = document.createElement('span');
    catArrow.textContent = '▸';
    catArrow.style.fontSize = '10px';
    catArrow.style.transition = 'transform .15s';
    catHeader.appendChild(catTitle);
    catHeader.appendChild(catArrow);
    bankSection.appendChild(catHeader);

    const itemsContainer = document.createElement('div');
    itemsContainer.className = 'bank-items-container collapsed';
    catHeader.addEventListener('click', () => {
      const collapsed = itemsContainer.classList.toggle('collapsed');
      catArrow.style.transform = collapsed ? '' : 'rotate(90deg)';
    });

    cat.options.forEach(item => {
      const itemRow = document.createElement('label');
      itemRow.className = `bank-item bank-item--${item.type}`;
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = state.selectedBank.has(item.text);
      cb.addEventListener('change', () => {
        if (cb.checked) state.selectedBank.add(item.text);
        else            state.selectedBank.delete(item.text);
      });
      itemRow.appendChild(cb);
      itemRow.appendChild(document.createTextNode(' ' + item.label));
      itemsContainer.appendChild(itemRow);
    });
    bankSection.appendChild(itemsContainer);
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

function normalizePerfForApi(val) {
  // API expects versioned values: good1, satisfactory1, average1, poor1, newstu1, etc.
  // Ensure the value has a version suffix; map needs_support → poor.
  let base = val.replace(/\d+$/, '');                // strip trailing digits
  const ver  = val.match(/\d+$/)?.[0] || '1';        // keep existing version or default to 1
  if (base === 'needs_support') base = 'poor';
  return `${base}${ver}`;
}

function formatPerfLabel(val) {
  const map = {
    good: 'good', satisfactory: 'satisfactory', average: 'average',
    newstu: 'new student', poor: 'needs support', needs_support: 'needs support'
  };
  const m = val.match(/^([a-z_]+?)(\d+)$/);
  if (m) return `${map[m[1]] || m[1]} version ${m[2]}`;
  return val.replace(/_/g, ' ');
}

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
    performanceLabel:        formatPerfLabel(resolvedPerf),
    needsSupport:            resolvedPerf === 'needs_support' || resolvedPerf.startsWith('poor'),
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

// ── Single Student mode ────────────────────────────────────────────────────────

function buildCommentBankSection(state) {
  const bankWrap = document.createElement('div');
  bankWrap.className = 'ss-section';
  const bankTitle = document.createElement('div');
  bankTitle.className = 'ss-section-title';
  bankTitle.textContent = 'Comment Bank';
  bankWrap.appendChild(bankTitle);

  COMMENT_BANK_MINI.forEach(cat => {
    const catHeader = document.createElement('div');
    catHeader.className = 'bank-category bank-category-header';
    const catSpan = document.createElement('span');
    catSpan.textContent = cat.title;
    const arrow = document.createElement('span');
    arrow.textContent = '▸';
    arrow.style.cssText = 'font-size:10px;transition:transform .15s';
    catHeader.appendChild(catSpan);
    catHeader.appendChild(arrow);
    bankWrap.appendChild(catHeader);

    const itemsWrap = document.createElement('div');
    itemsWrap.className = 'bank-items-container collapsed';
    catHeader.addEventListener('click', () => {
      const c = itemsWrap.classList.toggle('collapsed');
      arrow.style.transform = c ? '' : 'rotate(90deg)';
    });

    cat.options.forEach(item => {
      const row = document.createElement('label');
      row.className = `bank-item bank-item--${item.type}`;
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = state.selectedBankItems ? state.selectedBankItems.has(item.id) : state.selectedBank.has(item.text);
      cb.addEventListener('change', () => {
        if (state.selectedBankItems) {
          if (cb.checked) state.selectedBankItems.set(item.id, { id: item.id, categoryId: cat.id, text: item.text, type: item.type });
          else            state.selectedBankItems.delete(item.id);
        } else {
          if (cb.checked) state.selectedBank.add(item.text);
          else            state.selectedBank.delete(item.text);
        }
      });
      row.appendChild(cb);
      row.appendChild(document.createTextNode(' ' + item.label));
      itemsWrap.appendChild(row);
    });
    bankWrap.appendChild(itemsWrap);
  });
  return bankWrap;
}

function renderSingleStudentView() {
  if (!singleStudentData) return;
  const data = singleStudentData;

  // Parse name: Brightspace stores as "Last, First"
  const rawName  = data.studentName || 'Student';
  let firstName  = rawName;
  let lastName   = '';
  const commaIdx = rawName.indexOf(',');
  if (commaIdx > -1) {
    lastName  = rawName.slice(0, commaIdx).trim();
    firstName = rawName.slice(commaIdx + 1).trim();
  }
  const displayName = commaIdx > -1 ? `${firstName} ${lastName}` : rawName;

  const basePerfCode = data.finalPercent != null ? perfCode(data.finalPercent) : 'average';
  const perfLabels   = { good: 'Good', satisfactory: 'Satisfactory', average: 'Average', needs_support: 'Needs Support' };

  // Reset singleState on new load
  singleState.pronoun           = guessPronounFromName(firstName);
  singleState.perfOverride      = null;
  singleState.selectedBank      = new Set();
  singleState.selectedBankItems = new Map();
  singleState.customNote        = '';
  singleState.selectedAssigns   = [];   // unchecked by default
  singleState.selectedUpcoming  = null;

  singleContent.innerHTML = '';

  // ── Student info card ──
  const infoCard = document.createElement('div');
  infoCard.className = 'ss-info-card';
  const infoRow1 = document.createElement('div');
  infoRow1.className = 'ss-info-row';
  const nameEl = document.createElement('span');
  nameEl.className = 'ss-name';
  nameEl.textContent = displayName;
  const gradeEl = document.createElement('span');
  gradeEl.className = 'ss-grade';
  gradeEl.textContent = data.finalPercent != null ? `${data.finalPercent}%` : '—';
  const perfBadge = document.createElement('span');
  perfBadge.className = `ss-perf-badge perf-${basePerfCode}`;
  perfBadge.textContent = perfLabels[basePerfCode] || basePerfCode;
  infoRow1.appendChild(nameEl);
  infoRow1.appendChild(gradeEl);
  infoRow1.appendChild(perfBadge);
  infoCard.appendChild(infoRow1);
  singleContent.appendChild(infoCard);

  // ── Settings section ──
  const settingsWrap = document.createElement('div');
  settingsWrap.className = 'ss-section';
  const settingsTitle = document.createElement('div');
  settingsTitle.className = 'ss-section-title';
  settingsTitle.textContent = 'Settings';
  settingsWrap.appendChild(settingsTitle);

  // Term / Grade group / Structure dropdowns
  [
    { label: 'Term', id: 'ss-term', options: [['Term 1','Term 1'],['Term 2','Term 2'],['Term 3','Term 3']], state: 'term' },
    { label: 'Grade Grp', id: 'ss-gg', options: [['sk_gr2','SK – Gr 2'],['gr3_gr4','Gr 3 – Gr 4'],['gr5_gr8','Gr 5 – Gr 8'],['gr9_gr11','Gr 9 – Gr 11']], state: 'gg', default: 'gr5_gr8' },
    { label: 'Structure', id: 'ss-struct', options: [['sandwich_paragraph','Sandwich Paragraph'],['strengths_feedback_blocks','Strengths / Feedback'],['bullet_points','Bullet Points']], state: 'struct' }
  ].forEach(def => {
    const row = document.createElement('div');
    row.className = 'ss-control-row';
    const lbl = document.createElement('label');
    lbl.htmlFor = def.id;
    lbl.textContent = def.label;
    const sel = document.createElement('select');
    sel.id = def.id;
    sel.className = 'adv-select';
    def.options.forEach(([val, txt]) => {
      const o = document.createElement('option');
      o.value = val; o.textContent = txt;
      if (def.default && val === def.default) o.selected = true;
      sel.appendChild(o);
    });
    // Try to restore from bulk settings
    try {
      const saved = JSON.parse(localStorage.getItem('bb_settings') || '{}');
      if (def.state === 'term' && saved.term && sel.querySelector(`option[value="${saved.term}"]`)) sel.value = saved.term;
      if (def.state === 'gg' && saved.gradeGroup && sel.querySelector(`option[value="${saved.gradeGroup}"]`)) sel.value = saved.gradeGroup;
      if (def.state === 'struct' && saved.structure && sel.querySelector(`option[value="${saved.structure}"]`)) sel.value = saved.structure;
    } catch (_) {}
    row.appendChild(lbl);
    row.appendChild(sel);
    settingsWrap.appendChild(row);
  });

  // Pronoun radio
  const pronounRow = document.createElement('div');
  pronounRow.className = 'ss-control-row';
  const pronounLbl = document.createElement('label');
  pronounLbl.textContent = 'Pronoun';
  const pronounGrp = document.createElement('div');
  pronounGrp.className = 'ss-pronoun-group pronoun-radio-group';
  [{ val:'unknown', label:'Auto' },{ val:'he', label:'🚹 He' },{ val:'she', label:'🚺 She' }].forEach(opt => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'pronoun-radio-btn' + (singleState.pronoun === opt.val ? ' active' : '');
    btn.dataset.val = opt.val;
    btn.textContent = opt.label;
    btn.addEventListener('click', () => {
      singleState.pronoun = opt.val;
      pronounGrp.querySelectorAll('.pronoun-radio-btn').forEach(b => b.classList.toggle('active', b.dataset.val === opt.val));
    });
    pronounGrp.appendChild(btn);
  });
  pronounRow.appendChild(pronounLbl);
  pronounRow.appendChild(pronounGrp);
  settingsWrap.appendChild(pronounRow);

  // Performance override (versioned)
  const perfRow = document.createElement('div');
  perfRow.className = 'ss-control-row';
  const perfLbl = document.createElement('label');
  perfLbl.textContent = 'Performance';
  const perfSel = document.createElement('select');
  perfSel.className = 'adv-select';
  const autoO = document.createElement('option');
  autoO.value = ''; autoO.textContent = `Auto (${perfLabels[basePerfCode] || basePerfCode})`; autoO.selected = true;
  perfSel.appendChild(autoO);
  [{ group:'Good', base:'good' },{ group:'Satisfactory', base:'satisfactory' },{ group:'Average', base:'average' },{ group:'New Student', base:'newstu' },{ group:'Needs Support', base:'poor' }].forEach(({ group, base }) => {
    const og = document.createElement('optgroup'); og.label = group;
    [1,2,3].forEach(v => {
      const o = document.createElement('option'); o.value = `${base}${v}`; o.textContent = `${group} – Version ${v}`;
      og.appendChild(o);
    });
    perfSel.appendChild(og);
  });
  perfSel.addEventListener('change', () => { singleState.perfOverride = perfSel.value || null; });
  perfRow.appendChild(perfLbl);
  perfRow.appendChild(perfSel);
  settingsWrap.appendChild(perfRow);

  singleContent.appendChild(settingsWrap);

  // ── Assignments section ──
  const assignWrap = document.createElement('div');
  assignWrap.className = 'ss-section';
  const assignTitle = document.createElement('div');
  assignTitle.className = 'ss-section-title';
  assignTitle.textContent = 'Assignments (from page)';
  assignWrap.appendChild(assignTitle);

  const termSelEl = settingsWrap.querySelector('#ss-term');
  function filterAssignsByTerm() {
    const termCode = normalizeTermCode(termSelEl ? termSelEl.value : 'Term 1');
    const ranges   = { T1:[1,13], T2:[14,26], T3:[27,39] };
    const [lo, hi] = ranges[termCode] || [1,99];
    return (data.assignments || []).filter(a => {
      const m = a.name.match(/\bL(\d+)\b/i);
      if (!m) return false;
      const n = parseInt(m[1]);
      return n >= lo && n <= hi;
    });
  }

  const assignList = document.createElement('div');
  assignList.className = 'ss-assign-list';

  function renderAssignList() {
    assignList.innerHTML = '';
    const filtered = filterAssignsByTerm();
    if (!filtered.length) {
      assignList.innerHTML = '<p style="color:var(--muted);font-size:12px;padding:4px">No assignments found for selected term.</p>';
      singleState.selectedAssigns = [];
      return;
    }
    singleState.selectedAssigns = []; // unchecked by default
    filtered.forEach(a => {
      const row = document.createElement('div');
      const gradeTier = a.percent >= 85 ? 'high' : a.percent >= 70 ? 'mid' : 'low';
      row.className = `ss-assign-row ss-assign-row--${gradeTier}`;
      const cb = document.createElement('input');
      cb.type = 'checkbox'; cb.checked = false;
      cb.addEventListener('change', () => {
        if (cb.checked) { if (!singleState.selectedAssigns.includes(a)) singleState.selectedAssigns.push(a); }
        else            { singleState.selectedAssigns = singleState.selectedAssigns.filter(x => x !== a); }
      });
      const lbl = document.createElement('span');
      lbl.textContent = a.name;
      lbl.style.flex = '1';
      const pct = document.createElement('span');
      pct.className = 'ss-assign-pct';
      pct.textContent = `${a.percent}%`;
      row.appendChild(cb); row.appendChild(lbl); row.appendChild(pct);
      row.addEventListener('click', e => { if (e.target === cb) return; cb.checked = !cb.checked; cb.dispatchEvent(new Event('change')); });
      assignList.appendChild(row);
    });
  }

  renderAssignList();
  if (termSelEl) termSelEl.addEventListener('change', renderAssignList);

  assignWrap.appendChild(assignList);
  singleContent.appendChild(assignWrap);

  // ── Upcoming Tests section ──
  if (data.upcomingItems && data.upcomingItems.length > 0) {
    const upcomingWrap = document.createElement('div');
    upcomingWrap.className = 'ss-section';
    const upcomingTitle = document.createElement('div');
    upcomingTitle.className = 'ss-section-title';
    upcomingTitle.textContent = 'Upcoming Tests';
    const upcomingHint = document.createElement('p');
    upcomingHint.style.cssText = 'font-size:11px;color:var(--muted);margin:2px 0 6px';
    upcomingHint.textContent = 'Select one to mention at the end of the comment (optional).';
    upcomingWrap.appendChild(upcomingTitle);
    upcomingWrap.appendChild(upcomingHint);

    const upcomingList = document.createElement('div');
    upcomingList.className = 'ss-assign-list';

    // "None" option
    const noneRow = document.createElement('div');
    noneRow.className = 'ss-assign-row';
    const noneRb = document.createElement('input');
    noneRb.type = 'radio'; noneRb.name = 'ss-upcoming'; noneRb.value = '';
    noneRb.checked = true;
    noneRb.addEventListener('change', () => { if (noneRb.checked) singleState.selectedUpcoming = null; });
    const noneLbl = document.createElement('span');
    noneLbl.textContent = 'None';
    noneLbl.style.flex = '1';
    noneRow.appendChild(noneRb); noneRow.appendChild(noneLbl);
    noneRow.addEventListener('click', e => { if (e.target === noneRb) return; noneRb.checked = true; noneRb.dispatchEvent(new Event('change')); });
    upcomingList.appendChild(noneRow);

    data.upcomingItems.forEach(item => {
      const row = document.createElement('div');
      row.className = 'ss-assign-row';
      const rb = document.createElement('input');
      rb.type = 'radio'; rb.name = 'ss-upcoming'; rb.value = item.name;
      rb.addEventListener('change', () => { if (rb.checked) singleState.selectedUpcoming = item; });
      const lbl = document.createElement('span');
      lbl.textContent = item.name;
      lbl.style.flex = '1';
      row.appendChild(rb); row.appendChild(lbl);
      row.addEventListener('click', e => { if (e.target === rb) return; rb.checked = true; rb.dispatchEvent(new Event('change')); });
      upcomingList.appendChild(row);
    });

    upcomingWrap.appendChild(upcomingList);
    singleContent.appendChild(upcomingWrap);
  }

  // ── Comment bank ──
  singleContent.appendChild(buildCommentBankSection(singleState));

  // ── Custom note ──
  const noteWrap = document.createElement('div');
  noteWrap.className = 'ss-section';
  const noteTitle = document.createElement('div');
  noteTitle.className = 'ss-section-title';
  noteTitle.textContent = 'Custom Note';
  const noteTA = document.createElement('textarea');
  noteTA.className = 'ss-comment-out';
  noteTA.style.minHeight = '50px';
  noteTA.placeholder = 'e.g. "Recently joined from another school." — added to the prompt';
  noteTA.rows = 2;
  noteTA.addEventListener('input', () => { singleState.customNote = noteTA.value; });
  noteWrap.appendChild(noteTitle);
  noteWrap.appendChild(noteTA);
  singleContent.appendChild(noteWrap);

  // ── Generate + output ──
  const outWrap = document.createElement('div');
  outWrap.className = 'ss-section';
  const outTitle = document.createElement('div');
  outTitle.className = 'ss-section-title';
  outTitle.textContent = 'Generated Comment';
  const outTA = document.createElement('textarea');
  outTA.className = 'ss-comment-out';
  outTA.rows = 5;
  outTA.placeholder = 'Generated comment will appear here…';
  const actionRow = document.createElement('div');
  actionRow.className = 'ss-action-row';

  const genBtn = document.createElement('button');
  genBtn.type = 'button';
  genBtn.className = 'btn primary ss-gen-btn';
  genBtn.textContent = '▶ Generate Comment';

  const copyBtn = document.createElement('button');
  copyBtn.type = 'button';
  copyBtn.className = 'btn ss-copy-btn';
  copyBtn.textContent = 'Copy';

  copyBtn.addEventListener('click', () => {
    if (outTA.value.trim()) {
      navigator.clipboard.writeText(outTA.value).catch(() => {});
      copyBtn.textContent = 'Copied!';
      setTimeout(() => { copyBtn.textContent = 'Copy'; }, 1500);
    }
  });

  genBtn.addEventListener('click', async () => {
    const termSelEl2   = singleContent.querySelector('#ss-term');
    const ggSelEl      = singleContent.querySelector('#ss-gg');
    const structSelEl  = singleContent.querySelector('#ss-struct');
    const term         = termSelEl2 ? termSelEl2.value : 'Term 1';
    const gradeGroup   = ggSelEl    ? ggSelEl.value    : 'gr5_gr8';
    const structure    = structSelEl ? structSelEl.value : 'sandwich_paragraph';
    const resolvedPerf = singleState.perfOverride || basePerfCode;
    const resolvedPron = singleState.pronoun === 'unknown' ? guessPronounFromName(firstName) : singleState.pronoun;

    const payload = {
      reviseMode:       'refine_basic',
      basicStructure:   structure,
      draft:            '',
      studentName:      displayName,
      studentFirstName: firstName,
      pronoun:          resolvedPron === 'unknown' ? 'he' : resolvedPron,
      finalGrade:       data.finalPercent != null ? `${data.finalPercent}%` : '',
      performanceLevel: normalizePerfForApi(resolvedPerf),
      termLabel:        term,
      gradeGroup,
      assignmentFacts:  singleState.selectedAssigns.slice(0, 3).map(a => ({ label: a.name, scoreText: `${a.percent}%` })),
      futureAssignments: singleState.selectedUpcoming ? [singleState.selectedUpcoming.name] : [],
      lateAssignments:  [],
      selectedComments: [...singleState.selectedBankItems.values()].map(item => ({
        id: item.id, category: item.categoryId, text: item.text, type: item.type
      })),
      customComment:    singleState.customNote.trim()
    };

    genBtn.disabled = true;
    genBtn.textContent = '…';
    outTA.classList.add('loading');
    outTA.value = '';

    try {
      const res = await fetch(API_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (!res.ok) throw new Error(`Server error ${res.status}`);
      const responseData = await res.json();
      const comment = parseApiResponse(responseData);
      if (!comment) throw new Error('Empty response from server');
      outTA.classList.remove('loading');
      typewriterAnimate(outTA, comment, () => {
        outTA.classList.add('has-content');
        copyBtn.classList.add('has-content');
        copyBtn.textContent = 'Copy ✓';
        genBtn.disabled = false;
        genBtn.textContent = '↺ Regenerate';
      });
    } catch (err) {
      outTA.value = `⚠ Error: ${err.message}`;
      outTA.classList.remove('loading');
      genBtn.disabled = false;
      genBtn.textContent = '▶ Generate Comment';
    }
  });

  actionRow.appendChild(genBtn);
  actionRow.appendChild(copyBtn);
  outWrap.appendChild(outTitle);
  outWrap.appendChild(actionRow);
  outWrap.appendChild(outTA);
  singleContent.appendChild(outWrap);
}

async function loadSingleStudent() {
  setStatus('Reading student data from page…', 'loading');
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab) throw new Error('No active tab found.');

    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ['content.js']
    }).catch(() => {});

    const result = await chrome.tabs.sendMessage(tab.id, { action: 'scrapeSingleStudent' });
    if (!result)         throw new Error('No response from page.');
    if (!result.success) throw new Error(result.error);

    singleStudentData = result;
    renderSingleStudentView();
    setStatus(`✓ Loaded ${result.studentName}`, 'success');
  } catch (err) {
    setStatus(`⚠ ${err.message}`, 'error');
    // Revert to Bulk mode if scraping fails
    currentMode = 'bulk';
    modeBulkBtn.classList.add('active');
    modeSingleBtn.classList.remove('active');
    bulkPanel.style.display  = '';
    singlePanel.style.display = 'none';
  }
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

    if (result.paginationWarning) {
      // Pagination warning takes priority — stays visible so teacher acts on it
      setStatus(result.paginationWarning, 'error');
    } else {
      setStatus(`✓ ${students.length} students loaded`, 'success');
    }

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

modeBulkBtn.addEventListener('click', () => {
  if (currentMode === 'bulk') return;
  currentMode = 'bulk';
  modeBulkBtn.classList.add('active');
  modeSingleBtn.classList.remove('active');
  bulkPanel.style.display   = '';
  singlePanel.style.display = 'none';
});

modeSingleBtn.addEventListener('click', () => {
  if (currentMode === 'single') return;
  currentMode = 'single';
  modeSingleBtn.classList.add('active');
  modeBulkBtn.classList.remove('active');
  bulkPanel.style.display   = 'none';
  singlePanel.style.display = '';
  loadSingleStudent();
});

refreshBtn.addEventListener('click', () => {
  if (currentMode === 'single') loadSingleStudent();
  else loadGrades();
});
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
