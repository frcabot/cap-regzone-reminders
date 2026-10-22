/**
 * CAP Reg Zone reminder emails (Google Sheets + Apps Script)
 *
 * Reads the CAP Registration Zone (Reg Zone) exports for an activity (for example an encampment: staff, students, an advanced
 * training squadron) from a Google Drive folder and emails each person ONLY what they still owe, with due dates:
 *   - To-do list: staff, and students who are slotted. Unit and parent approval, promotion to C/Amn, payment, RST, CPPT,
 *     medical profile, Level I, linking a parent account, adding a parent email, and your own extra items (Extra items tab).
 *   - Waitlist update: alternates. Place in line (one line, or separate male and female lines), chances, what blocks them.
 *   - Unit commander list: once a week, one email per unit commander, with a table of who owes what, plus alternates.
 * Parents are emailed only for cadets under 18. The unit commander is copied only while unit approval is missing.
 * People whose list is empty get no email. Everything planned is shown on the "Reminder emails" tab before anything is sent.
 *
 * SET UP: see README.md. In short: make a Google Sheet, paste this into Extensions > Apps Script, fill in the SETTINGS below,
 * put the Reg Zone exports in a Drive folder, then run "setup". Use "Preview" and "Send test emails" before sending for real.
 *
 * PRIVACY: nothing is copied out of Reg Zone except what each email needs. The spreadsheet shows names, CAPIDs, email addresses
 * and the items each person still owes. No dates of birth, phone numbers or addresses are written to the sheet.
 */

// ================================================================ SETTINGS (fill these in)
var CONFIG = {
  // ---- the activity (shown in every email)
  ACTIVITY_NAME: '20XX Example Wing Winter Encampment',
  ACTIVITY_WORD: 'encampment',              // how the activity is called in a sentence: "welcome X to encampment"
  LOCATION: 'Camp Example',                 // where it is held
  DATES: { Staff: '26 Dec 20XX - 3 Jan 20XX', Encampment: '27 Dec 20XX - 3 Jan 20XX' },   // per group (same names as GROUPS)
  LAST_DAY: '20XX-01-03',                   // last day, YYYY-MM-DD: CPPT must be current through this day; "18 or older by" this day

  // ---- who runs it (signs every email; people reply to this address)
  CONTACT_NAME: 'Capt Jane Doe',
  CONTACT_TITLE: 'Encampment Administration Officer',
  CONTACT_EMAIL: 'jane.doe@example.org',
  // Shown as "QUESTIONS: email ..." at the end of every email
  HELP_CONTACTS: 'Capt Jane Doe, Encampment Administration Officer (jane.doe@example.org)',

  // ---- the Reg Zone exports: a Drive folder holding one .xlsx export per Reg Zone activity, saved under these exact file names
  // (replace them each time you pull fresh exports). Optional: the "Attendance Prerequisite" report for the same activity, which
  // holds RST dates, medical profile status, Level I and the real CPPT expiration date.
  REGZONE_FOLDER_ID: '',
  GROUPS: [
    // type 'staff': everyone registered is working the activity and gets a to-do list.
    // type 'students': only Slotted registrations get a to-do list; everyone else is an alternate and gets the waitlist update.
    //   lines: 'sex' = separate male and female lines (for all-male and all-female flights); 'one' = a single line.
    //   order: 'score' = Reg Zone SlotScore (highest first, then sign-up order); 'signup' = sign-up order (RegistrationID).
    //   roomForEveryone: true = no place numbers: alternates are slotted as soon as their approvals are done.
    { name: 'Staff', type: 'staff', file: 'STAFF_latest.xlsx', prereqFile: 'STAFF_prereq_latest.xlsx' },
    { name: 'Encampment', type: 'students', file: 'STUDENTS_latest.xlsx', prereqFile: 'STUDENTS_prereq_latest.xlsx',
      lines: 'sex', order: 'score', roomForEveryone: false }
  ],

  // ---- email limits and timing
  EMAIL_RESERVE: 20,                        // daily email recipients kept free (Gmail has a daily limit)
  HOUR: 9,                                  // automatic runs at about this hour (script time zone)
  RUN_DAYS: ['MONDAY', 'THURSDAY'],         // automatic runs on these days
  TODO_EVERY_DAYS: 2,                       // to-do list: never twice within this many days
  WEEKLY_EVERY_DAYS: 6,                     // waitlist updates and commander lists: about once a week
  MANUAL_MIN_HOURS: 12,                     // "send now" skips anyone emailed in the last 12 hours
  MAX_DATA_AGE_HOURS: 36,                   // automatic runs stop (and email you) if the exports in Drive are older than this
  COMMANDER_LISTS: true,                    // weekly unit commander lists
  WAITLIST_EMAILS: true,                    // weekly waitlist updates for alternates
  TEST_CC: [],                              // "Send test emails" also goes to these reviewers (personal links removed)
  TAB: 'Reminder emails'
};

// What to remind about. false = leave it out. A date ('YYYY-MM-DD') is the due date; '' = "as soon as possible".
var ITEMS = {
  unit: '',              // unit commander approval in Reg Zone
  parent: '',            // parent/guardian approval (cadets under 18)
  parentLink: '',        // parent approved but no parent account linked in eServices (cadets under 18)
  parentEmail: '',       // no parent email in eServices (cadets under 18)
  cab: '20XX-11-15',     // SLOTTED cadets who are C/AB must be C/Amn by this date (alternates: as soon as possible). false = no minimum grade
  pay: '20XX-11-20',     // payment (fees below)
  rst: '20XX-12-01',     // staff: Required Staff Training (needs the prerequisite report), completed on or after RST_FROM
  cppt: '20XX-12-01',    // everyone 18 or older by LAST_DAY: CPPT current through LAST_DAY (complete or renew)
  med: '20XX-12-01',     // medical profile current in CAP Health (needs the prerequisite report)
  level1: ''             // senior staff whose Reg Zone "Member Level" prerequisite is not met
};
var FEES = { student: 500, cadetStaff: 500, senior: 100 };   // what each person owes in Reg Zone
var RST_FROM = '20XX-10-01';   // RST counts only if completed on or after this date
// Cadets whose ONLY open item is payment are often waiting on financial assistance: remind them less often until this date.
var PAY_ONLY = { EVERY_DAYS: 14, UNTIL: '' };   // UNTIL 'YYYY-MM-DD'; blank = no special schedule
// Added under the payment item for cadets. Text inside **double stars** is shown bold red. Blank = nothing added.
var PAYMENT_NOTE = '';   // e.g. '**IF YOU APPLIED FOR FINANCIAL ASSISTANCE: DO NOT PAY until decisions are announced.**'

// Per-person adjustments (CAPID as text)
var ADJUST = {
  NO_EMAIL: {},           // CAPID: reason. Never emailed (reason shown on the tab).
  SKIP_ITEMS: {},         // CAPID: ['pay', 'cppt', ...]. Leave out these items for this person.
  SENIOR_PRICE: {},       // CAPID: 1. A cadet in a senior role who pays the senior fee.
  PAID_OVERRIDE: {},      // CAPID: amount. Use instead of Reg Zone's AmountPaid (for example a waiver that shows as a payment).
  LEVEL1_PENDING: {}      // CAPID: name. Senior allowed to register before finishing Level I: reminded until Reg Zone shows it done.
};

// Waitlist emails
var WAITLIST = {
  DROP_RATE: 0.05,        // share of slotted people expected not to come: sets the "good / possible / unlikely" chances
  OFFER_HOURS: 24,        // when a spot opens, the alternate has this long to accept
  FEMALE_CAP_NOTE: '',    // optional, shown with the female line, e.g. 'There are 36 female spots.'
  HOW_SPOTS_TEXT: '',     // optional: how spots were given (shown to alternates)
  OTHER_OPTIONS_TEXT: '', // optional: other activities to consider (shown when the chances are not good)
  GET_READY_TEXT: ''      // optional: what alternates can do now, in case a spot opens
};

// Optional extra paragraph for slotted students (for example: "Your acceptance packet will come in a separate email...").
var STUDENT_NOTE = '';

// Optional picture guides (blank = not shown)
var GUIDES = { parentApproval: '', parentLink: '', medicalProfile: '' };

// Warm closings before the signature. {who} = the person's grade and name.
var CLOSINGS = {
  seniorStaff: 'THANK YOU for volunteering your time to serve on staff. This activity does not happen without volunteers like you, and we look ' +
    'forward to working with you!',
  cadetStaff: 'THANK YOU for serving on cadre. The students will learn a lot from you, and we look forward to working with you!',
  students: 'WE LOOK FORWARD TO SEEING YOU: we are excited to welcome {who} to {activity}, and we look forward to a great time together!',
  alternates: 'THANK YOU for your patience. We know waiting is hard, and we truly hope to welcome {who} to {activity}.',
  commanders: 'THANK YOU for supporting {members}. This activity works because commanders like you help them get ready.'
};
var NO_REPLY = 'This is an automated email. There is no need to reply unless you have a question or a concern. We are always happy to help!';

// ================================================================ end of settings

var REGZONE_URL = 'https://www.capnhq.gov/CAP.Experiences.Web/';                  // Reg Zone ("My Events" > Cancel)
var CONTACT_INFO_URL = 'https://www.capnhq.gov/MyAccount/ContactInfo.aspx';         // eServices contact info (parent emails)
var RANK_FIX = { 'CADET': 'C/AB' };   // Reg Zone shows C/AB as "CADET"

// ---------------------------------------------------------------- setup, menu, checks
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Reg Zone reminders')
    .addItem('Preview (fills the tab, sends nothing)', 'previewReminders')
    .addItem('Send test emails (to you and TEST_CC only)', 'testReminders')
    .addSeparator()
    .addItem('Send to EVERYONE now', 'sendToEveryoneNow')
    .addItem('Send to-do lists now', 'sendTodoNow')
    .addItem('Send waitlist updates now', 'sendWaitlistNow')
    .addItem('Send unit commander lists now', 'sendCommanderNow')
    .addItem('Send to ONE group now...', 'sendOneGroupNow')
    .addSeparator()
    .addItem('Turn ON automatic emails', 'turnOnAutomatic')
    .addItem('Turn OFF automatic emails', 'turnOffAutomatic')
    .addItem('Automation status', 'showStatus')
    .addItem('Run automation from THIS account', 'takeOverAutomation')
    .addToUi();
}

function setup() {
  var ui = SpreadsheetApp.getUi(), probs = configProblems_();
  if (probs.length) { ui.alert('Fix these settings first (at the top of the script):\n\n- ' + probs.join('\n- ')); return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  props_().setProperty('SS_ID', ss.getId());
  if (!ss.getSheetByName('Extra items')) {
    var sh = ss.insertSheet('Extra items');
    sh.getRange(1, 1, 1, 5).setValues([['CAPID', 'Item (short title)', 'Details (what to do)', 'Due (YYYY-MM-DD, blank = as soon as possible)',
      'Done? (Yes = stop reminding)']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  props_().setProperty('AUTOMATION_OWNER', Session.getEffectiveUser().getEmail());
  ensureRetryTrigger_();
  onOpen();
  var d = readData_();
  ui.alert('Setup done.\n\nReg Zone folder: ' + (d.missing.length ? 'MISSING FILES: ' + d.missing.join(', ') : 'all files found') + '.\n' +
    'Registrations read: ' + d.people.length + '.\n\nNext: Reg Zone reminders > Preview, then Send test emails.');
}

function configProblems_() {
  var p = [], c = CONFIG;
  if (/20XX|Example/i.test(c.ACTIVITY_NAME)) p.push('CONFIG.ACTIVITY_NAME still has the example text');
  if (/example\.org/i.test(c.CONTACT_EMAIL) || c.CONTACT_EMAIL.indexOf('@') < 0) p.push('CONFIG.CONTACT_EMAIL is not set');
  if (/Jane Doe/.test(c.CONTACT_NAME)) p.push('CONFIG.CONTACT_NAME still has the example name');
  if (/Jane Doe|example\.org/i.test(c.HELP_CONTACTS)) p.push('CONFIG.HELP_CONTACTS still has the example contact');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(c.LAST_DAY)) p.push('CONFIG.LAST_DAY must be YYYY-MM-DD');
  if (!c.REGZONE_FOLDER_ID) p.push('CONFIG.REGZONE_FOLDER_ID is blank');
  if (!c.GROUPS.length) p.push('CONFIG.GROUPS is empty');
  c.GROUPS.forEach(function (g) {
    if (g.type !== 'staff' && g.type !== 'students') p.push('group "' + g.name + '": type must be staff or students');
    if (!g.file) p.push('group "' + g.name + '": no export file name');
  });
  Object.keys(ITEMS).forEach(function (k) {
    var v = ITEMS[k];
    if (v !== false && v !== '' && !/^\d{4}-\d{2}-\d{2}$/.test(String(v))) p.push('ITEMS.' + k + ' must be false, \'\' or YYYY-MM-DD (now: ' + v + ')');
  });
  if (ITEMS.rst !== false && !/^\d{4}-\d{2}-\d{2}$/.test(RST_FROM)) p.push('RST_FROM must be YYYY-MM-DD');
  return p;
}

// ---------------------------------------------------------------- small helpers
function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { var id = props_().getProperty('SS_ID'); return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return Session.getScriptTimeZone(); }
function todayIso_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function niceIso_(iso) { return iso ? Utilities.formatDate(new Date(iso + 'T12:00:00'), tz_(), 'EEE d MMM yyyy') : ''; }
function shortIso_(iso) { return iso ? Utilities.formatDate(new Date(iso + 'T12:00:00'), tz_(), 'd MMM') : ''; }
function asOfTxt_(d) { return d ? Utilities.formatDate(d, tz_(), 'EEE d MMM h:mm a') : 'latest'; }
function money_(n) { return '$' + Math.round(n).toLocaleString('en-US'); }
function andList_(a) { return a.length < 2 ? a.join('') : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1]; }
function who_(p) { return (p.rank + ' ' + p.first + ' ' + p.last).trim(); }
// true when the group name adds nothing (only one student group, or the activity name already says it: "Encampment")
function groupNamed_(name) {
  return CONFIG.GROUPS.filter(function (g) { return g.type === 'students'; }).length < 2 || CONFIG.ACTIVITY_NAME.toLowerCase().indexOf(String(name).toLowerCase()) >= 0;
}
function groupOf_(name) { return CONFIG.GROUPS.filter(function (g) { return g.name === name; })[0] || {}; }
function fill_(t, p, extra) {
  return String(t || '').replace(/\{who\}/g, p ? who_(p) : '').replace(/\{activity\}/g, CONFIG.ACTIVITY_WORD + ' at ' + CONFIG.LOCATION)
    .replace(/\{members\}/g, extra || '');
}
// Reg Zone dates: "31 Aug 2027", Excel serial numbers (46312), or ISO. Returns "yyyy-MM-dd" or ''.
function rzDate_(v) {
  v = String(v || '').trim();
  if (!v) return '';
  if (/^\d{4,6}(\.\d+)?$/.test(v)) return new Date(Date.UTC(1899, 11, 30) + Math.round(Number(v)) * 86400000).toISOString().slice(0, 10);
  var m = /^(\d{1,2})[ -]([A-Za-z]{3})[A-Za-z]*[ -](\d{4})/.exec(v);
  if (m) {
    var mo = 'jan feb mar apr may jun jul aug sep oct nov dec'.split(' ').indexOf(m[2].toLowerCase());
    if (mo >= 0) return m[3] + '-' + ('0' + (mo + 1)).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v);
  return m ? m[0] : '';
}
// "Maj Christopher D. Devine" -> "Maj Devine" (rank and last name; suffixes like Jr or IV are not the last name)
function rankLast_(full) {
  var t = String(full || '').trim().replace(/,?\s+(Jr\.?|Sr\.?|II|III|IV|V|VI)$/i, '');
  var m = /^((?:C\/)?(?:Lt Col|1st Lt|2d Lt|LtCol|1stLt|2dLt|[A-Za-z]+))\s+.*?(\S+)$/.exec(t);
  return m ? m[1] + ' ' + m[2] : t;
}
// Names typed all lowercase or ALL CAPS become Title Case; mixed case is kept, except words that start lowercase get a capital
// (particles like "van" or "de" stay as entered).
var NAME_PARTICLES = /^(van|von|der|den|de|del|della|degli|di|da|das|do|dos|du|la|le|ten|ter|y|e|bin|binti|bint|al|el|ibn)$/;
function nameCase_(v) {
  var t = String(v || '').replace(/\s+/g, ' ').trim();
  if (!/[A-Za-zÀ-ɏ]/.test(t)) return t;
  if (t !== t.toLowerCase() && t !== t.toUpperCase()) {
    return t.split(/(\s+)/).map(function (w) {
      if (!w || /^\s+$/.test(w) || NAME_PARTICLES.test(w)) return w;
      var c = w.charAt(0);
      return c !== c.toLowerCase() || c === c.toUpperCase() ? w : c.toUpperCase() + w.slice(1);
    }).join('');
  }
  var parts = t.toLowerCase().split(/([\s\-'’.]+)/), wordNo = 0;
  return parts.map(function (w) {
    if (!w || /^[\s\-'’.]+$/.test(w)) return w;
    wordNo++;
    if (/^(ii|iii|iv)$/.test(w) || (wordNo > 1 && /^(v|vi|vii|viii|ix|x)$/.test(w))) return w.toUpperCase();
    if (/^mc.{2,}/.test(w)) return 'Mc' + w.charAt(2).toUpperCase() + w.slice(3);
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join('');
}

// Minimal .xlsx reader (first worksheet only): rows as arrays of strings, header row first.
function xlsxRows_(blob) {
  var parts = Utilities.unzip(blob.setContentType('application/zip')), files = {};
  parts.forEach(function (b) { files[b.getName()] = b; });
  var text = function (name) { return files[name] ? files[name].getDataAsString('UTF-8') : ''; };
  var dec = function (t) {
    return t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (m, n) { return String.fromCharCode(Number(n)); }).replace(/&amp;/g, '&');
  };
  var shared = [], sst = text('xl/sharedStrings.xml'), m, re = /<si>([\s\S]*?)<\/si>/g;
  while ((m = re.exec(sst))) {
    var t = '', tm, tre = /<t[^>]*>([\s\S]*?)<\/t>/g;
    while ((tm = tre.exec(m[1]))) t += tm[1];
    shared.push(dec(t));
  }
  var sheetName = Object.keys(files).filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); }).sort()[0];
  var xml = text(sheetName), out = [], rre = /<row[^>]*>([\s\S]*?)<\/row>/g, rm;
  var colNum = function (letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n - 1; };
  while ((rm = rre.exec(xml))) {
    var row = [], cre = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, cm;
    while ((cm = cre.exec(rm[1]))) {
      var type = (/ t="([^"]+)"/.exec(cm[2]) || [])[1], body = cm[3] || '', val = '';
      if (type === 'inlineStr') { var it = /<t[^>]*>([\s\S]*?)<\/t>/.exec(body); val = it ? dec(it[1]) : ''; }
      else { var vm = /<v>([\s\S]*?)<\/v>/.exec(body); val = vm ? (type === 's' ? shared[Number(vm[1])] : dec(vm[1])) : ''; }
      row[colNum(cm[1])] = val;
    }
    for (var i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = '';
    out.push(row);
  }
  return out;
}

// ---------------------------------------------------------------- reading the Reg Zone exports
function readData_() {
  var folder = DriveApp.getFolderById(CONFIG.REGZONE_FOLDER_ID), missing = [], oldest = null;
  function read(name, optional) {
    if (!name) return [];
    var it = folder.getFilesByName(name);
    if (!it.hasNext()) { if (!optional) missing.push(name); return []; }
    var f = it.next(), upd = f.getLastUpdated();
    if (!oldest || upd < oldest) oldest = upd;
    var rows = xlsxRows_(f.getBlob()), H = rows.shift() || [];
    return rows.map(function (x) { var o = {}; H.forEach(function (k, j) { o[String(k).trim()] = String(x[j] === undefined ? '' : x[j]).trim(); }); return o; });
  }
  var people = [];
  CONFIG.GROUPS.forEach(function (g) {
    var pre = {};
    read(g.prereqFile, true).forEach(function (r) {
      var cap = String(r.RegistrantsCAPID).replace(/\.0$/, ''), e = pre[cap] = pre[cap] || {};
      if (r.PrereqType === 'Require Medical Profile Complete') e.med = r.MeetsRequirement;
      if (r.PrereqType === 'Training') { e.rstSeen = true; e.rst = rzDate_(r.CompletionDate); }
      if (r.PrereqType === 'Member Level') e.level = r.MeetsRequirement;
      // The registration export says "Needs CPPT" even when CPPT is current but expires before the end; this report keeps the date.
      if (r.PrereqType === 'Cadet Protection Current' && r.CPPTExpiration) e.cppt = rzDate_(r.CPPTExpiration);
    });
    var staff = g.type === 'staff';
    read(g.file).forEach(function (r) {
      var cap = String(r.RegistrantsCAPID).replace(/\.0$/, '');
      if (!/^[0-9]{5,7}$/.test(cap)) return;
      var st = r.RegistrationStatus;
      if (/withdr|cancel|disapprov|denied/i.test(st)) return;
      var age = Number(r.AgeAtEventStart), ageEnd = Number(r.AgeAtEventEnd || r.AgeAtEventStart);
      var senior = String(r.MbrType).toUpperCase() !== 'CADET', pr = pre[cap] || {};
      people.push({
        ev: g.name, staff: staff, cap: cap, status: staff ? 'Staff' : st, slotted: staff || st === 'Slotted',
        rank: RANK_FIX[r.Rank] || r.Rank, first: nameCase_(r.NameFirst), last: nameCase_(r.NameLast),
        sex: /^f/i.test(r.Gender) ? 'F' : /^m/i.test(r.Gender) ? 'M' : '', senior: senior,
        minor: !(age >= 18), adultByEnd: senior || ageEnd >= 18, minorNow: Number(r.Age) < 18, regId: Number(r.RegistrationID) || 0,
        score: r.SlotScore !== undefined && r.SlotScore !== '' && !isNaN(Number(r.SlotScore)) ? Number(r.SlotScore) : null,
        unitOK: r.UnitApproved === 'Yes', parentOK: !(age < 18) || r.ParentApproved === 'Approved', linked: r.LinkedToParentAccount,
        cab: String(r.Rank).toUpperCase() === 'CADET' && !senior,
        email: r.Email, parents: [r.CadetParentEmailPrimary, r.CadetParentEmailSecondary].filter(function (a) { return /@/.test(a); }),
        parentEmer: r.CadetParentEmailEmergency, noPrimaryParent: !/@/.test(r.CadetParentEmailPrimary), hasSecondaryParent: /@/.test(r.CadetParentEmailSecondary),
        addr: String(r.Addr1 || '').toLowerCase().replace(/[^a-z0-9]/g, ''), ecNum: String(r.EmergencyContactNumber || '').replace(/\D/g, '').slice(-10),
        ccName: r.UnitCCName, ccEmail: r.UnitCCEmail, unit: String(r.Unit).replace(/\.0$/, ''), wing: r.Wing,
        paid: ADJUST.PAID_OVERRIDE.hasOwnProperty(cap) ? Number(ADJUST.PAID_OVERRIDE[cap]) : (Number(r.AmountPaid) || 0), paidFull: r.PaidInFull === 'Yes',
        cppt: rzDate_(r.CPPTExpiration) || pr.cppt || '', med: pr.med || '', rst: pr.rst || '', rstKnown: !!pr.rstSeen,
        needLevel1: senior && staff && (pr.level === 'No' || (!!ADJUST.LEVEL1_PENDING[cap] && pr.level !== 'Yes'))
      });
    });
  });
  // Cadets under 18 with no parent email: use the emergency parent email; failing that, a brother's or sister's parent emails
  // (same last name AND same home address or emergency phone). Anyone still without one is listed in the run summary.
  people.forEach(function (p) {
    if (p.senior || !p.minor || p.parents.length) return;
    if (/@/.test(p.parentEmer)) { p.parents = [p.parentEmer]; p.parentSrc = 'emergency parent email'; return; }
    var sib = people.filter(function (q) {
      return q !== p && !q.senior && q.parents.length && q.parentSrc !== 'sibling' && q.last.toLowerCase() === p.last.toLowerCase() &&
        ((p.addr && q.addr === p.addr) || (p.ecNum && q.ecNum === p.ecNum));
    })[0];
    if (sib) { p.parents = sib.parents.slice(); p.parentSrc = 'sibling'; } else p.noParentEmail = true;
  });
  return { people: people, asOf: oldest, missing: missing, extra: readExtra_() };
}

// "Extra items" tab: one row per person and item (CAPID, title, details, due, done). Rows marked Done = Yes are skipped.
function readExtra_() {
  var out = {};
  try {
    var sh = ss_().getSheetByName('Extra items');
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
      var cap = String(r[0] || '').replace(/\.0$/, '').trim();
      if (!cap || !String(r[1] || '').trim() || /^y/i.test(String(r[4] || ''))) return;
      var due = r[3] instanceof Date ? Utilities.formatDate(r[3], tz_(), 'yyyy-MM-dd') : rzDate_(r[3]);
      (out[cap] = out[cap] || []).push({ title: String(r[1]).trim(), text: String(r[2] || '').trim(), due: due });
    });
  } catch (e) { console.log('Extra items not read: ' + e); }
  return out;
}

// ---------------------------------------------------------------- parent approval how-to (CAP eServices, any activity)
var PARENT_APPROVE_STEPS = 'Log in to the CAP Parent Portal, click "Pending Parent/Guardian Approvals", then "Manage Approvals". ' +
  'Select the registration, click "Approve", then click "Submit".';
var PARENT_LINK_STEPS =
  '     1. Log in to eServices (https://www.capnhq.gov) with your CADET\'s account. Click the cadet\'s name, then ' +
  '"Parent / Guardian Account Linking". Pick your email address and click "Request to Link Account". Write down the 6-digit ' +
  '"Parent Access Code".\n' +
  '     2. Open the email from "CAP eServices Notification System" (check your spam folder). Click "Click here to create a provisional ' +
  'account to be linked". Enter the Parent Access Code, the cadet\'s last name and date of birth, check "I certify", and click ' +
  '"Verify Cadet Access". Then register and create a password. SAVE the password: you will need it again.';
function parentHow_(p) {
  var g = GUIDES.parentApproval ? '\n  Step-by-step guide with pictures: ' + GUIDES.parentApproval : '';
  if (p.linked === 'Yes') return 'HOW TO APPROVE: your parent account is already linked to this cadet, so you only need to approve.\n  ' + PARENT_APPROVE_STEPS + g;
  if (p.linked === 'No') return 'HOW TO APPROVE: CAP records show NO parent account linked to this cadet yet. Link one first, then approve.\n' +
    PARENT_LINK_STEPS + '\n     3. Now approve: ' + PARENT_APPROVE_STEPS + g;
  return 'HOW TO APPROVE:\n  A. You ALREADY have a CAP Parent Portal account: ' + PARENT_APPROVE_STEPS + '\n  B. You do NOT have a parent account yet:\n' +
    PARENT_LINK_STEPS + '\n     3. Now approve: ' + PARENT_APPROVE_STEPS + g;
}

// ---------------------------------------------------------------- what each person still owes
function dueOf_(k) { var v = ITEMS[k]; return v ? [v, niceIso_(v)] : ['0000-00-00', 'as soon as possible']; }
function on_(k, p) { return ITEMS[k] !== false && (ADJUST.SKIP_ITEMS[p.cap] || []).indexOf(k) < 0; }

// Returns { items: [{k, due, dueTxt, title, text, short}], done: [...] } sorted by due date.
function todoItems_(p, extra, opts) {
  opts = opts || {};
  var items = [], done = [], who = who_(p), student = !p.staff, END_TXT = niceIso_(CONFIG.LAST_DAY);
  function add(k, title, text, short, dueIso) {
    var d = dueIso !== undefined ? (dueIso ? [dueIso, niceIso_(dueIso)] : ['0000-00-00', 'as soon as possible']) : dueOf_(k);
    items.push({ k: k, due: d[0], dueTxt: d[1], title: title, text: text, short: short || '' });
  }
  if (on_('unit', p)) {
    if (!p.unitOK) add('unit', 'UNIT COMMANDER APPROVAL', 'The unit commander' + (p.ccName ? ' (' + rankLast_(p.ccName) +
      (opts.cc ? ', copied on this email' : '') + ')' : '') + ' must approve this registration in Reg Zone. ' +
      (opts.cc ? 'Commander: please approve it in Reg Zone.' : 'Please ask the commander to approve it.'), 'unit approval');
    else done.push('unit approval');
  }
  if (on_('parent', p) && !p.senior && p.minor) {
    if (!p.parentOK) add('parent', 'PARENT/GUARDIAN APPROVAL', 'A parent or guardian must approve this registration in Reg Zone (CAP eServices).\n' + parentHow_(p), 'parent approval');
    else done.push('parent approval');
  }
  if (on_('cab', p) && !p.staff && p.cab) add('cab', 'PROMOTION TO C/Amn', who + ' is a C/AB (Cadet Airman Basic). Only cadets who are C/Amn (Cadet Airman) ' +
    'or higher can attend. Please work with the squadron to complete the promotion.', 'promotion to C/Amn', p.slotted ? undefined : '');
  if (opts.approvalsOnly) return finish();
  if (on_('parentLink', p) && p.parentOK && p.minor && !p.senior && p.linked === 'No') add('parentLink', 'LINK A PARENT ACCOUNT IN eSERVICES',
    'CAP records show NO parent account linked to ' + who + '. A parent or guardian needs a CAP parent account linked to the cadet. It lets the ' +
    'parent see the cadet\'s CAP activities and give approvals in eServices. How to link:\n' + PARENT_LINK_STEPS +
    (GUIDES.parentLink ? '\n  Step-by-step guide with pictures: ' + GUIDES.parentLink : ''), 'link a parent account');
  if (on_('parentEmail', p) && !p.senior && p.minor && p.noPrimaryParent) add('parentEmail', 'ADD A PARENT EMAIL IN eSERVICES', (p.hasSecondaryParent
      ? 'CAP records have only a secondary parent email for ' + who + ', and no primary one.'
      : 'CAP records have no parent email for ' + who + ', so parents may miss important emails.') +
    ' Log in to eServices with the cadet\'s account, go to ' + CONTACT_INFO_URL + ', and add a "Cadet Parent" email. Ideally add two: a primary and ' +
    'a secondary parent email.', 'add a parent email');
  // payment. Seniors: Reg Zone's "PaidInFull" is not reliable (it can show Yes at $0), so the amount paid is used.
  if (on_('pay', p)) {
    var senPrice = p.senior || !!ADJUST.SENIOR_PRICE[p.cap];
    var fee = student ? FEES.student : senPrice ? FEES.senior : FEES.cadetStaff;
    var paidUp = p.paid >= fee || (!senPrice && p.paidFull && !ADJUST.PAID_OVERRIDE.hasOwnProperty(p.cap));
    if (fee > 0 && !paidUp) add('pay', 'PAYMENT', (student ? 'Fee' : 'Staff fee') + ': ' + money_(fee) + '. ' + (p.paid > 0 ? 'Reg Zone shows ' + money_(p.paid) +
      ' paid, so ' + money_(fee - p.paid) + ' is still owed.' : 'Reg Zone shows no payment yet.') + ' Pay in Reg Zone (CAP eServices).' +
      (!p.senior && PAYMENT_NOTE ? '\n' + PAYMENT_NOTE : ''), 'payment');
    else if (fee > 0) done.push('payment');
  }
  if (on_('rst', p) && p.staff && p.rstKnown) {
    if (!(p.rst && p.rst >= RST_FROM)) add('rst', 'RST (REQUIRED STAFF TRAINING)', 'Complete the Cadet Programs RST online modules in eServices. Only RST ' +
      'completed on or after ' + niceIso_(RST_FROM) + ' counts. ' + (p.rst ? 'The RST on record is from ' + niceIso_(p.rst) + ', so it must be done again.'
      : 'Reg Zone shows no RST on record.'), 'RST');
    else done.push('RST');
  }
  if (on_('level1', p) && p.needLevel1) add('level1', 'LEVEL I (senior member training)', 'Please finish Level I in eServices. Reg Zone does not show it ' +
    'done yet.', 'Level I');
  if (on_('cppt', p) && p.adultByEnd) {
    if (!(p.cppt && p.cppt >= CONFIG.LAST_DAY)) {
      // A date = CPPT done but runs out before the end (renew). No date = no CPPT on record (complete it). First-time CPPT for cadets also
      // needs a summary conversation with the squadron commander; renewals do not.
      add('cppt', 'CADET PROTECTION PROGRAM TRAINING (CPPT): ' + (p.cppt ? 'Renew' : p.senior ? 'Complete' : 'Complete (first time)'),
        'Everyone who is 18 or older by ' + END_TXT + ' must have CPPT that is current through that day. ' +
        (p.cppt ? 'Reg Zone shows the CPPT expires on ' + niceIso_(p.cppt) + ', before then, so it must be RENEWED. Renew CPPT in eServices.'
          : p.senior ? 'Reg Zone does not show CPPT on record. Please complete CPPT in eServices.'
          : who + ' will be 18 by ' + END_TXT + ', so CPPT is required. ' + (p.minorNow ? 'Cadets can take CPPT at 17, so there is no need to wait ' +
            'for the birthday. ' : '') + 'Reg Zone does not show CPPT on record. First-time CPPT has two parts:\n' +
            '   a) The online CPPT course in eServices, AND\n   b) A summary conversation with the squadron commander.\n' +
            '**Please work with your squadron commander to complete both parts.**'),
        p.cppt ? 'CPPT renewal' : p.senior ? 'CPPT' : 'first-time CPPT (course + summary conversation with you)');
    } else done.push('CPPT');
  }
  if (on_('med', p)) {
    if (p.med === 'No') add('med', 'MEDICAL PROFILE (CAP Health)', 'Reg Zone shows the medical profile is not current. Please update it in CAP ' +
      'Health (in eServices), even if it was filled out before.' + (GUIDES.medicalProfile ? ' Step-by-step guide with pictures: ' + GUIDES.medicalProfile : ''),
      'medical profile update');
    else if (p.med === 'Yes') done.push('medical profile');
  }
  (extra[p.cap] || []).forEach(function (x, i) { add('x' + i, x.title.toUpperCase(), x.text || 'Please take care of this.', x.title, x.due || ''); });
  return finish();
  function finish() {
    items.sort(function (a, b) { return a.due < b.due ? -1 : a.due > b.due ? 1 : 0; });
    return { items: items, done: done };
  }
}

function recipients_(p) {   // cadets under 18 at the start: cadet + parents; 18+ and seniors: themselves
  var seen = {}, to = [];
  [p.email].concat(!p.senior && p.minor ? p.parents : []).forEach(function (a) {
    a = String(a || '').trim();
    if (a && a.indexOf('@') > 0 && !seen[a.toLowerCase()]) { seen[a.toLowerCase()] = 1; to.push(a); }
  });
  return to;
}
function numbered_(items, today) {
  return items.map(function (it, i) {
    var late = it.due !== '0000-00-00' && today > it.due;
    return (i + 1) + '. ' + it.title + '. ' + (late ? 'OVERDUE (was due ' + it.dueTxt + '). Please do this right away.' : 'Due: ' + it.dueTxt + '.') + '\n' + it.text;
  }).join('\n\n');
}
function cancelSteps_() {
  return '- Step 1: go to the Reg Zone main page: ' + REGZONE_URL + ' (log in with the member\'s eServices account if it asks).\n' +
    '- Step 2: under "My Events", find ' + CONFIG.ACTIVITY_NAME + ' and click "Cancel".';
}
function signature_() { return CONFIG.CONTACT_NAME + '\n' + CONFIG.CONTACT_TITLE + ', ' + CONFIG.ACTIVITY_NAME; }
function footer_() { return 'QUESTIONS: email ' + CONFIG.HELP_CONTACTS + '.\n\n' + signature_(); }
function closing_(p) {
  if (p.staff) return fill_(p.senior ? CLOSINGS.seniorStaff : CLOSINGS.cadetStaff, p);
  return fill_(p.slotted ? CLOSINGS.students : CLOSINGS.alternates, p);
}
function isPayOnly_(p, t) { return !p.senior && t.items.length > 0 && t.items.every(function (it) { return it.k === 'pay'; }); }
function payOnlyActive_() { return !!PAY_ONLY.UNTIL && todayIso_() < PAY_ONLY.UNTIL; }

function todoEmail_(p, t, asOf) {
  var who = who_(p), today = todayIso_(), n = t.items.length;
  var late = t.items.some(function (it) { return it.due !== '0000-00-00' && today > it.due; });
  var body = 'Hello,\n\n' + NO_REPLY + '\n\n' + who + ' (CAPID ' + p.cap + ') ' + (p.staff ? 'is on STAFF for ' : 'has a spot (is SLOTTED) at ') +
    CONFIG.ACTIVITY_NAME + (p.staff || groupNamed_(p.ev) ? '' : ' (' + p.ev + ')') + ', ' + CONFIG.LOCATION +
    (CONFIG.DATES[p.ev] ? ', ' + CONFIG.DATES[p.ev] : '') + '.\n\n' +
    (!p.staff && STUDENT_NOTE ? STUDENT_NOTE + '\n\n' : '') +
    'STILL TO DO: as of the Reg Zone report of ' + asOfTxt_(asOf) + ', ' + (n === 1 ? 'this item is' : 'these ' + n + ' items are') +
    ' not done yet. Each one has a due date.\n\n' + numbered_(t.items, today) + '\n\n' +
    (t.done.length ? 'ALREADY DONE: ' + andList_(t.done) + '. Thank you!\n\n' : '') +
    (p.staff ? '' : 'IF PLANS CHANGE: if ' + who + ' decides not to attend, **please cancel the registration RIGHT AWAY.** Someone on the ' +
      'waitlist may be waiting for the spot, and the sooner you cancel, the more time they have to get ready. How to cancel:\n\n' + cancelSteps_() + '\n\n') +
    'Reg Zone reports can be a few days behind. If you already finished something on this list, thank you, and please ignore it. It drops off ' +
    'this list on its own once Reg Zone shows it done, so there is no need to tell us. You will get ' +
    (isPayOnly_(p, t) && payOnlyActive_() ? 'a payment reminder about every ' + PAY_ONLY.EVERY_DAYS + ' days until ' + niceIso_(PAY_ONLY.UNTIL) +
      ', then more often until it is paid.' : 'this list about twice a week until everything is done.') + '\n\n' + closing_(p) + '\n\n' + footer_();
  return { to: recipients_(p), subject: (late ? 'OVERDUE - ' : '') + 'ACTION NEEDED: ' + n + ' item' + (n === 1 ? '' : 's') + ' to finish for ' + who +
    ' (' + CONFIG.ACTIVITY_NAME + ')', body: body };
}

// ---------------------------------------------------------------- waitlist
function waitlistOrder_(people) {
  people.forEach(function (p) {
    if (p.slotted) return;
    p.blockers = todoItems_(p, {}, { approvalsOnly: true }).items;
  });
  CONFIG.GROUPS.forEach(function (g) {
    if (g.type !== 'students' || g.roomForEveryone) return;
    var alts = people.filter(function (p) { return p.ev === g.name && !p.slotted; });
    var cmp = function (a, b) {
      if (g.order === 'score') {
        if ((a.score === null) !== (b.score === null)) return a.score === null ? 1 : -1;
        if (a.score !== b.score) return b.score - a.score;
      }
      return a.regId - b.regId;
    };
    (g.lines === 'sex' ? ['M', 'F'] : ['']).forEach(function (sx) {
      var inLine = function (p) { return !sx || p.sex === sx; };
      var slottedN = people.filter(function (p) { return p.ev === g.name && p.slotted && inLine(p); }).length;
      var open = Math.max(1, Math.round(slottedN * WAITLIST.DROP_RATE));
      var tier = function (place) { return !place ? '' : place <= open ? 'good' : place <= 2 * open ? 'possible' : 'unlikely'; };
      var line = alts.filter(inLine).sort(cmp), ok = line.filter(function (p) { return !p.blockers.length; });
      ok.forEach(function (p, i) { p.place = i + 1; p.line = sx; p.chance = tier(p.place); });
      line.filter(function (p) { return p.blockers.length; }).forEach(function (p) {
        p.placeIfFixed = ok.filter(function (q) { return cmp(q, p) < 0; }).length + 1; p.line = sx; p.chance = tier(p.placeIfFixed);
      });
    });
  });
}
function lineName_(p) { return p.line === 'F' ? 'female' : p.line === 'M' ? 'male' : ''; }
function waitlistEmail_(p, asOf) {
  var who = who_(p), today = todayIso_(), g = groupOf_(p.ev), blocked = p.blockers.length > 0, ln = lineName_(p);
  var need = andList_(p.blockers.map(function (x) { return x.short || x.title.toLowerCase(); }));
  var amongTxt = 'among the ' + (ln ? ln + ' ' : '') + 'alternates who can be slotted right now';
  var lineWhy = ln ? 'Male and female spots are filled separately, because there are all-male and all-female flights.' +
    (ln === 'female' && WAITLIST.FEMALE_CAP_NOTE ? ' ' + WAITLIST.FEMALE_CAP_NOTE : '') : '';
  var parts = ['Hello,', NO_REPLY, who + ' (CAPID ' + p.cap + ') is an ALTERNATE (on the waitlist) for ' + CONFIG.ACTIVITY_NAME +
    (groupNamed_(p.ev) ? '' : ' (' + p.ev + ')') + '. An alternate does not have a spot yet, but can be slotted if a spot opens.'];
  var subj;
  if (g.roomForEveryone) {
    if (!blocked) { parts.push('GOOD NEWS: there is room for everyone, and nothing is holding ' + who + ' back. We will slot ' + who + ' soon.'); subj = 'Update for ' + who; }
    else {
      parts.push('GOOD NEWS: there is room for everyone. ' + who + ' will be slotted as soon as ' + (p.blockers.length === 1 ? 'this item is' : 'these items are') + ' done:');
      parts.push(numbered_(p.blockers, today));
      subj = who + ' cannot be slotted until ' + need + (p.blockers.length === 1 ? ' is' : ' are') + ' done';
    }
  } else if (!blocked) {
    parts.push('PLACE IN LINE: #' + p.place + ' ' + amongTxt + '.');
    parts.push('- ' + (lineWhy ? lineWhy + ' ' : '') + 'When a ' + (ln ? ln + ' ' : '') + 'spot opens, the next alternate in line who can be slotted gets it.\n' +
      '- The number can change. Another alternate who becomes eligible (for example, by getting a missing approval or a promotion) may move ' +
      'ahead. If an alternate ahead leaves the waitlist, ' + who + ' moves up.');
    subj = 'Waitlist update for ' + who + ': #' + p.place + ' in line';
  } else {
    parts.push('CANNOT BE SLOTTED YET: ' + who + ' cannot be slotted, even if a spot opens, until ' + (p.blockers.length === 1 ? 'this item is' : 'these items are') + ' done:');
    parts.push(numbered_(p.blockers, today));
    if (p.placeIfFixed) parts.push('PLACE IN LINE ONCE DONE: based on today\'s list, ' + who + ' would be #' + p.placeIfFixed + ' ' + amongTxt.replace(' right now', '') + '. ' + lineWhy);
    subj = 'Waitlist: ' + who + ' cannot be slotted, even if a spot opens, until ' + need + (p.blockers.length === 1 ? ' is' : ' are') + ' done';
  }
  if (!g.roomForEveryone) {
    var after = blocked ? ' once ' + (p.blockers.length === 1 ? 'that item is' : 'those items are') + ' done' : '';
    if (p.chance === 'good') parts.push('CHANCES: spots often open up before the start. ' + who + ' is near the front of the line, so there is a good chance of ' +
      'getting a spot' + after + '. We cannot promise one. Please keep ' + (CONFIG.DATES[p.ev] || 'the dates') + ' open.');
    else if (p.chance === 'possible') parts.push('CHANCES: spots often open up before the start, and one may open for ' + who + after + ', but it is not certain.' +
      (WAITLIST.OTHER_OPTIONS_TEXT ? ' ' + WAITLIST.OTHER_OPTIONS_TEXT : ''));
    else if (p.chance === 'unlikely') parts.push('CHANCES: we are sorry. ' + who + ' is far back in the line, so a spot is unlikely. ' + who + ' will stay on ' +
      'the waitlist, and we will contact you if a spot opens up for this cadet.' + (WAITLIST.OTHER_OPTIONS_TEXT ? ' ' + WAITLIST.OTHER_OPTIONS_TEXT : ''));
    if (WAITLIST.HOW_SPOTS_TEXT) parts.push('HOW SPOTS WERE GIVEN: ' + WAITLIST.HOW_SPOTS_TEXT);
    parts.push('IF A SPOT OPENS: we will email you. You will then have ' + WAITLIST.OFFER_HOURS + ' hours to accept or turn down the spot, and to finish any ' +
      'requirement whose due date has already passed. If we do not hear back within ' + WAITLIST.OFFER_HOURS + ' hours, the spot goes to the next alternate in line.');
  }
  parts.push('LEAVING THE WAITLIST: if ' + who + ' decides not to stay on the waitlist, please cancel the registration right away. This lets us offer ' +
    'the spot to the next person in line. How to cancel:', cancelSteps_());
  if (WAITLIST.GET_READY_TEXT) parts.push('GET READY NOW, IN CASE A SPOT OPENS: ' + WAITLIST.GET_READY_TEXT);
  parts.push('This update uses the Reg Zone report of ' + asOfTxt_(asOf) + '. You will get an update about once a week while ' + who + ' is on the waitlist.');
  parts.push(closing_(p), footer_());
  return { to: recipients_(p), subject: subj + ' (' + CONFIG.ACTIVITY_NAME + ')', body: parts.join('\n\n') };
}

// ---------------------------------------------------------------- weekly unit commander list (plain text + HTML tables)
var CC_COLS = [['unit', 'Unit approval'], ['parent', 'Parent approval'], ['parentLink', 'Link parent account'], ['parentEmail', 'Add parent email'],
  ['cab', 'C/Amn'], ['pay', 'Payment'], ['rst', 'RST'], ['cppt', 'CPPT'], ['med', 'Medical profile'], ['level1', 'Level I'], ['x', 'Other']];
function commanderWeekly_(g, asOf, allSet) {
  var p0 = g[0].p, today = todayIso_(), name = p0.ccName ? rankLast_(p0.ccName) : 'Commander';
  var unit = p0.unit ? (p0.wing || '') + (p0.wing ? '-' : '') + ('000' + p0.unit).slice(-3) : 'your unit';
  var by = function (a, b) { return a.p.last.localeCompare(b.p.last) || a.p.first.localeCompare(b.p.first); };
  var mine = g.filter(function (e) { return !e.p.unitOK && on_('unit', e.p); }).sort(by);
  var go = g.filter(function (e) { return e.p.slotted; }).sort(by);
  var alts = g.filter(function (e) { return !e.p.slotted; });
  var altKey = function (e) { var p = e.p; return [p.ev, p.line === 'F' ? 1 : 0, p.place ? 0 : 1, p.place || p.placeIfFixed || 999, p.last]; };
  alts.sort(function (a, b) { var x = altKey(a), y = altKey(b); for (var i = 0; i < x.length; i++) if (x[i] !== y[i]) return x[i] < y[i] ? -1 : 1; return 0; });
  var everyone = g.map(function (e) { return e.p; }).concat(allSet), hasSr = everyone.some(function (p) { return p.senior; }),
      hasCdt = everyone.some(function (p) { return !p.senior; });
  var members = hasSr && hasCdt ? 'your senior members, cadets, and their families' : hasSr ? 'your senior members' : 'your cadets and their families';
  var dueShort = function (it) {
    var late = it.due !== '0000-00-00' && today > it.due;
    return it.due === '0000-00-00' ? '' : ' (' + (late ? 'OVERDUE, was due ' : 'due ') + shortIso_(it.due) + ')';
  };
  var line = function (e) {
    return '- ' + who_(e.p) + ' (' + e.p.cap + '), ' + (e.p.staff ? 'staff' : e.p.ev) + ': ' + e.items.map(function (it) {
      return (it.k === 'unit' ? 'YOUR unit approval in Reg Zone' : it.short || it.title.toLowerCase()) + dueShort(it); }).join('; ');
  };
  var altLine = function (e) {
    var p = e.p, gr = groupOf_(p.ev), need = andList_(e.items.map(function (it) { return it.k === 'unit' ? 'your unit approval' : it.short; }));
    var s = '- ' + who_(p) + ' (' + p.cap + '), ' + p.ev + ': ';
    if (gr.roomForEveryone) return s + (e.items.length ? 'will be slotted once this is done: ' + need : 'will be slotted soon');
    if (e.items.length) return s + 'CANNOT be slotted yet. Needs: ' + need + (p.placeIfFixed ? ' (if done today, would be #' + p.placeIfFixed +
      (lineName_(p) ? ' in the ' + lineName_(p) + ' line' : '') + ')' : '');
    return s + '#' + p.place + (lineName_(p) ? ' in the ' + lineName_(p) + ' line' : ' in line') + (p.chance ? '; chances: ' + p.chance : '');
  };
  var parts = ['Hello ' + name + ',', 'This is an automated weekly list. It updates itself from Reg Zone, so there is no need to reply when something ' +
    'gets done. Please reply only if you have a question or a concern.',
    'Below is this week\'s list for ' + unit + ' (' + CONFIG.ACTIVITY_NAME + ', ' + CONFIG.LOCATION + '). It uses the Reg Zone report of ' + asOfTxt_(asOf) +
    '. Each member (and, for cadets, their parents) also gets a personal email with step-by-step help.' +
    (go.length || mine.length ? ' Please help them finish, especially anything marked OVERDUE.' : '')];
  if (mine.length) parts.push('WAITING ON YOUR UNIT APPROVAL IN REG ZONE (' + mine.length + '):\n' + mine.map(function (e) {
    return '- ' + who_(e.p) + ' (' + e.p.cap + '), ' + (e.p.staff ? 'staff' : e.p.ev + ', ' + e.p.status); }).join('\n'));
  if (go.length) parts.push('STAFF AND SLOTTED: STILL TO DO (' + go.length + '):\n' + go.map(line).join('\n') +
    (allSet.length ? '\n(The other ' + allSet.length + ' staff and slotted member' + (allSet.length === 1 ? ' has' : 's have') + ' everything done.)' : ''));
  else if (allSet.length) parts.push('STAFF AND SLOTTED: ALL SET. ' + (allSet.length === 1 ? 'The 1 member' : allSet.length === 2 ? 'Both members' :
    'All ' + allSet.length + ' members') + ' of ' + unit + ' on staff or slotted ' + (allSet.length === 1 ? 'has' : 'have') + ' everything done as of this report.');
  if (alts.length) parts.push('ALTERNATES (WAITLIST) (' + alts.length + '): registered, but no spot yet. A cadet marked CANNOT be slotted yet will not ' +
    'get a spot, even if one opens, until the item listed is done. Each family gets a weekly waitlist email with the same information. Please feel ' +
    'free to talk with these cadets and their parents about where they stand.\n' + alts.map(altLine).join('\n'));
  if (PAYMENT_NOTE && g.some(function (e) { return !e.p.senior && e.items.some(function (it) { return it.k === 'pay'; }); }))
    parts.push('PAYMENT: ' + PAYMENT_NOTE.replace(/\*\*/g, ''));
  parts.push('Reg Zone reports can be a few days behind. If some of these are already done, thank you.', fill_(CLOSINGS.commanders, null, members), footer_());
  var subj = [];
  if (go.length) subj.push(go.length + ' with items to finish'); else if (allSet.length) subj.push('staff and slotted all set');
  if (alts.length) subj.push(alts.length + ' on the waitlist');
  if (mine.length) subj.push(mine.length + ' waiting on your approval');
  return { to: [String(p0.ccEmail).trim()], cc: [], subject: (mine.length ? 'ACTION NEEDED: ' : '') + 'Weekly list for ' + unit + ': ' + subj.join(', ') +
    ' (' + CONFIG.ACTIVITY_NAME + ')', body: parts.join('\n\n'),
    html: commanderHtml_({ name: name, unit: unit, asOf: asOf, mine: mine, go: go, alts: alts, allSet: allSet, today: today, members: members,
      payNote: PAYMENT_NOTE && g.some(function (e) { return !e.p.senior && e.items.some(function (it) { return it.k === 'pay'; }); }) }) };
}
function commanderHtml_(c) {
  var esc = function (t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var TH = 'style="background:#1f3864;color:#fff;font-weight:bold;font-size:12px;padding:6px 8px;border:1px solid #c9ced6;text-align:left;vertical-align:bottom"';
  var TD = 'padding:5px 8px;border:1px solid #d5d9e0;font-size:13px;vertical-align:top';
  var H = function (t) { return '<h3 style="font-size:16px;color:#1f3864;margin:22px 0 8px 0">' + t + '</h3>'; };
  var P = function (t) { return '<p style="margin:0 0 12px 0">' + t + '</p>'; };
  var NOTE = function (t) { return '<p style="margin:0 0 8px 0;font-size:13px;color:#5b6472">' + t + '</p>'; };
  var who = function (p) { return '<b>' + esc(who_(p)) + '</b><br><span style="color:#5b6472;font-size:12px">' + esc(p.cap) + ' &middot; ' + esc(p.staff ? 'Staff' : p.ev) + '</span>'; };
  var colOf = function (it) { return /^x\d+$/.test(it.k) ? 'x' : it.k; };
  var cell = function (its) {
    var late = its.some(function (it) { return it.due !== '0000-00-00' && c.today > it.due; });
    var words = its.map(function (it) {
      var w = it.k === 'unit' ? 'YOUR approval' : it.k === 'cppt' ? (/renewal/.test(it.short) ? 'Renew' : /first-time/.test(it.short) ? 'First time (course + talk with you)' : 'Needed')
        : /^x\d+$/.test(it.k) ? it.short : 'Needed';
      return w;
    }).join(', ');
    if (late) return '<td style="' + TD + ';background:#fde2e1;color:#b42318;font-weight:bold">OVERDUE<br><span style="font-weight:normal">' + esc(words) + '</span></td>';
    return '<td style="' + TD + ';background:#fff4d6">' + esc(words) + '</td>';
  };
  var out = [P('Hello ' + esc(c.name) + ','), P('<span style="color:#5b6472">This is an automated weekly list. It updates itself from Reg Zone, so there is ' +
    'no need to reply when something gets done. Please reply only if you have a question or a concern.</span>'),
    P('Here is this week\'s list for <b>' + esc(c.unit) + '</b> (' + esc(CONFIG.ACTIVITY_NAME) + ', ' + esc(CONFIG.LOCATION) + '), from the Reg Zone report of ' +
      esc(asOfTxt_(c.asOf)) + '. Each member (and, for cadets, their parents) also gets a personal email with step-by-step help.' +
      (c.go.length || c.mine.length ? ' Please help them finish, especially anything marked <b style="color:#b42318">OVERDUE</b>.' : ''))];
  var sum = [];
  if (c.mine.length) sum.push('<b style="color:#b42318">' + c.mine.length + ' waiting on YOUR unit approval</b>');
  if (c.go.length) sum.push('<b>' + c.go.length + '</b> on staff or slotted with items to finish');
  if (c.allSet.length) sum.push('<b>' + c.allSet.length + '</b> on staff or slotted with everything done');
  if (c.alts.length) sum.push('<b>' + c.alts.length + '</b> on the waitlist');
  out.push('<table role="presentation" style="border-collapse:collapse;margin:0 0 6px 0;background:#f3f5f9;width:100%"><tr><td style="padding:10px 12px;font-size:14px">' +
    sum.join(' &nbsp;|&nbsp; ') + '</td></tr></table>');
  if (c.mine.length) {
    out.push(H('Waiting on your unit approval in Reg Zone (' + c.mine.length + ')'));
    out.push('<table style="border-collapse:collapse"><tr><th ' + TH + '>Member</th><th ' + TH + '>Status</th></tr>' + c.mine.map(function (e) {
      return '<tr><td style="' + TD + '">' + who(e.p) + '</td><td style="' + TD + ';background:#fde2e1;color:#b42318;font-weight:bold">Approve in Reg Zone</td></tr>'; }).join('') + '</table>');
  }
  if (c.go.length) {
    var used = CC_COLS.filter(function (col) { return c.go.some(function (e) { return e.items.some(function (it) { return colOf(it) === col[0]; }); }); });
    var head = function (col) {
      var ds = {}; c.go.forEach(function (e) { e.items.forEach(function (it) { if (colOf(it) === col[0] && it.due !== '0000-00-00') ds[shortIso_(it.due)] = 1; }); });
      var k = Object.keys(ds); return col[1] + (k.length === 1 && col[0] !== 'x' ? ' (' + k[0] + ')' : '');
    };
    out.push(H('Staff and slotted: still to do (' + c.go.length + ')'));
    out.push(NOTE('One row per member. A blank cell means nothing is needed for that item. The first row shows how many still need each item.'));
    var totals = '<tr><td style="' + TD + ';background:#eef1f6;font-weight:bold">How many still need it</td>' + used.map(function (col) {
      var n = c.go.filter(function (e) { return e.items.some(function (it) { return colOf(it) === col[0]; }); }).length;
      return '<td style="' + TD + ';background:#eef1f6;font-weight:bold;text-align:center">' + n + '</td>'; }).join('') + '</tr>';
    out.push('<table style="border-collapse:collapse"><tr><th ' + TH + '>Member</th>' + used.map(function (col) { return '<th ' + TH + '>' + esc(head(col)) + '</th>'; }).join('') +
      '</tr>' + totals + c.go.map(function (e) {
        return '<tr><td style="' + TD + '">' + who(e.p) + '</td>' + used.map(function (col) {
          var its = e.items.filter(function (it) { return colOf(it) === col[0]; });
          return its.length ? cell(its) : '<td style="' + TD + '"></td>';
        }).join('') + '</tr>';
      }).join('') + '</table>');
    if (c.allSet.length) out.push('<p style="margin:8px 0 0 0;font-size:13px;color:#17795a"><b>&#10003; The other ' + c.allSet.length + ' staff and slotted member' +
      (c.allSet.length === 1 ? ' has' : 's have') + ' everything done.</b></p>');
  } else if (c.allSet.length) {
    out.push(H('Staff and slotted'));
    out.push('<p style="margin:0 0 12px 0;padding:10px 12px;background:#e6f4ea;color:#17795a;font-size:14px"><b>&#10003; ALL SET.</b> ' +
      (c.allSet.length === 1 ? 'The 1 member' : c.allSet.length === 2 ? 'Both members' : 'All ' + c.allSet.length + ' members') + ' of ' + esc(c.unit) +
      ' on staff or slotted ' + (c.allSet.length === 1 ? 'has' : 'have') + ' everything done as of this report.</p>');
  }
  if (c.alts.length) {
    var anyLine = c.alts.some(function (e) { return !groupOf_(e.p.ev).roomForEveryone; }), anyRoom = c.alts.some(function (e) { return groupOf_(e.p.ev).roomForEveryone; });
    out.push(H('Alternates on the waitlist (' + c.alts.length + ')'));
    out.push(NOTE('Registered, but no spot yet. ' + (anyLine ? '"#3" means third among the alternates in that line who can be slotted right now. A cadet who ' +
      '<b>cannot be slotted yet</b> will not get a spot, even if one opens, until the item listed is done. ' : '') +
      (anyRoom ? 'Where there is room for everyone, a cadet is slotted as soon as the item listed is done. ' : '') +
      'Each family gets a weekly waitlist email with the same information. Please feel free to talk with these cadets and their parents about where they stand.'));
    out.push('<table style="border-collapse:collapse"><tr><th ' + TH + '>Cadet</th><th ' + TH + '>Line</th><th ' + TH + '>Place</th><th ' + TH +
      '>Chances</th><th ' + TH + '>Needs before a spot can be given</th></tr>' + c.alts.map(function (e) {
        var p = e.p, room = groupOf_(p.ev).roomForEveryone, need = e.items.map(function (it) { return it.k === 'unit' ? 'your unit approval' : it.short; }).join(', ');
        var line = room ? 'Room for everyone' : p.line === 'F' ? 'Female' : p.line === 'M' ? 'Male' : 'Waitlist';
        var place = room ? '' : p.place ? '<b>#' + p.place + '</b>' : 'Not yet' + (p.placeIfFixed ? '<br><span style="font-size:12px;color:#5b6472">(would be #' +
          p.placeIfFixed + ' if done today)</span>' : '');
        var ch = room ? '' : p.chance === 'good' ? '<span style="color:#17795a;font-weight:bold">Good</span>' : p.chance === 'possible' ? 'Possible' :
          p.chance === 'unlikely' ? '<span style="color:#5b6472">Unlikely</span>' : '';
        return '<tr><td style="' + TD + '">' + who(p) + '</td><td style="' + TD + '">' + line + '</td><td style="' + TD + '">' + place + '</td><td style="' + TD + '">' + ch +
          '</td><td style="' + TD + (need ? ';background:#fde2e1;color:#b42318;font-weight:bold' : '') + '">' + (need ? esc(need.charAt(0).toUpperCase() + need.slice(1)) :
          (room ? 'Nothing: will be slotted' : 'Nothing: waiting for a spot')) + '</td></tr>';
      }).join('') + '</table>');
  }
  if (c.payNote) out.push(H('Payment') + P(esc(PAYMENT_NOTE.replace(/\*\*/g, ''))));
  out.push('<div style="height:14px"></div>' + P('<span style="color:#5b6472;font-size:13px">Reg Zone reports can be a few days behind. If some of these are already done, thank you.</span>'));
  out.push(P(esc(fill_(CLOSINGS.commanders, null, c.members)).replace(/^THANK YOU/, '<b>THANK YOU</b>')));
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.45;color:#1b1f27;max-width:900px">' + out.join('') +
    textToHtml_(capAfterColon_(footer_())).replace(/^<div[^>]*>/, '').replace(/<\/div>$/, '') + '</div>';
}

// ---------------------------------------------------------------- plan: who gets what, written to the tab
function plan_(mode) {
  var d = readData_(), all = props_().getProperties(), now = Date.now(), rows = [];
  waitlistOrder_(d.people);
  var gapH = { todo: CONFIG.TODO_EVERY_DAYS * 24 - 3, wait: CONFIG.WEEKLY_EVERY_DAYS * 24 - 3, cc: CONFIG.WEEKLY_EVERY_DAYS * 24 - 3 };
  gapH.todoPay = payOnlyActive_() ? PAY_ONLY.EVERY_DAYS * 24 - 3 : gapH.todo;
  function hist(key) { try { return JSON.parse(all['EM_' + key] || 'null'); } catch (e) { return null; } }
  function due(kind, key, to) {
    if (!to.length) return 'No (no email address)';
    var h = hist(key), lim = mode === 'manual' ? CONFIG.MANUAL_MIN_HOURS : gapH[kind];
    if (h && (now - new Date(h.last).getTime()) / 3600000 < lim) return 'No (emailed ' + Utilities.formatDate(new Date(h.last), tz_(), 'EEE d MMM h:mm a') + ')';
    return 'Yes';
  }
  var seen = {};
  d.people.forEach(function (p) {
    var key = p.ev + ':' + p.cap;
    if (seen[key]) return; seen[key] = 1;
    if (ADJUST.NO_EMAIL[p.cap]) { rows.push({ kind: 'done', p: p, key: key, items: 'Do-not-email: ' + ADJUST.NO_EMAIL[p.cap], to: [], send: 'No' }); return; }
    var ccIt = !p.unitOK && on_('unit', p) && /@/.test(String(p.ccEmail || '')) ? String(p.ccEmail).trim() : '';
    var addCC = function (m) { m.cc = ccIt && m.to.map(function (a) { return a.toLowerCase(); }).indexOf(ccIt.toLowerCase()) < 0 ? [ccIt] : []; return m; };
    if (p.slotted) {
      var t = todoItems_(p, d.extra, { cc: !!ccIt });
      p.todoItems = t.items;
      if (!t.items.length) { rows.push({ kind: 'done', p: p, key: key, items: '', to: [], send: 'No (nothing missing)' }); return; }
      var m = addCC(todoEmail_(p, t, d.asOf)), payOnly = isPayOnly_(p, t) && payOnlyActive_();
      rows.push({ kind: 'todo', p: p, key: 'T:' + key, m: m, items: t.items.map(function (x) { return x.short || x.k; }).join(', ') +
        (payOnly ? ' (payment only: every ' + PAY_ONLY.EVERY_DAYS + ' days)' : ''), to: m.to, send: due(payOnly ? 'todoPay' : 'todo', 'T:' + key, m.to) });
    } else if (CONFIG.WAITLIST_EMAILS) {
      var w = addCC(waitlistEmail_(p, d.asOf));
      var info = groupOf_(p.ev).roomForEveryone ? (p.blockers.length ? 'Needs ' + p.blockers.map(function (x) { return x.short; }).join(', ') : 'Ready to slot')
        : p.blockers.length ? 'Cannot be slotted yet: ' + p.blockers.map(function (x) { return x.short; }).join(', ') + (p.placeIfFixed ? ' (would be #' + p.placeIfFixed + ')' : '')
        : '#' + p.place + (p.line ? ' (' + p.line + ' line)' : '');
      if (p.chance) info += ' | chances: ' + p.chance;
      rows.push({ kind: 'wait', p: p, key: 'W:' + key, m: w, items: info, to: w.to, send: due('wait', 'W:' + key, w.to) });
    }
  });
  if (CONFIG.COMMANDER_LISTS) {
    var byCC = {}, doneBy = {};
    d.people.forEach(function (p) {
      if (ADJUST.NO_EMAIL[p.cap]) return;
      var k = String(p.ccEmail || '').trim().toLowerCase();
      if (k.indexOf('@') < 0) return;
      var list = p.slotted ? (p.todoItems || []) : (p.blockers || []);
      if (p.slotted && !list.length) { (doneBy[k] = doneBy[k] || []).push(p); return; }
      (byCC[k] = byCC[k] || []).push({ p: p, items: list });
    });
    Object.keys(byCC).forEach(function (k) {
      var g = byCC[k], m = commanderWeekly_(g, d.asOf, doneBy[k] || []);
      rows.push({ kind: 'cc', p: { ev: 'Commander', cap: '', rank: '', first: g[0].p.ccName, last: '' }, key: 'C:' + k, m: m,
        items: g.length + ' member' + (g.length === 1 ? '' : 's') + ' listed', to: m.to, send: due('cc', 'C:' + k, m.to) });
    });
  }
  var ss = ss_(), sh = ss.getSheetByName(CONFIG.TAB) || ss.insertSheet(CONFIG.TAB);
  sh.clearContents();
  var label = { todo: 'To-do', wait: 'Waitlist', cc: 'Commander list', done: 'All done' }, order = { todo: 0, wait: 1, cc: 2, done: 3 };
  var head = ['Email', 'Group', 'CAPID', 'Name', 'Reg Zone status', 'Still to do / place in line', 'To', 'Times emailed', 'Last emailed',
    mode === 'manual' ? 'Send now?' : 'Send on next automatic run?'];
  rows.sort(function (a, b) { return order[a.kind] - order[b.kind] || String(a.p.ev).localeCompare(String(b.p.ev)) || String(a.p.last).localeCompare(String(b.p.last)); });
  var out = rows.map(function (x) {
    var h = hist(x.key);
    return [label[x.kind], x.p.ev, x.p.cap, who_(x.p), x.p.status || '', x.items, (x.to || []).join(', ') + (x.m && x.m.cc && x.m.cc.length ? '  CC: ' + x.m.cc.join(', ') : '') +
      (x.p.parentSrc ? '  (parent email from ' + x.p.parentSrc + ')' : x.p.noParentEmail ? '  (NO parent email in Reg Zone)' : ''), h ? h.n : 0,
      h ? Utilities.formatDate(new Date(h.last), tz_(), 'yyyy-MM-dd HH:mm') : '', x.send];
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (out.length) sh.getRange(2, 1, out.length, head.length).setNumberFormat('@').setValues(out);
  sh.getRange(out.length + 3, 1).setValue('Reg Zone data as of ' + (d.asOf ? Utilities.formatDate(d.asOf, tz_(), 'EEE d MMM yyyy h:mm a') : '?') +
    ' (the oldest export in the Drive folder).' + (d.missing.length ? ' MISSING FILES: ' + d.missing.join(', ') + '.' : '') +
    ' Built ' + Utilities.formatDate(new Date(), tz_(), 'EEE d MMM h:mm a') + '.');
  sh.setFrozenRows(1);
  return { rows: rows, asOf: d.asOf, missing: d.missing };
}

// ---------------------------------------------------------------- sending
function counts_(r, kinds, ev) {
  var c = { todo: 0, wait: 0, cc: 0 };
  r.rows.forEach(function (x) { if (c.hasOwnProperty(x.kind) && kinds[x.kind] && x.send === 'Yes' && (!ev || x.p.ev === ev)) c[x.kind]++; });
  return c;
}
var RUN_SECONDS = 240;   // Google stops scripts at 6 minutes: send for 4 minutes, then continue 1 minute later
function send_(mode, kinds, ev) {
  var start = Date.now(), r = plan_(mode), sent = { todo: 0, wait: 0, cc: 0 }, queued = 0, left = 0;
  var go = r.rows.filter(function (x) { return kinds[x.kind] && x.send === 'Yes' && (!ev || x.p.ev === ev); });
  var nowIso = new Date().toISOString(), P = props_();
  for (var i = 0; i < go.length; i++) {
    if ((Date.now() - start) / 1000 > RUN_SECONDS) { left = go.length - i; break; }
    var x = go[i], opts = { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME };
    if (x.m.cc && x.m.cc.length) opts.cc = x.m.cc.join(',');
    if (x.m.html) opts.htmlBody = x.m.html;
    var lowQuota = MailApp.getRemainingDailyQuota() - x.m.to.length - (x.m.cc ? x.m.cc.length : 0) <= CONFIG.EMAIL_RESERVE;
    if (!safeSend_(x.m.to.join(','), x.m.subject, x.m.body, opts, lowQuota)) queued++;
    var h = null; try { h = JSON.parse(P.getProperty('EM_' + x.key) || 'null'); } catch (e) {}
    P.setProperty('EM_' + x.key, JSON.stringify({ last: nowIso, n: (h ? h.n : 0) + 1 }));
    sent[x.kind]++;
  }
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'continueSending') ScriptApp.deleteTrigger(t); });
  if (left) {
    P.setProperty('EM_CONTINUE', JSON.stringify({ mode: mode, kinds: kinds, ev: ev || '' }));
    ScriptApp.newTrigger('continueSending').timeBased().after(60 * 1000).create();
  } else P.deleteProperty('EM_CONTINUE');
  plan_(mode);   // refresh the tab with the new "last emailed" times
  return { sent: sent, queued: queued, left: left, r: r };
}
function continueSending() {
  if (!isAutomationOwner_()) return;
  var c = null; try { c = JSON.parse(props_().getProperty('EM_CONTINUE') || 'null'); } catch (e) {}
  if (!c) return;
  var res = send_(c.mode, c.kinds, c.ev || '');
  if (!res.left) sendMail_(CONFIG.CONTACT_EMAIL, 'Reminder emails: finished sending', summary_(res));
}
function summary_(res) {
  var noEmail = res.r.rows.filter(function (x) { return x.send === 'No (no email address)'; }).map(function (x) { return who_(x.p) + ' (' + x.p.ev + ')'; });
  var noParent = res.r.rows.filter(function (x) { return x.p.noParentEmail && x.kind !== 'cc'; }).map(function (x) { return who_(x.p) + ' (' + x.p.ev + ')'; });
  return 'Sent in this batch: ' + res.sent.todo + ' to-do lists, ' + res.sent.wait + ' waitlist updates, ' + res.sent.cc + ' unit commander lists.' +
    (res.queued ? '\n' + res.queued + ' are waiting on the Email queue tab (daily email limit) and will go out automatically.' : '') +
    (res.left ? '\n' + res.left + ' more will be sent in about 1 minute (Google time limit).' : '') +
    (noEmail.length ? '\n\nNO EMAIL ADDRESS in Reg Zone (not emailed): ' + noEmail.join(', ') : '') +
    (noParent.length ? '\n\nCadets under 18 with NO parent email in Reg Zone (only the cadet was emailed): ' + noParent.join(', ') : '') +
    '\n\nReg Zone data as of ' + asOfTxt_(res.r.asOf) + '. Details: "' + CONFIG.TAB + '" tab.';
}

// Every email goes out as formatted HTML plus plain text. "**text**" = bold red in HTML (stars removed from the plain text).
function sendMail_(to, subject, body, opts) {
  body = capAfterColon_(body);
  var o = {}; Object.keys(opts || {}).forEach(function (k) { o[k] = opts[k]; });
  if (!o.htmlBody) o.htmlBody = textToHtml_(body);
  MailApp.sendEmail(to, subject, String(body).replace(/\*\*/g, ''), o);
}
function safeSend_(to, subject, body, opts, forceQueue) {
  try {
    if (!forceQueue && MailApp.getRemainingDailyQuota() > 0) { sendMail_(to, subject, body, opts || {}); logEmail_(to, opts, subject, 'Sent'); return true; }
  } catch (err) { console.log('send failed, queuing: ' + err); }
  var sh = ss_().getSheetByName('Email queue') || ss_().insertSheet('Email queue');
  if (sh.getLastRow() === 0) sh.appendRow(['Queued', 'To', 'Subject', 'Body', 'Options', 'Sent']);
  sh.appendRow([new Date(), to, subject, body, JSON.stringify(opts || {}), '']);
  ensureRetryTrigger_();
  logEmail_(to, opts, subject, 'Queued (daily limit)');
  return false;
}
function logEmail_(to, opts, subject, status) {
  try {
    var sh = ss_().getSheetByName('Email log');
    if (!sh) { sh = ss_().insertSheet('Email log'); sh.appendRow(['Time', 'To', 'CC', 'Subject', 'Status']); sh.getRange(1, 1, 1, 5).setFontWeight('bold'); }
    sh.appendRow([Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), to, (opts && opts.cc) || '', subject, status]);
  } catch (e) { console.log('email log failed: ' + e); }
}
function ensureRetryTrigger_() {
  if (!ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'sendQueuedEmails'; }))
    ScriptApp.newTrigger('sendQueuedEmails').timeBased().everyHours(1).create();
}
function sendQueuedEmails() {
  if (!isAutomationOwner_()) return;
  var sh = ss_().getSheetByName('Email queue');
  if (!sh || sh.getLastRow() < 2) return;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < v.length; i++) {
    if (v[i][5]) continue;
    if (MailApp.getRemainingDailyQuota() < 1) break;
    try {
      var o = JSON.parse(v[i][4] || '{}');
      sendMail_(v[i][1], v[i][2], v[i][3], o);
      sh.getRange(i + 2, 6).setValue(new Date());
      logEmail_(v[i][1], o, v[i][2], 'Sent from queue');
    } catch (err) { console.log('queued send failed: ' + err); break; }
  }
}
function textToHtml_(body) {
  var esc = function (t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var link = function (t) {
    return t.replace(/(https?:\/\/[^\s<]+)/g, function (u) {
      var trail = (u.match(/[.,);]+$/) || [''])[0], url = u.slice(0, u.length - trail.length);
      var shown = url.length <= 60 ? url : 'Click here to open the link';
      return '<a href="' + url + '" style="color:#1a56c4;' + (shown !== url ? 'font-weight:bold' : 'word-break:break-all') + '">' + shown + '</a>' + trail;
    });
  };
  var label = function (t) {   // "STILL TO DO:", "PLACE IN LINE:" -> bold
    t = t.replace(/^((?:[A-Z0-9][A-Z0-9 ,'()\/&-]{2,}?)(?::|(?= -)))/, '<b>$1</b>');
    if (t.indexOf('<b>') !== 0) t = t.replace(/^([A-Z][A-Z]+(?: [A-Z]+)*\.?)(?=[ .]|$)/, function (m0) { return m0.replace(/[^A-Z]/g, '').length >= 5 && m0 !== 'CAPID' ? '<b>' + m0 + '</b>' : m0; });
    return t.replace(/\*\*(.+?)\*\*/g, '<b style="color:#b42318">$1</b>');
  };
  var html = String(body || '').replace(/\r/g, '').split(/\n\s*\n/).map(function (b) {
    var lines = b.split('\n');
    if (lines.every(function (l) { return /^\s*- /.test(l); })) return '<ul style="margin:0 0 14px 0;padding-left:22px">' + lines.map(function (l) {
      return '<li style="margin:0 0 4px 0">' + label(link(esc(l.replace(/^\s*- /, '')))) + '</li>'; }).join('') + '</ul>';
    var m = /^(\d+)\. ([\s\S]*)$/.exec(b);
    if (m) return '<p style="margin:0 0 14px 0;padding-left:26px;text-indent:-26px"><b style="display:inline-block;width:26px;text-indent:0">' + m[1] + '.</b>' +
      lines.map(function (l, i) {
        if (i) return label(link(esc(l)));
        var t = esc(l.replace(/^\d+\. /, '')), tm = /^(.+?\.) ((?:Due|OVERDUE)\b[\s\S]*)$/.exec(t);
        return tm ? '<b>' + tm[1] + '</b> ' + (/^OVERDUE/.test(tm[2]) ? '<b style="color:#b42318">' + tm[2] + '</b>' : tm[2]) : label(link(t));
      }).join('<br>') + '</p>';
    return '<p style="margin:0 0 14px 0">' + lines.map(function (l) { return label(link(esc(l))); }).join('<br>') + '</p>';
  }).join('');
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.5;color:#1b1f27;max-width:640px">' + html + '</div>';
}
// "CHANCES: we are sorry" -> "CHANCES: We are sorry" (capital after a label's colon; web addresses and emails left alone)
function capAfterColon_(body) {
  return String(body || '').replace(/([A-Za-z)]):([ \t]+)(?!https?:|www\.)([a-z])(?![a-z0-9._-]*@)/g, function (m0, a, sp, ch) { return a + ':' + sp + ch.toUpperCase(); });
}

// ---------------------------------------------------------------- menu actions
function previewReminders() {
  var ui = SpreadsheetApp.getUi(), r = plan_('auto');
  try { ss_().getSheetByName(CONFIG.TAB).activate(); } catch (e) {}
  var c = counts_(r, { todo: 1, wait: 1, cc: 1 }), all = { todo: 0, wait: 0, done: 0, cc: 0 };
  r.rows.forEach(function (x) { if (all.hasOwnProperty(x.kind)) all[x.kind]++; });
  ui.alert('Reminder emails (preview, nothing sent)', (r.missing.length ? 'WARNING: missing in the Reg Zone folder: ' + r.missing.join(', ') + '\n\n' : '') +
    'To-do lists (staff + slotted with something left): ' + all.todo + ' (' + c.todo + ' due on the next automatic run)\n' +
    'Nothing left (not emailed): ' + all.done + '\nWaitlist updates: ' + all.wait + ' (' + c.wait + ' due)\nUnit commander lists: ' + all.cc + ' (' + c.cc + ' due)\n\n' +
    'Reg Zone data as of ' + asOfTxt_(r.asOf) + '. Every person and email is on the "' + CONFIG.TAB + '" tab.', ui.ButtonSet.OK);
}
function testReminders() {
  var ui = SpreadsheetApp.getUi(), r = plan_('manual');
  var pick = function (kind, fn) {
    var l = r.rows.filter(function (x) { return x.kind === kind && (!fn || fn(x)); });
    var n = function (x) { return ((x.kind === 'todo' ? x.p.todoItems : x.p.blockers) || []).length * 100000 + (x.m ? x.m.body.length : 0); };
    return l.sort(function (a, b) { return n(b) - n(a); })[0];
  };
  var samples = [
    [pick('todo', function (x) { return x.p.staff; }), 'STAFF TO-DO LIST: everything a staff member still owes, with due dates. Goes to the member (cadets under ' +
      '18: also their parents), plus the unit commander only while unit approval is missing.'],
    [pick('todo', function (x) { return !x.p.staff; }), 'STUDENT TO-DO LIST (slotted only): same idea, plus how to cancel if plans change.'],
    [pick('wait', function (x) { return x.p.place; }), 'WAITLIST UPDATE (alternate who can be slotted): place in line, chances, what happens if a spot opens.'],
    [pick('wait', function (x) { return x.p.blockers && x.p.blockers.length; }), 'WAITLIST UPDATE (alternate who cannot be slotted yet): what blocks them, and ' +
      'where they would be in line once it is fixed.'],
    [pick('cc'), 'UNIT COMMANDER LIST (weekly): a table of who in the unit still owes what, plus every alternate and their place in line.']
  ].filter(function (s) { return s[0]; });
  var cc = (CONFIG.TEST_CC || []).join(',');
  samples.forEach(function (s, i) {
    var x = s[0], intro = 'TEST EMAIL ' + (i + 1) + ' OF ' + samples.length + ': for review only. Nobody else was emailed.\n\nWHAT THIS TESTS: ' + s[1] +
      '\n\nThis sample uses real Reg Zone data.\n\n- - - - - - - - - - (the email starts below) - - - - - - - - - -';
    var o = { name: CONFIG.CONTACT_NAME };
    if (cc) o.cc = cc;
    if (x.m.html) o.htmlBody = textToHtml_(capAfterColon_(intro)) + x.m.html;
    sendMail_(CONFIG.CONTACT_EMAIL, '[TEST ' + (i + 1) + ' of ' + samples.length + '] ' + x.m.subject, intro + '\n\n' + x.m.body, o);
  });
  ui.alert('Sent ' + samples.length + ' test emails to ' + CONFIG.CONTACT_EMAIL + (cc ? ', copied to ' + cc : '') + '. Nobody else was emailed.');
}
function manualSend_(kinds, what, ev) {
  var ui = SpreadsheetApp.getUi(), r = plan_('manual');
  if (r.missing.length) { ui.alert('Not sent. Missing in the Reg Zone folder: ' + r.missing.join(', ')); return; }
  var c = counts_(r, kinds, ev), ageH = r.asOf ? (new Date() - r.asOf) / 3600000 : 999;
  if (!c.todo && !c.wait && !c.cc) { ui.alert('Nobody to email right now (see the "' + CONFIG.TAB + '" tab).'); return; }
  if (ui.alert('Send ' + what, 'Email now:\n- ' + c.todo + ' to-do lists\n- ' + c.wait + ' waitlist updates\n- ' + c.cc + ' unit commander lists\n\n' +
      'Anyone emailed in the last ' + CONFIG.MANUAL_MIN_HOURS + ' hours is skipped.' +
      (ageH > CONFIG.MAX_DATA_AGE_HOURS ? '\n\nWARNING: the Reg Zone data is ' + Math.round(ageH) + ' hours old. People who already finished may be emailed.' : '') +
      '\n\nSend now?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert(summary_(send_('manual', kinds, ev)));
}
function sendToEveryoneNow() { manualSend_({ todo: 1, wait: 1, cc: 1 }, 'to EVERYONE'); }
function sendTodoNow() { manualSend_({ todo: 1 }, 'to-do lists'); }
function sendWaitlistNow() { manualSend_({ wait: 1 }, 'waitlist updates'); }
function sendCommanderNow() { manualSend_({ cc: 1 }, 'unit commander lists'); }
function sendOneGroupNow() {
  var ui = SpreadsheetApp.getUi(), names = CONFIG.GROUPS.map(function (g) { return g.name; });
  var r = ui.prompt('Send to one group', 'Type one of: ' + names.join(', '), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var ev = names.filter(function (n) { return n.toLowerCase() === r.getResponseText().trim().toLowerCase(); })[0];
  if (!ev) { ui.alert('No group called "' + r.getResponseText() + '".'); return; }
  manualSend_({ todo: 1, wait: 1 }, 'to ' + ev + ' only (to-do lists + waitlist updates)', ev);
}

// ---------------------------------------------------------------- automatic runs
function autoReminders() {
  if (!isAutomationOwner_()) return;
  var d;
  try { d = readData_(); } catch (e) { sendMail_(CONFIG.CONTACT_EMAIL, 'Reminder emails NOT sent', 'Could not read the Reg Zone folder: ' + e); return; }
  var ageH = d.asOf ? (new Date() - d.asOf) / 3600000 : 999;
  if (d.missing.length || ageH > CONFIG.MAX_DATA_AGE_HOURS) {
    sendMail_(CONFIG.CONTACT_EMAIL, 'Reminder emails NOT sent: Reg Zone data is ' + (d.missing.length ? 'missing' : Math.round(ageH) + ' hours old'),
      'The automatic reminder emails did not go out.\n\n' + (d.missing.length ? 'Missing in the Reg Zone folder: ' + d.missing.join(', ') + '\n\n' : '') +
      'Oldest export in Drive: ' + asOfTxt_(d.asOf) + '. Save fresh Reg Zone exports to the Drive folder, then use Reg Zone reminders > Send to EVERYONE now, ' +
      'or wait for the next automatic run.');
    return;
  }
  var res = send_('auto', { todo: 1, wait: 1, cc: 1 });
  if (!res.left) sendMail_(CONFIG.CONTACT_EMAIL, 'Reminder emails: ' + (res.sent.todo + res.sent.wait + res.sent.cc) + ' sent', summary_(res));
}
function turnOnAutomatic() {
  var ui = SpreadsheetApp.getUi();
  if (!isAutomationOwner_()) { ui.alert('Automation runs from ' + automationOwner_() + '. Sign in as that account, or use "Run automation from THIS account".'); return; }
  turnOffAutomatic(true);
  CONFIG.RUN_DAYS.forEach(function (d) { ScriptApp.newTrigger('autoReminders').timeBased().onWeekDay(ScriptApp.WeekDay[d]).atHour(CONFIG.HOUR).create(); });
  props_().setProperty('AUTO', 'on');
  ui.alert('Automatic reminder emails are ON (' + CONFIG.RUN_DAYS.join(' and ') + ', about ' + CONFIG.HOUR + ':00).\n- To-do lists: never twice within ' +
    CONFIG.TODO_EVERY_DAYS + ' days.\n- Waitlist updates and commander lists: about once a week.\n\nEmails stop and you get a warning if the exports in ' +
    'Drive are more than ' + CONFIG.MAX_DATA_AGE_HOURS + ' hours old. You get a summary email after each run.');
}
function turnOffAutomatic(quiet) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (['autoReminders', 'continueSending'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  if (quiet !== true) { props_().deleteProperty('AUTO'); SpreadsheetApp.getUi().alert('Automatic reminder emails are OFF. You can still send from the menu.'); }
}
function showStatus() {
  var t = ScriptApp.getProjectTriggers().filter(function (x) { return x.getHandlerFunction() === 'autoReminders'; }).length;
  SpreadsheetApp.getUi().alert('Automatic emails: ' + (t ? 'ON (' + t + ' runs a week)' : 'OFF') + '\nRuns as: ' + (automationOwner_() || '(not set)') +
    '\nYou are: ' + Session.getEffectiveUser().getEmail() + (isAutomationOwner_() ? '' : '\n\nThis account does not run the automation.'));
}
// Triggers run as the account that made them, and emails come from that account. Run this as the account that should send.
function automationOwner_() { return props_().getProperty('AUTOMATION_OWNER') || ''; }
function isAutomationOwner_() {
  var owner = automationOwner_(), me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) {}
  return !owner || !me || owner.toLowerCase() === me.toLowerCase();
}
function takeOverAutomation() {
  var ui = SpreadsheetApp.getUi(), me = Session.getEffectiveUser().getEmail(), wasOn = props_().getProperty('AUTO') === 'on';
  try { DriveApp.getFolderById(CONFIG.REGZONE_FOLDER_ID).getName(); } catch (e) { ui.alert('This account cannot open the Reg Zone folder. Share it with ' + me + ' first.'); return; }
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  props_().setProperty('AUTOMATION_OWNER', me);
  ensureRetryTrigger_();
  if (wasOn) CONFIG.RUN_DAYS.forEach(function (d) { ScriptApp.newTrigger('autoReminders').timeBased().onWeekDay(ScriptApp.WeekDay[d]).atHour(CONFIG.HOUR).create(); });
  ui.alert('Automation now runs from ' + me + '. Automatic emails: ' + (wasOn ? 'ON' : 'OFF') + '.');
}
