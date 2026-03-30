/* ================= CONFIG ================= */
const SHEET_NAMES = {
  PERMISSIONS: 'PERMISSIONS',
  LABS: 'LABS'
};

const RANGE_SETTINGS_START_ROW = 2;
const RANGE_SETTINGS_END_ROW = 50;

function testEmail(e){
  return JSON.stringify(e);
}

/* ================= ENTRY ================= */
function doGet(e) {
  const email = resolveEmail_(e);
  const user = getUserDataByEmail_(email);

  if (!email) {
    return HtmlService.createHtmlOutput(`
      <div style="padding:30px;text-align:center;font-family:sans-serif">
        يرجى تسجيل الدخول بحساب Google
      </div>
    `);
  }

  if (!user.authorized) {
    const t = HtmlService.createTemplateFromFile('Unauthorized');
    t.email = email;
    return t.evaluate().setTitle('غير مصرح');
  }

  const t = HtmlService.createTemplateFromFile('Index');
  t.initialEmail = email;
  t.authToken = createAuthSession_(user);

  return t.evaluate().setTitle('IPA Dashboard');
}

/* ================= INCLUDE ================= */
function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/* ================= CONTEXT / AUTH ================= */
function resolveRequestContext_(e) {
  const email = resolveEmail_(e);
  const user = getUserDataByEmail_(email);

  return {
    email: email || '',
    authorized: !!(user && user.authorized),
    user: user
  };
}

function resolveEmail_(e) {
  try {
    const email = Session.getActiveUser().getEmail();
    return (email || '').trim().toLowerCase();
  } catch (err) {
    return '';
  }
}

function resolveAuthEmail_(email) {
  const passedEmail = String(email || '').trim().toLowerCase();
  return passedEmail || resolveEmail_();
}

function createAuthSession_(user) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('auth:' + token, JSON.stringify(user), 21600);
  return token;
}

function getAuthorizedUser_(email, authToken) {
  const token = String(authToken || '').trim();

  if (token) {
    const cachedUser = CacheService.getScriptCache().get('auth:' + token);
    if (cachedUser) {
      const user = JSON.parse(cachedUser);
      if (user && user.authorized) {
        return user;
      }
    }
  }

  const authEmail = resolveAuthEmail_(email);
  const user = getUserDataByEmail_(authEmail);
  if (!user.authorized) {
    throw new Error('غير مصرح');
  }
  return user;
}

function normalizeText_(value) {
  return String(value || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeComparable_(value) {
  return normalizeText_(value).toLowerCase();
}

function isAllValue_(value) {
  return normalizeComparable_(value) === 'all';
}

function isNoneValue_(value) {
  const normalized = normalizeComparable_(value);
  return normalized === '' || normalized === 'none' || normalized === '-';
}

function parseAllowedValues_(value) {
  return normalizeText_(value)
    .split(/[\n\r,،;]+/)
    .map(x => x.trim())
    .filter(Boolean);
}

function matchesAllowedLab_(labName, allowedValues) {
  const normalizedLabName = normalizeComparable_(labName);

  return allowedValues.some(value => {
    const normalizedValue = normalizeComparable_(value);
    if (!normalizedValue) return false;

    return normalizedLabName === normalizedValue || normalizedLabName.includes(normalizedValue);
  });
}

function getUserDataByEmail_(email) {
  if (!email) {
    return {
      authorized: false
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PERMISSIONS);
  if (!sheet) {
    throw new Error('PERMISSIONS sheet not found');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { authorized: false };
  }

  data.shift(); // remove headers

  const normalizedEmail = normalizeComparable_(email);

  for (const row of data) {
    const rowEmail = normalizeComparable_(row[0]);
    if (!rowEmail) continue;

    if (rowEmail === normalizedEmail) {
      return {
        email: normalizeText_(row[0]),
        name: normalizeText_(row[1]),
        school: normalizeText_(row[2]) || 'NONE',
        labs: normalizeText_(row[3]) || 'NONE',
        shareLink: normalizeText_(row[4]),
        authorized: true
      };
    }
  }

  return {
    email: email,
    authorized: false
  };
}

function getCurrentUser_() {
  const email = resolveEmail_();
  const user = getUserDataByEmail_(email);
  if (!user.authorized) {
    throw new Error('غير مصرح');
  }
  return user;
}

/* ================= PUBLIC API ================= */
function getUserData(email, authToken) {
  return getAuthorizedUser_(email, authToken);
}

function getLabs(email, authToken) {
  const user = getAuthorizedUser_(email, authToken);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.LABS);

  const data = sheet.getDataRange().getValues();
  data.shift();

  let labs = data
    .filter(r => r[0])
    .map(r => {
  const fullName = normalizeText_(r[0]);

  const isStudent = fullName.toUpperCase().includes("STUDENT");
  const isTeacher = fullName.toUpperCase().includes("TEACHER");

  const baseName = fullName
    .replace(/-STUDENT/i, "")
    .replace(/-TEACHER/i, "")
    .trim();

  return {
    labName: baseName,
    type: isStudent ? "STUDENT" : isTeacher ? "TEACHER" : "UNKNOWN",
    school: normalizeText_(r[1]),
    sheetId: normalizeText_(r[2])
  };
});
  if (!isAllValue_(user.school) && !isNoneValue_(user.school)) {
    const allowedSchools = parseAllowedValues_(user.school).map(normalizeComparable_);
    labs = labs.filter(l => allowedSchools.includes(normalizeComparable_(l.school)));
  }

  if (!isAllValue_(user.labs) && !isNoneValue_(user.labs)) {
    const allowed = parseAllowedValues_(user.labs);
    labs = labs.filter(l => matchesAllowedLab_(l.labName, allowed));
  }

  return labs;
}

function getLabSettings(sheetId, email, authToken) {
  const user = getAuthorizedUser_(email, authToken);
  const labs = getLabs(user.email, authToken);

  const allowed = labs.find(l => l.sheetId === sheetId);
  if (!allowed) throw new Error('غير مصرح');

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];

  const values = sheet.getRange(2,1,49,2).getValues();
  const validations = sheet.getRange(2,1,49,2).getDataValidations();

  const result = [];

  for (let i=0;i<values.length;i++){
    if(!values[i][0]) continue;

    let options = [];
    const rule = validations[i][1];
    if(rule){
      options = rule.getCriteriaValues()[0] || [];
    }

    result.push({
      setting: values[i][0],
      value: values[i][1],
      options: options
    });
  }

  return result;
}

function updateSetting(sheetId, name, value, email, authToken) {
  const user = getAuthorizedUser_(email, authToken);
  const labs = getLabs(user.email, authToken);

  const allowed = labs.find(l => l.sheetId === sheetId);
  if (!allowed) throw new Error('غير مصرح');

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];

  const data = sheet.getRange(2,1,49,2).getValues();

  for (let i=0;i<data.length;i++){
    if(data[i][0] === name){
      sheet.getRange(2+i,2).setValue(value);
      return true;
    }
  }

  throw new Error('Setting not found');
}

/* ================= OPTIONAL HELPER ================= */
function getAuthorizedUserSummary() {
  const user = getCurrentUser_();
  return {
    email: user.email,
    name: user.name,
    school: user.school,
    labs: user.labs
  };
}
