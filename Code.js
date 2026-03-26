/* ================= CONFIG ================= */
const SHEET_NAMES = {
  PERMISSIONS: 'PERMISSIONS',
  LABS: 'LABS'
};

const RANGE_SETTINGS_START_ROW = 2;
const RANGE_SETTINGS_END_ROW = 50;

/* ================= ENTRY ================= */
function doGet(e) {
  const ctx = resolveRequestContext_(e);

  if (!ctx.authorized) {
    const t = HtmlService.createTemplateFromFile('Unauthorized');
    t.email = ctx.email || 'غير معروف';
    return t.evaluate().setTitle('غير مصرح');
  }

  const t = HtmlService.createTemplateFromFile('Index');
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
  let email = '';

  try {
    email = Session.getActiveUser().getEmail() || '';
  } catch (err) {
    email = '';
  }

  if (!email && e && e.parameter && e.parameter.email) {
    email = String(e.parameter.email).trim();
  }

  return email;
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

  const normalizedEmail = String(email).trim().toLowerCase();

  for (const row of data) {
    const rowEmail = String(row[0] || '').trim().toLowerCase();
    if (!rowEmail) continue;

    if (rowEmail === normalizedEmail) {
      return {
        email: row[0] || '',
        name: row[1] || '',
        school: row[2] || 'NONE',
        labs: row[3] || 'NONE',
        shareLink: row[4] || '',
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
function getUserData() {
  const email = resolveEmail_();
  return getUserDataByEmail_(email);
}

function getLabs() {
  const user = getCurrentUser_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.LABS);
  if (!sheet) {
    throw new Error('LABS sheet not found');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  data.shift();

  let labs = data
    .filter(row => row[0] && row[1] && row[2])
    .map(row => ({
      labName: String(row[0]).trim(),
      school: String(row[1]).trim(),
      sheetId: String(row[2]).trim()
    }));

  if (user.school !== 'ALL') {
    labs = labs.filter(lab => lab.school === user.school);
  }

  if (user.labs !== 'ALL') {
    const allowedLabs = String(user.labs)
      .split(',')
      .map(x => x.trim())
      .filter(Boolean);

    labs = labs.filter(lab => allowedLabs.includes(lab.labName));
  }

  return labs;
}

function getLabSettings(sheetId) {
  const labs = getLabs();
  const allowed = labs.find(l => String(l.sheetId) === String(sheetId));
  if (!allowed) {
    throw new Error('غير مصرح');
  }

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];
  if (!sheet) {
    throw new Error('Lab sheet not found');
  }

  const range = sheet.getRange(RANGE_SETTINGS_START_ROW, 1, RANGE_SETTINGS_END_ROW - RANGE_SETTINGS_START_ROW + 1, 2);
  const values = range.getValues();
  const validations = range.getDataValidations();

  const result = [];

  for (let i = 0; i < values.length; i++) {
    const settingName = values[i][0];
    const settingValue = values[i][1];

    if (!settingName) continue;

    let options = [];
    const rule = validations[i][1];

    if (rule) {
      const type = rule.getCriteriaType();
      const args = rule.getCriteriaValues();

      if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        options = (args[0] || []).map(x => String(x));
      }
    }

    result.push({
      setting: String(settingName).trim(),
      value: settingValue === null || settingValue === undefined ? '' : String(settingValue),
      options: options
    });
  }

  return result;
}

function updateSetting(sheetId, name, value) {
  const labs = getLabs();
  const allowed = labs.find(l => String(l.sheetId) === String(sheetId));
  if (!allowed) {
    throw new Error('غير مصرح');
  }

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];
  if (!sheet) {
    throw new Error('Lab sheet not found');
  }

  const data = sheet.getRange(RANGE_SETTINGS_START_ROW, 1, RANGE_SETTINGS_END_ROW - RANGE_SETTINGS_START_ROW + 1, 2).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(name).trim()) {
      sheet.getRange(RANGE_SETTINGS_START_ROW + i, 2).setValue(value);
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