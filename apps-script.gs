const STATUS_SHEET_NAME = "_break_board_status";
const SEAT_COLUMN = 1;
const NAME_COLUMN = 2;
const HEADER_ROW = 1;
const FIRST_STUDENT_ROW = 2;
const FIRST_ASSIGNMENT_COLUMN = 3;
const RED_COLORS = ["#ff0000", "#f4cccc", "#ea9999", "#e06666", "#ea4335", "rgb(255, 0, 0)"];
const YELLOW_COLORS = ["#ffff00", "#fff2cc", "#ffd966", "#ffe599", "rgb(255, 255, 0)"];

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const result = handleAction(params.action || "getStatus", params);
  return outputJson(result, params.callback);
}

function doPost(e) {
  let body = {};
  try {
    body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  } catch (error) {
    return outputJson({ ok: false, error: "POST_JSON_ERROR", message: "POST body is not valid JSON." });
  }
  return outputJson(handleAction(body.action || "getStatus", body));
}

function handleAction(action, params) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    if (action === "getStatus") return getStatus();
    if (action === "updateStatus") return updateStatus(params);
    if (action === "updateAllStatus") return updateAllStatus(params);
    return { ok: false, error: "UNKNOWN_ACTION", message: "Unknown action: " + action };
  } finally {
    lock.releaseLock();
  }
}

function getStatus() {
  const manual = readStatusStore();
  const scanned = scanStudentsAndMissing();
  const classCalm = manual.classCalm === true;

  const students = scanned.students.map(student => {
    const manualStatus = manual.students[student.seat] || "auto";
    return {
      seat: student.seat,
      name: student.name,
      manualStatus,
      missingAssignments: student.missingAssignments,
      missingCorrections: student.missingCorrections,
      missing: student.missingAssignments,
      calm: manualStatus === "calm"
    };
  });

  return {
    ok: true,
    action: "getStatus",
    updatedAt: new Date().toISOString(),
    classCalm,
    students,
    debugColors: scanned.debugColors
  };
}

function updateStatus(params) {
  const seat = String(params.seat || "").trim();
  const status = normalizeManualStatus(params.status);
  if (!isStudentSeat(seat)) {
    return { ok: false, error: "BAD_SEAT", message: "seat must be a number." };
  }

  const sheet = ensureStatusSheet();
  const row = findSeatRow(sheet, seat);
  if (row) {
    sheet.getRange(row, 1, 1, 3).setValues([[seat, status, new Date()]]);
  } else {
    sheet.appendRow([seat, status, new Date()]);
  }
  removeDuplicateSeatRows(sheet, seat, row || sheet.getLastRow());
  return getStatus();
}

function updateAllStatus(params) {
  const classCalm = parseBoolean(params.classCalm);
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("classCalm", classCalm ? "true" : "false");
  props.setProperty("classCalmUpdatedAt", new Date().toISOString());
  return getStatus();
}

function readStatusStore() {
  const sheet = ensureStatusSheet();
  const values = sheet.getDataRange().getDisplayValues();
  const props = PropertiesService.getDocumentProperties();
  const store = { classCalm: parseBoolean(props.getProperty("classCalm")), students: {} };

  for (let r = 1; r < values.length; r++) {
    const first = values[r][0];
    const second = values[r][1];
    const third = values[r][2];

    if (first === "meta" && second === "classCalm") {
      store.classCalm = parseBoolean(third);
      continue;
    }
    if (first === "student" && isStudentSeat(second)) {
      store.students[second] = normalizeManualStatus(third);
      continue;
    }
    if (isStudentSeat(first)) store.students[first] = normalizeManualStatus(second);
  }
  return store;
}

function ensureStatusSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STATUS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STATUS_SHEET_NAME);
    sheet.hideSheet();
    sheet.appendRow(["seat", "status", "updatedAt"]);
  } else {
    normalizeStatusSheetHeader(sheet);
  }
  return sheet;
}

function normalizeStatusSheetHeader(sheet) {
  const header = sheet.getRange(1, 1, 1, Math.max(3, sheet.getLastColumn())).getDisplayValues()[0];
  if (header[0] === "seat" && header[1] === "status" && header[2] === "updatedAt") return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["seat", "status", "updatedAt"]);
    return;
  }
  sheet.getRange(1, 1, 1, 3).setValues([["seat", "status", "updatedAt"]]);
}

function findSeatRow(sheet, seat) {
  const values = sheet.getDataRange().getDisplayValues();
  for (let r = 1; r < values.length; r++) {
    if (values[r][0] === seat) return r + 1;
    if (values[r][0] === "student" && values[r][1] === seat) return r + 1;
  }
  return 0;
}

function removeDuplicateSeatRows(sheet, seat, keepRow) {
  const values = sheet.getDataRange().getDisplayValues();
  for (let r = values.length - 1; r >= 1; r--) {
    const row = r + 1;
    const currentSeat = values[r][0] === "student" ? values[r][1] : values[r][0];
    if (row !== keepRow && currentSeat === seat) sheet.deleteRow(row);
  }
}

function scanStudentsAndMissing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bySeat = {};
  const debugColors = [];

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === STATUS_SHEET_NAME || sheetName.charAt(0) === "_") return;

    const range = sheet.getDataRange();
    const values = range.getDisplayValues();
    const backgrounds = range.getBackgrounds();
    if (!values.length) return;

    const headers = values[HEADER_ROW - 1] || [];
    for (let r = FIRST_STUDENT_ROW - 1; r < values.length; r++) {
      const seat = String(values[r][SEAT_COLUMN - 1] || "").trim();
      const rawName = String(values[r][NAME_COLUMN - 1] || "").trim();
      if (!isStudentSeat(seat) || !rawName) continue;

      if (!bySeat[seat]) {
        bySeat[seat] = {
          seat,
          studentNumber: Number(seat),
          name: seat + "號 " + rawName,
          missingAssignments: [],
          missingCorrections: []
        };
      }

      for (let c = FIRST_ASSIGNMENT_COLUMN - 1; c < backgrounds[r].length; c++) {
        const color = backgrounds[r][c];
        const title = titleForCell(headers, c, sheetName);
        if (String(color || "").trim()) {
          debugColors.push({
            sheetName,
            row: r + 1,
            column: c + 1,
            title,
            color
          });
        }
        if (isRedLikeColor(color)) bySeat[seat].missingAssignments.push(title);
        else if (isYellowLikeColor(color)) bySeat[seat].missingCorrections.push(title);
      }
    }
  });

  const students = Object.values(bySeat)
    .sort((a, b) => a.studentNumber - b.studentNumber)
    .map(student => ({
      seat: student.seat,
      name: student.name,
      missingAssignments: [...new Set(student.missingAssignments)],
      missingCorrections: [...new Set(student.missingCorrections)]
    }));

  return { students, debugColors };
}

function titleForCell(headers, col, sheetName) {
  const header = String(headers[col] || "").trim();
  return sheetName + "｜" + (header || "未命名作業");
}

function isRedLikeColor(color) {
  const value = String(color || "").trim().toLowerCase();
  if (RED_COLORS.includes(value)) return true;
  const rgb = parseColorToRgb(value);
  if (!rgb) return false;
  return isStrictRed(rgb.r, rgb.g, rgb.b);
}

function isYellowLikeColor(color) {
  const value = String(color || "").trim().toLowerCase();
  if (YELLOW_COLORS.includes(value)) return true;
  const rgb = parseColorToRgb(value);
  if (!rgb) return false;
  return isStrictYellow(rgb.r, rgb.g, rgb.b);
}

function parseColorToRgb(value) {
  const hex = String(value || "").match(/^#([0-9a-f]{6})$/i);
  if (hex) return { r: parseInt(hex[1].slice(0,2),16), g: parseInt(hex[1].slice(2,4),16), b: parseInt(hex[1].slice(4,6),16) };
  const rgb = String(value || "").match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (rgb) return { r: Number(rgb[1]), g: Number(rgb[2]), b: Number(rgb[3]) };
  return null;
}

function isStrictRed(r, g, b) {
  if (r >= 220 && g <= 110 && b <= 110) return true;
  if (r >= 180 && g <= 140 && b <= 140 && r - g >= 50 && r - b >= 50) return true;
  return false;
}

function isStrictYellow(r, g, b) {
  return r >= 170 && g >= 150 && b <= 130 && Math.abs(r - g) <= 85;
}

function isRedCellColor(color) { return isRedLikeColor(color); }
function isYellowCellColor(color) { return isYellowLikeColor(color); }

function isStudentSeat(value) {
  return /^\d+$/.test(String(value || "").trim());
}

function normalizeManualStatus(value) {
  const status = String(value || "auto").trim();
  if (status === "ok") return "free";
  if (status === "free" || status === "calm" || status === "auto") return status;
  return "auto";
}

function parseBoolean(value) {
  return String(value || "").toLowerCase() === "true" || value === true;
}

function outputJson(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService
      .createTextOutput(String(callback).replace(/[^\w.$]/g, "") + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}


