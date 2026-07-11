/**
 * Classroom OS mirror read integration for the legacy break board Apps Script.
 *
 * Paste target:
 * - Replace the contents of board.gs with this file.
 *
 * Keep unchanged:
 * - Keep break-board-sync.gs.
 * - Keep webapp.gs.gs if it is the only top-level doPost(e) used for
 *   Classroom OS -> Google Sheet writes.
 *
 * Do not paste this beside the old board.gs content. Apps Script uses one
 * shared global namespace, so keeping both copies would duplicate getStatus().
 */

const CLASSROOM_BREAK_BOARD_STATUS_SHEET_NAME = "_break_board_status";
const CLASSROOM_BREAK_BOARD_SEAT_COLUMN = 1;
const CLASSROOM_BREAK_BOARD_NAME_COLUMN = 2;
const CLASSROOM_BREAK_BOARD_HEADER_ROW = 1;
const CLASSROOM_BREAK_BOARD_FIRST_STUDENT_ROW = 2;
const CLASSROOM_BREAK_BOARD_FIRST_ASSIGNMENT_COLUMN = 3;
const CLASSROOM_BREAK_BOARD_RED_COLORS = ["#ff0000", "#f4cccc", "#ea9999", "#e06666", "#ea4335", "rgb(255, 0, 0)"];
const CLASSROOM_BREAK_BOARD_YELLOW_COLORS = ["#ffff00"];

const CLASSROOM_OS_MIRROR_READ_SPREADSHEET_ID = "17yodPxaX0iUZTajiJzapnMqnebFZ9ja7TREeMiHmkwc";
const CLASSROOM_OS_MIRROR_READ_SHEET_NAME = "_ClassroomOS_BreakBoard";

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const result = handleAction(params.action || "getStatus", params);
  return outputJson(result, params.callback);
}

function legacyBoardDoPostHandler(e) {
  let body = {};
  try {
    body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  } catch (error) {
    return outputJson({
      ok: false,
      error: "POST_JSON_ERROR",
      message: "POST body is not valid JSON.",
    });
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
    return {
      ok: false,
      error: "UNKNOWN_ACTION",
      message: "Unknown action: " + action,
    };
  } finally {
    lock.releaseLock();
  }
}

function getStatus() {
  const manual = classroomBreakBoardReadStatusStore();
  const classroomOsMirror = classroomOsMirrorReadSheet();
  const classCalm = manual.classCalm === true;

  if (classroomOsMirror && classroomOsMirror.students.length) {
    return classroomBreakBoardBuildStatusResponse({
      source: "ClassroomOSMirror",
      classroomOSMirror: true,
      classCalm,
      manual,
      updatedAt: classroomOsMirror.updatedAt,
      students: classroomOsMirror.students,
      debugColors: [],
    });
  }

  const scanned = classroomBreakBoardScanStudentsAndMissing();
  return classroomBreakBoardBuildStatusResponse({
    source: "LegacyColorScan",
    classroomOSMirror: false,
    classCalm,
    manual,
    updatedAt: new Date().toISOString(),
    students: scanned.students,
    debugColors: scanned.debugColors,
  });
}

function classroomBreakBoardBuildStatusResponse(input) {
  const students = input.students.map(function(student) {
    const manualStatus = input.manual.students[student.seat] || "auto";
    const missingAssignments = classroomOsMirrorSplitItems(student.missingAssignments);
    const missingCorrections = classroomOsMirrorSplitItems(student.missingCorrections);
    const incompleteAssignments = classroomOsMirrorSplitItems(student.incompleteAssignments);

    return {
      seat: student.seat,
      name: student.name,
      manualStatus,
      missingAssignments,
      missingCorrections,
      incompleteAssignments,
      missing: missingAssignments,
      calm: manualStatus === "calm",
      canBreak: classroomOsMirrorBoolean(student.canBreak, true),
      tabletBlocked: classroomOsMirrorBoolean(student.tabletBlocked, false),
      syncedAt: student.syncedAt || "",
    };
  });

  return {
    ok: true,
    action: "getStatus",
    updatedAt: input.updatedAt || new Date().toISOString(),
    source: input.source,
    classroomOSMirror: input.classroomOSMirror,
    classCalm: input.classCalm,
    students,
    debugColors: input.debugColors || [],
  };
}

function updateStatus(params) {
  const seat = String(params.seat || "").trim();
  const status = normalizeManualStatus(params.status);
  if (!classroomBreakBoardIsStudentSeat(seat)) {
    return { ok: false, error: "BAD_SEAT", message: "seat must be a number." };
  }

  const sheet = classroomBreakBoardEnsureStatusSheet();
  const row = classroomBreakBoardFindSeatRow(sheet, seat);
  if (row) {
    sheet.getRange(row, 1, 1, 3).setValues([[seat, status, new Date()]]);
  } else {
    sheet.appendRow([seat, status, new Date()]);
  }
  classroomBreakBoardRemoveDuplicateSeatRows(sheet, seat, row || sheet.getLastRow());
  return getStatus();
}

function updateAllStatus(params) {
  const classCalm = parseBoolean(params.classCalm);
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("classCalm", classCalm ? "true" : "false");
  props.setProperty("classCalmUpdatedAt", new Date().toISOString());
  return getStatus();
}

function classroomBreakBoardReadStatusStore() {
  const sheet = classroomBreakBoardEnsureStatusSheet();
  const values = sheet.getDataRange().getDisplayValues();
  const props = PropertiesService.getDocumentProperties();
  const store = { classCalm: parseBoolean(props.getProperty("classCalm")), students: {} };

  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const first = values[rowIndex][0];
    const second = values[rowIndex][1];
    const third = values[rowIndex][2];

    if (first === "meta" && second === "classCalm") {
      store.classCalm = parseBoolean(third);
      continue;
    }
    if (first === "student" && classroomBreakBoardIsStudentSeat(second)) {
      store.students[second] = normalizeManualStatus(third);
      continue;
    }
    if (classroomBreakBoardIsStudentSeat(first)) {
      store.students[first] = normalizeManualStatus(second);
    }
  }

  return store;
}

function classroomBreakBoardEnsureStatusSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CLASSROOM_BREAK_BOARD_STATUS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CLASSROOM_BREAK_BOARD_STATUS_SHEET_NAME);
    sheet.hideSheet();
    sheet.appendRow(["seat", "status", "updatedAt"]);
  } else {
    classroomBreakBoardNormalizeStatusSheetHeader(sheet);
  }
  return sheet;
}

function classroomBreakBoardNormalizeStatusSheetHeader(sheet) {
  const header = sheet.getRange(1, 1, 1, Math.max(3, sheet.getLastColumn())).getDisplayValues()[0];
  if (header[0] === "seat" && header[1] === "status" && header[2] === "updatedAt") return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["seat", "status", "updatedAt"]);
    return;
  }
  sheet.getRange(1, 1, 1, 3).setValues([["seat", "status", "updatedAt"]]);
}

function classroomBreakBoardFindSeatRow(sheet, seat) {
  const values = sheet.getDataRange().getDisplayValues();
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = rowIndex + 1;
    if (values[rowIndex][0] === seat) return row;
    if (values[rowIndex][0] === "student" && values[rowIndex][1] === seat) return row;
  }
  return 0;
}

function classroomBreakBoardRemoveDuplicateSeatRows(sheet, seat, keepRow) {
  const values = sheet.getDataRange().getDisplayValues();
  for (let rowIndex = values.length - 1; rowIndex >= 1; rowIndex -= 1) {
    const row = rowIndex + 1;
    const currentSeat = values[rowIndex][0] === "student" ? values[rowIndex][1] : values[rowIndex][0];
    if (row !== keepRow && currentSeat === seat) sheet.deleteRow(row);
  }
}

function classroomBreakBoardScanStudentsAndMissing() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const bySeat = {};
  const debugColors = [];

  spreadsheet.getSheets().forEach(function(sheet) {
    const sheetName = sheet.getName();
    if (sheetName === CLASSROOM_BREAK_BOARD_STATUS_SHEET_NAME || sheetName.charAt(0) === "_") return;

    const range = sheet.getDataRange();
    const values = range.getDisplayValues();
    const backgrounds = range.getBackgrounds();
    if (!values.length) return;

    const headers = values[CLASSROOM_BREAK_BOARD_HEADER_ROW - 1] || [];
    for (let rowIndex = CLASSROOM_BREAK_BOARD_FIRST_STUDENT_ROW - 1; rowIndex < values.length; rowIndex += 1) {
      const seat = String(values[rowIndex][CLASSROOM_BREAK_BOARD_SEAT_COLUMN - 1] || "").trim();
      const rawName = String(values[rowIndex][CLASSROOM_BREAK_BOARD_NAME_COLUMN - 1] || "").trim();
      if (!classroomBreakBoardIsStudentSeat(seat) || !rawName) continue;

      if (!bySeat[seat]) {
        bySeat[seat] = {
          seat,
          studentNumber: Number(seat),
          name: seat + "號 " + rawName,
          missingAssignments: [],
          missingCorrections: [],
          incompleteAssignments: [],
          canBreak: true,
          tabletBlocked: false,
          syncedAt: "",
        };
      }

      for (let columnIndex = CLASSROOM_BREAK_BOARD_FIRST_ASSIGNMENT_COLUMN - 1; columnIndex < backgrounds[rowIndex].length; columnIndex += 1) {
        const color = backgrounds[rowIndex][columnIndex];
        const title = classroomBreakBoardTitleForCell(headers, columnIndex, sheetName);
        if (classroomBreakBoardIsDebuggableColor(color)) {
          debugColors.push({
            sheetName,
            row: rowIndex + 1,
            column: columnIndex + 1,
            title,
            color,
          });
        }
        if (classroomBreakBoardIsRedLikeColor(color)) bySeat[seat].missingAssignments.push(title);
        else if (classroomBreakBoardIsYellowLikeColor(color)) bySeat[seat].missingCorrections.push(title);
      }
    }
  });

  const students = Object.values(bySeat)
    .sort(function(a, b) {
      return a.studentNumber - b.studentNumber;
    })
    .map(function(student) {
      return {
        seat: student.seat,
        name: student.name,
        missingAssignments: Array.from(new Set(student.missingAssignments)),
        missingCorrections: Array.from(new Set(student.missingCorrections)),
        incompleteAssignments: [],
        canBreak: true,
        tabletBlocked: false,
        syncedAt: "",
      };
    });

  return { students, debugColors };
}

function classroomOsMirrorReadSheet() {
  const spreadsheet = SpreadsheetApp.openById(CLASSROOM_OS_MIRROR_READ_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(CLASSROOM_OS_MIRROR_READ_SHEET_NAME);
  if (!sheet) return null;

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return null;

  const headerRow = values[0];
  const seatIndex = classroomOsMirrorFindHeader(headerRow, ["\u5ea7\u865f", "seat", "seatNumber"]);
  const nameIndex = classroomOsMirrorFindHeader(headerRow, ["\u59d3\u540d", "name"]);
  if (seatIndex < 0 || nameIndex < 0) {
    throw new Error("Classroom OS mirror sheet is missing \u5ea7\u865f or \u59d3\u540d header.");
  }

  const missingIndex = classroomOsMirrorFindHeader(headerRow, ["\u672a\u4ea4", "missing", "missingAssignments"]);
  const revisionIndex = classroomOsMirrorFindHeader(headerRow, ["\u8a02\u6b63", "revision", "missingCorrections", "needsRevision", "needs_correction"]);
  const incompleteIndex = classroomOsMirrorFindHeader(headerRow, ["\u672a\u5b8c\u6210", "incomplete", "incompleteAssignments", "pending"]);
  const canBreakIndex = classroomOsMirrorFindHeader(headerRow, ["\u53ef\u4e0b\u8ab2", "canBreak"]);
  const tabletBlockedIndex = classroomOsMirrorFindHeader(headerRow, ["\u7981\u7528\u5e73\u677f", "tabletBlocked"]);
  const syncedAtIndex = classroomOsMirrorFindHeader(headerRow, ["\u540c\u6b65\u6642\u9593", "syncedAt", "updatedAt"]);

  const students = [];
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    const seat = String(classroomOsMirrorCell(row, seatIndex) || "").trim();
    const rawName = String(classroomOsMirrorCell(row, nameIndex) || "").trim();
    if (!classroomBreakBoardIsStudentSeat(seat) || !rawName) continue;

    students.push({
      seat,
      studentNumber: classroomOsMirrorSeatNumber(seat),
      name: seat + "\u865f " + rawName.replace(/^\d+\s*\u865f\s*/, ""),
      missingAssignments: classroomOsMirrorSplitItems(classroomOsMirrorCell(row, missingIndex)),
      missingCorrections: classroomOsMirrorSplitItems(classroomOsMirrorCell(row, revisionIndex)),
      incompleteAssignments: classroomOsMirrorSplitItems(classroomOsMirrorCell(row, incompleteIndex)),
      canBreak: classroomOsMirrorBoolean(classroomOsMirrorCell(row, canBreakIndex), true),
      tabletBlocked: classroomOsMirrorBoolean(classroomOsMirrorCell(row, tabletBlockedIndex), false),
      syncedAt: String(classroomOsMirrorCell(row, syncedAtIndex) || "").trim(),
    });
  }

  if (!students.length) return null;
  students.sort(function(a, b) {
    if (a.studentNumber !== b.studentNumber) return a.studentNumber - b.studentNumber;
    return a.name.localeCompare(b.name);
  });

  return {
    students,
    updatedAt: classroomOsMirrorLatestSyncedAt(students) || new Date().toISOString(),
  };
}

function classroomOsMirrorFindHeader(headerRow, names) {
  const normalizedTargets = names.map(function(name) {
    return classroomOsMirrorNormalizeHeader(name);
  });

  for (let index = 0; index < headerRow.length; index += 1) {
    const normalized = classroomOsMirrorNormalizeHeader(headerRow[index]);
    if (normalizedTargets.indexOf(normalized) >= 0) return index;
  }

  return -1;
}

function classroomOsMirrorCell(row, index) {
  if (index < 0) return "";
  return row[index];
}

function classroomOsMirrorSplitItems(value) {
  if (Array.isArray(value)) return value.map(String).map(function(item) { return item.trim(); }).filter(Boolean);
  return String(value || "")
    .split(/\r?\n|[、，,]/)
    .map(function(item) {
      return item.trim();
    })
    .filter(Boolean);
}

function classroomOsMirrorBoolean(value, fallback) {
  if (value === true) return true;
  if (value === false) return false;
  const text = String(value || "").trim().toLowerCase();
  if (text === "true" || text === "是" || text === "1" || text === "yes" || text === "y") return true;
  if (text === "false" || text === "否" || text === "0" || text === "no" || text === "n") return false;
  return fallback;
}

function classroomOsMirrorSeatNumber(value) {
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : 9999;
}

function classroomOsMirrorNormalizeHeader(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, "");
}

function classroomOsMirrorLatestSyncedAt(students) {
  return students
    .map(function(student) {
      return String(student.syncedAt || "").trim();
    })
    .filter(Boolean)
    .sort()
    .pop() || "";
}

function testClassroomOsMirrorRead() {
  const result = classroomOsMirrorReadSheet();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function classroomBreakBoardTitleForCell(headers, columnIndex, sheetName) {
  const header = String(headers[columnIndex] || "").trim();
  return sheetName + "\uff5c" + (header || "\u672a\u547d\u540d\u4f5c\u696d");
}

function classroomBreakBoardIsRedLikeColor(color) {
  const value = String(color || "").trim().toLowerCase();
  if (CLASSROOM_BREAK_BOARD_RED_COLORS.indexOf(value) >= 0) return true;
  const rgb = classroomBreakBoardParseColorToRgb(value);
  if (!rgb) return false;
  return classroomBreakBoardIsStrictRed(rgb.r, rgb.g, rgb.b);
}

function classroomBreakBoardIsYellowLikeColor(color) {
  const value = String(color || "").trim().toLowerCase();
  return CLASSROOM_BREAK_BOARD_YELLOW_COLORS.indexOf(value) >= 0;
}

function classroomBreakBoardParseColorToRgb(value) {
  const hex = String(value || "").match(/^#([0-9a-f]{6})$/i);
  if (hex) {
    return {
      r: parseInt(hex[1].slice(0, 2), 16),
      g: parseInt(hex[1].slice(2, 4), 16),
      b: parseInt(hex[1].slice(4, 6), 16),
    };
  }
  const rgb = String(value || "").match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (rgb) return { r: Number(rgb[1]), g: Number(rgb[2]), b: Number(rgb[3]) };
  return null;
}

function classroomBreakBoardIsStrictRed(red, green, blue) {
  if (red >= 220 && green <= 110 && blue <= 110) return true;
  if (red >= 180 && green <= 140 && blue <= 140 && red - green >= 50 && red - blue >= 50) return true;
  return false;
}

function classroomBreakBoardIsDebuggableColor(color) {
  const value = String(color || "").trim().toLowerCase();
  if (!value) return false;
  if (value === "#ffffff" || value === "white" || value === "rgb(255, 255, 255)" || value === "rgba(255, 255, 255, 1)") return false;
  return true;
}

function isRedCellColor(color) {
  return classroomBreakBoardIsRedLikeColor(color);
}

function isYellowCellColor(color) {
  return classroomBreakBoardIsYellowLikeColor(color);
}

function classroomBreakBoardIsStudentSeat(value) {
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
