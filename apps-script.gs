const SEAT_COLUMN = 1;
const NAME_COLUMN = 2;
const HEADER_ROW = 1;
const FIRST_STUDENT_ROW = 2;
const FIRST_ASSIGNMENT_COLUMN = 3;
const RED_COLORS = ["#ff0000", "#ea4335", "rgb(255, 0, 0)"];

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bySeat = {};

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const range = sheet.getDataRange();
    const values = range.getDisplayValues();
    const backgrounds = range.getBackgrounds();
    if (!values.length) return;

    const headers = values[HEADER_ROW - 1] || [];
    for (let r = FIRST_STUDENT_ROW - 1; r < values.length; r++) {
      const seat = String(values[r][SEAT_COLUMN - 1] || "").trim();
      if (!isStudentSeat(seat)) continue;

      const rawName = String(values[r][NAME_COLUMN - 1] || "").trim();
      if (!rawName) continue;
      const studentNumber = Number(seat) || r;
      const displayName = seat + "號 " + rawName;
      if (!bySeat[seat]) bySeat[seat] = { seat, studentNumber, name: displayName, calm: false, missing: [] };

      for (let c = FIRST_ASSIGNMENT_COLUMN - 1; c < backgrounds[r].length; c++) {
        if (!isRedCellColor(backgrounds[r][c])) continue;
        bySeat[seat].calm = true;
        bySeat[seat].missing.push(titleForCell(headers, c, sheetName));
      }
    }
  });

  const students = Object.values(bySeat)
    .sort((a, b) => a.studentNumber - b.studentNumber)
    .map(student => ({
      seat: student.seat,
      name: student.name,
      calm: student.calm,
      missing: [...new Set(student.missing)]
    }));

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, updatedAt: new Date().toISOString(), students }))
    .setMimeType(ContentService.MimeType.JSON);
}

function titleForCell(headers, col, sheetName) {
  const header = String(headers[col] || "").trim();
  return sheetName + "：" + (header || "未命名作業");
}

function isRedCellColor(color) {
  const value = String(color || "").trim().toLowerCase();
  if (RED_COLORS.includes(value)) return true;
  const hex = value.match(/^#([0-9a-f]{6})$/);
  if (hex) return isStrictRed(parseInt(hex[1].slice(0, 2), 16), parseInt(hex[1].slice(2, 4), 16), parseInt(hex[1].slice(4, 6), 16));
  const rgb = value.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (rgb) return isStrictRed(Number(rgb[1]), Number(rgb[2]), Number(rgb[3]));
  return false;
}

function isStrictRed(r, g, b) {
  return r >= 220 && g <= 85 && b <= 85 && r > g * 2.4 && r > b * 2.4;
}

function isStudentSeat(value) {
  return /^\d+$/.test(String(value || "").trim());
}
