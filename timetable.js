// Builds each teacher's weekly timetable from the school's schedule workbook.
// It reads the class sheets (G6 … G12): one block per class, with the subject and
// instructor in each cell. Returns { academic_year, semester, slots }, where each slot is
// { teacher_key, teacher_name, day, period, grade, subject }.
// Pure functions only (no DOM), so they can also be tested with Node.

const TT_DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
// Period 1–8 follows the school's bell times: 8–11 AM, then 1–4 PM.
const TT_PERIOD_BY_HOUR = { 8: 1, 9: 2, 10: 3, 11: 4, 13: 5, 14: 6, 15: 7, 16: 8 };
const TT_ZW_RE = /[​‌‍﻿]/g;
const TT_TITLE_RE = /^(Mr|Ms|Mrs|Dr)\.?\s*/i;

function ttText(v) {
  return String(v ?? '').replace(TT_ZW_RE, '').replace(/\s+/g, ' ').trim();
}

// Name key used to match a teacher across files: no title, upper case, single spaces.
function teacherKey(name) {
  return ttText(name).replace(TT_TITLE_RE, '').toUpperCase().trim();
}

function parseTimetableWorkbook(wb) {
  let academicYear = null;
  let semester = null;
  const slots = [];

  for (const sheetName of wb.SheetNames) {
    if (!/^G\d+$/.test(sheetName)) continue;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '', blankrows: false });

    let grade = null;
    let dayCols = null; // column index -> day name, set from each block's header row
    for (const r of rows) {
      const cells = r.map(ttText);
      const joined = cells.join(' ');

      if (!academicYear) {
        const m = joined.match(/Academic Year\s*(\d{4})\s*-\s*(\d{4})/i);
        if (m) academicYear = `${m[1]}-${m[2]}`;
      }
      if (!semester) {
        const m = joined.match(/Semester\s*(\d)/i);
        if (m) semester = Number(m[1]);
      }

      const label = cells[0];
      const g = label.match(/^Grade\s+(\d{1,2}[A-Z]+)/i);
      if (g) {
        grade = g[1].toUpperCase();
        dayCols = null;
        continue;
      }

      if (cells.includes('Monday')) {
        dayCols = {};
        cells.forEach((c, i) => { if (TT_DAYS.includes(c)) dayCols[i] = c; });
        continue;
      }

      const tm = label.match(/^(\d{1,2}):(\d{2})\s*-\s*\d{1,2}:\d{2}/);
      if (!tm || !grade || !dayCols) continue;
      let hour = Number(tm[1]);
      if (hour <= 7) hour += 12; // "1:00 PM" → 13
      const period = TT_PERIOD_BY_HOUR[hour];
      if (!period) continue;

      for (const [ci, day] of Object.entries(dayCols)) {
        const lines = String(r[ci] ?? '').split(/\r?\n/).map(ttText).filter(Boolean);
        if (!lines.length) continue;
        const subject = lines[0];
        const instructor = lines.slice(1).find(l => TT_TITLE_RE.test(l));
        if (!instructor) continue; // a cell without an instructor cannot be assigned to a teacher
        slots.push({
          teacher_key: teacherKey(instructor),
          teacher_name: instructor.replace(TT_TITLE_RE, '').trim(),
          day,
          period,
          grade,
          subject
        });
      }
    }
  }

  return { academic_year: academicYear, semester, slots };
}
