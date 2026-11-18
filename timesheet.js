// Reads the school's schedule workbook (sheet "TIMESHEET") and returns one
// record per active teacher: { name_en, subject, classes }.
// Pure functions only (no DOM), so they can also be tested with Node.

const ZW_RE = /[​‌‍﻿]/g;          // invisible characters
const CAMPUS_TAG_RE = /\s*\((STD|SSK|TSK|BCH|CAR|SOW|CKD\d*|BTB\d|VSB|BKK|Southwest2|Sunway)\)/g;

function cleanText(v) {
  return String(v ?? '').replace(ZW_RE, '').replace(/\s+/g, ' ').trim();
}

// Returns an array of teachers, or null if the sheet is not in the TIMESHEET layout.
function parseTimesheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });

  // 1. Find the header row of the teacher list (has "Name" and "Subject").
  const hIdx = rows.findIndex(r =>
    r.some(c => cleanText(c).toLowerCase() === 'name') &&
    r.some(c => cleanText(c).toLowerCase() === 'subject'));
  if (hIdx < 0) return null;

  const header = rows[hIdx].map(c => cleanText(c).toLowerCase());
  const colNo = header.findIndex(c => c === 'nº' || c === 'no' || c === 'n°');
  const colName = header.indexOf('name');
  const colSubject = header.indexOf('subject');
  const colGrade = header.indexOf('grade');
  if (colNo < 0 || colName < 0) return null;

  // 2. Walk the list. A numbered row starts a teacher; rows below it add
  //    more subjects / grades for the same teacher.
  const people = new Map();
  let current = null;

  for (let i = hIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const rowText = row.map(cleanText).join(' ').toLowerCase();

    if (/resign|registrar|maternity/.test(rowText)) break;                          // end of active list
    if (row.some(c => cleanText(c).toLowerCase() === 'nº')) break;                  // next table header

    const no = cleanText(row[colNo]);
    const name = cleanText(row[colName]);
    if (/^\d+$/.test(no) && name) {
      const key = name.toUpperCase();
      if (!people.has(key)) people.set(key, { name_en: name, subjects: [], grades: [] });
      current = people.get(key);
    }
    if (!current) continue;

    if (colSubject >= 0) {
      const subject = cleanText(row[colSubject]).replace(CAMPUS_TAG_RE, '').replace(/\s+/g, ' ').trim();
      if (subject && !current.subjects.includes(subject)) current.subjects.push(subject);
    }
    if (colGrade >= 0) {
      const grade = cleanText(row[colGrade]);
      if (grade && !current.grades.includes(grade)) current.grades.push(grade);
    }
  }

  return [...people.values()].map(p => ({
    name_en: p.name_en,
    subject: p.subjects.join(', ') || null,
    classes: p.grades.join(' | ') || null
  }));
}
