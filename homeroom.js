// Reads the school's "Master list homeroom teacher" workbook (.xls or .xlsx).
// Returns { academic_year, rows } where each row is one class:
// { campus, grade, name_en, name_km, subject, remark }.
// Pure functions only (no DOM), so they can also be tested with Node.

const HR_ZW_RE = /[​‌‍﻿]/g;

function hrText(v) {
  return String(v ?? '').replace(HR_ZW_RE, '').replace(/\s+/g, ' ').trim();
}

// Returns null when the sheet does not have the homeroom layout.
function parseHomeroomSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });

  let academicYear = null;
  for (const r of rows) {
    const m = hrText(r.join(' ')).match(/Academic Year:?\s*(\d{4})\s*-\s*(\d{4})/i);
    if (m) { academicYear = `${m[1]}-${m[2]}`; break; }
  }

  // Header row: has "Name" (English) and "Grade".
  const hIdx = rows.findIndex(r => {
    const h = r.map(c => hrText(c).toLowerCase());
    return h.includes('name') && h.includes('grade');
  });
  if (hIdx < 0) return null;

  const header = rows[hIdx].map(c => hrText(c).toLowerCase());
  const colKm = header.indexOf('ឈ្មោះ');
  const colName = header.indexOf('name');
  const colSubject = header.indexOf('subject');
  const colGrade = header.indexOf('grade');
  const colRemark = header.indexOf('remarks');

  const out = [];
  let campus = null;
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const first = hrText(r[0]);

    // Section rows such as "STD 1 CAMPUS" set the campus for the rows below.
    if (/CAMPUS/i.test(first) && !hrText(r[colGrade])) {
      campus = first;
      continue;
    }
    // Stop at the footer notes.
    if (/Note|បញ្ឈប់|គ្រូបន្ទុកថ្នាក់ត្រូ/i.test(first)) break;

    const grade = hrText(r[colGrade]).toUpperCase();
    const nameEn = hrText(r[colName]);
    if (!/^\d{1,2}[A-Z]+$/.test(grade) || !nameEn) continue;

    out.push({
      campus,
      grade,
      name_en: nameEn,
      name_km: colKm >= 0 ? hrText(r[colKm]) || null : null,
      subject: colSubject >= 0 ? hrText(r[colSubject]) || null : null,
      remark: colRemark >= 0 ? hrText(r[colRemark]) || null : null
    });
  }

  return { academic_year: academicYear, rows: out };
}

// Picks the first sheet that has the homeroom layout.
function parseHomeroomWorkbook(wb) {
  for (const name of wb.SheetNames) {
    const res = parseHomeroomSheet(wb.Sheets[name]);
    if (res && res.rows.length) return { ...res, sheet: name };
  }
  return null;
}
