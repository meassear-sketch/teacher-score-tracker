const cfg = window.APP_CONFIG;
const sb = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_KEY);
const $ = id => document.getElementById(id);

const STR = {
  km: {
    title: 'ប្រព័ន្ធតាមដានការប្រគល់ពិន្ទុរបស់គ្រូ',
    titleHr: 'ប្រព័ន្ធតាមដានគ្រូបន្ទុក',
    subtitle: 'គ្រប់គ្រងឈ្មោះគ្រូ កាលបរិច្ឆេទប្រគល់ពិន្ទុ និងគណនាថ្ងៃយឺតដោយស្វ័យប្រវត្តិ',
    refresh: 'ធ្វើបច្ចុប្បន្នភាព', export: 'ទាញយក Excel', langBtn: 'English',
    dueLabel: 'ថ្ងៃកំណត់ប្រគល់ពិន្ទុ (រួម)', calcLabel: 'ថ្ងៃគោលគណនា (ថ្ងៃនេះ)',
    applyAll: 'ដាក់គ្រប់គ្នា', importBtn: 'បញ្ចូល File Excel', addBtn: 'បន្ថែមឈ្មោះគ្រូ',
    resetBtn: 'កំណត់ការប្រគល់ឡើងវិញ', clearBtn: 'លុបទិន្នន័យគ្រូទាំងអស់',
    total: 'គ្រូសរុបទាំងអស់', onTime: 'ប្រគល់រួច (ទាន់ពេល)', onTimeHint: 'ប្រគល់មុន ឬចំថ្ងៃកំណត់',
    late: 'ប្រគល់យឺត', notSub: 'មិនទាន់ប្រគល់', notHint: 'ហួសថ្ងៃកំណត់', pct: 'បានប្រគល់សរុប', avgLate: 'មធ្យមភាគថ្ងៃយឺត',
    fAll: 'ទាំងអស់', fOn: 'ប្រគល់រួច', fLate: 'ប្រគល់យឺត', fNot: 'មិនទាន់ប្រគល់',
    vTable: 'តារាងតាមដាន', vBoard: 'ផ្ទាំងគ្រូមិនទាន់', vCharts: 'ក្រាហ្វិក',
    search: 'ស្វែងរកឈ្មោះគ្រូ ឬមុខវិជ្ជា...',
    sortName: 'តម្រៀបតាមឈ្មោះ', sortSub: 'តម្រៀបតាមមុខវិជ្ជា', sortStatus: 'តម្រៀបតាមស្ថានភាព', sortDays: 'តម្រៀបតាមចំនួនថ្ងៃ',
    no: '#', nameEn: 'ឈ្មោះ (EN)', nameKm: 'ឈ្មោះ (ខ្មែរ)', subject: 'មុខវិជ្ជា', classes: 'ថ្នាក់',
    dueCol: 'ថ្ងៃកំណត់', dueOwn: 'ថ្ងៃកំណត់ (ផ្ទាល់ខ្លួន, ទុកទទេ = ប្រើថ្ងៃរួម)', submittedCol: 'ថ្ងៃបានប្រគល់',
    status: 'ស្ថានភាព', days: 'ការគណនាថ្ងៃ', remark: 'កំណត់ចំណាំ', actions: 'សកម្មភាព',
    edit: 'កែ', del: 'លុប', save: 'រក្សាទុក', cancel: 'បោះបង់',
    addTitle: 'បន្ថែមគ្រូ', editTitle: 'កែប្រែព័ត៌មានគ្រូ',
    stOn: 'ប្រគល់រួច', stLate: 'ប្រគល់យឺត', stNot: 'មិនទាន់ប្រគល់',
    dOnTime: 'ចំថ្ងៃកំណត់', dEarly: 'មុនកំណត់', dLate: 'យឺត', dOver: 'ហួស', dLeft: 'នៅសល់', dToday: 'ថ្ងៃនេះជាថ្ងៃកំណត់',
    days_u: 'ថ្ងៃ', empty: 'មិនមានទិន្នន័យ',
    saved: 'បានរក្សាទុករួច', errSave: 'រក្សាទុកមិនបាន', errLoad: 'ទាញទិន្នន័យមិនបាន',
    cloudOk: '☁️ បានភ្ជាប់ Cloud', cloudErr: '⚠️ ភ្ជាប់ Cloud មិនបាន', cloudLoading: 'កំពុងភ្ជាប់...',
    confirmDel: 'លុបគ្រូនេះមែនទេ?', confirmReset: 'កំណត់ថ្ងៃបានប្រគល់ទាំងអស់ឲ្យទៅទទេមែនទេ?', confirmClear: 'លុបគ្រូទាំងអស់មែនទេ? មិនអាចត្រឡប់វិញបានទេ!',
    importDone: 'បញ្ចូលរួច', noRows: 'File គ្មានទិន្នន័យ', viewerNo: 'អ្នកមើលមិនអាចកែប្រែបានទេ',
    signOut: 'ត្រឡប់ទៅចូលប្រើ',
    vHome: 'គ្រូបន្ទុកថ្នាក់', hrImport: 'នាំចូលគ្រូបន្ទុកថ្នាក់ (Excel)', hrSearch: 'ស្វែងរកថ្នាក់ ឬឈ្មោះគ្រូ...',
    hrNone: 'មិនមានទិន្នន័យគ្រូបន្ទុកថ្នាក់', hrError: 'ទាញទិន្នន័យគ្រូបន្ទុកថ្នាក់មិនបាន (ពិនិត្យ database)',
    hrClass: 'ថ្នាក់', hrTeacherEn: 'គ្រូបន្ទុក (EN)', hrTeacherKm: 'គ្រូបន្ទុក (ខ្មែរ)', hrCount: 'ចំនួនថ្នាក់',
    phone: 'លេខទូរស័ព្ទ',
    qRowLbl: 'ត្រីមាស (ថ្ងៃកំណត់)', qLabel1: 'ត្រីមាសទី១', qLabel2: 'ត្រីមាសទី២', qLabel3: 'ត្រីមាសទី៣', qLabel4: 'ត្រីមាសទី៤',
    qNone: 'សូមកំណត់ថ្ងៃសម្រាប់ត្រីមាសនេះជាមុនសិន',
    ttTotal: 'ម៉ោងសរុប', ttNone: 'មិនមានកាលវិភាគសម្រាប់គ្រូនេះ', ttImported: 'នាំចូលកាលវិភាគ',
    ttSave: 'រក្សាទុកម៉ោង', ttSaved: 'បានរក្សាទុកម៉ោងការងារបន្ទុកថ្នាក់',
    hrEdSave: 'រក្សាទុកព័ត៌មាន', hrEdSaved: 'បានរក្សាទុកព័ត៌មានគ្រូ', hrEdNameReq: 'សូមបញ្ចូលឈ្មោះគ្រូ (EN)', hrEdNoRow: 'មិនអាចរក្សាទុកបាន',
    ttDutyCol: 'ម៉ោងសរុប', ttActCol: 'សកម្មភាព', ttHours: 'ម៉ោង', ttEditHours: 'កែម៉ោង', ttSetHours: 'កំណត់ម៉ោង',
    ttClear: 'លុបម៉ោងការងារ', ttClearConfirm: 'លុបម៉ោងការងារបន្ទុកថ្នាក់របស់គ្រូនេះទាំងអស់?', ttCleared: 'បានលុបម៉ោងការងារ',
    hrSemLbl: 'ឆមាស', hrBadFile: 'ឯកសារនេះមិនមានទម្រង់គ្រូបន្ទុកថ្នាក់',
    hrConfirm: 'ជំនួសទិន្នន័យឆ្នាំ {y} ឆមាស {s} ដោយថ្នាក់ {n} ពីឯកសារនេះ?',
    hrImported: 'បាននាំចូលគ្រូបន្ទុកថ្នាក់',
    chartTitle: 'សមាមាត្រតាមមុខវិជ្ជា', lgOn: 'រួច', lgLate: 'យឺត', lgNot: 'មិនទាន់'
  },
  en: {
    title: 'Teacher Score Submission Tracker',
    titleHr: 'Homeroom Teacher Tracker',
    subtitle: 'Manage teacher names, submission dates and late days automatically',
    refresh: 'Refresh', export: 'Export Excel', langBtn: 'ភាសាខ្មែរ',
    dueLabel: 'Submission deadline (all)', calcLabel: 'Calculation date (today)',
    applyAll: 'Apply to all', importBtn: 'Import Excel', addBtn: 'Add teacher',
    resetBtn: 'Reset submissions', clearBtn: 'Clear all teachers',
    total: 'Total teachers', onTime: 'Submitted on time', onTimeHint: 'Submitted on or before deadline',
    late: 'Submitted late', notSub: 'Not yet submitted', notHint: 'Overdue', pct: 'submitted', avgLate: 'avg late days',
    fAll: 'All', fOn: 'On time', fLate: 'Late', fNot: 'Not submitted',
    vTable: 'Tracker table', vBoard: 'Missing board', vCharts: 'Charts',
    search: 'Search teacher or subject...',
    sortName: 'Sort by name', sortSub: 'Sort by subject', sortStatus: 'Sort by status', sortDays: 'Sort by days',
    no: '#', nameEn: 'Name (EN)', nameKm: 'Name (Khmer)', subject: 'Subject', classes: 'Classes',
    dueCol: 'Deadline', dueOwn: 'Own deadline (empty = use the shared one)', submittedCol: 'Submitted on',
    status: 'Status', days: 'Days', remark: 'Remark', actions: 'Actions',
    edit: 'Edit', del: 'Delete', save: 'Save', cancel: 'Cancel',
    addTitle: 'Add teacher', editTitle: 'Edit teacher',
    stOn: 'On time', stLate: 'Late', stNot: 'Not submitted',
    dOnTime: 'On deadline', dEarly: 'early', dLate: 'late', dOver: 'overdue', dLeft: 'days left', dToday: 'Due today',
    days_u: 'days', empty: 'No data',
    saved: 'Saved', errSave: 'Could not save', errLoad: 'Could not load data',
    cloudOk: '☁️ Cloud connected', cloudErr: '⚠️ Cloud connection failed', cloudLoading: 'Connecting...',
    confirmDel: 'Delete this teacher?', confirmReset: 'Clear all submission dates?', confirmClear: 'Delete ALL teachers? This cannot be undone!',
    importDone: 'Imported', noRows: 'File has no rows', viewerNo: 'View-only access',
    signOut: 'Back to sign-in',
    vHome: 'Homeroom teachers', hrImport: 'Import homeroom (Excel)', hrSearch: 'Search class or teacher...',
    hrNone: 'No homeroom data yet', hrError: 'Could not load homeroom data (check database)',
    hrClass: 'Class', hrTeacherEn: 'Homeroom teacher (EN)', hrTeacherKm: 'Homeroom teacher (Khmer)', hrCount: 'Classes',
    phone: 'Phone',
    qRowLbl: 'Quarters (due date)', qLabel1: 'Quarter 1', qLabel2: 'Quarter 2', qLabel3: 'Quarter 3', qLabel4: 'Quarter 4',
    qNone: 'Set a date for this quarter first',
    ttTotal: 'Total hours', ttNone: 'No timetable for this teacher', ttImported: 'Timetable imported',
    ttSave: 'Save hours', ttSaved: 'Homeroom duty hours saved',
    hrEdSave: 'Save details', hrEdSaved: 'Teacher details saved', hrEdNameReq: 'Please enter the teacher name (EN)', hrEdNoRow: 'Could not save',
    ttDutyCol: 'Duty hours', ttActCol: 'Actions', ttHours: 'hrs', ttEditHours: 'Edit hours', ttSetHours: 'Set hours',
    ttClear: 'Delete duty hours', ttClearConfirm: 'Delete all homeroom duty hours of this teacher?', ttCleared: 'Duty hours deleted',
    hrSemLbl: 'Semester', hrBadFile: 'This file is not in the homeroom teacher format',
    hrConfirm: 'Replace academic year {y}, semester {s} with {n} classes from this file?',
    hrImported: 'Homeroom teachers imported',
    chartTitle: 'Status by subject', lgOn: 'On time', lgLate: 'Late', lgNot: 'Not submitted'
  }
};

const state = {
  lang: localStorage.getItem('lang') || 'km',
  theme: localStorage.getItem('theme') || 'light',
  filter: 'all', view: 'table', search: '', sort: 'name',
  teachers: [], settings: {}, editId: null, role: 'admin',
  hr: { rows: [], sel: null, error: false, duty: {} }
};

const t = k => (STR[state.lang][k] ?? k);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const dayNum = s => { if (!s) return null; const [y, m, d] = s.split('-').map(Number); return Date.UTC(y, m - 1, d) / 86400000; };
// Today's date from the device's own clock (YYYY-MM-DD).
const todayStr = () => {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
};
const fmt = s => { if (!s) return ''; const [y, m, d] = s.split('-'); return `${d}/${m}/${y}`; };
const calcDue = tc => tc.due_date || state.settings.due_date || null;
const STATUS_KEY = { ontime: 'stOn', late: 'stLate', notsub: 'stNot' };

// Returns { key: 'ontime' | 'late' | 'notsub', n }
// ontime: n = days early (0 = on the deadline day)
// late:   n = days late
// notsub: n > 0 = days overdue, n <= 0 = days left before deadline
function evaluate(tc) {
  const due = dayNum(calcDue(tc));
  if (tc.submitted_date) {
    if (due === null) return { key: 'ontime', n: 0 };
    const diff = dayNum(tc.submitted_date) - due;
    return diff > 0 ? { key: 'late', n: diff } : { key: 'ontime', n: -diff };
  }
  if (due === null) return { key: 'notsub', n: 0 };
  const calc = dayNum(todayStr());
  return { key: 'notsub', n: calc - due };
}

function dayLabel(ev) {
  if (ev.key === 'ontime') return ev.n === 0 ? t('dOnTime') : `${ev.n} ${t('days_u')} ${t('dEarly')}`;
  if (ev.key === 'late') return `${ev.n} ${t('days_u')} ${t('dLate')}`;
  if (ev.n > 0) return `${ev.n} ${t('days_u')} ${t('dOver')}`;
  if (ev.n === 0) return t('dToday');
  return `${-ev.n} ${t('days_u')} ${t('dLeft')}`;
}

function computeRows() {
  return state.teachers.map(tc => ({ ...tc, ev: evaluate(tc) }));
}

function visibleRows(rows) {
  const q = state.search.trim().toLowerCase();
  const list = rows
    .filter(r => state.filter === 'all' || r.ev.key === state.filter)
    .filter(r => !q || [r.name_en, r.name_km, r.subject, r.classes].some(v => (v || '').toLowerCase().includes(q)));
  const cmp = {
    name: (a, b) => a.name_en.localeCompare(b.name_en),
    subject: (a, b) => (a.subject || '').localeCompare(b.subject || ''),
    status: (a, b) => a.ev.key.localeCompare(b.ev.key),
    days: (a, b) => b.ev.n - a.ev.n
  }[state.sort];
  return list.sort(cmp);
}

function applyStatic() {
  document.body.classList.toggle('km', state.lang === 'km');
  document.documentElement.lang = state.lang;
  document.documentElement.setAttribute('data-theme', state.theme);
  document.querySelectorAll('[data-i18n]').forEach(el => { el.textContent = t(el.dataset.i18n); });
  // Homeroom view: shorter title and no subtitle line.
  const hrView = state.view === 'homeroom';
  document.querySelector('header h1').textContent = hrView ? t('titleHr') : t('title');
  document.querySelector('header p').style.display = hrView ? 'none' : '';
  $('langBtn').textContent = t('langBtn');
  $('search').placeholder = t('search');
  document.querySelectorAll('[data-filter]').forEach(b => b.classList.toggle('on', b.dataset.filter === state.filter));
  document.querySelectorAll('[data-view]').forEach(b => b.classList.toggle('on', b.dataset.view === state.view));
}

function render() {
  applyStatic();
  const rows = computeRows();
  const c = { total: rows.length, ontime: 0, late: 0, notsub: 0, lateSum: 0, overdue: 0 };
  rows.forEach(r => {
    c[r.ev.key]++;
    if (r.ev.key === 'late') c.lateSum += r.ev.n;
    if (r.ev.key === 'notsub' && r.ev.n > 0) c.overdue++;
  });
  $('sTotal').textContent = c.total;
  $('sOn').textContent = c.ontime;
  $('sLate').textContent = c.late;
  $('sNot').textContent = c.notsub;
  $('sPct').textContent = c.total ? `${Math.round((c.ontime + c.late) / c.total * 100)}% ${t('pct')}` : '';
  $('sLateAvg').textContent = c.late ? `${t('avgLate')} ${(c.lateSum / c.late).toFixed(1)}` : '';
  $('sNotHint').textContent = `${t('notHint')}: ${c.overdue}`;

  const counts = { all: c.total, ontime: c.ontime, late: c.late, notsub: c.notsub };
  document.querySelectorAll('[data-filter] .cnt').forEach(el => { el.textContent = counts[el.parentElement.dataset.filter]; });

  // The homeroom view has its own list, so hide the teacher-only controls there.
  const isHr = state.view === 'homeroom';
  $('stats').style.display = isHr ? 'none' : '';
  $('filterBar').style.display = isHr ? 'none' : '';
  $('sort').style.display = isHr ? 'none' : '';
  $('hrBar').style.display = isHr ? 'flex' : 'none';
  $('search').placeholder = isHr ? t('hrSearch') : t('search');

  let html;
  if (isHr) html = renderHomeroom();
  else if (state.view === 'table') html = renderTable(visibleRows(rows));
  else if (state.view === 'board') html = renderBoard(rows.filter(r => r.ev.key === 'notsub'));
  else html = renderCharts(rows);
  $('content').innerHTML = html;
}

function renderTable(rows) {
  if (!rows.length) return `<div class="empty">${t('empty')}</div>`;
  return `<div class="table-wrap tbl-scroll"><table><thead><tr>
    <th>${t('no')}</th><th>${t('nameEn')}</th><th>${t('nameKm')}</th><th>${t('subject')}</th><th>${t('classes')}</th>
    <th>${t('dueCol')}</th><th>${t('submittedCol')}</th><th>${t('status')}</th><th>${t('days')}</th>
    <th>${t('remark')}</th><th>${t('actions')}</th>
  </tr></thead><tbody>
  ${rows.map((r, i) => `<tr>
    <td class="idx">${i + 1}</td>
    <td><strong>${esc(r.name_en)}</strong></td>
    <td>${esc(r.name_km)}</td>
    <td>${esc(r.subject)}</td>
    <td><span class="code">${esc(r.classes)}</span></td>
    <td>${fmt(calcDue(r))}</td>
    <td><input type="date" value="${r.submitted_date || ''}" data-sub="${r.id}" ${state.role === 'viewer' ? 'disabled' : ''} /></td>
    <td><span class="pill ${r.ev.key}">${t(STATUS_KEY[r.ev.key])}</span></td>
    <td class="days ${r.ev.key}">${dayLabel(r.ev)}</td>
    <td class="muted">${esc(r.remark)}</td>
    <td><button class="btn sm" data-edit="${r.id}">${t('edit')}</button>
        <button class="btn sm danger" data-del="${r.id}">${t('del')}</button></td>
  </tr>`).join('')}
  </tbody></table></div>`;
}

function renderBoard(rows) {
  if (!rows.length) return `<div class="empty">${t('empty')}</div>`;
  return `<div class="board">${rows.map(r => `
    <div class="bcard">
      <strong>${esc(r.name_en)}</strong>
      ${r.name_km ? `<div class="muted">${esc(r.name_km)}</div>` : ''}
      <div class="muted">${esc(r.subject)} · ${esc(r.classes)}</div>
      <div class="days notsub" style="margin-top:6px">${dayLabel(r.ev)}</div>
    </div>`).join('')}</div>`;
}

function renderCharts(rows) {
  if (!rows.length) return `<div class="empty">${t('empty')}</div>`;
  const groups = {};
  rows.forEach(r => {
    const k = r.subject || '—';
    groups[k] ||= { ontime: 0, late: 0, notsub: 0, total: 0 };
    groups[k][r.ev.key]++;
    groups[k].total++;
  });
  const pct = (x, total) => (total ? (x / total * 100) : 0).toFixed(2) + '%';
  const bars = Object.entries(groups).sort((a, b) => b[1].total - a[1].total).map(([k, g]) => `
    <div style="margin:14px 0">
      <div style="display:flex;justify-content:space-between;margin-bottom:4px"><strong>${esc(k)}</strong><span class="muted">${g.total}</span></div>
      <div class="bar">
        <span style="width:${pct(g.ontime, g.total)};background:#1b9a5e"></span>
        <span style="width:${pct(g.late, g.total)};background:#e0931c"></span>
        <span style="width:${pct(g.notsub, g.total)};background:#d33a4b"></span>
      </div>
    </div>`).join('');
  return `<h3 style="margin:4px 0">${t('chartTitle')}</h3>${bars}
    <div class="legend">
      <span><i class="dot" style="background:#1b9a5e"></i>${t('lgOn')}</span>
      <span><i class="dot" style="background:#e0931c"></i>${t('lgLate')}</span>
      <span><i class="dot" style="background:#d33a4b"></i>${t('lgNot')}</span>
    </div>`;
}

function toast(msg) {
  const el = $('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove('show'), 2200);
}

function fail(err) {
  console.error(err);
  toast(`${t('errSave')}: ${err.message || err}`);
}

function setCloud(s) {
  $('cloud').textContent = s === 'ok' ? t('cloudOk') : s === 'error' ? t('cloudErr') : t('cloudLoading');
}

async function loadAll() {
  setCloud('loading');
  const [s, tc] = await Promise.all([
    sb.from('app_settings').select('*').eq('id', 1).maybeSingle(),
    sb.from('teachers').select('*').order('name_en')
  ]);
  if (s.error || tc.error) {
    setCloud('error');
    toast(t('errLoad'));
    console.error(s.error || tc.error);
    return;
  }
  state.settings = s.data || {};
  state.teachers = tc.data || [];
  renderQuarters();
  $('calcDate').value = todayStr();
  setCloud('ok');
  await loadHomeroom();
  render();
}

async function saveSetting(patch) {
  if (guardWrite()) return;
  const { error } = await sb.from('app_settings').update(patch).eq('id', 1);
  if (error) return fail(error);
  Object.assign(state.settings, patch);
  toast(t('saved'));
  render();
}

async function updateTeacher(id, patch) {
  if (guardWrite()) return;
  const { error } = await sb.from('teachers').update({ ...patch, updated_at: new Date().toISOString() }).eq('id', id);
  if (error) return fail(error);
  Object.assign(state.teachers.find(x => x.id === id), patch);
  toast(t('saved'));
  render();
}

async function removeTeacher(id) {
  if (guardWrite()) return;
  if (!confirm(t('confirmDel'))) return;
  const { error } = await sb.from('teachers').delete().eq('id', id);
  if (error) return fail(error);
  await loadAll();
}

const FIELDS = ['name_en', 'name_km', 'subject', 'classes', 'due_date', 'submitted_date', 'remark'];

function openForm(id) {
  if (guardWrite()) return;
  const f = $('teacherForm');
  f.reset();
  state.editId = id || null;
  $('dlgTitle').textContent = id ? t('editTitle') : t('addTitle');
  if (id) {
    const tc = state.teachers.find(x => x.id === id);
    FIELDS.forEach(k => { f.elements[k].value = tc[k] || ''; });
  }
  $('dlg').showModal();
}

$('teacherForm').addEventListener('submit', async e => {
  e.preventDefault();
  if (guardWrite()) return;
  const f = e.target;
  const data = {};
  FIELDS.forEach(k => { data[k] = f.elements[k].value.trim() || null; });
  if (!data.name_en) return;
  const res = state.editId
    ? await sb.from('teachers').update(data).eq('id', state.editId)
    : await sb.from('teachers').insert(data);
  if (res.error) return fail(res.error);
  $('dlg').close();
  await loadAll();
  toast(t('saved'));
});

$('cancelBtn').onclick = () => $('dlg').close();
$('addBtn').onclick = () => openForm(null);

$('content').addEventListener('change', e => {
  const id = e.target.dataset.sub;
  if (id) updateTeacher(id, { submitted_date: e.target.value || null });
});

$('content').addEventListener('click', e => {
  const dc = e.target.closest('[data-duty-clear]');
  if (dc) return clearHrDuty(dc.dataset.dutyClear);
  const hr = e.target.closest('[data-hr]');
  if (hr) return openHrDialog(hr.dataset.hr);
  const ed = e.target.closest('[data-edit]');
  if (ed) return openForm(ed.dataset.edit);
  const dl = e.target.closest('[data-del]');
  if (dl) return removeTeacher(dl.dataset.del);
});

document.querySelectorAll('[data-filter]').forEach(b => b.onclick = () => { state.filter = b.dataset.filter; render(); });
document.querySelectorAll('[data-view]').forEach(b => b.onclick = () => { state.view = b.dataset.view; render(); });
$('search').oninput = e => { state.search = e.target.value; render(); };
$('sort').onchange = e => { state.sort = e.target.value; render(); };
$('hrSem').onchange = e => { state.hr.sel = e.target.value; render(); };
$('hrDlgClose').onclick = () => $('hrDlg').close();
$('hrFile').addEventListener('change', async e => {
  const f = e.target.files[0];
  e.target.value = '';
  if (f) await importHomeroom(f);
});

// Quarter dates (set by admin) are stored in app_settings.quarters, so every user sees the same ones.
function renderQuarters() {
  const q = state.settings.quarters || {};
  [1, 2, 3, 4].forEach(n => {
    const input = $('qDate' + n);
    input.value = q[n] || '';
    input.disabled = state.role === 'viewer';
    document.querySelector(`[data-q="${n}"]`).disabled = state.role === 'viewer';
  });
}

// Sets the due date for all teachers (clears each teacher's own date first).
// quarters: optional new quarter map to save with it.
async function applyDueDate(due, quarters) {
  if (guardWrite()) return;
  const { error } = await sb.from('teachers').update({ due_date: null }).not('id', 'is', null);
  if (error) return fail(error);
  await saveSetting(quarters ? { due_date: due, quarters } : { due_date: due });
}

[1, 2, 3, 4].forEach(n => {
  // Picking a quarter date saves it and applies it to all teachers (admin only).
  $('qDate' + n).addEventListener('change', e => {
    if (guardWrite()) { renderQuarters(); return; }
    const date = e.target.value || null;
    const quarters = { ...(state.settings.quarters || {}), [n]: date };
    if (date) applyDueDate(date, quarters);
    else saveSetting({ quarters });
  });
  // Clicking a quarter label applies that quarter's saved date to all teachers.
  document.querySelector(`[data-q="${n}"]`).addEventListener('click', () => {
    const date = (state.settings.quarters || {})[n];
    if (!date) return toast(t('qNone'));
    applyDueDate(date);
  });
});

// The calculation date is always the device's date, so it is not saved or editable.
// Refresh it if the page stays open past midnight.
let lastDay = todayStr();
setInterval(() => {
  const now = todayStr();
  if (now !== lastDay) {
    lastDay = now;
    $('calcDate').value = now;
    render();
  }
}, 60000);

$('resetBtn').onclick = async () => {
  if (guardWrite()) return;
  if (!confirm(t('confirmReset'))) return;
  const { error } = await sb.from('teachers').update({ submitted_date: null }).not('id', 'is', null);
  if (error) return fail(error);
  await loadAll();
};

$('clearBtn').onclick = async () => {
  if (guardWrite()) return;
  if (!confirm(t('confirmClear'))) return;
  const { error } = await sb.from('teachers').delete().not('id', 'is', null);
  if (error) return fail(error);
  await loadAll();
};

$('refreshBtn').onclick = loadAll;

$('langBtn').onclick = () => {
  state.lang = state.lang === 'km' ? 'en' : 'km';
  localStorage.setItem('lang', state.lang);
  render();
};

$('themeBtn').onclick = () => {
  state.theme = state.theme === 'dark' ? 'light' : 'dark';
  localStorage.setItem('theme', state.theme);
  applyStatic();
};

// Imports the school's full schedule workbook (TIMESHEET layout).
// Returns false when the file is not in that layout, so the simple importer runs instead.
async function importTimesheet(file) {
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  const order = [...wb.SheetNames].sort((a, b) => (b === 'TIMESHEET') - (a === 'TIMESHEET'));
  let parsed = null;
  for (const name of order) {
    const p = parseTimesheet(wb.Sheets[name]);
    if (p && p.length) { parsed = p; break; }
  }
  if (!parsed) return false;

  const byName = new Map(state.teachers.map(x => [x.name_en.toLowerCase(), x]));
  const ops = [];
  const inserts = [];
  let updated = 0;
  parsed.forEach(r => {
    const ex = byName.get(r.name_en.toLowerCase());
    if (ex) {
      // The schedule file is the source of truth for subject and classes.
      ops.push(sb.from('teachers').update({ subject: r.subject, classes: r.classes }).eq('id', ex.id));
      updated++;
    } else {
      inserts.push({ name_en: r.name_en, subject: r.subject, classes: r.classes });
    }
  });

  const inFile = new Set(parsed.map(r => r.name_en.toLowerCase()));
  const notInFile = state.teachers.filter(x => !inFile.has(x.name_en.toLowerCase())).length;

  const results = await Promise.all(ops);
  const bad = results.find(x => x.error);
  if (bad) { fail(bad.error); return true; }
  if (inserts.length) {
    const { error } = await sb.from('teachers').insert(inserts);
    if (error) { fail(error); return true; }
  }
  const tt = await importTimetable(wb);
  if (tt && tt.error) { fail(tt.error); return true; }
  await loadAll();
  const ttMsg = typeof tt === 'number' && tt > 0 ? `, ${t('ttImported')}: ${tt}` : '';
  toast(`${t('importDone')}: ${updated} updated, ${inserts.length} added, ${notInFile} not in file${ttMsg}`);
  return true;
}

// Saves the weekly timetable found in the class sheets (G6 … G12) of the schedule file.
// Replaces the same academic year + semester. Returns the number of slots, or { error }.
async function importTimetable(wb) {
  const tt = parseTimetableWorkbook(wb);
  if (!tt.slots.length || !tt.academic_year || !tt.semester) return 0;
  const del = await sb.from('teacher_timetable').delete()
    .eq('academic_year', tt.academic_year).eq('semester', tt.semester);
  if (del.error) return { error: del.error };
  const payload = tt.slots.map(s => ({ ...s, academic_year: tt.academic_year, semester: tt.semester }));
  for (let i = 0; i < payload.length; i += 400) {
    const { error } = await sb.from('teacher_timetable').insert(payload.slice(i, i + 400));
    if (error) return { error };
  }
  return payload.length;
}

$('importFile').addEventListener('change', async e => {
  if (guardWrite()) { e.target.value = ''; return; }
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    if (await importTimesheet(file)) return;
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    if (!rows.length) return toast(t('noRows'));

    const keys = Object.keys(rows[0]);
    const pick = re => keys.find(k => re.test(k));
    const nameKeys = keys.filter(k => /ឈ្មោះ|name/i.test(k));
    const enKey = nameKeys.find(k => /\ben\b|english|អង់គ្លេស/i.test(k)) || nameKeys[0];
    const kmKey = nameKeys.find(k => k !== enKey && /ខ្មែរ|khmer|km/i.test(k)) || nameKeys.find(k => k !== enKey);
    const subjKey = pick(/subject|មុខវិជ្ជា/i);
    const classKey = pick(/class|ថ្នាក់/i);
    const remarkKey = pick(/remark|ចំណាំ|កំណត់ហេតុ/i);

    const seen = new Set();
    const incoming = rows.map(r => {
      const rec = {
        name_en: enKey ? String(r[enKey]).trim() : '',
        name_km: kmKey ? String(r[kmKey]).trim() : '',
        subject: subjKey ? String(r[subjKey]).trim() : '',
        classes: classKey ? String(r[classKey]).trim() : '',
        remark: remarkKey ? String(r[remarkKey]).trim() : ''
      };
      return rec;
    }).filter(r => {
      const key = r.name_en.toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    const byName = new Map(state.teachers.map(x => [x.name_en.toLowerCase(), x]));
    const ops = [];
    const inserts = [];
    incoming.forEach(r => {
      // Only write non-empty values so blanks in Excel never erase existing data.
      const clean = Object.fromEntries(Object.entries(r).filter(([, v]) => v !== ''));
      const ex = byName.get(r.name_en.toLowerCase());
      if (ex) ops.push(sb.from('teachers').update(clean).eq('id', ex.id));
      else inserts.push(clean);
    });

    const results = await Promise.all(ops);
    const bad = results.find(x => x.error);
    if (bad) return fail(bad.error);
    if (inserts.length) {
      const { error } = await sb.from('teachers').insert(inserts);
      if (error) return fail(error);
    }
    await loadAll();
    toast(`${t('importDone')}: ${incoming.length}`);
  } catch (err) {
    fail(err);
  }
});

$('exportBtn').onclick = () => {
  const data = visibleRows(computeRows()).map((r, i) => ({
    '#': i + 1,
    [t('nameEn')]: r.name_en,
    [t('nameKm')]: r.name_km || '',
    [t('subject')]: r.subject || '',
    [t('classes')]: r.classes || '',
    [t('dueCol')]: fmt(calcDue(r)),
    [t('submittedCol')]: fmt(r.submitted_date),
    [t('status')]: t(STATUS_KEY[r.ev.key]),
    [t('days')]: dayLabel(r.ev),
    [t('remark')]: r.remark || ''
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Scores');
  XLSX.writeFile(wb, `score-submission-${todayStr()}.xlsx`);
};

// ---- Homeroom teachers (គ្រូបន្ទុកថ្នាក់) ----
const hrCombo = r => `${r.academic_year}|${r.semester}`;
const gradeKey = g => { const m = /^(\d+)/.exec(g || ''); return m ? Number(m[1]) : 0; };

async function loadHomeroom() {
  const { data, error } = await sb.from('homeroom_assignments').select('*').order('grade');
  if (error) {
    console.error(error);
    state.hr.rows = [];
    state.hr.error = true;
    return;
  }
  state.hr.rows = data || [];
  state.hr.error = false;
  await loadHrDuty();
  renderHrControls();
}

// Counts the saved duty days per homeroom teacher: key = "academic_year|semester|teacher_key".
async function loadHrDuty() {
  const { data, error } = await sb.from('homeroom_duty').select('academic_year, semester, teacher_key, day, period').limit(5000);
  if (error) { console.error(error); state.hr.duty = {}; return; }
  const map = {};
  (data || []).forEach(d => {
    const k = `${d.academic_year}|${d.semester}|${d.teacher_key}`;
    const e = map[k] || (map[k] = { n: 0, byDay: {} });
    e.n++;
    e.byDay[d.day] = d.period;
  });
  state.hr.duty = map;
}

// Deletes all duty hours of one homeroom teacher for the selected academic year + semester.
async function clearHrDuty(id) {
  if (guardWrite()) return;
  const r = state.hr.rows.find(x => String(x.id) === String(id));
  if (!r || !confirm(t('ttClearConfirm'))) return;
  const [ay, sem] = hrCombo(r).split('|');
  const del = await sb.from('homeroom_duty').delete()
    .eq('academic_year', ay).eq('semester', Number(sem)).eq('teacher_key', teacherKey(r.name_en));
  if (del.error) return fail(del.error);
  await loadHrDuty();
  render();
  toast(t('ttCleared'));
}

// Two table cells per teacher: duty-hour count (pill) and actions (admin only).
const HR_DUTY_TOTAL = 2; // duty hours each homeroom teacher must set per week

// Weekday cells (chosen period in an amber pill), the total pill, and actions for admin.
function hrDutyCells(r) {
  const d = state.hr.duty[`${hrCombo(r)}|${teacherKey(r.name_en)}`] || { n: 0, byDay: {} };
  const n = d.n;
  const done = n >= HR_DUTY_TOTAL;
  const gear = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M4.9 19.1L7 17M17 7l2.1-2.1"/></svg>`;
  const days = TT_DAYS.map(day => {
    const p = d.byDay[day];
    if (!p) return '<td><span class="muted">-</span></td>';
    const label = TT_UI_PERIODS[p - 1].replace(/ (AM|PM)$/, '');
    return `<td><span class="duty-slot">${gear}${label}</span></td>`;
  }).join('');
  const pill = `<td><span class="duty-pill${done ? ' ok' : ''}">${done ? '✓ ' : ''}${n}/${HR_DUTY_TOTAL} ${t('ttHours')}</span></td>`;
  let actions = '';
  if (state.role === 'admin') {
    const pencil = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M4 20h4L19 9l-4-4L4 16z"/></svg>`;
    const trash = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg>`;
    actions = `<div class="duty-actions">
        <button class="btn ghost sm" type="button" data-hr="${esc(r.id)}">${pencil}${n ? t('ttEditHours') : t('ttSetHours')}</button>
        ${n ? `<button class="btn ghost sm duty-del" type="button" data-duty-clear="${esc(r.id)}" title="${t('ttClear')}">${trash}</button>` : ''}
      </div>`;
  }
  return days + pill + (state.role === 'admin' ? `<td>${actions}</td>` : '');
}

// Fills the year/semester dropdown: years found in the data, plus the current year's two semesters.
function renderHrControls() {
  const combos = new Set(state.hr.rows.map(hrCombo));
  const now = new Date();
  const y = now.getFullYear();
  const ay = now.getMonth() >= 7 ? `${y}-${y + 1}` : `${y - 1}-${y}`;
  combos.add(`${ay}|1`);
  combos.add(`${ay}|2`);
  const list = [...combos].sort().reverse();
  // Default to the latest semester that already has data.
  const withData = [...new Set(state.hr.rows.map(hrCombo))].sort().reverse();
  if (!list.includes(state.hr.sel)) state.hr.sel = withData[0] || list[0];
  $('hrSem').innerHTML = list.map(c => {
    const [a, s] = c.split('|');
    return `<option value="${c}">${a} · ${t('hrSemLbl')} ${s}</option>`;
  }).join('');
  $('hrSem').value = state.hr.sel;
}

function renderHomeroom() {
  if (state.hr.error) return `<div class="empty">${t('hrError')}</div>`;
  const q = state.search.trim().toLowerCase();
  const rows = state.hr.rows
    .filter(r => hrCombo(r) === state.hr.sel)
    .filter(r => !q || [r.grade, r.name_en, r.name_km, r.subject].some(v => (v || '').toLowerCase().includes(q)))
    .sort((a, b) => gradeKey(a.grade) - gradeKey(b.grade) || a.grade.localeCompare(b.grade));
  if (!rows.length) return `<div class="empty">${t('hrNone')}</div>`;

  // Summary pills for the selected semester (not affected by the search box).
  const all = state.hr.rows.filter(r => hrCombo(r) === state.hr.sel);
  const doneCount = all.filter(r => (state.hr.duty[`${hrCombo(r)}|${teacherKey(r.name_en)}`]?.n || 0) >= HR_DUTY_TOTAL).length;
  const svgUsers = `<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M16 11c1.7 0 3-1.3 3-3s-1.3-3-3-3-3 1.3-3 3 1.3 3 3 3zM8 11c1.7 0 3-1.3 3-3S9.7 5 8 5 5 6.3 5 8s1.3 3 3 3zm0 2c-2.3 0-7 1.2-7 3.5V19h14v-2.5C15 14.2 10.3 13 8 13zm8 0c-.3 0-.6 0-1 .1 1.2.9 2 2 2 3.4V19h6v-2.5c0-2.3-4.7-3.5-7-3.5z"/></svg>`;
  const svgDone = `<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 2a10 10 0 100 20 10 10 0 000-20zm-2 15l-5-5 1.4-1.4L10 14.2l7.6-7.6L19 8z"/></svg>`;
  const svgPending = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>`;
  const statsHtml = `<div class="hr-stats">
      <span class="hr-stat">${svgUsers}<span>គ្រូបន្ទុកសរុប: <b>${all.length}</b> នាក់</span></span>
      <span class="hr-stat ok">${svgDone}<span>បានដាក់ម៉ោងរួច: <b>${doneCount}</b> នាក់</span></span>
      <span class="hr-stat pending">${svgPending}<span>មិនទាន់ដាក់ម៉ោង: <b>${all.length - doneCount}</b> នាក់</span></span>
    </div>`;
  // Phone and actions are admin only. Campus is not shown in this table.
  const admin = state.role === 'admin';
  const dayHead = ['ថ្ងៃចន្ទ (Mon)', 'ថ្ងៃអង្គារ (Tue)', 'ថ្ងៃពុធ (Wed)', 'ថ្ងៃព្រហស្បតិ៍ (Thu)', 'ថ្ងៃសុក្រ (Fri)'];
  let body = '';
  rows.forEach((r, i) => {
    body += `<tr>
      <td class="idx">${i + 1}</td>
      <td><span class="class-pill">${esc(r.grade)}</span></td>
      <td><button class="hr-name hr-stack" data-hr="${esc(r.id)}"><span class="hr-km">${esc(r.name_km)}</span><span class="hr-en">${esc(r.name_en)}</span></button></td>
      <td>${esc(r.subject)}</td>
      ${admin ? `<td>${esc(r.phone)}</td>` : ''}
      ${hrDutyCells(r)}
    </tr>`;
  });
  return `${statsHtml}<div class="table-wrap hr-scroll"><table class="hr-duty"><thead><tr>
      <th>ល.រ</th><th>ថ្នាក់</th><th>ឈ្មោះគ្រូបន្ទុកថ្នាក់</th><th>មុខវិជ្ជា</th>${admin ? '<th>លេខទូរស័ព្ទ (Phone)</th>' : ''}
      ${dayHead.map(d => `<th>${d}</th>`).join('')}
      <th>ម៉ោងសរុប</th>${admin ? '<th>សកម្មភាព</th>' : ''}
    </tr></thead><tbody>${body}</tbody></table></div>
    <div class="legend">${t('hrCount')}: ${rows.length}</div>`;
}

// Labels for the timetable pop-up (days Mon–Fri; periods 1–8 by bell time).
const TT_UI_DAYS = ['ថ្ងៃចន្ទ', 'ថ្ងៃអង្គារ', 'ថ្ងៃពុធ', 'ថ្ងៃព្រហស្បតិ៍', 'ថ្ងៃសុក្រ'];
const TT_UI_PERIODS = ['8:00-8:50 AM', '9:00-9:50 AM', '10:00-10:50 AM', '11:00-11:50 AM',
  '1:00-1:50 PM', '2:00-2:50 PM', '3:00-3:50 PM', '4:00-4:50 PM'];

// One row per weekday: the homeroom duty period (admin picks and saves it).
// rows = the teacher's timetable slots, shown under each dropdown as a hint.
// duty = saved picks [{ day, period }]. editable = admin only.
function ttDutyHtml(rows, duty, editable) {
  const byDay = {};
  rows.forEach(x => { (byDay[x.day] = byDay[x.day] || []).push(x); });
  const dutyBy = {};
  duty.forEach(d => { dutyBy[d.day] = d.period; });
  const icon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M7 2h2v2h6V2h2v2h3v18H4V4h3V2zm13 8H4v10h16V10z"/></svg>`;
  const body = TT_DAYS.map((day, i) => {
    const teach = (byDay[day] || []).slice().sort((a, b) => a.period - b.period)
      .map(x => `${TT_UI_PERIODS[x.period - 1]} · ${esc(x.subject)} ${esc(x.grade)}`).join('<br>');
    const opts = TT_UI_PERIODS.map((label, p) =>
      `<option value="${p + 1}"${dutyBy[day] === p + 1 ? ' selected' : ''}>${label}</option>`).join('');
    const select = `<select class="tt-select" data-day="${day}"${editable ? '' : ' disabled'}>
      <option value="">-- ជ្រើសរើសម៉ោង --</option>${opts}</select>`;
    return `<tr><td><span class="tt-day">${icon}${TT_UI_DAYS[i]}</span></td>
      <td>${select}${teach ? `<div class="tt-hint">${teach}</div>` : ''}</td></tr>`;
  }).join('');
  return `<div class="tt-wrap"><table class="tt"><thead><tr>
      <th>ថ្ងៃ (Day)</th><th>ម៉ោងការងារបន្ទុកថ្នាក់ (Drop List)</th>
    </tr></thead><tbody>${body}</tbody></table></div>
    ${editable ? `<div class="tt-actions"><button class="btn primary" type="button" id="ttSave">${t('ttSave')}</button></div>` : ''}
    ${rows.length ? `<div class="legend">${t('ttTotal')}: ${rows.length}</div>` : ''}`;
}

// Saves the duty periods of one homeroom teacher for the selected academic year + semester.
// Replaces the teacher's previous picks: delete, then insert the days that have a period.
async function saveHrDuty(r, ay, sem) {
  if (guardWrite()) return;
  const key = teacherKey(r.name_en);
  const picks = [...$('ttBox').querySelectorAll('.tt-select')]
    .map(s => ({ day: s.dataset.day, period: Number(s.value) }))
    .filter(p => p.period);
  const btn = $('ttSave');
  if (btn) btn.disabled = true;
  try {
    const del = await sb.from('homeroom_duty').delete()
      .eq('academic_year', ay).eq('semester', sem).eq('teacher_key', key);
    if (del.error) return fail(del.error);
    if (picks.length) {
      const ins = await sb.from('homeroom_duty').insert(picks.map(p => ({
        academic_year: ay, semester: sem, teacher_key: key, day: p.day, period: p.period
      })));
      if (ins.error) return fail(ins.error);
    }
    await loadHrDuty();
    render();
    $('hrDlg').close();
    toast(t('ttSaved'));
  } finally {
    if (btn) btn.disabled = false;
  }
}

// Admin saves a homeroom teacher's name, subject and phone. If the English name changes,
// the teacher's saved duty hours move to the new name so they are not lost.
async function saveHrInfo(r) {
  if (guardWrite()) return;
  const val = id => $(id).value.replace(/\s+/g, ' ').trim();
  const name_en = val('hrEdEn');
  if (!name_en) return toast(t('hrEdNameReq'));
  const payload = {
    name_en,
    name_km: val('hrEdKm') || null,
    subject: val('hrEdSubj') || null,
    phone: val('hrEdPhone') || null
  };
  const btn = $('hrEdSave');
  if (btn) btn.disabled = true;
  try {
    const up = await sb.from('homeroom_assignments').update(payload).eq('id', r.id).select('id');
    if (up.error) return fail(up.error);
    if (!up.data || !up.data.length) return toast(t('hrEdNoRow'));
    const oldKey = teacherKey(r.name_en);
    const newKey = teacherKey(name_en);
    if (oldKey !== newKey) {
      const [ay, sm] = hrCombo(r).split('|');
      const mv = await sb.from('homeroom_duty').update({ teacher_key: newKey })
        .eq('academic_year', ay).eq('semester', Number(sm)).eq('teacher_key', oldKey);
      if (mv.error) return fail(mv.error);
    }
    Object.assign(r, payload);
    await loadHomeroom(); // refreshes the table rows and duty counts
    render();
    toast(t('hrEdSaved'));
  } finally {
    if (btn) btn.disabled = false;
  }
}

let hrDlgSeq = 0;

// Pop-up for a homeroom teacher: subject, phone (admin only) and the weekly timetable.
async function openHrDialog(id) {
  const r = state.hr.rows.find(x => String(x.id) === String(id));
  if (!r) return;
  const seq = ++hrDlgSeq;
  $('hrDlgTitle').textContent = `${t('hrClass')} ${r.grade}`;
  const line = (label, val, big) => `
    <div class="hr-line">
      <span class="hr-label">${esc(label)}</span>
      <span class="hr-val ${big ? 'big' : ''}">${esc(val) || '—'}</span>
    </div>`;
  let html;
  if (state.role === 'admin') {
    // Admin can change the teacher's details here.
    html = `<div class="hr-edit">
      <label>${esc(t('hrTeacherEn'))}<input id="hrEdEn" value="${esc(r.name_en)}"></label>
      <label>${esc(t('hrTeacherKm'))}<input id="hrEdKm" value="${esc(r.name_km)}"></label>
      <label>${esc(t('subject'))}<input id="hrEdSubj" value="${esc(r.subject)}"></label>
      <label>${esc(t('phone'))}<input id="hrEdPhone" value="${esc(r.phone)}" inputmode="tel"></label>
      <div class="tt-actions"><button class="btn primary" type="button" id="hrEdSave">${esc(t('hrEdSave'))}</button></div>
    </div>`;
  } else {
    html = line(t('hrTeacherEn'), r.name_en, true) + line(t('hrTeacherKm'), r.name_km) + line(t('subject'), r.subject);
  }
  html += `<div id="ttBox"></div>`;
  $('hrDlgBody').innerHTML = html;
  const edBtn = $('hrEdSave');
  if (edBtn) edBtn.onclick = () => saveHrInfo(r);
  $('hrDlg').showModal();

  const [ay, smStr] = state.hr.sel.split('|');
  const sem = Number(smStr) || 1;
  const key = teacherKey(r.name_en);
  const [tt, duty] = await Promise.all([
    sb.from('teacher_timetable')
      .select('academic_year, semester, day, period, grade, subject')
      .eq('teacher_key', key),
    sb.from('homeroom_duty')
      .select('day, period')
      .eq('academic_year', ay).eq('semester', sem).eq('teacher_key', key)
  ]);
  if (seq !== hrDlgSeq) return; // another teacher was opened meanwhile
  const box = $('ttBox');
  if (tt.error || duty.error) { box.innerHTML = `<div class="empty">${t('hrError')}</div>`; return; }
  // Hints use the latest academic year and semester that has timetable data for this teacher.
  const data = tt.data || [];
  const latest = data.reduce((best, x) =>
    (!best || x.academic_year > best.academic_year || (x.academic_year === best.academic_year && x.semester > best.semester)) ? x : best, null);
  const rows = latest ? data.filter(x => x.academic_year === latest.academic_year && x.semester === latest.semester) : [];
  box.innerHTML = ttDutyHtml(rows, duty.data || [], state.role === 'admin');
  const btn = $('ttSave');
  if (btn) btn.onclick = () => saveHrDuty(r, ay, sem);
}

// Replaces one academic year + semester with the classes in the chosen Excel file.
// Each class has one homeroom teacher, and the database rejects a teacher listed twice.
async function importHomeroom(file) {
  if (guardWrite()) return;
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  const res = parseHomeroomWorkbook(wb);
  if (!res || !res.academic_year) return toast(t('hrBadFile'));

  const sem = Number(state.hr.sel.split('|')[1]) || 1;
  const msg = t('hrConfirm').replace('{y}', res.academic_year).replace('{s}', sem).replace('{n}', res.rows.length);
  if (!confirm(msg)) return;

  const del = await sb.from('homeroom_assignments').delete()
    .eq('academic_year', res.academic_year).eq('semester', sem);
  if (del.error) return fail(del.error);

  const payload = res.rows.map(r => ({
    academic_year: res.academic_year, semester: sem, campus: r.campus, grade: r.grade,
    name_en: r.name_en, name_km: r.name_km, subject: r.subject, remark: r.remark, phone: r.phone
  }));
  const ins = await sb.from('homeroom_assignments').insert(payload);
  if (ins.error) return fail(ins.error);

  state.hr.sel = `${res.academic_year}|${sem}`;
  await loadHomeroom();
  render();
  toast(`${t('hrImported')}: ${payload.length}`);
}

// ---- Roles: admin (full) and viewer (read-only, UI level) ----
function setRole(r) {
  state.role = r;
  document.body.classList.toggle('viewer', r === 'viewer');
}

// Returns true (and shows a message) when the current user may not change data.
function guardWrite() {
  if (state.role === 'viewer') {
    toast(t('viewerNo'));
    return true;
  }
  return false;
}

// ---- Shared-password gate ----
// Stores only the SHA-256 hash of the password, never the plain text.
const HASHES = {
  admin: '5313e5bf17148de844ff74be3663d47c6e361ca469b30a36337701233c89a15e'  // full access
};
const AUTH_KEY = 'tst_auth';

async function sha256(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('');
}

function startApp() {
  applyStatic();
  loadAll();
}

const savedRole = sessionStorage.getItem(AUTH_KEY);
if (savedRole === 'admin' || savedRole === 'viewer') {
  setRole(savedRole);
  startApp();
} else {
  $('gate').style.display = 'flex';
  $('gatePass').focus();
}

$('gateForm').addEventListener('submit', async e => {
  e.preventDefault();
  const hash = await sha256($('gatePass').value);
  if (hash === HASHES.admin) {
    sessionStorage.setItem(AUTH_KEY, 'admin');
    setRole('admin');
    $('gate').style.display = 'none';
    startApp();
  } else {
    $('gateErr').textContent = 'Password មិនត្រឹមត្រូវ / Wrong password';
    $('gatePass').value = '';
    $('gatePass').focus();
  }
});

// Back to the sign-in screen (admin or view only) from any mode.
$('signOutBtn').addEventListener('click', () => {
  sessionStorage.removeItem(AUTH_KEY);
  location.reload();
});

// "View only" needs no password: entering it gives read-only access to one page.
document.querySelectorAll('[data-view-only]').forEach(b => b.addEventListener('click', () => {
  sessionStorage.setItem(AUTH_KEY, 'viewer');
  setRole('viewer');
  state.view = b.dataset.viewOnly;
  $('gate').style.display = 'none';
  startApp();
}));
