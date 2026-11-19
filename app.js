const cfg = window.APP_CONFIG;
const sb = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_KEY);
const $ = id => document.getElementById(id);

const STR = {
  km: {
    title: 'ប្រព័ន្ធតាមដានការប្រគល់ពិន្ទុរបស់គ្រូ',
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
    chartTitle: 'សមាមាត្រតាមមុខវិជ្ជា', lgOn: 'រួច', lgLate: 'យឺត', lgNot: 'មិនទាន់'
  },
  en: {
    title: 'Teacher Score Submission Tracker',
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
    chartTitle: 'Status by subject', lgOn: 'On time', lgLate: 'Late', lgNot: 'Not submitted'
  }
};

const state = {
  lang: localStorage.getItem('lang') || 'km',
  theme: localStorage.getItem('theme') || 'light',
  filter: 'all', view: 'table', search: '', sort: 'name',
  teachers: [], settings: {}, editId: null, role: 'admin'
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

  let html;
  if (state.view === 'table') html = renderTable(visibleRows(rows));
  else if (state.view === 'board') html = renderBoard(rows.filter(r => r.ev.key === 'notsub'));
  else html = renderCharts(rows);
  $('content').innerHTML = html;
}

function renderTable(rows) {
  if (!rows.length) return `<div class="empty">${t('empty')}</div>`;
  return `<div class="table-wrap"><table><thead><tr>
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
  $('dueDate').value = state.settings.due_date || '';
  $('calcDate').value = todayStr();
  setCloud('ok');
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
  const ed = e.target.closest('[data-edit]');
  if (ed) return openForm(ed.dataset.edit);
  const dl = e.target.closest('[data-del]');
  if (dl) return removeTeacher(dl.dataset.del);
});

document.querySelectorAll('[data-filter]').forEach(b => b.onclick = () => { state.filter = b.dataset.filter; render(); });
document.querySelectorAll('[data-view]').forEach(b => b.onclick = () => { state.view = b.dataset.view; render(); });
$('search').oninput = e => { state.search = e.target.value; render(); };
$('sort').onchange = e => { state.sort = e.target.value; render(); };

$('applyDue').onclick = async () => {
  if (guardWrite()) return;
  const due = $('dueDate').value || null;
  const { error } = await sb.from('teachers').update({ due_date: null }).not('id', 'is', null);
  if (error) return fail(error);
  await saveSetting({ due_date: due });
};

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
  await loadAll();
  toast(`${t('importDone')}: ${updated} updated, ${inserts.length} added, ${notInFile} not in file`);
  return true;
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

// ---- Roles: admin (full) and viewer (read-only, UI level) ----
function setRole(r) {
  state.role = r;
  document.body.classList.toggle('viewer', r === 'viewer');
  $('dueDate').disabled = r === 'viewer';
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

// "View only" needs no password: entering it gives read-only access.
$('viewOnlyBtn').addEventListener('click', () => {
  sessionStorage.setItem(AUTH_KEY, 'viewer');
  setRole('viewer');
  $('gate').style.display = 'none';
  startApp();
});
