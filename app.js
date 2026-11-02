/* مصروفات مجمع ابن تيمية — الواجهة
 * الأقسام: ١) أدوات عامة  ٢) الدخول  ٣) قائمة الفواتير  ٤) الرفع  ٥) نافذة الفاتورة
 *          ٦) الكشوف والطباعة و Excel  ٧) السجل  ٨) الإعدادات والمستخدمون
 */
'use strict';

// =====================================================================
// ١) أدوات عامة
// =====================================================================
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const KINDS = ['ضريبية', 'بدون ضريبة', 'ورقة مكتوبة'];
const LS = { api: 'mx_api', token: 'mx_token', user: 'mx_user', rep: 'mx_report' };

const S = {                 // حالة البرنامج
  token: localStorage.getItem(LS.token) || '',
  user: JSON.parse(localStorage.getItem(LS.user) || 'null'),
  boot: null, settings: {}, invoices: [], users: [], report: null
};

function apiUrl() {
  const u = (localStorage.getItem(LS.api) || window.API_URL || '').trim();
  return /^https:\/\/script\.google\.com\//.test(u) ? u : '';
}
async function api(action, body = {}) {
  const url = apiUrl();
  if (!url) throw new Error('لم يُضبط رابط الخادم');
  const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action, token: S.token, ...body }), redirect: 'follow' });
  if (!res.ok) throw new Error('تعذّر الاتصال بالخادم (' + res.status + ')');
  const j = await res.json();
  if (!j.ok) {
    if (j.auth) { logout(true); }
    throw new Error(j.error || 'خطأ غير معروف');
  }
  return j.data;
}
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const money = n => (n === '' || n == null || isNaN(n)) ? '' : Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const round2 = n => Math.round(Number(n || 0) * 100) / 100;
function toast(msg, err) {
  const t = $('#toast'); t.textContent = msg; t.className = 'toast show' + (err ? ' err' : '');
  clearTimeout(t._h); t._h = setTimeout(() => t.className = 'toast', err ? 5000 : 2800);
}
function can(p) { return !!(S.user && S.user.perms && (S.user.perms.admin || S.user.perms[p])); }
function ymd(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
function gDate(s) { if (!s) return ''; const [y, m, d] = s.split('-').map(Number); return y + '/' + m + '/' + d; }
function hDate(s) {
  if (!s) return '';
  try { return new Intl.DateTimeFormat('ar-SA-u-ca-islamic-umalqura-nu-latn', { day: 'numeric', month: 'numeric', year: 'numeric' }).format(new Date(s + 'T12:00:00')).replace(/\s?هـ/, '') + 'هـ'; }
  catch (e) { return ''; }
}
function showDate(s, mode) { mode = mode || S.settings.dateMode || 'ميلادي'; return mode === 'هجري' ? hDate(s) : mode === 'كلاهما' ? (gDate(s) + (s ? ' — ' + hDate(s) : '')) : gDate(s); }
function stamp() { const d = new Date(); return ymd(d) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); }
function cats() { return String(S.settings.categories || S.boot?.categories || '').split('\n').map(s => s.trim()).filter(Boolean); }
function kindBadge(k) { const c = { 'ضريبية': 'b-tax', 'بدون ضريبة': 'b-notax', 'ورقة مكتوبة': 'b-paper' }[k] || 'b-notax'; return k ? `<span class="badge ${c}">${esc(k)}</span>` : ''; }
function statusBadge(s, dupOf, x) {
  if (s === 'مرفق') { const p = x && S.invoices.find(v => v.id === x.attach_to); return `<span class="badge b-att">📎 مرفق${p ? ' للفاتورة ' + esc(p.code) : ''}</span>`; }
  if (s === 'مكررة') return `<span class="badge b-dup">مكررة${dupOf ? ' من ' + esc(dupOf) : ''}</span>`;
  if (s === 'تحتاج مراجعة') return '<span class="badge b-review">تحتاج مراجعة</span>';
  return '';
}
const isTaalumTax = x => x.kind === 'ضريبية' && x.taalum === 'نعم'; // الفواتير الضريبية المكتوب عليها جمعية تعلم أو المركز الخيري لتعليم القرآن وعلومه
function logoSrc(side) { const v = S.settings['logo' + (side === 'right' ? 'Right' : 'Left')] || S.boot?.['logo' + (side === 'right' ? 'Right' : 'Left')]; return v || (side === 'right' ? window.DEFAULT_LOGO_RIGHT : window.DEFAULT_LOGO_LEFT); }
function thumbUrl(inv, size = 400) { return inv.file_id ? `https://drive.google.com/thumbnail?id=${inv.file_id}&sz=w${size}` : ''; }
function isPdf(inv) { return /pdf/i.test(inv.mime || ''); }
function fileB64(file) { return new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(',')[1]); r.onerror = no; r.readAsDataURL(file); }); }
async function sha256(buf) { try { const h = await crypto.subtle.digest('SHA-256', buf); return [...new Uint8Array(h)].map(b => b.toString(16).padStart(2, '0')).join(''); } catch (e) { return ''; } }
// تصغير الصورة قبل الرفع لتسريعه (لا يغيّر وضوح الفاتورة)
async function prepareFile(file) {
  const buf = await file.arrayBuffer();
  const hash = await sha256(buf);
  if (/pdf/i.test(file.type)) {
    if (file.size > 12 * 1024 * 1024) throw new Error('ملف PDF أكبر من 12 ميجا');
    return { file: { mime: 'application/pdf', name: file.name, data: await fileB64(file) }, hash, preview: '' };
  }
  const img = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => no(new Error('صيغة الصورة غير مدعومة')); i.src = URL.createObjectURL(file); });
  const max = 2000, k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
  const c = document.createElement('canvas'); c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
  const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
  const url = c.toDataURL('image/jpeg', 0.85);
  return { file: { mime: 'image/jpeg', name: file.name, data: url.split(',')[1] }, hash, preview: img.src };
}
// صورة مصغرة: من Drive مباشرة، وإن تعذّر فمن الخادم
const thumbCache = {};
function imgEl(inv, size = 300) {
  if (!inv.file_id) return '<span>بلا صورة</span>';
  if (thumbCache[inv.id]) return `<img src="${thumbCache[inv.id]}" alt="">`;
  return `<img src="${thumbUrl(inv, size)}" data-id="${esc(inv.id)}" loading="lazy" alt="" onerror="thumbFallback(this)">`;
}
window.thumbFallback = async function (el) {
  el.onerror = null;
  const id = el.dataset.id, big = el.dataset.size === 'big', key = big ? id + '_big' : id;
  try {
    if (!thumbCache[key]) { const d = await api('fileData', { id, thumb: true, size: big ? 1600 : 400 }); thumbCache[key] = `data:${d.mime};base64,${d.data}`; }
    el.src = thumbCache[key];
  } catch (e) { el.replaceWith(Object.assign(document.createElement('span'), { textContent: 'PDF' })); }
};

// النافذة العامة
const modal = {
  open(html, cls) { const b = $('#modalBox'); b.innerHTML = html; b.className = 'modal-box ' + (cls || ''); const d = $('#modal'); if (!d.open) d.showModal(); return b; },
  close() { const d = $('#modal'); if (d.open) d.close(); }
};
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal' || e.target.closest('[data-close]')) modal.close(); });
function head(title) { return `<div class="m-head"><h2>${title}</h2><button class="x" data-close aria-label="إغلاق">×</button></div>`; }
function confirmBox(msg, okText = 'تأكيد', danger) {
  return new Promise(res => {
    const b = modal.open(head('تأكيد') + `<div class="m-body"><p>${msg}</p></div><div class="m-foot"><button class="btn ${danger ? 'danger' : 'primary'}" id="cOk">${okText}</button><button class="btn" data-close>إلغاء</button></div>`);
    $('#cOk', b).onclick = () => { modal.close(); res(true); };
    $('#modal').addEventListener('close', () => res(false), { once: true });
  });
}
function busy(btn, on, text) { if (!btn) return; if (on) { btn._t = btn.innerHTML; btn.disabled = true; btn.innerHTML = text || '…'; } else { btn.disabled = false; if (btn._t) btn.innerHTML = btn._t; } }

// =====================================================================
// ٢) الدخول
// =====================================================================
function showAuth(which, msg) {
  $('#appView').hidden = true; $('#authView').hidden = false;
  ['apiForm', 'firstForm', 'loginForm', 'registerForm'].forEach(f => $('#' + f).hidden = f !== which);
  $('#toRegister').hidden = !(S.boot && S.boot.allowRegister);
  $('#authMsg').textContent = msg || '';
  $('#authMsg').style.color = '';
}
function applyBrand() {
  const name = S.settings.appName || S.boot?.appName || 'مصروفات مجمع ابن تيمية';
  $$('[data-app-name]').forEach(e => e.textContent = name);
  document.title = name;
  $$('img[data-logo]').forEach(i => i.src = logoSrc(i.dataset.logo));
}
async function start() {
  applyBrand();
  if (!apiUrl()) return showAuth('apiForm');
  try { S.boot = await api('boot'); } catch (e) { return showAuth('apiForm', 'تعذّر الاتصال بالخادم: ' + e.message); }
  applyBrand();
  if (S.boot.needsSetup) return showAuth('firstForm');
  if (!S.token) return showAuth('loginForm');
  try { S.user = await api('me'); localStorage.setItem(LS.user, JSON.stringify(S.user)); } catch (e) { return showAuth('loginForm', e.message); }
  enterApp();
}
function onAuthSubmit(form, action, after) {
  $(form).addEventListener('submit', async e => {
    e.preventDefault();
    const btn = $('button.primary', e.target); busy(btn, true, 'جارٍ…');
    try {
      const d = Object.fromEntries(new FormData(e.target));
      const r = await api(action, d);
      after(r, e.target);
    } catch (err) { $('#authMsg').textContent = err.message; }
    finally { busy(btn, false); }
  });
}
$('#apiForm').addEventListener('submit', e => { e.preventDefault(); const u = e.target.url.value.trim(); if (!/^https:\/\/script\.google\.com\/.+\/exec/.test(u)) { $('#authMsg').textContent = 'الرابط يجب أن يبدأ بـ https://script.google.com وينتهي بـ /exec'; return; } localStorage.setItem(LS.api, u); start(); });
const signedIn = r => { S.token = r.token; S.user = r.user; localStorage.setItem(LS.token, r.token); localStorage.setItem(LS.user, JSON.stringify(r.user)); enterApp(); };
onAuthSubmit('#firstForm', 'firstAdmin', signedIn);
onAuthSubmit('#loginForm', 'login', signedIn);
onAuthSubmit('#registerForm', 'register', (r, f) => { f.reset(); showAuth('loginForm'); $('#authMsg').style.color = 'var(--green)'; $('#authMsg').textContent = r.msg; });
$('#toRegister').onclick = () => showAuth('registerForm');
$('#toLogin').onclick = () => showAuth('loginForm');
function logout(expired) {
  S.token = ''; S.user = null; localStorage.removeItem(LS.token); localStorage.removeItem(LS.user);
  modal.close(); showAuth('loginForm', expired ? 'انتهت الجلسة، سجّل الدخول من جديد' : '');
}

async function enterApp() {
  $('#authView').hidden = true; $('#appView').hidden = false;
  $('#whoami').textContent = S.user.name;
  S.invoices = []; S.report = null; Q.items = []; renderQueue();
  $('#reportWrap').innerHTML = '<p class="muted center">اختر نوع الكشف ثم اضغط «عرض الكشف».</p>';
  $('#reportNote').innerHTML = '';
  $('#btnPrint').disabled = $('#btnExcel').disabled = true;
  switchTab('list');
  $$('[data-perm]').forEach(el => el.hidden = !can(el.dataset.perm));
  try { S.settings = await api('getSettings'); } catch (e) { toast(e.message, true); }
  applyBrand(); fillCatSelects(); restoreReportPrefs();
  await loadInvoices();
  if (can('admin')) refreshPendingBadge();
}
function fillCatSelects() {
  const opts = cats().map(c => `<option>${esc(c)}</option>`).join('');
  [['#fCat', 'كل التصنيفات'], ['#upCat', 'تلقائي (يقترحه البرنامج)'], ['#rCat', 'كل التصنيفات']].forEach(([s, first]) => {
    const el = $(s), v = el.value; el.innerHTML = `<option value="">${first}</option>` + opts; el.value = v;
  });
}

// التبويبات
$('#tabs').addEventListener('click', e => { const b = e.target.closest('button[data-tab]'); if (b) switchTab(b.dataset.tab); });
function switchTab(t) {
  if (t === 'list' && S.invoices.length) renderList();
  $$('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  $$('.view').forEach(v => v.hidden = v.id !== 'view-' + t);
  if (t === 'log') loadLog();
}

// قائمة الحساب
$('#btnUser').onclick = e => {
  e.stopPropagation();
  const old = $('.menu'); if (old) return old.remove();
  const m = document.createElement('div'); m.className = 'menu';
  m.innerHTML = `<div class="who">${esc(S.user.name)} — ${esc(S.user.username)}</div><button data-a="pw">تغيير كلمة المرور</button><button data-a="out">تسجيل الخروج</button>`;
  document.body.appendChild(m);
  m.onclick = ev => { const a = ev.target.dataset.a; m.remove(); if (a === 'out') logout(); if (a === 'pw') changePwDialog(); };
  setTimeout(() => document.addEventListener('click', () => m.remove(), { once: true }));
};
function changePwDialog() {
  const b = modal.open(head('تغيير كلمة المرور') + `<form class="m-body stack" id="pwF"><label>كلمة المرور الحالية<input type="password" name="oldPassword" required></label><label>كلمة المرور الجديدة<input type="password" name="newPassword" minlength="6" required></label><button class="btn primary">حفظ</button></form>`);
  $('#pwF', b).onsubmit = async e => {
    e.preventDefault();
    try { const r = await api('changePassword', Object.fromEntries(new FormData(e.target))); S.token = r.token; localStorage.setItem(LS.token, r.token); modal.close(); toast('تم تغيير كلمة المرور'); }
    catch (err) { toast(err.message, true); }
  };
}

// =====================================================================
// ٣) قائمة الفواتير
// =====================================================================
async function loadInvoices() {
  $('#invList').innerHTML = '<div class="empty">جارٍ التحميل…</div>';
  try { S.invoices = await api('list'); } catch (e) { $('#invList').innerHTML = `<div class="empty">${esc(e.message)}</div>`; return; }
  renderList();
}
// الترتيب حسب تاريخ الفاتورة: الأقدم في الأعلى (افتراضياً) أو الأحدث في الأعلى، والفواتير بلا تاريخ في الآخر
function sortInv(a) {
  const dir = $('#fSort').value === 'desc' ? -1 : 1;
  return a.sort((x, y) => {
    if (!x.date !== !y.date) return x.date ? -1 : 1;
    return dir * ((x.date || '').localeCompare(y.date || '') || (x.created_at || '').localeCompare(y.created_at || ''));
  });
}
function filtered() {
  const q = $('#fSearch').value.trim().toLowerCase(), k = $('#fKind').value, c = $('#fCat').value, st = $('#fStatus').value, f = $('#fFrom').value, t = $('#fTo').value;
  return sortInv(S.invoices.filter(x =>
    (!k || (k === 'taalum' ? isTaalumTax(x) : x.kind === k)) && (!c || x.category === c) && (!st || (x.status || 'سليمة') === st) &&
    (!f || (x.date && x.date >= f)) && (!t || (x.date && x.date <= t)) &&
    (!q || [x.vendor, x.description, x.number, x.code, x.vat_no, x.notes, x.category, String(x.total)].join(' ').toLowerCase().includes(q))
  ));
}
function renderList() {
  const un = unread(S.invoices);
  const rb = $('#btnRereadAll');
  if (rb && !REREADING) { rb.hidden = !(un.length && can('edit')); rb.textContent = `↻ إعادة قراءة ما لم يُقرأ (${un.length})`; }
  const list = filtered();
  const valid = list.filter(x => x.status !== 'مكررة' && x.status !== 'مرفق');
  const sum = valid.reduce((a, x) => a + Number(x.total || 0), 0);
  const vat = valid.reduce((a, x) => a + Number(x.vat || 0), 0);
  const dups = list.filter(x => x.status === 'مكررة').length, rev = list.filter(x => x.status === 'تحتاج مراجعة').length;
  const cur = esc(S.settings.currency || 'ريال');
  $('#stats').innerHTML = `
    <div class="stat main"><small>الإجمالي (بدون المكررة)</small><b>${money(sum)}</b> <small style="display:inline">${cur}</small></div>
    <div class="stat"><small>عدد الفواتير</small><b>${valid.length}</b></div>
    <div class="stat"><small>إجمالي الضريبة</small><b>${money(vat)}</b></div>
    <div class="stat warn"><small>تحتاج مراجعة</small><b>${rev}</b></div>
    <div class="stat dup"><small>مكررة مستبعدة</small><b>${dups}</b></div>`;
  if (!list.length) { $('#invList').innerHTML = `<div class="empty">${S.invoices.length ? 'لا توجد فواتير بهذه الشروط' : 'لا توجد فواتير بعد — ابدأ من تبويب «رفع فواتير»'}</div>`; return; }
  const shown = list.slice(0, 400);
  $('#invList').innerHTML = shown.map(x => `
    <div class="inv ${x.status === 'مكررة' ? 'is-dup' : ''}" data-id="${esc(x.id)}">
      <div class="thumb">${imgEl(x)}</div>
      <div class="inv-main">
        <div class="inv-title"><span class="code">${esc(x.code)}</span><span>${esc(x.vendor || 'بدون اسم مؤسسة')}</span>${kindBadge(x.kind)}${x.taalum === 'نعم' ? '<span class="badge b-taalum">جمعية تعلم / المركز الخيري</span>' : ''}${statusBadge(x.status, x.dup_of, x)}</div>
        <div class="inv-desc">${esc(x.description || '—')}${x.category ? ' · <b>' + esc(x.category) + '</b>' : ''}</div>
        ${reviewReason(x) ? `<div class="inv-why">⚠️ ${esc(reviewReason(x))}</div>` : ''}
        <div class="inv-sub"><span>📅 ${esc(showDate(x.date) || 'بلا تاريخ')}</span>${x.number ? `<span>رقم: ${esc(x.number)}</span>` : ''}<span>أدخلها ${esc(x.created_by)} — ${esc(x.created_at)}</span>${x.updated_by ? `<span>عدّلها ${esc(x.updated_by)} — ${esc(x.updated_at)}</span>` : ''}</div>
      </div>
      <div class="amount">${x.status === 'مرفق' ? '<small>بدون مبلغ</small>' : money(x.total)}<small>${Number(x.vat) ? 'ضريبة ' + money(x.vat) : ''}</small></div>
    </div>`).join('') + (list.length > shown.length ? `<p class="muted center">يظهر أول ${shown.length} من ${list.length} — استخدم البحث أو الفلاتر.</p>` : '');
}
['#fSearch', '#fKind', '#fCat', '#fStatus', '#fFrom', '#fTo'].forEach(s => $(s).addEventListener('input', renderList));
try { const v = localStorage.getItem('mx_sort'); if (v) $('#fSort').value = v; } catch (e) { }
$('#fSort').addEventListener('change', () => { try { localStorage.setItem('mx_sort', $('#fSort').value); } catch (e) { } renderList(); });
$('#fClear').onclick = () => { ['#fSearch', '#fKind', '#fCat', '#fStatus', '#fFrom', '#fTo'].forEach(s => $(s).value = ''); renderList(); };
$('#btnRefresh').onclick = loadInvoices;
$('#btnRereadAll').onclick = e => rereadMany(unread(S.invoices).map(x => x.id), e.currentTarget);
$('#invList').addEventListener('click', e => { const r = e.target.closest('.inv'); if (r) openInvoice(r.dataset.id); });

// =====================================================================
// ٤) رفع الفواتير
// =====================================================================
const drop = $('#drop');
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => enqueue([...e.dataTransfer.files]));
$('#fileInput').onchange = e => { enqueue([...e.target.files]); e.target.value = ''; };
$('#camInput').onchange = e => { enqueue([...e.target.files]); e.target.value = ''; };
$('#btnManual').onclick = () => openInvoice(null);

const Q = { items: [], running: 0, MAX: 1 }; // فاتورة واحدة في كل مرة حتى لا نتجاوز حد Gemini المجاني
const isPdfFile = f => /pdf$/i.test(f.type) || /\.pdf$/i.test(f.name);
async function enqueue(files) {
  files = files.filter(f => /^image\//i.test(f.type) || isPdfFile(f) || /\.(jpe?g|png|webp|heic)$/i.test(f.name));
  if (!files.length) return toast('اختر صوراً أو ملفات PDF', true);
  const opts = { kind: $('input[name=upKind]:checked').value, cat: $('#upCat').value, ai: $('#upAi').checked };
  const split = $('input[name=pdfMode]:checked').value === 'pages';
  const all = [];
  for (const f of files) {
    if (split && isPdfFile(f)) {
      try { all.push(...await pdfToPages(f)); }
      catch (e) { toast(`تعذّر تقسيم «${f.name}» إلى صفحات، فرُفع كفاتورة واحدة`, true); all.push(f); }
    } else all.push(f);
  }
  all.forEach(f => Q.items.unshift({ id: Math.random().toString(36).slice(2), file: f, opts, state: 'wait', msg: 'في الانتظار…' }));
  renderQueue(); pump();
}
// تحويل كل صفحة من ملف PDF إلى صورة مستقلة (فاتورة لكل صفحة)
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  const base = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/';
  return new Promise((ok, no) => {
    const sc = document.createElement('script'); sc.src = base + 'pdf.min.js';
    sc.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = base + 'pdf.worker.min.js'; ok(window.pdfjsLib); };
    sc.onerror = () => no(new Error('تعذّر تحميل قارئ PDF'));
    document.head.appendChild(sc);
  });
}
async function pdfToPages(file) {
  const lib = await loadPdfJs();
  const pdf = await lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const base = file.name.replace(/\.pdf$/i, '');
  if (pdf.numPages > 1) toast(`«${base}»: ${pdf.numPages} صفحات — كل صفحة فاتورة مستقلة`);
  const out = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const v1 = page.getViewport({ scale: 1 });
    const scale = Math.min(3, 2000 / Math.max(v1.width, v1.height));
    const vp = page.getViewport({ scale });
    const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.88));
    out.push(new File([blob], `${base} - صفحة ${n}.jpg`, { type: 'image/jpeg' }));
    page.cleanup && page.cleanup();
  }
  pdf.destroy && pdf.destroy();
  return out;
}
function pump() {
  while (Q.running < Q.MAX) {
    const it = Q.items.slice().reverse().find(i => i.state === 'wait');
    if (!it) break;
    Q.running++; processItem(it).finally(() => { Q.running--; pump(); if (!Q.running && !Q.items.some(i => i.state === 'wait')) loadInvoices(); });
  }
}
async function processItem(it) {
  try {
    it.state = 'busy'; it.msg = 'تجهيز الصورة…'; renderQueue();
    const p = await prepareFile(it.file); it.preview = p.preview;
    it.msg = it.opts.ai ? 'رفع وقراءة الفاتورة…' : 'رفع…'; renderQueue();
    const r = await api('add', { file: p.file, fileHash: p.hash, useAi: it.opts.ai, forceKind: it.opts.kind, forceCategory: it.opts.cat });
    it.inv = r; it.state = r.status === 'مكررة' ? 'dup' : r.status === 'تحتاج مراجعة' ? 'review' : 'ok';
    it.msg = r.status === 'مكررة' ? `مكررة (${r._dupReason}) — أصلها ${r.dup_of}` :
      r._aiError ? 'حُفظت — تعذّرت القراءة الآلية: ' + r._aiError :
        r.status === 'تحتاج مراجعة' ? 'حُفظت — أكمل البيانات الناقصة' : 'تم الحفظ';
    upsertLocal(r);
    // إن تعذّرت القراءة بسبب ازدحام أو حد مؤقت: نعيد المحاولة تلقائياً
    if (r._aiError && r.status !== 'مكررة' && can('edit')) {
      for (let k = 1; k <= 3; k++) {
        const wait = 15 * k;
        for (let sec = wait; sec > 0; sec--) { it.msg = `تعذّرت القراءة (${r._aiError}) — إعادة المحاولة تلقائياً بعد ${sec} ث (${k} من 3)`; renderQueue(); await sleep(1000); }
        it.msg = 'إعادة القراءة…'; renderQueue();
        try {
          const r2 = await api('reread', { id: r.id, save: true });
          upsertLocal(r2); it.inv = r2;
          it.state = r2.status === 'مكررة' ? 'dup' : r2.status === 'تحتاج مراجعة' ? 'review' : 'ok';
          it.msg = r2.status === 'تحتاج مراجعة' ? 'قُرئت — أكمل البيانات الناقصة' : 'قُرئت وحُفظت ✓';
          break;
        } catch (e2) { r._aiError = e2.message; it.msg = 'تعذّرت القراءة: ' + e2.message; }
      }
    }
    // المرحلة الأخيرة: إن بقيت ناقصة نكبّر الصورة ونقرأ أجزاءها
    if (it.opts.ai && incomplete(it.inv) && !it.inv.ai_error && can('edit') && it.inv.file_id) {
      it.state = 'busy'; it.msg = 'قراءة متأنية بتكبير الصورة…'; renderQueue();
      try {
        const z = await zoomRead(it.inv.id, p.file.mime, p.file.data);
        it.inv = z; it.state = z.status === 'مكررة' ? 'dup' : incomplete(z) ? 'review' : 'ok';
        it.msg = incomplete(z) ? 'حُفظت — لم تتضح كل البيانات، أكملها يدوياً' : 'قُرئت بعد التكبير ✓';
      } catch (e3) { it.state = 'review'; it.msg = 'حُفظت — تعذّرت القراءة بالتكبير: ' + e3.message; }
    }
  } catch (e) { it.state = 'err'; it.msg = e.message; }
  renderQueue();
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
// ---------- القراءة بالتكبير: للفواتير التي لم تُقرأ كاملة ----------
const incomplete = x => x && x.status === 'تحتاج مراجعة' && (!Number(x.total) || !x.date);
async function sourceImage(mime, data) {
  if (/pdf/i.test(mime)) {
    const bytes = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    const pages = await pdfToPages(new File([bytes], 'f.pdf', { type: 'application/pdf' }));
    data = await fileB64(pages[0]); mime = 'image/jpeg';
  }
  return await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => no(new Error('تعذّر فتح الصورة')); i.src = `data:${mime};base64,${data}`; });
}
// مقاطع مكبّرة ومحسّنة الوضوح: أعلى الفاتورة ووسطها وأسفلها، ونسخة كاملة بتباين أعلى
function zoomCrops(img) {
  const W = img.naturalWidth, H = img.naturalHeight, out = [];
  const regions = H > W * 1.1 ? [[0, 0, 1, 0.42], [0, 0.3, 1, 0.72], [0, 0.58, 1, 1]] : [[0, 0, 0.55, 1], [0.45, 0, 1, 1]];
  const draw = (sx, sy, sw, sh, enhance) => {
    const k = Math.min(4, 1568 / Math.max(sw, sh));
    const c = document.createElement('canvas'); c.width = Math.round(sw * k); c.height = Math.round(sh * k);
    const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
    x.imageSmoothingQuality = 'high';
    if (enhance) x.filter = 'grayscale(1) contrast(1.6) brightness(1.08)';
    x.drawImage(img, sx, sy, sw, sh, 0, 0, c.width, c.height);
    return { mime: 'image/jpeg', data: c.toDataURL('image/jpeg', 0.9).split(',')[1] };
  };
  out.push(draw(0, 0, W, H, true));
  regions.forEach(([a, b, c2, d]) => out.push(draw(W * a, H * b, W * (c2 - a), H * (d - b), true)));
  return out;
}
async function zoomRead(id, mime, data) {
  if (!data) { const d = await api('fileData', { id }); mime = d.mime; data = d.data; }
  const img = await sourceImage(mime, data);
  const r = await api('reread', { id, save: true, extras: zoomCrops(img) });
  upsertLocal(r);
  return r;
}
// لماذا تحتاج الفاتورة مراجعة؟
function reviewReason(x) {
  if (x.status !== 'تحتاج مراجعة') return '';
  if (x.ai_error) return 'لم تُقرأ: ' + x.ai_error;
  const miss = [];
  if (!Number(x.total)) miss.push('المبلغ');
  if (!x.date) miss.push('التاريخ');
  return miss.length ? 'ينقصها: ' + miss.join(' و') : '';
}
function upsertLocal(r) {
  const clean = Object.fromEntries(Object.entries(r).filter(([k]) => !k.startsWith('_')));
  const i = S.invoices.findIndex(x => x.id === clean.id);
  if (i >= 0) S.invoices[i] = clean; else S.invoices.push(clean);
  if (!$("#view-list").hidden) renderList();
}
// فواتير لها صورة لكنها لم تُقرأ (بلا مبلغ)
function unread(list) { return list.filter(x => x.file_id && x.status !== 'مكررة' && x.status !== 'مرفق' && (!Number(x.total) || !x.date || x.ai_error)); }
let REREADING = false;
async function rereadMany(ids, btn) {
  if (REREADING || !ids.length) return;
  if (!can('edit')) return toast('إعادة القراءة تحتاج صلاحية التعديل', true);
  REREADING = true; let ok = 0, fail = 0, lastErr = '';
  try {
    for (let i = 0; i < ids.length; i++) {
      if (btn) btn.textContent = `جارٍ القراءة ${i + 1} من ${ids.length}…`;
      try {
        if (i) await sleep(3000);
        let r = await api('reread', { id: ids[i], save: true });
        upsertLocal(r);
        if (incomplete(r) && r.file_id) {
          if (btn) btn.textContent = `قراءة متأنية بالتكبير ${i + 1} من ${ids.length}…`;
          try { r = await zoomRead(ids[i]); } catch (e) { }
        }
        renderQueueSync(r, incomplete(r) ? 'لم تتضح كل البيانات' : 'قُرئت من جديد'); if (incomplete(r)) { fail++; lastErr = 'بعضها يحتاج إكمالاً يدوياً'; } else ok++;
      } catch (e) { fail++; lastErr = e.message; }
    }
  } finally { REREADING = false; }
  renderList(); renderQueue();
  toast(fail ? `قُرئت ${ok}، وتعذّرت ${fail}: ${lastErr}` : `تمت قراءة ${ok} فاتورة ✓`, !!fail);
}
function renderQueue() {
  if (!Q.items.length) { $('#queue').innerHTML = ''; return; }
  const done = Q.items.filter(i => !['wait', 'busy'].includes(i.state)).length;
  const qUnread = done === Q.items.length ? unread(Q.items.filter(i => i.inv).map(i => S.invoices.find(x => x.id === i.inv.id) || i.inv)) : [];
  $('#queue').innerHTML = `<div class="row gap" style="justify-content:space-between"><b>المرفوعة: ${done} من ${Q.items.length}</b><span class="row gap">${qUnread.length && can('edit') ? `<button class="btn primary sm" id="qReread">↻ إعادة قراءة ما لم يُقرأ (${qUnread.length})</button>` : ''}${done === Q.items.length ? '<button class="btn ghost sm" id="qClear">مسح القائمة</button>' : ''}</span></div>` +
    Q.items.map(it => {
      const inv = it.inv;
      const color = { ok: 'var(--green)', dup: 'var(--red)', review: 'var(--amber)', err: 'var(--red)' }[it.state] || 'var(--muted)';
      return `<div class="q-item ${it.state === 'err' ? 'err' : ''}" data-q="${it.id}">
        <div class="thumb">${it.preview ? `<img src="${it.preview}" alt="">` : 'PDF'}</div>
        <div><div><b>${esc(inv ? (inv.code + ' — ' + (inv.vendor || 'بدون اسم')) : it.file.name)}</b> ${inv ? kindBadge(inv.kind) + statusBadge(inv.status, inv.dup_of, inv) : ''}</div>
          ${inv ? `<div class="small muted">${esc(inv.description || '')} ${inv.category ? '· ' + esc(inv.category) : ''} ${inv.date ? '· ' + esc(showDate(inv.date)) : ''}</div>` : ''}
          <div class="q-state" style="color:${color}">${esc(it.msg)}</div>
          ${it.state === 'busy' ? '<div class="bar busy"><i></i></div>' : ''}</div>
        <div class="row gap">${inv ? `<b class="amount">${money(inv.total)}</b>${it.state === 'review' && inv.file_id && !Number(inv.total) && can('edit') ? `<button class="btn sm" data-reread="${esc(inv.id)}">↻ إعادة القراءة</button>` : ''}<button class="btn sm" data-open="${esc(inv.id)}">مراجعة / تعديل</button>` : ''}${it.state === 'err' ? `<button class="btn sm" data-retry="${it.id}">إعادة</button>` : ''}</div>
      </div>`;
    }).join('');
}
$('#queue').addEventListener('click', e => {
  const o = e.target.closest('[data-open]'); if (o) return openInvoice(o.dataset.open);
  const r = e.target.closest('[data-retry]'); if (r) { const it = Q.items.find(i => i.id === r.dataset.retry); it.state = 'wait'; it.msg = 'في الانتظار…'; renderQueue(); pump(); }
  if (e.target.id === 'qClear') { Q.items = []; renderQueue(); }
  const rr = e.target.closest('[data-reread]'); if (rr) rereadMany([rr.dataset.reread], rr);
  if (e.target.id === 'qReread') rereadMany(unread(Q.items.filter(i => i.inv).map(i => S.invoices.find(x => x.id === i.inv.id) || i.inv)).map(x => x.id), e.target);
});

// =====================================================================
// ٥) نافذة الفاتورة (عرض / تعديل / إدخال يدوي)
// =====================================================================
async function openInvoice(id) {
  const inv = id ? S.invoices.find(x => x.id === id) : { kind: 'ورقة مكتوبة', date: ymd(new Date()) };
  if (!inv) return toast('الفاتورة غير موجودة', true);
  const isNew = !id, editable = isNew ? can('add') : can('edit');
  const catOpts = ['', ...cats()].map(c => `<option ${c === inv.category ? 'selected' : ''} value="${esc(c)}">${c ? esc(c) : '— اختر —'}</option>`).join('');
  const kindOpts = KINDS.map(k => `<option ${k === inv.kind ? 'selected' : ''}>${k}</option>`).join('');
  const dis = editable ? '' : 'disabled';
  const b = modal.open(head(isNew ? 'إدخال يدوي' : `الفاتورة ${esc(inv.code)}`) + `
    <div class="m-body">
      ${inv.status === 'مكررة' ? `<div class="dup-box"><span>⚠️ هذه الفاتورة <b>مكررة</b> من الفاتورة ذات الرمز <b>${esc(inv.dup_of)}</b>، ومستبعدة من الإجماليات والكشوف.</span>${can('edit') ? '<span class="row gap"><button class="btn sm" id="asAtt">📎 مرفق للفاتورة الأصلية</button><button class="btn sm" id="notDup">ليست مكررة</button></span>' : ''}</div>` : ''}
      ${inv.status === 'مرفق' ? `<div class="att-box"><span>📎 هذه <b>مرفق</b> للفاتورة <b>${esc((S.invoices.find(v => v.id === inv.attach_to) || {}).code || '')}</b> — بدون مبلغ، وتظهر صورتها مع فاتورتها في الكشوف.</span>${can('edit') ? '<button class="btn sm" id="unAtt">ليست مرفقاً</button>' : ''}</div>` : ''}
      ${!isNew && can('edit') && inv.status !== 'مرفق' && inv.status !== 'مكررة' ? `<details class="att-pick"><summary>📎 جعلها مرفقاً لفاتورة أخرى</summary><div class="row gap"><select id="attParent"><option value="">— اختر الفاتورة الأصلية —</option>${sortInv(S.invoices.filter(v => v.id !== inv.id && v.status !== 'مرفق' && v.status !== 'مكررة').slice()).map(v => `<option value="${esc(v.id)}">${esc(v.code)} — ${esc(v.vendor || 'بدون اسم')} — ${money(v.total)}</option>`).join('')}</select><button class="btn sm" id="doAtt">تأكيد</button></div></details>` : ''}
      ${inv.not_dup === 'نعم' ? '<p class="note">أكّد مستخدم أن هذه الفاتورة ليست مكررة.</p>' : ''}
      <div class="edit-grid">
        <form id="invF" class="form-grid">
          <label>المبلغ (الإجمالي)<input name="total" inputmode="decimal" value="${esc(inv.total)}" ${dis} required></label>
          <label>نوع المستند<select name="kind" ${dis}>${kindOpts}</select></label>
          <label>المبلغ قبل الضريبة<input name="subtotal" inputmode="decimal" value="${esc(inv.subtotal)}" ${dis}></label>
          <label>الضريبة<input name="vat" inputmode="decimal" value="${esc(inv.vat)}" ${dis}></label>
          <label>التصنيف<select name="category" ${dis}>${catOpts}</select></label>
          <label>تاريخ الفاتورة<input type="date" name="date" value="${esc(inv.date)}" ${dis}><small class="muted" id="hijri">${esc(hDate(inv.date))}</small></label>
          <label class="full">بيان الصرف<input name="description" value="${esc(inv.description)}" ${dis}></label>
          <label class="full">اسم المؤسسة (عملية شراء إن وجد)<input name="vendor" value="${esc(inv.vendor)}" ${dis}></label>
          <label>رقم الفاتورة<input name="number" dir="ltr" value="${esc(inv.number)}" ${dis}></label>
          <label>الرقم الضريبي للمورد<input name="vat_no" dir="ltr" value="${esc(inv.vat_no)}" ${dis}></label>
          <label>الرمز (تلقائي حسب التاريخ)<input value="${esc(inv.code || 'يُحدَّد بعد الحفظ')}" disabled></label>
          <label>ملاحظات<input name="notes" value="${esc(inv.notes)}" ${dis}></label>
          <label>مكتوب عليها جمعية تعلم أو المركز الخيري<select name="taalum" ${dis}>${[['', 'لم تُفحص'], ['نعم', 'نعم'], ['لا', 'لا']].map(([v, t]) => `<option value="${v}" ${v === (inv.taalum || '') ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
        </form>
        <div>
          <div class="viewer" id="viewer">${inv.file_id ? (isPdf(inv) ? `<a class="btn" href="${esc(inv.file_url)}" target="_blank" rel="noopener">فتح ملف PDF</a>` : `<img src="${thumbUrl(inv, 1600)}" data-id="${esc(inv.id)}" data-size="big" onerror="thumbFallback(this)" alt="صورة الفاتورة">`) : '<span class="muted">لا توجد صورة</span>'}</div>
          <div class="row gap" style="margin-top:8px">
            ${inv.file_url ? `<a class="btn sm ghost" href="${esc(inv.file_url)}" target="_blank" rel="noopener">فتح الأصل في Drive</a>` : ''}
            ${editable ? `<label class="btn sm ghost"><input type="file" id="newFile" accept="image/*,application/pdf" hidden>${inv.file_id ? 'استبدال الصورة' : 'إرفاق صورة'}</label>` : ''}
            ${!isNew && inv.file_id && can('edit') ? '<button class="btn sm ghost" id="reread">إعادة القراءة الآلية</button>' : ''}
          </div>
          <div id="newFileName" class="small muted"></div>
          ${isNew ? '' : `<div class="meta-line"><span>أدخلها: <b>${esc(inv.created_by)}</b> — ${esc(inv.created_at)}</span>${inv.updated_by ? `<span>آخر تعديل: <b>${esc(inv.updated_by)}</b> — ${esc(inv.updated_at)}</span>` : ''}</div>
          <div class="history" id="hist"><span class="muted">سجل التعديلات…</span></div>`}
        </div>
      </div>
    </div>
    <div class="m-foot">
      ${editable ? `<button class="btn primary" id="save">${isNew ? 'حفظ' : 'حفظ التعديل'}</button>` : ''}
      <button class="btn" data-close>إغلاق</button>
      <span class="spacer"></span>
      ${!isNew && can('delete') ? '<button class="btn danger" id="del">حذف الفاتورة</button>' : ''}
    </div>`);
  const F = $('#invF', b);
  F.date.addEventListener('input', () => $('#hijri', b).textContent = hDate(F.date.value));
  // حساب الضريبة تلقائياً عند تغيير المبلغ
  const rate = () => (Number(S.settings.vatRate) || 15) / 100;
  F.total.addEventListener('change', () => { if (F.kind.value === 'ضريبية' && F.total.value) { const t = Number(F.total.value); F.subtotal.value = round2(t / (1 + rate())); F.vat.value = round2(t - Number(F.subtotal.value)); } else if (F.kind.value !== 'ضريبية') { F.subtotal.value = F.total.value; F.vat.value = 0; } });
  F.kind.addEventListener('change', () => F.total.dispatchEvent(new Event('change')));
  $('#viewer img', b)?.addEventListener('click', e => e.currentTarget.parentElement.classList.toggle('zoom'));
  let newFile = null;
  $('#newFile', b)?.addEventListener('change', async e => {
    const f = e.target.files[0]; if (!f) return;
    try { newFile = await prepareFile(f); $('#newFileName', b).textContent = 'سيُرفع: ' + f.name; if (newFile.preview) $('#viewer', b).innerHTML = `<img src="${newFile.preview}" alt="">`; }
    catch (err) { toast(err.message, true); }
  });
  if (!isNew) api('history', { id }).then(h => { const el = $('#hist', b); if (el) el.innerHTML = h.length ? '<b>سجل الفاتورة</b>' + h.map(x => `<div><b>${esc(x.action)}</b> — ${esc(x.user)} — ${esc(x.at)}${x.details ? `<br><span class="muted">${esc(x.details)}</span>` : ''}</div>`).join('') : ''; }).catch(() => { });
  $('#reread', b)?.addEventListener('click', async e => {
    busy(e.target, true, 'جارٍ القراءة…');
    try {
      const ai = await api('reread', { id });
      let n = 0; Object.entries(ai).forEach(([k, v]) => { if (F[k] && v !== '' && v != null && String(F[k].value) !== String(v)) { F[k].value = v; F[k].style.background = '#fff7d6'; n++; } });
      toast(n ? `حُدّث ${n} حقلاً (باللون الأصفر) — راجعها ثم احفظ` : 'لم تتغير البيانات');
    } catch (err) { toast(err.message, true); } finally { busy(e.target, false); }
  });
  const attach = async (parentId) => {
    try { await api('setAttach', { id, parentId }); modal.close(); await loadInvoices(); toast(parentId ? 'أصبحت مرفقاً للفاتورة الأصلية ✓' : 'أصبحت فاتورة مستقلة — أدخل مبلغها'); } catch (err) { toast(err.message, true); }
  };
  $('#asAtt', b)?.addEventListener('click', () => attach(inv.dup_of_id || (S.invoices.find(v => v.code === inv.dup_of) || {}).id));
  $('#unAtt', b)?.addEventListener('click', () => attach(null));
  $('#doAtt', b)?.addEventListener('click', () => { const v = $('#attParent', b).value; if (!v) return toast('اختر الفاتورة الأصلية', true); attach(v); });
  $('#notDup', b)?.addEventListener('click', async () => {
    try { const r = await api('setNotDup', { id }); modal.close(); await loadInvoices(); renderQueueSync(r); toast(r.status === 'مكررة' ? 'ما زالت مطابقة لفاتورة أخرى' : 'أُلغيت علامة التكرار'); } catch (err) { toast(err.message, true); }
  });
  $('#del', b)?.addEventListener('click', async () => {
    if (!(await confirmBox(`حذف الفاتورة <b>${esc(inv.code)}</b> بمبلغ ${money(inv.total)}؟ يبقى أثرها في السجل.`, 'حذف', true))) return;
    try { await api('remove', { id }); modal.close(); await loadInvoices(); toast('حُذفت الفاتورة'); } catch (err) { toast(err.message, true); }
  });
  $('#save', b)?.addEventListener('click', async e => {
    if (!F.reportValidity()) return;
    const data = Object.fromEntries(new FormData(F));
    busy(e.target, true, 'جارٍ الحفظ…');
    try {
      const r = isNew
        ? await api('add', { data, file: newFile?.file, fileHash: newFile?.hash, useAi: false })
        : await api('update', { id, data, file: newFile?.file, fileHash: newFile?.hash });
      modal.close(); await loadInvoices(); renderQueueSync(S.invoices.find(x => x.id === r.id) || r);
      toast(r.status === 'مكررة' ? `حُفظت لكنها مكررة من ${r.dup_of} (${r._dupReason})` : 'تم الحفظ', r.status === 'مكررة');
    } catch (err) { toast(err.message, true); } finally { busy(e.target, false); }
  });
}
function renderQueueSync(r, msg) { const it = Q.items.find(i => i.inv && i.inv.id === r.id); if (it) { it.inv = r; it.state = r.status === 'مكررة' ? 'dup' : r.status === 'تحتاج مراجعة' ? 'review' : 'ok'; it.msg = r.status === 'مكررة' ? `مكررة — أصلها ${r.dup_of}` : (msg || 'عُدّلت'); renderQueue(); } }

// =====================================================================
// ٦) الكشوف
// =====================================================================
const RP = ['rFrom', 'rTo', 'rCat', 'rGroup', 'rImages', 'rVat', 'rVatNo', 'rKindCol', 'rWho', 'rDups', 'rSign', 'rDate'];
function reportOpts() {
  const o = { kind: $('input[name=rKind]:checked').value, logo: $('input[name=rLogo]:checked').value };
  RP.forEach(k => { const el = $('#' + k); o[k] = el.type === 'checkbox' ? el.checked : el.value; });
  return o;
}
function restoreReportPrefs() {
  let p = {}; try { p = JSON.parse(localStorage.getItem(LS.rep) || '{}'); } catch (e) { }
  if (p.v !== 3) delete p.rImages; // صورة الفاتورة تظهر افتراضياً في كل كشف
  RP.forEach(k => { if (k in p && k !== 'rFrom' && k !== 'rTo') { const el = $('#' + k); if (el.type === 'checkbox') el.checked = p[k]; else el.value = p[k]; } });
  if (!p.rDate) $('#rDate').value = S.settings.dateMode || 'ميلادي';
  if (p.logo) { const r = $(`input[name=rLogo][value="${p.logo}"]`); if (r) r.checked = true; }
}
$('#rQuick').addEventListener('click', e => {
  const q = e.target.dataset.q; if (!q) return;
  const d = new Date(), y = d.getFullYear(), m = d.getMonth();
  const set = (a, b) => { $('#rFrom').value = a; $('#rTo').value = b; };
  if (q === 'all') set('', '');
  if (q === 'month') set(ymd(new Date(y, m, 1)), ymd(new Date(y, m + 1, 0)));
  if (q === 'last') set(ymd(new Date(y, m - 1, 1)), ymd(new Date(y, m, 0)));
  if (q === 'year') set(ymd(new Date(y, 0, 1)), ymd(new Date(y, 11, 31)));
});
const inReportScope = (x, o) => (!o.rCat || x.category === o.rCat) && (!o.rFrom || (x.date && x.date >= o.rFrom)) && (!o.rTo || (x.date && x.date <= o.rTo));
function reportData(o) {
  let rows = S.invoices.filter(x =>
    (!o.kind || (o.kind === 'taalum' ? isTaalumTax(x) : x.kind === o.kind)) && inReportScope(x, o) &&
    x.status !== 'مرفق' && (o.rDups || x.status !== 'مكررة'));
  rows.sort((a, b) => (a.date || '9').localeCompare(b.date || '9') || String(a.code).localeCompare(String(b.code), 'ar', { numeric: true }));
  return rows;
}
function reportTitle(o) {
  const base = S.settings.reportTitle || 'مصروفات مجمع ابن تيمية';
  const kindT = { '': 'كشف بجميع الفواتير', 'ضريبية': 'كشف الفواتير الضريبية', 'بدون ضريبة': 'كشف الفواتير بدون ضريبة', 'ورقة مكتوبة': 'كشف الأوراق المكتوبة بدل الفاتورة', taalum: 'كشف الفواتير الضريبية باسم جمعية تعلم أو المركز الخيري لتعليم القرآن وعلومه' }[o.kind];
  return { main: base, sub: kindT + (o.rGroup ? ' — حسب التصنيف' : '') };
}
function columns(o) {
  const c = [{ k: 'idx', t: 'م', cls: 'num' }, { k: 'total', t: 'المبلغ', cls: 'amt' }, { k: 'category', t: 'التصنيف' }, { k: 'description', t: 'بيان الصرف', cls: 'txt' }, { k: 'vendor', t: 'اسم المؤسسة (عملية شراء إن وجد)', cls: 'txt' }, { k: 'number', t: 'رقم الفاتورة', cls: 'num' }, { k: 'date', t: 'تاريخ الفاتورة', cls: 'num' }];
  if (o.rImages) c.push({ k: 'image', t: 'الفاتورة', cls: 'img' });
  c.push({ k: 'code', t: 'الرمز', cls: 'num' });
  if (o.rVat || o.kind === 'taalum') c.splice(2, 0, { k: 'subtotal', t: 'قبل الضريبة', cls: 'amt' }, { k: 'vat', t: 'الضريبة', cls: 'amt' });
  if (o.rVatNo) c.push({ k: 'vat_no', t: 'الرقم الضريبي', cls: 'num' });
  if (o.rKindCol) c.push({ k: 'kind', t: 'نوع المستند' });
  if (o.rWho) c.push({ k: 'who', t: 'أدخلها' });
  return c;
}
function cellText(x, k, i, o) {
  if (k === 'idx') return i;
  if (['total', 'subtotal', 'vat'].includes(k)) return money(x[k]);
  if (k === 'date') return showDate(x.date, o.rDate);
  if (k === 'who') return (x.created_by || '') + (x.created_at ? ' ' + x.created_at.slice(0, 16) : '');
  if (k === 'description') return (x.description || '') + (x.status === 'مكررة' ? ` (مكررة من ${x.dup_of})` : '');
  return x[k] ?? '';
}
function sums(rows) {
  const v = rows.filter(x => x.status !== 'مكررة' && x.status !== 'مرفق');
  return { n: v.length, total: v.reduce((a, x) => a + Number(x.total || 0), 0), vat: v.reduce((a, x) => a + Number(x.vat || 0), 0), sub: v.reduce((a, x) => a + Number(x.subtotal || x.total || 0), 0) };
}
function buildReport() {
  const o = reportOpts();
  try { localStorage.setItem(LS.rep, JSON.stringify({ ...o, v: 3 })); } catch (e) { }
  const rows = reportData(o), cols = columns(o), T = reportTitle(o), cur = S.settings.currency || 'ريال';
  const period = (o.rFrom || o.rTo) ? `الفترة: من ${o.rFrom ? showDate(o.rFrom, o.rDate) : 'البداية'} إلى ${o.rTo ? showDate(o.rTo, o.rDate) : 'اليوم'}` : 'الفترة: جميع الفواتير';
  const L = s => (o.logo === 'both' || o.logo === s) ? `<img src="${esc(logoSrc(s))}" alt="">` : '';
  const thead = `<thead><tr>${cols.map(c => `<th class="${c.cls === 'amt' ? 'amt' : ''}">${c.t}</th>`).join('')}</tr></thead>`;
  const docImg = (x, label) => x.file_id ? `<a href="${esc(x.file_url)}" target="_blank" rel="noopener" data-view="${esc(x.id)}" title="اضغط لعرض ${label}">${isPdf(x) ? '<span class="pdf-tag">PDF</span>' : `<img src="${thumbCache[x.id] || thumbUrl(x, 300)}" data-id="${esc(x.id)}" onerror="thumbFallback(this)" alt="${label} ${esc(x.code)}">`}</a>` : '';
  const atts = x => S.invoices.filter(a => a.attach_to === x.id && a.status === 'مرفق');
  const tr = (x, i) => `<tr class="${x.status === 'مكررة' ? 'dup' : ''}">${cols.map(c => c.k === 'image'
    ? `<td class="img">${x.file_id || atts(x).length ? `<div class="imgs">${docImg(x, 'الفاتورة')}${atts(x).map(a => docImg(a, 'المرفق')).join('')}</div>${atts(x).length ? `<small class="att-n">+ ${atts(x).length} مرفق</small>` : ''}` : '<span class="muted">—</span>'}</td>`
    : `<td class="${c.cls || ''}">${esc(cellText(x, c.k, i, o))}</td>`).join('')}</tr>`;
  const totRow = (s, label) => {
    const ti = cols.findIndex(c => c.k === 'total');
    return `<tr class="tot">${cols.map((c, j) => j === 0 ? `<td>${label}</td>` : c.k === 'total' ? `<td class="amt">${money(s.total)}</td>` : c.k === 'vat' ? `<td class="amt">${money(s.vat)}</td>` : c.k === 'subtotal' ? `<td class="amt">${money(s.sub)}</td>` : j < ti ? '' : '<td></td>').join('')}</tr>`;
  };
  let body = '', i = 0;
  if (o.rGroup) {
    const groups = {};
    rows.forEach(x => (groups[x.category || 'غير مصنّف'] ||= []).push(x));
    const names = Object.keys(groups).sort((a, b) => cats().indexOf(a) - cats().indexOf(b));
    // ملخص التصنيفات
    const all = sums(rows);
    body += `<h3 class="r-sec">ملخص حسب التصنيف</h3><table><thead><tr><th>م</th><th>التصنيف</th><th>عدد الفواتير</th><th class="amt">الضريبة</th><th class="amt">الإجمالي</th><th>النسبة</th></tr></thead><tbody>` +
      names.map((n, k) => { const s = sums(groups[n]); return `<tr><td class="num">${k + 1}</td><td>${esc(n)}</td><td class="num">${s.n}</td><td class="amt">${money(s.vat)}</td><td class="amt">${money(s.total)}</td><td class="num">${all.total ? (s.total / all.total * 100).toFixed(1) + '%' : ''}</td></tr>`; }).join('') +
      `<tr class="tot"><td colspan="2">الإجمالي</td><td>${all.n}</td><td class="amt">${money(all.vat)}</td><td class="amt">${money(all.total)}</td><td>100%</td></tr></tbody></table>`;
    body += `<h3 class="r-sec">التفاصيل</h3><table>${thead}<tbody>` + names.map(n => {
      const g = groups[n];
      return `<tr class="grp"><td colspan="${cols.length}">${esc(n)} — ${sums(g).n} فاتورة</td></tr>` + g.map(x => tr(x, ++i)).join('') + totRow(sums(g), 'مجموع ' + esc(n)).replace('class="tot"', 'class="sub"');
    }).join('') + totRow(all, 'الإجمالي') + '</tbody></table>';
  } else {
    body = `<table>${thead}<tbody>${rows.map(x => tr(x, ++i)).join('') || `<tr><td colspan="${cols.length}">لا توجد فواتير بهذه الشروط</td></tr>`}${rows.length ? totRow(sums(rows), 'الإجمالي') : ''}</tbody></table>`;
  }
  const s = sums(rows);
  let summary = `<span>عدد الفواتير: <b>${s.n}</b></span>`;
  if (o.kind === 'taalum') summary += `<span>الإجمالي بدون الضريبة: <b>${money(s.sub)}</b> ${esc(cur)}</span><span>إجمالي الضريبة: <b>${money(s.vat)}</b> ${esc(cur)}</span><span>الإجمالي شامل الضريبة: <b>${money(s.total)}</b> ${esc(cur)}</span>`;
  else summary += `<span>الإجمالي: <b>${money(s.total)}</b> ${esc(cur)}</span>`;
  if (s.vat && o.kind !== 'taalum') summary += `<span>إجمالي الضريبة: <b>${money(s.vat)}</b></span>`;
  if (!o.kind && !o.rGroup) KINDS.forEach(k => { const ks = sums(rows.filter(x => x.kind === k)); if (ks.n) summary += `<span>${k}: <b>${money(ks.total)}</b> (${ks.n})</span>`; });
  const dupN = rows.filter(x => x.status === 'مكررة').length;
  if (dupN) summary += `<span>مكررة مستبعدة: <b>${dupN}</b></span>`;
  const signs = o.rSign ? String(S.settings.signatures || '').split('\n').map(x => x.trim()).filter(Boolean) : [];
  const html = `<div class="sheet">
    ${`<div class="r-head"><div class="lg r">${L('right')}</div><div class="r-title"><h2>${esc(T.main)}</h2><div>${esc(T.sub)}</div></div><div class="lg l">${L('left')}</div></div>`}
    <div class="r-meta"><span>${esc(period)}</span>${o.rCat ? `<span>التصنيف: ${esc(o.rCat)}</span>` : ''}</div>
    ${body}
    <div class="r-sum">${summary}</div>
    ${S.settings.reportFooter ? `<p style="text-align:center">${esc(S.settings.reportFooter)}</p>` : ''}
    ${signs.length ? `<div class="r-sign">${signs.map(x => `<div>${esc(x)}</div>`).join('')}</div>` : ''}
    <div class="r-foot"><span>أُصدر بواسطة: ${esc(S.user.name)}</span><span>${esc(S.settings.appName || '')}</span><span>تاريخ الإصدار: ${stamp()}</span></div>
  </div>`;
  S.report = { o, rows, cols, T, period, html };
  $('#reportWrap').innerHTML = html;
  $('#btnPrint').disabled = $('#btnExcel').disabled = false;
  renderTaalumNote();
}
$('#btnReport').onclick = buildReport;
// ---------- فحص الفواتير الضريبية السابقة: هل مكتوب عليها جمعية تعلم أو المركز الخيري؟ ----------
let TCHECK = false;
function taalumPending(o) { return S.invoices.filter(x => x.kind === 'ضريبية' && !x.taalum && x.status !== 'مكررة' && x.status !== 'مرفق' && inReportScope(x, o)); }
function renderTaalumNote() {
  const box = $('#reportNote'), o = reportOpts();
  if (o.kind !== 'taalum') { box.innerHTML = ''; return; }
  const pend = taalumPending(o), withImg = pend.filter(x => x.file_id).length, noImg = pend.length - withImg;
  if (!pend.length) { box.innerHTML = ''; return; }
  box.innerHTML = `<div class="note"><span>⚠️ <b>${pend.length}</b> فاتورة ضريبية في هذا النطاق لم تُفحص بعد، ولن تدخل في الكشف حتى تُفحص.${noImg ? ` منها ${noImg} بلا صورة تُحدَّد من نافذة الفاتورة.` : ''}</span>${withImg ? (can('edit') ? `<button class="btn primary sm" id="btnCheckTaalum" ${TCHECK ? 'disabled' : ''}>${TCHECK ? 'جارٍ الفحص…' : `🔍 فحص ${withImg} فاتورة (يستخدم رصيد القراءة الآلية)`}</button>` : '<span class="small">الفحص يحتاج صلاحية التعديل.</span>') : ''}</div>`;
  const btn = $('#btnCheckTaalum');
  if (btn) btn.onclick = () => checkTaalumMany(btn);
}
async function checkTaalumMany(btn) {
  if (TCHECK) return;
  const ids = taalumPending(reportOpts()).filter(x => x.file_id).map(x => x.id);
  TCHECK = true; btn.disabled = true;
  let yes = 0, no = 0, unclear = 0, fail = 0, lastErr = '';
  try {
    for (let i = 0; i < ids.length; i++) {
      const b = $('#btnCheckTaalum'); if (b) b.textContent = `جارٍ الفحص ${i + 1} من ${ids.length}…`;
      try {
        if (i) await sleep(1500);
        let r = await api('checkTaalum', { id: ids[i] });
        if (r.unclear) { // صورة غير واضحة: نكبّرها ونعيد الفحص
          const d = await api('fileData', { id: ids[i] });
          r = await api('checkTaalum', { id: ids[i], extras: zoomCrops(await sourceImage(d.mime, d.data)) });
        }
        upsertLocal(r.inv);
        if (r.unclear) unclear++; else if (r.inv.taalum === 'نعم') yes++; else no++;
      } catch (e) {
        fail++; lastErr = e.message;
        if (/مفتاح|رصيد/.test(e.message)) break; // لا فائدة من المتابعة
      }
    }
  } finally { TCHECK = false; }
  renderList();
  if (S.report && S.report.o.kind === 'taalum') buildReport(); else renderTaalumNote();
  const msg = `باسم الجمعية أو المركز: ${yes} — بغير ذلك: ${no}` + (unclear ? ` — لم تتضح: ${unclear} (حدّدها من نافذة الفاتورة)` : '') + (fail ? ` — تعذّر: ${fail} (${lastErr})` : '');
  toast(msg, !!(fail || unclear));
}
$('#rKind').addEventListener('change', renderTaalumNote);
['#rFrom', '#rTo', '#rCat'].forEach(s => $(s).addEventListener('input', renderTaalumNote));
$('#rQuick').addEventListener('click', () => setTimeout(renderTaalumNote));
// الضغط على صورة الفاتورة في الكشف يعرضها كبيرة
$('#reportWrap').addEventListener('click', e => {
  const a = e.target.closest('[data-view]'); if (!a) return;
  e.preventDefault(); viewInvoiceImage(a.dataset.view);
});
function viewInvoiceImage(id) {
  const x = S.invoices.find(v => v.id === id); if (!x) return;
  const b = modal.open(head(`الفاتورة ${esc(x.code)} — ${esc(x.vendor || '')} — ${money(x.total)}`) + `
    <div class="m-body"><div class="viewer big" id="bigView">${isPdf(x)
      ? `<iframe src="https://drive.google.com/file/d/${esc(x.file_id)}/preview" title="الفاتورة"></iframe>`
      : `<img src="${thumbCache[x.id + '_big'] || thumbUrl(x, 1600)}" data-id="${esc(x.id)}" data-size="big" onerror="thumbFallback(this)" alt="صورة الفاتورة">`}</div></div>
    <div class="m-foot"><a class="btn" href="${esc(x.file_url)}" target="_blank" rel="noopener">فتح الأصل في Drive</a><button class="btn" id="toInv">بيانات الفاتورة</button><span class="spacer"></span><button class="btn" data-close>إغلاق</button></div>`);
  $('#bigView img', b)?.addEventListener('click', ev => ev.currentTarget.parentElement.classList.toggle('zoom'));
  $('#toInv', b).onclick = () => openInvoice(id);
}
$('#btnPrint').onclick = () => {
  if (!S.report) return;
  $('#printArea').innerHTML = $('#reportWrap').innerHTML;
  const imgs = $$('#printArea img');
  Promise.all(imgs.map(i => i.complete ? 1 : new Promise(r => { i.onload = i.onerror = r; setTimeout(r, 4000); }))).then(() => window.print());
};

function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve();
  return new Promise((ok, no) => { const sc = document.createElement('script'); sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js'; sc.onload = ok; sc.onerror = () => no(new Error('تعذّر تحميل مكتبة Excel، تأكد من الاتصال')); document.head.appendChild(sc); });
}
// تصدير Excel بنفس ترتيب الكشف المعتمد، مع الصور والشعارين
async function toB64FromUrl(src) {
  if (src.startsWith('data:')) return { data: src.split(',')[1], ext: /png/.test(src.slice(0, 30)) ? 'png' : 'jpeg' };
  const r = await fetch(src); const b = await r.blob();
  return { data: await fileB64(b), ext: /png/.test(b.type) ? 'png' : 'jpeg' };
}
$('#btnExcel').onclick = async e => {
  if (!S.report) return;
  const btn = e.currentTarget; busy(btn, true, 'جارٍ التجهيز…');
  try {
    await loadExcelJS();
    const { o, rows, cols, T, period } = S.report;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('الكشف', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 5 }], pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    const font = { name: 'Traditional Arabic', size: 14, bold: true };
    const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    const widths = { idx: 6, total: 14, subtotal: 14, vat: 12, category: 22, description: 34, vendor: 30, number: 18, date: 16, image: 22, code: 10, vat_no: 20, kind: 14, who: 24 };
    ws.columns = cols.map(c => ({ key: c.k, width: widths[c.k] || 14 }));
    const n = cols.length;
    ws.getRow(1).height = 70;
    ws.mergeCells(1, 1, 1, n);
    ws.getCell(1, 1).value = '';
    ws.mergeCells(2, 1, 2, n); ws.getCell(2, 1).value = T.main + ' — ' + T.sub; ws.getCell(2, 1).font = { name: 'Traditional Arabic', size: 22, bold: true, color: { argb: 'FF14513F' } };
    ws.mergeCells(3, 1, 3, n); ws.getCell(3, 1).value = period + (o.rCat ? '   |   التصنيف: ' + o.rCat : ''); ws.getCell(3, 1).font = { name: 'Traditional Arabic', size: 14 };
    [2, 3].forEach(r => ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' });
    const hr = ws.getRow(5);
    cols.forEach((c, j) => { const cell = hr.getCell(j + 1); cell.value = c.t; cell.font = font; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAD3' } }; cell.border = border; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
    hr.height = 32;
    // الشعاران
    if (o.logo !== 'none') {
      const add = async (side, col) => { try { const l = await toB64FromUrl(logoSrc(side)); const id = wb.addImage({ base64: l.data, extension: l.ext }); ws.addImage(id, { tl: { col, row: 0.05 }, ext: { width: 120, height: 85 } }); } catch (err) { } };
      if (o.logo === 'both' || o.logo === 'right') await add('right', 0.1);
      if (o.logo === 'both' || o.logo === 'left') await add('left', Math.max(n - 1.6, 1));
    }
    // الصفوف
    let r = 6, i = 0;
    const imgCol = cols.findIndex(c => c.k === 'image');
    const imgs = [];
    for (const x of rows) {
      i++;
      const row = ws.getRow(r);
      cols.forEach((c, j) => {
        const cell = row.getCell(j + 1);
        if (['total', 'subtotal', 'vat'].includes(c.k)) { cell.value = x[c.k] === '' ? null : Number(x[c.k]); cell.numFmt = '#,##0.00'; }
        else if (c.k === 'date') { cell.value = showDate(x.date, o.rDate); }
        else if (c.k === 'image') { cell.value = x.file_url ? { text: isPdf(x) ? 'فتح الفاتورة (PDF)' : 'فتح الفاتورة', hyperlink: x.file_url } : ''; }
        else cell.value = cellText(x, c.k, i, o);
        cell.font = { name: 'Traditional Arabic', size: 14, bold: true, color: x.status === 'مكررة' ? { argb: 'FF9B1C1C' } : undefined, strike: x.status === 'مكررة' && c.k === 'total' };
        cell.border = border; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        if (c.k === 'total') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      });
      if (imgCol >= 0 && x.file_id && !isPdf(x)) { row.height = 95; imgs.push({ x, r }); }
      r++;
    }
    // صف الإجمالي بمعادلة مثل الكشف الأصلي
    const s = sums(rows), tr = ws.getRow(r);
    cols.forEach((c, j) => {
      const cell = tr.getCell(j + 1); cell.border = border; cell.font = { name: 'Traditional Arabic', size: 16, bold: true }; cell.alignment = { horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      if (j === 0) cell.value = 'الإجمالي';
      if (['total', 'subtotal', 'vat'].includes(c.k)) {
        const L = ws.getColumn(j + 1).letter;
        const v = s[c.k === 'subtotal' ? 'sub' : c.k];
        // المكررة تظهر في الكشف لكن لا تُجمع، لذلك يُكتب الإجمالي رقماً إن ظهرت
        cell.value = o.rDups ? v : { formula: `SUM(${L}6:${L}${r - 1})`, result: v };
        cell.numFmt = '#,##0.00';
      }
    });
    // صور الفواتير (تُجلب من الخادم)
    if (imgs.length) {
      let done = 0;
      const work = imgs.slice();
      const worker = async () => {
        while (work.length) {
          const { x, r: rr } = work.shift();
          try {
            if (!thumbCache[x.id]) { const d = await api('fileData', { id: x.id, thumb: true, size: 300 }); thumbCache[x.id] = `data:${d.mime};base64,${d.data}`; }
            const id = wb.addImage({ base64: thumbCache[x.id].split(',')[1], extension: 'jpeg' });
            ws.addImage(id, { tl: { col: imgCol + 0.1, row: rr - 1 + 0.05 }, ext: { width: 140, height: 118 }, hyperlinks: { hyperlink: x.file_url, tooltip: 'اضغط لفتح الفاتورة' } });
          } catch (err) { }
          btn.innerHTML = `الصور ${++done}/${imgs.length}`;
        }
      };
      await Promise.all([worker(), worker(), worker()]);
    }
    const buf = await wb.xlsx.writeBuffer();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = `${T.sub} ${o.rFrom || ''}${o.rTo ? '_' + o.rTo : ''}.xlsx`.replace(/[\\/:*?"<>|]/g, '-');
    document.body.appendChild(a); a.click(); a.remove();
    toast('تم تنزيل ملف Excel');
  } catch (err) { toast('تعذّر التصدير: ' + err.message, true); }
  finally { busy(btn, false); }
};

// =====================================================================
// ٧) السجل
// =====================================================================
let LOG = [];
async function loadLog() {
  $('#logList').innerHTML = '<p class="muted center">جارٍ التحميل…</p>';
  try { LOG = await api('log', { limit: 1500 }); renderLog(); } catch (e) { $('#logList').innerHTML = `<p class="center">${esc(e.message)}</p>`; }
}
function renderLog() {
  const q = $('#logSearch').value.trim().toLowerCase();
  const rows = LOG.filter(x => !q || Object.values(x).join(' ').toLowerCase().includes(q));
  $('#logList').innerHTML = `<table class="plain"><thead><tr><th>الوقت</th><th>المستخدم</th><th>العملية</th><th>الرمز</th><th>التفاصيل</th></tr></thead><tbody>${rows.map(x => `<tr><td style="white-space:nowrap">${esc(x.at)}</td><td>${esc(x.user)}</td><td><b>${esc(x.action)}</b></td><td>${x.id ? `<a href="#" data-inv="${esc(x.id)}">${esc(x.code)}</a>` : esc(x.code)}</td><td class="small">${esc(x.details)}</td></tr>`).join('')}</tbody></table>`;
}
$('#logSearch').addEventListener('input', renderLog);
$('#logReload').onclick = loadLog;
$('#logList').addEventListener('click', e => { const a = e.target.closest('[data-inv]'); if (a) { e.preventDefault(); openInvoice(a.dataset.inv); } });

// =====================================================================
// ٨) الإعدادات والمستخدمون
// =====================================================================
async function refreshPendingBadge() {
  try {
    S.users = await api('users');
    const n = S.users.filter(u => u.status === 'بانتظار الموافقة').length;
    const btn = $('#btnSettings'); $('.dot', btn)?.remove();
    if (n) btn.insertAdjacentHTML('beforeend', `<span class="dot">${n}</span>`);
  } catch (e) { }
}
$('#btnSettings').onclick = () => openSettings();
const SET_TABS = [['general', 'عام'], ['cats', 'التصنيفات'], ['logos', 'الشعارات'], ['users', 'المستخدمون'], ['ai', 'القراءة الآلية'], ['dups', 'الفواتير المكررة'], ['reports', 'الكشوف'], ['links', 'الروابط والخادم']];
async function openSettings(tab = 'general') {
  try { S.settings = await api('getSettings'); } catch (e) { return toast(e.message, true); }
  const b = modal.open(head('⚙️ الإعدادات') + `<div class="set-layout"><nav class="set-nav" id="setNav">${SET_TABS.map(([k, t]) => `<button data-t="${k}">${t}${k === 'users' ? '<span id="uCount"></span>' : ''}</button>`).join('')}</nav><div class="set-pane" id="setPane"></div></div>`);
  $('#setNav', b).onclick = e => { const x = e.target.closest('button'); if (x) showSet(x.dataset.t); };
  showSet(tab);
}
function showSet(t) {
  $$('#setNav button').forEach(x => x.classList.toggle('on', x.dataset.t === t));
  const P = $('#setPane'), st = S.settings;
  const yes = k => st[k] === 'نعم' ? 'checked' : '';
  const saveBtn = '<div class="row gap end"><button class="btn primary" id="setSave">حفظ</button></div>';
  if (t === 'general') {
    P.innerHTML = `<h3>عام</h3><form class="form-grid" id="sf">
      <label class="full">اسم البرنامج<input name="appName" value="${esc(st.appName)}"></label>
      <label class="full">عنوان الكشوف<input name="reportTitle" value="${esc(st.reportTitle)}"></label>
      <label>اسم الجهة (يمين)<input name="orgRight" value="${esc(st.orgRight)}"></label>
      <label>اسم الجهة (يسار)<input name="orgLeft" value="${esc(st.orgLeft)}"></label>
      <label>العملة<input name="currency" value="${esc(st.currency)}"></label>
      <label>نسبة ضريبة القيمة المضافة %<input name="vatRate" inputmode="decimal" value="${esc(st.vatRate)}"></label>
      <label>بادئة الرمز (اختياري)<input name="codePrefix" value="${esc(st.codePrefix)}" placeholder="مثال: 1448-"></label>
      <label>عرض التاريخ<select name="dateMode">${['ميلادي', 'هجري', 'كلاهما'].map(x => `<option ${x === st.dateMode ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label class="check full"><input type="checkbox" name="allowRegister" ${yes('allowRegister')}> السماح بطلب حساب جديد من شاشة الدخول (لا يدخل إلا بعد موافقتك)</label>
    </form>${saveBtn}`;
    bindSave(P);
  }
  if (t === 'cats') {
    let list = cats();
    const draw = () => { $('#catList').innerHTML = list.map((c, i) => `<span class="cat">${esc(c)}<button data-i="${i}" title="حذف">×</button></span>`).join('') || '<span class="muted">لا توجد تصنيفات</span>'; };
    P.innerHTML = `<h3>التصنيفات</h3><p class="muted small">تظهر في القوائم ويختار منها البرنامج عند القراءة الآلية. حذف تصنيف لا يغيّر الفواتير القديمة.</p>
      <div class="cat-list" id="catList"></div>
      <form class="row gap" id="catAdd"><input name="c" placeholder="تصنيف جديد" style="flex:1"><button class="btn">إضافة</button></form><br>${saveBtn}`;
    draw();
    $('#catList').onclick = e => { const i = e.target.dataset.i; if (i != null) { list.splice(Number(i), 1); draw(); } };
    $('#catAdd').onsubmit = e => { e.preventDefault(); const v = e.target.c.value.trim(); if (v && !list.includes(v)) list.push(v); e.target.reset(); draw(); };
    $('#setSave').onclick = () => saveSettings({ categories: list.join('\n') }).then(fillCatSelects);
  }
  if (t === 'logos') {
    P.innerHTML = `<h3>الشعارات</h3><p class="muted small">شعار جمعية تعلّم على يمين الكشف، وشعار مجمع ابن تيمية على يساره. اختيار الإظهار أو الإخفاء يكون عند إصدار الكشف.</p>
      <div class="logo-slots">${['right', 'left'].map(s => `<div class="logo-slot"><b>${s === 'right' ? 'الشعار الأيمن' : 'الشعار الأيسر'}</b><img src="${esc(logoSrc(s))}" id="lg-${s}" alt=""><div class="row gap" style="justify-content:center">
        <label class="btn sm"><input type="file" accept="image/*" hidden data-side="${s}">استبدال</label><button class="btn sm ghost" data-reset="${s}">الافتراضي</button></div></div>`).join('')}</div>`;
    P.onchange = async e => {
      const s = e.target.dataset.side; if (!s) return;
      try { const url = await logoToDataUrl(e.target.files[0]); await saveSettings({ ['logo' + (s === 'right' ? 'Right' : 'Left')]: url }); $('#lg-' + s).src = url; applyBrand(); } catch (err) { toast(err.message, true); }
    };
    P.onclick = async e => { const s = e.target.dataset.reset; if (!s) return; await saveSettings({ ['logo' + (s === 'right' ? 'Right' : 'Left')]: '' }); $('#lg-' + s).src = logoSrc(s); applyBrand(); };
  }
  if (t === 'ai') {
    const keyField = (name, label, ph) => `<label class="full">مفتاح ${label} ${st[name + 'Set'] ? `<small class="muted">(محفوظ: <span dir="ltr">${esc(st[name + 'Hint'])}</span>) — اتركه فارغاً للإبقاء عليه</small>` : '<small style="color:var(--red)">(غير مضبوط)</small>'}<input name="${name}" dir="ltr" autocomplete="off" placeholder="${ph}"></label>`;
    P.innerHTML = `<h3>القراءة الآلية للفواتير</h3>
      <p class="note">القارئ المستخدم الآن: <b>${esc(st.aiActive || '—')}</b></p><br>
      <form class="form-grid" id="sf">
        <label class="check full"><input type="checkbox" name="aiEnabled" ${yes('aiEnabled')}> تفعيل القراءة الآلية</label>
        <div class="field full">من يقرأ الفواتير
          <div class="seg">${[['تلقائي', 'تلقائي (Claude إن وُجد مفتاحه)'], ['Claude', 'Claude'], ['Gemini', 'Gemini']].map(([v, t]) => `<label><input type="radio" name="aiProvider" value="${v}" ${st.aiProvider === v ? 'checked' : ''}><span>${t}</span></label>`).join('')}</div>
        </div>
        <h4 class="full" style="margin:10px 0 0">Claude (من Anthropic)</h4>
        <p class="muted small full" style="margin:0">مدفوع بحسب الاستخدام (يحتاج رصيداً مسبقاً). أنشئ المفتاح من <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noopener">console.anthropic.com</a> — يبدأ بـ sk-ant-</p>
        ${keyField('claudeKey', 'Claude', 'sk-ant-…')}
        <label>نموذج Claude<input name="claudeModel" dir="ltr" value="${esc(st.claudeModel)}"></label><div></div>
        <h4 class="full" style="margin:10px 0 0">Gemini (من جوجل)</h4>
        <p class="muted small full" style="margin:0">مجاني بحدود يومية. المفتاح من <a href="https://aistudio.google.com/apikey" target="_blank" rel="noopener">Google AI Studio</a></p>
        ${keyField('geminiKey', 'Gemini', 'AQ… أو AIza…')}
        <label>نموذج Gemini<input name="aiModel" dir="ltr" value="${esc(st.aiModel)}"></label><div></div>
      </form>
      <p class="muted small">المفاتيح تُحفظ في الخادم فقط ولا تظهر لأي مستخدم.</p>
      <div class="row gap end"><button class="btn" id="aiTest">اختبار القارئ الحالي</button>${st.claudeKeySet ? '<button class="btn danger" data-clear="claudeKey">حذف مفتاح Claude</button>' : ''}${st.geminiKeySet ? '<button class="btn danger" data-clear="geminiKey">حذف مفتاح Gemini</button>' : ''}<button class="btn primary" id="setSave">حفظ</button></div>`;
    $('#setSave', P).onclick = async () => {
      const f = $('#sf', P), o = {};
      $$('input,select,textarea', f).forEach(el => {
        if (!el.name) return;
        if (el.type === 'radio') { if (el.checked) o[el.name] = el.value; return; }
        o[el.name] = el.type === 'checkbox' ? (el.checked ? 'نعم' : 'لا') : el.value;
      });
      ['claudeKey', 'geminiKey'].forEach(k => { if (!o[k]) delete o[k]; });
      try { await saveSettings(o); showSet('ai'); } catch (e) { }
    };
    $('#aiTest').onclick = async e => { busy(e.target, true, 'جارٍ…'); try { const r = await api('testAi'); toast(r.ok ? `القراءة الآلية تعمل ✓ (${r.provider})` : 'رد غير متوقع'); } catch (err) { toast(err.message, true); } finally { busy(e.target, false); } };
    $$('[data-clear]', P).forEach(b => b.onclick = () => saveSettings({ [b.dataset.clear]: '__clear__' }).then(() => showSet('ai')));
  }
  if (t === 'dups') {
    P.innerHTML = `<h3>اكتشاف الفواتير المكررة</h3><p class="muted small">الفاتورة المكررة تُحفظ وعليها علامة «مكررة»، ولا تدخل في الإجماليات ولا الكشوف. يمكن لمن لديه صلاحية التعديل إلغاء العلامة.</p>
      <form class="stack" id="sf">
        <label class="check"><input type="checkbox" name="dupHash" ${yes('dupHash')}> نفس ملف الصورة رُفع مرتين</label>
        <label class="check"><input type="checkbox" name="dupNumber" ${yes('dupNumber')}> نفس رقم الفاتورة مع نفس الرقم الضريبي أو المورد أو المبلغ</label>
        <label class="check"><input type="checkbox" name="dupAmount" ${yes('dupAmount')}> نفس المورد ونفس التاريخ ونفس المبلغ (للأوراق المكتوبة بلا رقم)</label>
      </form><br>${saveBtn}`;
    bindSave(P);
  }
  if (t === 'reports') {
    P.innerHTML = `<h3>الكشوف</h3><form class="form-grid" id="sf">
      <label class="full">خانات التوقيع (كل سطر خانة)<textarea name="signatures" rows="4">${esc(st.signatures)}</textarea></label>
      <label class="full">نص أسفل الكشف (اختياري)<input name="reportFooter" value="${esc(st.reportFooter)}"></label></form>${saveBtn}`;
    bindSave(P);
  }
  if (t === 'links') {
    P.innerHTML = `<h3>الروابط والخادم</h3><div class="stack">
      <div id="lnk" class="row gap"><span class="muted">جارٍ جلب الروابط…</span></div>
      <label>رابط الخادم على هذا الجهاز<input id="apiIn" dir="ltr" value="${esc(apiUrl())}"></label>
      <div class="row gap"><button class="btn" id="apiSave">حفظ الرابط</button></div>
      <p class="muted small">التحديثات تُرفع إلى GitHub وتصل تلقائياً للخادم خلال 10 دقائق، وللواجهة عند إعادة فتح الصفحة.</p></div>`;
    api('links').then(l => $('#lnk').innerHTML = `<a class="btn" href="${esc(l.sheet)}" target="_blank" rel="noopener">📊 فتح جدول الكشف (Google Sheets)</a><a class="btn" href="${esc(l.folder)}" target="_blank" rel="noopener">📁 مجلد صور الفواتير</a>`).catch(e => $('#lnk').textContent = e.message);
    $('#apiSave').onclick = () => { const u = $('#apiIn').value.trim(); if (!/^https:\/\/script\.google\.com\//.test(u)) return toast('رابط غير صحيح', true); localStorage.setItem(LS.api, u); toast('حُفظ الرابط'); };
  }
  if (t === 'users') drawUsers(P);
}
function bindSave(P) {
  $('#setSave', P).onclick = () => {
    const f = $('#sf', P), o = {};
    $$('input,select,textarea', f).forEach(el => { if (!el.name) return; o[el.name] = el.type === 'checkbox' ? (el.checked ? 'نعم' : 'لا') : el.value; });
    if ('geminiKey' in o && !o.geminiKey) delete o.geminiKey;
    saveSettings(o);
  };
}
async function saveSettings(o) {
  try { S.settings = await api('saveSettings', { settings: o }); applyBrand(); toast('حُفظت الإعدادات'); if ('geminiKey' in o && $('#setNav .on')?.dataset.t === 'ai') showSet('ai'); }
  catch (e) { toast(e.message, true); throw e; }
}
function logoToDataUrl(file) {
  return new Promise((ok, no) => {
    const i = new Image();
    i.onload = () => {
      for (const max of [360, 280, 200, 150]) {
        const k = Math.min(1, max / Math.max(i.naturalWidth, i.naturalHeight)), c = document.createElement('canvas');
        c.width = Math.round(i.naturalWidth * k); c.height = Math.round(i.naturalHeight * k);
        c.getContext('2d').drawImage(i, 0, 0, c.width, c.height);
        let u = c.toDataURL('image/png'); if (u.length > 45000) { const x = c.getContext('2d'); x.globalCompositeOperation = 'destination-over'; x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); u = c.toDataURL('image/jpeg', 0.88); }
        if (u.length <= 45000) return ok(u);
      }
      no(new Error('الصورة كبيرة جداً'));
    };
    i.onerror = () => no(new Error('صورة غير صالحة'));
    i.src = URL.createObjectURL(file);
  });
}
const PERM_NAMES = { add: 'إضافة ورفع الفواتير', edit: 'تعديل الفواتير', delete: 'حذف الفواتير', reports: 'إصدار الكشوف', admin: 'مدير (كل الصلاحيات والإعدادات)' };
async function drawUsers(P) {
  P.innerHTML = '<h3>المستخدمون</h3><p class="muted">جارٍ التحميل…</p>';
  try { S.users = await api('users'); } catch (e) { P.innerHTML = esc(e.message); return; }
  const pend = S.users.filter(u => u.status === 'بانتظار الموافقة');
  const card = u => `<div class="u-card ${u.status === 'بانتظار الموافقة' ? 'pending' : ''}" data-u="${esc(u.id)}">
    <div class="u-head"><div><b>${esc(u.name)}</b> <span class="muted" dir="ltr">@${esc(u.username)}</span> <span class="badge ${u.status === 'مفعّل' ? 'b-ok' : u.status === 'موقوف' ? 'b-dup' : 'b-review'}">${esc(u.status)}</span></div>
      <div class="small muted">${u.status === 'بانتظار الموافقة' ? 'طلب في ' + esc(u.requestedAt) : (u.approvedBy ? 'وافق عليه ' + esc(u.approvedBy) + ' — ' + esc(u.approvedAt) : '')}${u.lastLogin ? ' · آخر دخول ' + esc(u.lastLogin) : ''}</div></div>
    <div class="u-perms">${Object.entries(PERM_NAMES).map(([p, t]) => `<label class="check"><input type="checkbox" data-p="${p}" ${u.perms[p] ? 'checked' : ''}> ${t}</label>`).join('')}</div>
    <div class="row gap">${u.status === 'بانتظار الموافقة' ? '<button class="btn primary sm" data-a="approve">✓ موافقة بالصلاحيات المحددة</button><button class="btn danger sm" data-a="del">رفض الطلب</button>'
      : `<button class="btn sm" data-a="save">حفظ الصلاحيات</button><button class="btn sm ghost" data-a="pw">كلمة مرور جديدة</button>${u.id === S.user.id ? '' : (u.status === 'موقوف' ? '<button class="btn sm ghost" data-a="on">تفعيل</button>' : '<button class="btn sm ghost" data-a="off">إيقاف</button>') + '<button class="btn sm danger" data-a="del">حذف</button>'}`}</div></div>`;
  P.innerHTML = `<h3>المستخدمون</h3>
    ${pend.length ? `<h4>طلبات بانتظار موافقتك (${pend.length})</h4>${pend.map(card).join('')}<h4>المستخدمون الحاليون</h4>` : ''}
    ${S.users.filter(u => u.status !== 'بانتظار الموافقة').map(card).join('')}
    <details class="card" style="margin-top:12px"><summary><b>➕ إضافة مستخدم مباشرة</b></summary><form id="uAdd" class="form-grid" style="margin-top:10px">
      <label>الاسم<input name="name" required></label><label>اسم الدخول<input name="username" dir="ltr" required></label>
      <label>كلمة المرور<input name="password" minlength="6" required></label><div></div>
      <div class="u-perms full">${Object.entries(PERM_NAMES).map(([p, t]) => `<label class="check"><input type="checkbox" name="${p}" ${p === 'add' ? 'checked' : ''}> ${t}</label>`).join('')}</div>
      <div class="full"><button class="btn primary">إضافة</button></div></form></details>`;
  refreshPendingBadge();
  P.onclick = async e => {
    const a = e.target.dataset.a; if (!a) return;
    const c = e.target.closest('[data-u]'), id = c.dataset.u, u = S.users.find(x => x.id === id);
    const perms = {}; $$('[data-p]', c).forEach(x => perms[x.dataset.p] = x.checked);
    try {
      if (a === 'approve') S.users = await api('saveUser', { user: { id, perms, status: 'مفعّل' } });
      if (a === 'save') S.users = await api('saveUser', { user: { id, perms } });
      if (a === 'off' || a === 'on') S.users = await api('saveUser', { user: { id, status: a === 'off' ? 'موقوف' : 'مفعّل' } });
      if (a === 'pw') { const pw = prompt('كلمة المرور الجديدة لـ ' + u.name + ' (6 أحرف على الأقل):'); if (!pw) return; S.users = await api('saveUser', { user: { id, password: pw } }); }
      if (a === 'del') { if (!confirm('حذف ' + u.name + '؟')) return; S.users = await api('deleteUser', { id }); }
      toast('تم'); drawUsers(P);
    } catch (err) { toast(err.message, true); }
  };
  $('#uAdd', P).onsubmit = async e => {
    e.preventDefault(); const f = e.target, perms = {};
    Object.keys(PERM_NAMES).forEach(p => perms[p] = f[p].checked);
    try { await api('saveUser', { user: { name: f.name.value, username: f.username.value, password: f.password.value, perms } }); toast('أُضيف المستخدم'); drawUsers(P); } catch (err) { toast(err.message, true); }
  };
}

start();
