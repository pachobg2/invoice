'use strict';
const KEY = 'faktura.v1';
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const num = v => { const n = parseFloat(String(v).replace(',', '.')); return isFinite(n) ? n : 0; };
const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const money = n => (isFinite(n) ? n : 0).toLocaleString('bg-BG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const eur = n => money(n) + '\u00a0€';
const padNo = n => /^\d+$/.test(String(n).trim()) ? String(n).trim().padStart(10, '0') : String(n);
const qtyFmt = v => v === '' || v == null ? '' : String(num(v));
const today = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
const dmy = iso => /^\d{4}-\d\d-\d\d$/.test(iso || '') ? iso.slice(8) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4) : (iso || '');
const todayDmy = () => dmy(today());

/* ---------------- number to Bulgarian words (replaces VBA Slov) ---------------- */
const U = ['', 'едно', 'две', 'три', 'четири', 'пет', 'шест', 'седем', 'осем', 'девет', 'десет', 'единадесет', 'дванадесет', 'тринадесет', 'четиринадесет', 'петнадесет', 'шестнадесет', 'седемнадесет', 'осемнадесет', 'деветнадесет'];
const T = ['', '', 'двадесет', 'тридесет', 'четиридесет', 'петдесет', 'шестдесет', 'седемдесет', 'осемдесет', 'деветдесет'];
const H = ['', 'сто', 'двеста', 'триста', 'четиристотин', 'петстотин', 'шестстотин', 'седемстотин', 'осемстотин', 'деветстотин'];
function groupTokens(n, gender) { // n 1..999 -> number words
  const t = [], h = Math.floor(n / 100), rest = n % 100;
  if (h) t.push(H[h]);
  if (rest) {
    if (rest < 20) t.push(fix(U[rest], rest, gender));
    else { t.push(T[Math.floor(rest / 10)]); if (rest % 10) t.push(fix(U[rest % 10], rest % 10, gender)); }
  }
  return t;
}
function fix(w, n, gender) { // gender: 'f' (хиляди), 'm' (милиона)
  if (gender === 'f' && n % 10 === 1 && n !== 11) return 'една';
  if (gender === 'm' && n % 10 === 1 && n !== 11) return 'един';
  if ((gender === 'f' || gender === 'm') && n % 10 === 2 && n !== 12) return 'две'.replace('две', gender === 'm' ? 'два' : 'две');
  return w;
}
function intWords(n) {
  if (n === 0) return 'нула';
  const g = []; // [{tokens, scale}]
  const bil = Math.floor(n / 1e9), mil = Math.floor(n / 1e6) % 1000, th = Math.floor(n / 1e3) % 1000, un = n % 1000;
  if (bil) g.push(bil === 1 ? { t: ['един'], s: 'милиард' } : { t: groupTokens(bil, 'm'), s: 'милиарда' });
  if (mil) g.push(mil === 1 ? { t: ['един'], s: 'милион' } : { t: groupTokens(mil, 'm'), s: 'милиона' });
  if (th) g.push(th === 1 ? { t: [], s: 'хиляда' } : { t: groupTokens(th, 'f'), s: 'хиляди' });
  if (un) g.push({ t: groupTokens(un, 'n'), s: '' });
  const last = g[g.length - 1];
  const out = [];
  g.forEach(x => {
    const t = x.t.slice();
    if (x === last) { // "и" before the last numeric word of the final group
      if (t.length > 1) t.splice(t.length - 1, 0, 'и');
      else if (t.length === 1 && g.length > 1) t.unshift('и');
      else if (t.length === 0 && g.length > 1) { /* "хиляда" after millions */ out.push('и'); }
    }
    out.push(...t); if (x.s) out.push(x.s);
  });
  return out.join(' ');
}
function slov(v) {
  v = r2(v);
  const neg = v < 0; v = Math.abs(v);
  const whole = Math.floor(v + 1e-9), cents = Math.round((v - whole) * 100);
  let s = (neg ? 'минус ' : '') + intWords(whole) + ' евро';
  if (cents) s += ' и ' + String(cents).padStart(2, '0') + ' евро цента';
  return s;
}

/* ---------------- state ---------------- */
let S;
const blankLines = n => Array.from({ length: n }, () => ({ name: '', code: '', qty: '', price: '' }));
function freshState() {
  // no company or customer data ships with the program: an empty start (fill in Настройки or import the Excel workbook)
  const sd = { units: ['БР.', 'КГ.', 'К-Т', 'М.', 'МЕС.', 'ОП.', 'ЧАСА', 'ДНИ', 'ЛИТРА'], payments: ['В БРОЙ', 'ПО СМЕТКА'], compilers: [] };
  return {
    clients: [], suppliers: [], articles: [], invoices: [], units: sd.units, payments: sd.payments, compilers: sd.compilers,
    company: { name: '', addr: '', vat: '', eik: '', mol: '', bank: '', bic: '', iban: '' },
    logo: true, logoImage: '', nextNo: 1, rate: 20,
    inv: { client: '', date: today(), tax: today(), taxManual: false, pay: sd.payments[0] || '', comp: sd.compilers[0] || '', rate: 20, orig: 'ОРИГИНАЛ', lines: blankLines(HALF) },
    stock: { sup: '', date: today(), comp: sd.compilers[0] || '', orig: 'ОРИГИНАЛ', lines: Array.from({ length: 8 }, () => ({ name: '', unit: 'БР.', qty: '', price: '' })) },
    receipts: [], nextRcpt: 1
  };
}
const HOST = window.chrome && window.chrome.webview ? window.chrome.webview : null;
let hostSeq = 0; const hostPending = {};
if (HOST) HOST.addEventListener('message', e => { const m = e.data || {}; if (m.ev === 'upd') { updProgress(m.pct); return; } if (m.ev === 'drive') { driveCheck('arrive'); return; } const p = hostPending[m.id]; if (!p) return; delete hostPending[m.id]; m.error ? p.rej(new Error(m.error)) : p.res(m.result); });
const hostCall = (cmd, arg) => new Promise((res, rej) => {
  if (!HOST) return rej(new Error('no host'));
  const id = ++hostSeq; hostPending[id] = { res, rej }; HOST.postMessage({ id, cmd, arg });
  setTimeout(() => { if (hostPending[id]) { delete hostPending[id]; rej(new Error('timeout')); } }, cmd === 'pdf' || cmd === 'setBackupDir' || cmd === 'updateInstall' ? 600000 : 20000);
});
async function load() {
  if (HOST) { try { const j = await hostCall('load'); if (j) { S = JSON.parse(j); return; } } catch (e) { } }
  try { const j = localStorage.getItem(KEY); if (j) { S = JSON.parse(j); return; } } catch (e) { }
  S = freshState();
}
function persist() {
  const j = JSON.stringify(S);
  if (HOST) HOST.postMessage({ cmd: 'save', arg: j });
  else try { localStorage.setItem(KEY, j); } catch (e) { alert('Не може да се запише: ' + e.message); }
}
let saveT;
function save() { clearTimeout(saveT); saveT = setTimeout(persist, 250); }
const saveNow = () => { clearTimeout(saveT); persist(); };
window.__state = () => JSON.stringify(S);
window.addEventListener('beforeunload', saveNow);
const backupNow = async () => { if (!HOST) return null; try { return await hostCall('backup', JSON.stringify(S)); } catch (e) { return null; } };

/* ---------------- dialog helper ---------------- */
function ask(title, fields, vals, validate) {
  return new Promise(res => {
    const dlg = $('#dlg'); $('#dlg_t').textContent = title;
    $('#dlg_b').innerHTML = fields.map(f => {
      const v = vals[f.k] ?? '';
      let inp;
      if (f.type === 'select') inp = `<select data-k="${f.k}">${f.opts.map(o => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
      else if (f.type === 'check') inp = `<input type="checkbox" data-k="${f.k}" ${v ? 'checked' : ''}>`;
      else inp = `<input data-k="${f.k}" type="${f.type || 'text'}" value="${esc(v)}" ${f.list ? `list="${f.list}"` : ''} ${f.max ? `maxlength="${f.max}"` : ''}>`;
      return `<div class="f"><label>${esc(f.label)}</label>${inp}</div>`;
    }).join('') + '<div id="dlg_e" class="err hide"></div>';
    const get = () => { const o = {}; $$('#dlg_b [data-k]').forEach(e => o[e.dataset.k] = e.type === 'checkbox' ? e.checked : e.value.trim()); return o; };
    const ok = $('#dlg_ok');
    ok.onclick = ev => { const o = get(); const m = validate && validate(o); if (m) { ev.preventDefault(); const e = $('#dlg_e'); e.textContent = m; e.classList.remove('hide'); } else { dlg.close(); res(o); } };
    $('#dlg_x').onclick = ev => { ev.preventDefault(); dlg.close(); res(null); };
    dlg.oncancel = () => res(null);
    dlg.showModal();
    const first = $('#dlg_b input,#dlg_b select'); if (first) first.focus();
  });
}

const grow = el => { if (el.offsetParent === null) return; el.style.height = 'auto'; el.style.height = el.scrollHeight + 'px'; };
const growAll = root => $$('textarea[data-f=name]', root).forEach(grow);
addEventListener('resize', () => growAll(document));

/* ---------------- search dropdown (клиенти, доставчици, артикули) ---------------- */
const SEARCH = {
  clients: { items: () => S.clients, hay: c => [c.name, c.eik, c.vat, c.addr, c.mol].join(' '), sub: c => [c.eik, c.addr].filter(Boolean).join(' · ') },
  sups: { items: () => S.suppliers, hay: c => [c.name, c.eik, c.vat, c.addr, c.mol].join(' '), sub: c => [c.eik, c.addr].filter(Boolean).join(' · ') },
  arts: { items: () => S.articles, hay: a => a.name, sub: a => [a.unit, a.price !== '' && a.price != null ? eur(num(a.price)) : ''].filter(Boolean).join(' · ') }
};
(function () {
  const dd = document.createElement('div'); dd.id = 'sdd'; dd.className = 'hide'; document.body.appendChild(dd);
  let cur = null, list = [], sel = 0, kind;
  const hide = () => { dd.classList.add('hide'); cur = null; };
  function show(input) {
    cur = input; kind = SEARCH[input.dataset.search];
    const toks = input.value.toLowerCase().split(/\s+/).filter(Boolean);
    list = kind.items().filter(x => { const h = kind.hay(x).toLowerCase(); return toks.every(t => h.includes(t)); });
    if (sel >= list.length) sel = 0;
    dd.innerHTML = list.length ? `<div class="sdh">${list.length} ${list.length === 1 ? 'съвпадение' : 'съвпадения'}</div>` + list.map((x, i) => `<div data-i="${i}" class="${i === sel ? 'on' : ''}">${esc(x.name)}<small>${esc(kind.sub(x))}</small></div>`).join('')
      : '<div class="none">Няма намерени. Може да въведете свободен текст.</div>';
    // every match is listed, with its full name; the box is as wide as needed and uses the free space below (or above) the field
    const r = input.getBoundingClientRect(), w = Math.min(Math.max(r.width, 460), innerWidth - 16);
    dd.style.width = w + 'px'; dd.style.left = Math.max(8, Math.min(r.left, innerWidth - w - 8)) + 'px';
    dd.style.maxHeight = 'none'; dd.classList.remove('hide');
    const want = dd.scrollHeight, below = innerHeight - r.bottom - 10, above = r.top - 10;
    if (below >= Math.min(want, 300) || below >= above) { dd.style.maxHeight = Math.max(120, below) + 'px'; dd.style.top = (r.bottom + 2) + 'px'; }
    else { dd.style.maxHeight = Math.max(120, above) + 'px'; dd.style.top = Math.max(4, r.top - 2 - dd.offsetHeight) + 'px'; }
  }
  function pick(i) {
    const x = list[i]; if (!x || !cur) return;
    const inp = cur; inp.value = x.name; hide(); inp.dispatchEvent(new Event('input', { bubbles: true })); inp.dispatchEvent(new Event('change', { bubbles: true }));
  }
  document.addEventListener('focusin', e => { if (e.target.dataset && e.target.dataset.search) { sel = 0; show(e.target); } else if (!dd.contains(e.target)) hide(); });
  document.addEventListener('input', e => { if (e.isTrusted && e.target.dataset && e.target.dataset.search) { sel = 0; show(e.target); } }, true);
  document.addEventListener('keydown', e => {
    if (!cur || dd.classList.contains('hide') || e.target !== cur) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); sel = (sel + (e.key === 'ArrowDown' ? 1 : -1) + list.length) % Math.max(list.length, 1); show(cur); const on = dd.querySelector('.on'); if (on) on.scrollIntoView({ block: 'nearest' }); }
    else if (e.key === 'Enter' && list.length) { e.preventDefault(); e.stopPropagation(); pick(sel); }
    else if (e.key === 'Escape' || e.key === 'Tab') hide();
  }, true);
  dd.addEventListener('mousedown', e => { e.preventDefault(); const r = e.target.closest('[data-i]'); if (r) pick(+r.dataset.i); });
  document.addEventListener('mousedown', e => { if (cur && e.target !== cur && !dd.contains(e.target)) hide(); });
  addEventListener('scroll', e => { if (cur && !dd.contains(e.target)) show(cur); }, true); addEventListener('resize', hide);
})();

/* ---------------- tabs ---------------- */
const GROUP = { inv: 'inv', prev: 'inv', stock: 'stock', sprev: 'stock', reg: 'issued', sreg: 'issued' };
const SEGS = { inv: [['inv', 'Данни'], ['prev', 'Преглед и печат']], stock: [['stock', 'Данни'], ['sprev', 'Преглед и печат']], issued: [['reg', 'Издадени фактури'], ['sreg', 'Издадени разписки']] };
let lastIssued = 'reg';
$$('.seg').forEach(el => { el.innerHTML = SEGS[el.dataset.seg].map(([t, l]) => `<button data-t="${t}">${l}</button>`).join(''); });
function showTab(t) {
  const g = GROUP[t] || t; if (g === 'issued') lastIssued = t;
  $$('#side button[data-g]').forEach(b => b.classList.toggle('on', b.dataset.g === g));
  $$('.seg button').forEach(b => b.classList.toggle('on', b.dataset.t === t));
  $$('.tab').forEach(s => s.classList.toggle('hide', s.id !== 'tab-' + t));
  if (t === 'reg') renderReg();
  if (t === 'vat') { renderVat(); renderJournal(); }
  if (t === 'clients') lists.clients.render();
  if (t === 'sups') lists.sups.render();
  if (t === 'arts') lists.arts.render();
  if (t === 'set') renderSettings();
  if (t === 'inv' || t === 'prev') renderInvoice();
  if (t === 'stock' || t === 'sprev') renderStock();
  if (t === 'sreg') renderSReg();
  try { sessionStorage.setItem('tab', t); } catch (e) { }
  $('#scroller').scrollTop = 0;
}
$$('#side button[data-t]').forEach(b => b.onclick = () => showTab(b.dataset.g === 'issued' ? lastIssued : b.dataset.t));
$$('.seg button').forEach(b => b.onclick = () => showTab(b.dataset.t));

/* ---------------- light / dark theme ---------------- */
const ICON_SUN = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';
const ICON_MOON = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>';
function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  $('#themeBtn .ic').innerHTML = t === 'dark' ? ICON_SUN : ICON_MOON;
  $('#themeBtn').title = $('#themeLbl').textContent = t === 'dark' ? 'Светъл режим' : 'Тъмен режим';
}
function setTheme(t) { applyTheme(t); try { localStorage.setItem('fk.theme', t); } catch (e) { } if (S) { S.theme = t; save(); } }
$('#themeBtn').onclick = () => setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');
applyTheme(document.documentElement.dataset.theme || 'light');

/* ---------------- clear (✕) and edit links next to Получател / Доставчик ---------------- */
async function editPartyDlg(list, name, title, done) {
  const x = list.find(c => c.name === name); if (!x) return;
  const wo = warnOnce();
  const o = await ask(title, [{ k: 'name', label: 'Име', max: 51 }, { k: 'addr', label: 'Адрес' }, { k: 'eik', label: 'ЕИК / БУЛСТАТ' }, { k: 'vat', label: 'Номер по ЗДДС' }, { k: 'mol', label: 'МОЛ' }], x,
    v => !v.name ? 'Въведете име.' : list.some(c => c !== x && c.name === v.name) ? 'Вече съществува такова име.' : wo(v.eik, v.vat));
  if (!o) return;
  Object.assign(x, { name: o.name, addr: o.addr, eik: o.eik, vat: o.vat, mol: o.mol });
  saveNow(); fillDatalists(); done(o.name);
}
function partyTools() {
  $('#i_clientX').classList.toggle('hide', !S.inv.client); $('#i_editClient').classList.toggle('hide', !findClient(S.inv.client));
  $('#s_supX').classList.toggle('hide', !S.stock.sup); $('#s_editSup').classList.toggle('hide', !findSup(S.stock.sup));
}
/* ---------------- ЕИК / ЕГН / ЗДДС number checks (a warning, the user can still save) ---------------- */
function egnValid(s) {
  if (!/^\d{10}$/.test(s)) return false;
  const w = [2, 4, 8, 5, 10, 9, 7, 3, 6]; let sum = 0; for (let i = 0; i < 9; i++) sum += +s[i] * w[i];
  let r = sum % 11; if (r === 10) r = 0; return r === +s[9];
}
function eikValid(s) {
  s = String(s || '').trim(); if (!/^\d+$/.test(s)) return false;
  const d = s.split('').map(Number);
  const nine = () => { let sum = 0; for (let i = 0; i < 8; i++) sum += d[i] * (i + 1); let r = sum % 11; if (r === 10) { sum = 0; for (let i = 0; i < 8; i++) sum += d[i] * (i + 3); r = sum % 11; if (r === 10) r = 0; } return r === d[8]; };
  if (s.length === 9) return nine();
  if (s.length === 13) { if (!nine()) return false; let r = (d[8] * 2 + d[9] * 7 + d[10] * 3 + d[11] * 5) % 11; if (r === 10) { r = (d[8] * 4 + d[9] * 9 + d[10] * 5 + d[11] * 7) % 11; if (r === 10) r = 0; } return r === d[12]; }
  if (s.length === 10) return egnValid(s);
  return false;
}
function partyWarn(eik, vat) {
  const w = [], e = String(eik || '').trim(), v = String(vat || '').trim().toUpperCase();
  if (e) {
    if (!/^\d+$/.test(e)) w.push('ЕИК/БУЛСТАТ трябва да съдържа само цифри');
    else if (![9, 10, 13].includes(e.length)) w.push(`ЕИК/ЕГН трябва да е 9, 10 или 13 цифри (въведени са ${e.length})`);
    else if (!eikValid(e)) w.push('контролната цифра на ЕИК/ЕГН не съвпада — вероятна грешка при въвеждане');
  }
  if (v.startsWith('BG')) {
    const g = v.slice(2);
    if (!/^\d{9,10}$/.test(g)) w.push('номерът по ЗДДС трябва да е BG и 9 или 10 цифри');
    else if (g.length === 9 && !eikValid(g)) w.push('контролната цифра на номера по ЗДДС не съвпада');
    else if (g.length === 9 && /^\d{9,13}$/.test(e) && e.slice(0, 9) !== g) w.push('номерът по ЗДДС не съвпада с ЕИК');
  }
  return w;
}
// validate callback for the dialogs: the first attempt shows the warning, pressing „Запази“ again accepts
function warnOnce() { let seen = ''; return (eik, vat) => { const w = partyWarn(eik, vat); if (!w.length) return ''; const k = w.join('|'); if (seen === k) return ''; seen = k; return 'Проверка: ' + w.join('; ') + '. Натиснете „Запази“ отново, ако данните са верни.'; }; }

/* ---------------- datalists ---------------- */
function fillDatalists() {
  $('#dl_clients').innerHTML = S.clients.map(c => `<option value="${esc(c.name)}">`).join('');
  $('#dl_sups').innerHTML = S.suppliers.map(c => `<option value="${esc(c.name)}">`).join('');
  $('#dl_arts').innerHTML = S.articles.map(a => `<option value="${esc(a.name)}">`).join('');
  $('#dl_units').innerHTML = S.units.map(u => `<option value="${esc(u)}">`).join('');
}
const unitOpts = cur => ['', ...S.units, ...(cur && !S.units.includes(cur) ? [cur] : [])].map(u => `<option value="${esc(u)}" ${u === (cur || '') ? 'selected' : ''}>${esc(u)}</option>`).join('');
const unitSel = (f, cur) => `<select data-f="${f}">${unitOpts(cur)}</select>`;
const setUnit = (el, v) => { el.innerHTML = unitOpts(v); el.value = v || ''; };
const findClient = n => S.clients.find(c => c.name === n);
const findSup = n => S.suppliers.find(c => c.name === n);
const findArt = n => S.articles.find(a => a.name === n);
const fillSelect = (el, arr, val) => { el.innerHTML = arr.map(o => `<option ${o === val ? 'selected' : ''}>${esc(o)}</option>`).join(''); if (val && !arr.includes(val)) el.insertAdjacentHTML('afterbegin', `<option selected>${esc(val)}</option>`); };

/* ---------------- INVOICE ---------------- */
function calcInv(d = S.inv) {
  const lines = d.lines.map(l => ({ ...l, val: r2(num(l.qty) * num(l.price)), used: (l.name || '').trim() !== '' || num(l.qty) * num(l.price) !== 0 }));
  const base = r2(lines.reduce((s, l) => s + l.val, 0));
  const vat = r2(base * num(d.rate) / 100);
  return { lines, items: lines.filter(l => l.used), base, vat, total: r2(base + vat) };
}
function buildInvoiceForm() {
  const L = S.inv.lines;
  $('#i_items tbody').innerHTML = L.map((l, i) => `<tr data-i="${i}">
   <td>${i + 1}</td>
   <td><textarea rows="1" data-f="name" data-search="arts" autocomplete="off">${esc(l.name)}</textarea></td>
   <td>${unitSel('code', l.code)}</td>
   <td><input data-f="qty" inputmode="decimal" style="text-align:right" value="${esc(l.qty)}"></td>
   <td><input data-f="price" inputmode="decimal" style="text-align:right" value="${esc(l.price)}"></td>
   <td class="n v"></td>
   <td><button class="xbtn" data-del="${i}" title="Премахни реда">✕</button></td></tr>`).join('');
  $('#i_rowInfo').textContent = L.length + ' реда';
  growAll($('#i_items'));
}
function addInvRow(focus) {
  S.inv.lines.push({ name: '', code: '', qty: '', price: '' }); save(); buildInvoiceForm(); updateInvoice();
  if (focus) { const r = $$('#i_items tbody tr').pop(); $('[data-f=name]', r).focus(); }
}
function syncInvForm() {
  const d = S.inv;
  if (d.taxManual === undefined) d.taxManual = !!d.tax && d.tax !== d.date;
  if (!d.taxManual && d.date) d.tax = d.date;
  $('#i_client').value = d.client; $('#i_no').value = padNo(S.nextNo); $('#i_date').value = d.date; $('#i_tax').value = d.tax;
  $('#i_taxReset').classList.toggle('hide', !d.taxManual);
  fillSelect($('#i_pay'), S.payments, d.pay); fillSelect($('#i_comp'), S.compilers, d.comp);
  if (!d.edit) d.rate = S.rate; // VAT rate comes from Настройки (a reopened invoice keeps its own)
  $('#i_orig').value = d.orig;
  if (!d.lines.length) d.lines = blankLines(1);
  const eb = $('#i_edit'); eb.classList.toggle('hide', !d.edit);
  if (d.edit) $('#i_editTxt').textContent = `Редакция на записана фактура № ${d.edit.no}. „Печат и запис“ и „Само запис“ обновяват записа в регистъра — не се създава нов номер.`;
  buildInvoiceForm();
}
function clientInfoHtml(c) {
  if (!c) return '';
  return `${esc(c.addr)}${c.vat ? ' · ЗДДС: ' + esc(c.vat) : ''}${c.eik ? ' · ЕИК: ' + esc(c.eik) : ''}${c.mol ? ' · МОЛ: ' + esc(c.mol) : ''}`;
}
function renderInvoice() {
  fillDatalists(); syncInvForm(); updateInvoice();
}
const HALF = 7, RHALF = 8, MID = 34, LAST = 24; // rows: invoice, two copies per sheet (7) / warehouse receipt, one copy on half a sheet (8) / full page without totals / full page with totals
function copyHtml(d, c, calc, orig, items, o) { // o: {start, pad, last, pg}
  const co = S.company, sup = esc(co.name), lg = logoTag();
  let rows = items.map((l, i) => `<tr><td class="c">${o.start + i + 1}</td><td>${esc(l.name)}</td><td class="c">${esc(l.code)}</td><td class="c">${qtyFmt(l.qty)}</td><td class="n">${eur(num(l.price))}</td><td class="n">${eur(l.val)}</td></tr>`).join('');
  for (let i = items.length; i < o.pad; i++) rows += '<tr><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>';
  const party = (lbl, nm, p) => `<div class="pbox"><div class="lbl">${lbl}</div><div class="nm">${nm}</div><div class="kv"><span>Адрес:</span><span>${esc(p?.addr)}</span><span>№ по ЗДДС:</span><span>${esc(p?.vat)}</span><span>ЕИК/ЕГН:</span><span>${esc(p?.eik)}</span><span>МОЛ:</span><span>${esc(p?.mol)}</span></div></div>`;
  const tail = o.last ? `<div class="iv-sum"><div class="iv-words">ДДС словом: <b>${esc(slov(calc.vat))}</b><br>Сума за плащане словом: <b>${esc(slov(calc.total))}</b></div>
   <table class="iv-tot"><tr><td>дан. основа:</td><td>${eur(calc.base)}</td></tr><tr><td>ДДС ставка:</td><td>${num(d.rate)}%</td></tr><tr><td>ДДС:</td><td>${eur(calc.vat)}</td></tr><tr class="g"><td>сума за плащане:</td><td>${eur(calc.total)}</td></tr></table></div>
  <div class="iv-foot"><div class="bank"><span>плащане:</span> <b>${esc(d.pay)}</b><br><span>банка:</span> ${esc(co.bank)}<br><span>BIC код:</span> ${esc(co.bic)}<br><span>IBAN:</span> ${esc(co.iban)}</div>
   <div class="sg">подпис на получателя<b>&nbsp;</b></div>
   <div class="sg">съставител<b>${esc(d.comp)}</b></div></div>`
    : `<div class="cont">продължава на следващата страница…</div>`;
  return `<div class="copy">${o.pg ? `<div class="pg">стр. ${o.pg}</div>` : ''}
  <div class="iv-top"><div class="iv-l"><h1>ФАКТУРА</h1>${orig ? `<span class="tag">${esc(orig)}</span>` : ''}</div><div class="iv-c">${lg}</div>
   <div class="iv-meta"><div class="no"><span>№</span> ${esc(padNo(S.nextNo))}</div><div><span>дата на издаване:</span> ${dmy(d.date)}</div><div><span>дата на данъчно събитие:</span> ${dmy(d.tax || d.date)}</div></div></div>
  <div class="iv-parties">${party('П О Л У Ч А Т Е Л', esc(d.client), c)}${party('Д О С Т А В Ч И К', sup, co)}</div>
  <table class="ib"><colgroup><col style="width:8mm"><col><col style="width:15mm"><col style="width:20mm"><col style="width:24mm"><col style="width:26mm"></colgroup>
   <thead><tr><th>№</th><th>наименование на стоката и услугата</th><th>код</th><th>количество</th><th>ед. цена</th><th>стойност</th></tr></thead><tbody>${rows}</tbody></table>
  ${tail}</div>`;
}
const CPL = 50; // approx. characters of an item name per printed line
const wOf = l => Math.max(1, Math.ceil((l.name || '').length / CPL));
function chunkRows(items) { // pages for a full-page copy, packed by weight (long names take several lines)
  const pages = []; let cur = [], w = 0;
  items.forEach(l => { const x = wOf(l); if (cur.length && w + x > MID) { pages.push(cur); cur = []; w = 0; } cur.push(l); w += x; });
  pages.push(cur);
  const wt = p => p.reduce((a, l) => a + wOf(l), 0);
  const last = pages[pages.length - 1];
  if (wt(last) > LAST) { // totals need room: move a leading part of the last page onto its own page
    let k = 1; while (k < last.length && wt(last.slice(k)) > LAST) k++;
    pages.splice(pages.length - 1, 1, last.slice(0, k), last.slice(k));
    if (!pages[pages.length - 1].length) pages.pop();
  }
  return pages;
}
function invoiceSheetHtml(d, c, calc) {
  const items = calc.items, W = items.reduce((a, l) => a + wOf(l), 0);
  if (W <= HALF)
    return `<div class="page">${copyHtml(d, c, calc, d.orig, items, { start: 0, pad: HALF - W + items.length, last: true })}${copyHtml(d, c, calc, '', items, { start: 0, pad: HALF - W + items.length, last: true })}</div>`;
  const pages = chunkRows(items);
  return [d.orig, ''].map(orig => { let at = 0;
    return pages.map((pg, pi) => { const html = `<div class="page full">${copyHtml(d, c, calc, orig, pg, { start: at, pad: 0, last: pi === pages.length - 1, pg: pages.length > 1 ? `${pi + 1} от ${pages.length}` : '' })}</div>`; at += pg.length; return html; }).join('');
  }).join('');
}
function updateInvoice() {
  const d = S.inv, calc = calcInv(), c = findClient(d.client);
  $$('#i_items tbody tr').forEach((tr, i) => { if (calc.lines[i]) tr.querySelector('.v').textContent = calc.lines[i].used ? eur(calc.lines[i].val) : ''; });
  $('#t_vatl').textContent = 'ДДС ' + num(d.rate) + '%'; $('#t_base').textContent = eur(calc.base); $('#t_vat').textContent = eur(calc.vat); $('#t_tot').textContent = eur(calc.total);
  const ci = $('#i_clientInfo');
  ci.innerHTML = d.client ? (c ? clientInfoHtml(c) : '<span style="color:#b3261e">Клиентът не е в списъка. <a href="#" id="i_nfAdd">Добави го като нов клиент</a></span>') : '';
  $('#invSheet').innerHTML = invoiceSheetHtml(d, c, calc);
  const np = $$('#invSheet .page').length;
  $('#p_info').textContent = `${calc.items.length} реда · ${np} ${np === 1 ? 'лист (оригинал и копие)' : 'листа'}`;
  partyTools();
}
function invMsg(m, ok) { const h = m ? `<div class="${ok ? 'okmsg' : 'err'}">${esc(m)}</div>` : ''; $('#i_msg').innerHTML = h; $('#p_msg').innerHTML = h; }
function validateInv() {
  const d = S.inv;
  if (!d.client.trim()) return 'Няма избран клиент!';
  if (!findClient(d.client)) return 'Клиентът не е в списъка „Клиенти“. Добавете го първо.';
  if (!d.date) return 'Няма дата!';
  if (!d.lines.some(l => l.name.trim())) return 'Няма добавени стоки!';
  return '';
}
function recordInvoice() {
  const d = S.inv, calc = calcInv(), c = findClient(d.client), no = /^\d+$/.test(String(S.nextNo).trim()) ? +S.nextNo : S.nextNo;
  const rec = {
    date: d.date, tax: d.tax, no, client: d.client, vat: c.vat || '', eik: c.eik || '',
    lines: calc.items.map(l => l.name), base: calc.base, dds: calc.vat, total: calc.total, pay: d.pay, comp: d.comp, rate: d.rate,
    items: calc.items.map(l => ({ name: l.name, code: l.code, qty: l.qty, price: l.price }))
  };
  if (d.edit && S.invoices[d.edit.idx]) S.invoices[d.edit.idx] = rec; else S.invoices.push(rec);
}
function clearInvoice() {
  S.inv.client = ''; S.inv.date = today(); S.inv.tax = S.inv.date; S.inv.taxManual = false; S.inv.lines = blankLines(HALF);
  S.nextNo = (parseInt(S.nextNo) || 0) + 1;
}
function printSheet(cls) {
  return new Promise(res => {
    document.body.classList.add(cls);
    setTimeout(() => {
      let done = false;
      const fin = () => { if (done) return; done = true; removeEventListener('afterprint', fin); document.body.classList.remove(cls); res(); };
      addEventListener('afterprint', fin); window.print(); setTimeout(fin, 400);
    }, 60);
  });
}
const draftEmpty = () => !S.inv.client && S.inv.lines.every(l => !l.name && !l.qty && !l.price);
function endEdit() { const e = S.inv.edit; if (!e) return; S.inv = e.stash; S.nextNo = e.prevNext; }
function afterRecord() { if (S.inv.edit) endEdit(); else clearInvoice(); }
function openInvoice(idx) {
  const r = S.invoices[idx]; if (!r || !r.items) return;
  if (S.inv.edit) { if (!confirm('Излизате от режим „редакция“ на друга фактура без запис. Продължавам?')) return; endEdit(); }
  const stash = JSON.parse(JSON.stringify(S.inv)); delete stash.edit;
  S.inv = { client: r.client, date: r.date, tax: r.tax || r.date, taxManual: !!r.tax && r.tax !== r.date, pay: r.pay, comp: r.comp || stash.comp, rate: r.rate ?? S.rate, orig: stash.orig, lines: r.items.map(x => ({ ...x })), edit: { idx, no: r.no, stash, prevNext: S.nextNo } };
  S.nextNo = r.no; saveNow(); showTab('inv'); invMsg('');
}
function copyInvoice(idx) {
  const r = S.invoices[idx]; if (!r) return;
  if (S.inv.edit) { if (!confirm('Излизате от режим „редакция“ без запис. Продължавам?')) return; endEdit(); }
  else if (!draftEmpty() && !confirm('Текущата чернова ще бъде заменена. Продължавам?')) return;
  let lines;
  if (r.items) lines = r.items.map(x => ({ ...x }));
  else { const names = r.lines.filter(Boolean); lines = names.map(n => ({ name: n, code: S.units[0] || '', qty: '1', price: names.length === 1 && r.base ? String(r.base) : '' })); }
  Object.assign(S.inv, { client: r.client, date: today(), tax: today(), taxManual: false, pay: r.pay || S.inv.pay, comp: r.comp || S.inv.comp, lines: lines.length ? lines : blankLines(HALF) });
  saveNow(); showTab('inv');
  invMsg(r.items ? `Копирана е фактура № ${r.no}. Проверете количествата и цените.` : `Копирани са наименованията от стария запис № ${r.no} (старите записи нямат количества и цени) — попълнете ги.`, true);
}
function newInvoice() {
  if (S.inv.edit) endEdit();
  S.inv.client = ''; S.inv.date = today(); S.inv.tax = S.inv.date; S.inv.taxManual = false; S.inv.lines = blankLines(HALF);
  saveNow(); showTab('inv'); invMsg(''); $('#i_client').focus();
}

/* checks before a record is saved: duplicate / skipped number, date order */
function numberWarnings() {
  const d = S.inv, no = parseInt(S.nextNo), self = d.edit ? S.invoices[d.edit.idx] : null, w = [];
  if (!isFinite(no)) return ['Номерът на фактурата не е число.'];
  const others = S.invoices.filter(r => r !== self && /^\d+$/.test(String(r.no)));
  const dup = others.some(r => +r.no === no);
  if (dup) w.push(`Фактура № ${no} вече съществува в регистъра (ще се получи дубликат).`);
  const max = others.reduce((m, r) => Math.max(m, +r.no), 0);
  if (max && no > max + 1) w.push(`Пропускате ${no - 1 === max + 1 ? 'номер ' + (max + 1) : 'номера ' + (max + 1) + '–' + (no - 1)} (последният издаден е № ${max}).`);
  if (max && no < max && !dup) w.push(`Номерът е по-малък от последния издаден (№ ${max}).`);
  const prev = others.filter(r => +r.no < no).sort((a, b) => b.no - a.no)[0], next = others.filter(r => +r.no > no).sort((a, b) => a.no - b.no)[0];
  if (prev && prev.date && d.date && prev.date > d.date) w.push(`Датата (${dmy(d.date)}) е по-ранна от тази на фактура № ${prev.no} (${dmy(prev.date)}).`);
  if (next && next.date && d.date && next.date < d.date) w.push(`Датата (${dmy(d.date)}) е по-късна от тази на фактура № ${next.no} (${dmy(next.date)}), която има по-голям номер.`);
  if (d.tax && d.date && d.tax > d.date) w.push('Датата на данъчното събитие е след датата на издаване.');
  return w;
}
const confirmWarnings = () => { const w = numberWarnings(); return !w.length || confirm('Проверка преди запис:\n\n• ' + w.join('\n• ') + '\n\nДа продължа ли?'); };

/* last price a client was invoiced for an item */
function lastDeal(client, name) {
  let best = null;
  S.invoices.forEach(r => {
    if (r.client !== client) return;
    let hit = null;
    if (r.items) hit = r.items.find(x => x.name === name);
    else { const ls = r.lines.filter(Boolean); if (ls.length === 1 && ls[0] === name && r.base) hit = { name, code: '', qty: '1', price: r.base }; }
    if (hit && num(hit.price) && (!best || (r.date || '') >= best.date)) best = { date: r.date || '', ...hit };
  });
  return best;
}

$('#tab-inv').addEventListener('input', e => {
  const t = e.target, d = S.inv;
  if (t.id === 'i_client') d.client = t.value;
  else if (t.id === 'i_no') { S.nextNo = t.value.replace(/\D/g, '').replace(/^0+(?=\d)/, ''); }
  else if (t.id === 'i_date') { d.date = t.value; if (!d.taxManual) { d.tax = d.date; $('#i_tax').value = d.tax; } }
  else if (t.id === 'i_tax') { d.tax = t.value; d.taxManual = !!d.tax && d.tax !== d.date; if (!d.tax) { d.tax = d.date; t.value = d.tax; } $('#i_taxReset').classList.toggle('hide', !d.taxManual); }
  else if (t.id === 'i_pay') d.pay = t.value;
  else if (t.id === 'i_comp') d.comp = t.value;
  else if (t.id === 'i_orig') d.orig = t.value;
  else if (t.dataset.f) {
    if (t.tagName === 'TEXTAREA') { if (/\n/.test(t.value)) t.value = t.value.replace(/\s*\n+\s*/g, ' '); grow(t); }
    const tr = t.closest('tr'), l = d.lines[tr.dataset.i]; l[t.dataset.f] = t.value;
    if (t.dataset.f === 'name') {
      const a = findArt(t.value), deal = d.client && t.value ? lastDeal(d.client, t.value) : null, pe = $('[data-f=price]', tr);
      if (a && a.unit && !l.code) { l.code = a.unit; setUnit($('[data-f=code]', tr), a.unit); }
      if (!num(l.price)) {
        if (deal) {
          l.price = String(deal.price); pe.value = l.price; pe.title = 'Последна цена за този клиент: ' + dmy(deal.date);
          if (!l.qty) { l.qty = String(deal.qty || 1); $('[data-f=qty]', tr).value = l.qty; }
          if (!l.code && deal.code) { l.code = deal.code; setUnit($('[data-f=code]', tr), deal.code); }
        } else if (a && a.price !== '') { l.price = a.price; pe.value = a.price; }
      }
    }
    if (t.dataset.f === 'name' && !l.code && t.value) { l.code = S.units[0] || ''; setUnit($('[data-f=code]', tr), l.code); }
  } else return;
  invMsg(''); updateInvoice(); save();
});
$('#i_no').addEventListener('change', e => { e.target.value = padNo(S.nextNo); });
$('#i_items').addEventListener('click', e => {
  const b = e.target.closest('[data-del]'); if (!b) return;
  const L = S.inv.lines, i = +b.dataset.del;
  if (L.length > 1) L.splice(i, 1); else L[0] = { name: '', code: '', qty: '', price: '' };
  save(); buildInvoiceForm(); updateInvoice();
});
$('#s_items').addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.tagName === 'TEXTAREA') e.preventDefault(); });
$('#i_items').addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.tagName === 'TEXTAREA') { e.preventDefault(); return; }
  if (e.key !== 'Enter' || e.defaultPrevented || e.target.dataset.f !== 'price') return;
  const tr = e.target.closest('tr'); e.preventDefault();
  if (tr.nextElementSibling) $('[data-f=name]', tr.nextElementSibling).focus(); else addInvRow(true);
});
$('#i_addRow').onclick = () => addInvRow(true);
$('#i_save').onclick = async () => {
  const m = validateInv(); if (m) return invMsg(m);
  if (!confirmWarnings()) return;
  const wasEdit = !!S.inv.edit, no = S.nextNo;
  recordInvoice(); saveNow(); backupNow(); updateInvoice();
  await printSheet('p-inv');
  afterRecord(); saveNow(); showTab('inv'); invMsg(wasEdit ? `Фактура № ${no} е обновена в регистъра.` : 'Фактурата е записана в „Издадени фактури“.', true);
};
$('#i_saveonly').onclick = () => {
  const m = validateInv(); if (m) return invMsg(m);
  if (!confirmWarnings()) return;
  const wasEdit = !!S.inv.edit, no = S.nextNo;
  recordInvoice(); afterRecord(); saveNow(); backupNow(); showTab('inv'); invMsg(wasEdit ? `Фактура № ${no} е обновена в регистъра.` : 'Фактурата е записана (без печат).', true);
};
$('#i_print').onclick = () => { const m = validateInv(); if (m) return invMsg(m); printSheet('p-inv'); };
$('#i_prev').onclick = () => showTab('prev');
$('#p_back').onclick = () => showTab('inv');
$('#p_save').onclick = () => $('#i_save').onclick();
$('#p_print').onclick = () => $('#i_print').onclick();
$('#i_clear').onclick = () => { if (S.inv.edit) { endEdit(); } else { S.inv.client = ''; S.inv.tax = S.inv.date; S.inv.taxManual = false; S.inv.lines = blankLines(HALF); } save(); invMsg(''); syncInvForm(); updateInvoice(); };
$('#i_taxReset').onclick = e => { e.preventDefault(); S.inv.taxManual = false; S.inv.tax = S.inv.date; save(); syncInvForm(); updateInvoice(); };
$('#i_cancelEdit').onclick = () => { endEdit(); saveNow(); invMsg(''); syncInvForm(); updateInvoice(); };
$('#p_pdf').onclick = async () => {
  const m = validateInv(); if (m) return invMsg(m);
  if (!HOST) { invMsg('Изберете „Save as PDF“ като принтер в прозореца за печат.', true); await printSheet('p-inv'); return; }
  const nm = `${S.nextNo}_${(S.inv.client || '').replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 40)}_${S.inv.date}.pdf`.replace(/\s+/g, ' ');
  document.body.classList.add('p-inv'); await new Promise(r => setTimeout(r, 150));
  try { const r = await hostCall('pdf', { name: nm }); invMsg(r ? 'PDF е запазен: ' + r : '', !!r); }
  catch (e) { invMsg('Грешка при запис на PDF: ' + e.message); }
  finally { document.body.classList.remove('p-inv'); }
};

/* ---- quick add client ---- */
async function addClientDlg(prefill) {
  const wo = warnOnce();
  const o = await ask('Нов клиент', [{ k: 'name', label: 'Име (ГЛАВНИ БУКВИ, без кавички)', max: 51 }, { k: 'addr', label: 'Адрес' }, { k: 'eik', label: 'ЕИК / БУЛСТАТ' }, { k: 'vat', label: 'Номер по ЗДДС (по избор)' }, { k: 'isvat', label: 'Регистриран по ЗДДС (№ = BG + ЕИК)', type: 'check' }, { k: 'mol', label: 'МОЛ' }], { name: prefill || '' },
    o => !o.name ? 'Въведете име.' : findClient(o.name) ? 'Клиент с това име вече съществува.' : wo(o.eik, o.vat || (o.isvat ? 'BG' + o.eik.replace(/\D/g, '') : '')));
  if (!o) return;
  const eik = o.eik.replace(/\D/g, '');
  S.clients.push({ name: o.name, addr: o.addr, vat: o.vat || (o.isvat && eik ? 'BG' + eik : ''), eik: eik || o.eik, mol: o.mol });
  S.inv.client = o.name; saveNow(); fillDatalists(); syncInvForm(); updateInvoice();
}
$('#i_clientX').onclick = () => { S.inv.client = ''; $('#i_client').value = ''; invMsg(''); updateInvoice(); save(); $('#i_client').focus(); };
$('#i_editClient').onclick = e => { e.preventDefault(); editPartyDlg(S.clients, S.inv.client, 'Редакция на клиент', n => { S.inv.client = n; syncInvForm(); updateInvoice(); }); };
$('#i_addClient').onclick = () => addClientDlg(S.inv.client && !findClient(S.inv.client) ? S.inv.client : '');
$('#i_clientInfo').addEventListener('click', e => { if (e.target.id === 'i_nfAdd') { e.preventDefault(); addClientDlg(S.inv.client); } });


/* ---------------- STOCK RECEIPT (same workflow as the invoice) ---------------- */
const sBlank = n => Array.from({ length: n }, () => ({ name: '', unit: 'БР.', qty: '', price: '' }));
function ensureStock() { // older drafts had a free-text date and no numbering
  const d = S.stock;
  if (!d.lines || !d.lines.length || (!d.edit && d.lines.length > HALF && d.lines.every(l => !l.name && !l.qty && !l.price))) d.lines = sBlank(RHALF);
  if (!/^\d{4}-\d\d-\d\d$/.test(d.date || '')) d.date = today();
  if (d.orig === undefined) d.orig = 'ОРИГИНАЛ';
  if (!S.receipts) S.receipts = [];
}
function calcStock(d = S.stock) {
  const lines = d.lines.map(l => ({ ...l, val: r2(num(l.qty) * num(l.price)), used: (l.name || '').trim() !== '' || num(l.qty) * num(l.price) !== 0 }));
  return { lines, items: lines.filter(l => l.used), total: r2(lines.reduce((a, l) => a + l.val, 0)) };
}
function buildStockForm() {
  const L = S.stock.lines;
  $('#s_items tbody').innerHTML = L.map((l, i) => `<tr data-i="${i}">
   <td>${i + 1}</td>
   <td><textarea rows="1" data-f="name" data-search="arts" autocomplete="off">${esc(l.name)}</textarea></td>
   <td>${unitSel('unit', l.unit)}</td>
   <td><input data-f="qty" inputmode="decimal" style="text-align:right" value="${esc(l.qty)}"></td>
   <td><input data-f="price" inputmode="decimal" style="text-align:right" value="${esc(l.price)}"></td>
   <td class="n v"></td>
   <td><button class="xbtn" data-del="${i}" title="Премахни реда">✕</button></td></tr>`).join('');
  $('#s_rowInfo').textContent = L.length + ' реда';
  growAll($('#s_items'));
}
function addStockRow(focus) {
  S.stock.lines.push({ name: '', unit: 'БР.', qty: '', price: '' }); save(); buildStockForm(); updateStock();
  if (focus) { const r = $$('#s_items tbody tr').pop(); $('[data-f=name]', r).focus(); }
}
function syncStockForm() {
  const d = S.stock; ensureStock();
  $('#s_sup').value = d.sup; $('#s_date').value = d.date; fillSelect($('#s_comp'), S.compilers, d.comp);
  $('#s_edit').classList.toggle('hide', !d.edit);
  if (d.edit) $('#s_editTxt').textContent = `Редакция на записана разписка от ${dmy(d.date)}. „Печат и запис“ и „Само запис“ обновяват записа в регистъра — не се създава нов номер.`;
  buildStockForm();
}
function renderStock() { fillDatalists(); syncStockForm(); updateStock(); }
function rcpCopyHtml(d, s, calc, orig, items, o) {
  const co = S.company, logo = esc(co.name);   // the warehouse receipt has no logo
  let rows = items.map((l, i) => `<tr><td class="c">${o.start + i + 1}</td><td>${esc(l.name)}</td><td class="c">${esc(l.unit)}</td><td class="c">${qtyFmt(l.qty)}</td><td class="n">${eur(num(l.price))}</td><td class="n">${eur(l.val)}</td></tr>`).join('');
  for (let i = items.length; i < o.pad; i++) rows += '<tr><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>';
  const party = (lbl, nm, p) => `<div class="pbox"><div class="lbl">${lbl}</div><div class="nm">${nm}</div><div class="kv"><span>Адрес:</span><span>${esc(p?.addr)}</span><span>№ по ЗДДС:</span><span>${esc(p?.vat)}</span><span>ЕИК/ЕГН:</span><span>${esc(p?.eik)}</span><span>МОЛ:</span><span>${esc(p?.mol)}</span></div></div>`;
  const tail = o.last ? `<div class="iv-sum"><div class="iv-words">Сума словом: <b>${esc(slov(calc.total))}</b></div>
   <table class="iv-tot"><tr class="g"><td>всичко:</td><td>${eur(calc.total)}</td></tr></table></div>
  <div class="iv-foot"><div></div><div class="sg">подпис на доставчика<b>&nbsp;</b></div><div class="sg">съставил /подпис/<b>${esc(d.comp)}</b></div></div>`
    : `<div class="cont">продължава на следващата страница…</div>`;
  return `<div class="copy">${o.pg ? `<div class="pg">стр. ${o.pg}</div>` : ''}
  <div class="iv-top"><div class="iv-l"><h1 style="font-size:19px;letter-spacing:3px">СКЛАДОВА РАЗПИСКА</h1>${orig ? `<span class="tag">${esc(orig)}</span>` : ''}</div><div class="iv-c"></div>
   <div class="iv-meta"><div class="no"><span>дата</span> ${dmy(d.date)}</div></div></div>
  <div class="iv-parties">${party('Д О С Т А В Ч И К', esc(d.sup), s)}${party('И З Д А В А', logo, co)}</div>
  <table class="ib"><colgroup><col style="width:8mm"><col><col style="width:15mm"><col style="width:20mm"><col style="width:24mm"><col style="width:26mm"></colgroup>
   <thead><tr><th>№</th><th>наименование на стоката</th><th>мярка</th><th>количество</th><th>ед. цена</th><th>стойност</th></tr></thead><tbody>${rows}</tbody></table>
  ${tail}</div>`;
}
function stockSheetHtml(d, s, calc) { // ONE copy per receipt: half an A4 sheet when it is short, a full page (continued if needed) when it is long
  const items = calc.items, W = items.reduce((a, l) => a + wOf(l), 0);
  if (W <= RHALF) // short receipt: a single copy on the TOP HALF of the sheet (the lower half stays free)
    return `<div class="page one">${rcpCopyHtml(d, s, calc, '', items, { start: 0, pad: RHALF - W + items.length, last: true })}</div>`;
  const pages = W <= LAST ? [items] : chunkRows(items);
  let at = 0;
  return pages.map((pg, pi) => {
    const last = pi === pages.length - 1, w = pg.reduce((a, l) => a + wOf(l), 0);
    const html = `<div class="page full">${rcpCopyHtml(d, s, calc, '', pg, { start: at, pad: pages.length === 1 ? Math.max(0, 12 - w) + pg.length : 0, last, pg: pages.length > 1 ? `${pi + 1} от ${pages.length}` : '' })}</div>`;
    at += pg.length; return html;
  }).join('');
}
function updateStock() {
  const d = S.stock, calc = calcStock(), s = findSup(d.sup);
  $$('#s_items tbody tr').forEach((tr, i) => { if (calc.lines[i]) tr.querySelector('.v').textContent = calc.lines[i].used ? eur(calc.lines[i].val) : ''; });
  $('#st_tot').textContent = eur(calc.total);
  $('#s_supInfo').innerHTML = d.sup ? (s ? clientInfoHtml(s) : '<span style="color:#b3261e">Доставчикът не е в списъка. <a href="#" id="s_nfAdd">Добави го като нов доставчик</a></span>') : '';
  $('#stockSheet').innerHTML = stockSheetHtml(d, s, calc);
  const np = $$('#stockSheet .page').length;
  $('#sp_info').textContent = `${calc.items.length} реда · ${np} ${np === 1 ? 'лист' : 'листа'} (един екземпляр)`;
  partyTools();
}
function sMsg(m, ok) { const x = m ? `<div class="${ok ? 'okmsg' : 'err'}">${esc(m)}</div>` : ''; $('#s_msg').innerHTML = x; $('#sp_msg').innerHTML = x; }
function validateStock() {
  const d = S.stock;
  if (!d.sup.trim()) return 'Няма избран доставчик!';
  if (!findSup(d.sup)) return 'Доставчикът не е в списъка „Доставчици“. Добавете го първо.';
  if (!d.date) return 'Няма дата!';
  if (!d.lines.some(l => l.name.trim())) return 'Няма добавени стоки!';
  return '';
}
function recordReceipt() {
  const d = S.stock, calc = calcStock(), sp = findSup(d.sup);
  const rec = { date: d.date, sup: d.sup, vat: sp.vat || '', eik: sp.eik || '', lines: calc.items.map(l => l.name), total: calc.total, comp: d.comp,
    items: calc.items.map(l => ({ name: l.name, unit: l.unit, qty: l.qty, price: l.price })) };
  if (d.edit && S.receipts[d.edit.idx]) S.receipts[d.edit.idx] = rec; else S.receipts.push(rec);
}
function clearStock() { S.stock.sup = ''; S.stock.date = today(); S.stock.lines = sBlank(RHALF); }
const stockEmpty = () => !S.stock.sup && S.stock.lines.every(l => !l.name && !l.qty && !l.price);
function endStockEdit() { const e = S.stock.edit; if (!e) return; S.stock = e.stash; }
function afterReceipt() { if (S.stock.edit) endStockEdit(); else clearStock(); }
function openReceipt(idx) {
  const r = S.receipts[idx]; if (!r) return;
  if (S.stock.edit) { if (!confirm('Излизате от режим „редакция“ на друга разписка без запис. Продължавам?')) return; endStockEdit(); }
  const stash = JSON.parse(JSON.stringify(S.stock)); delete stash.edit;
  S.stock = { sup: r.sup, date: r.date, comp: r.comp || stash.comp, orig: stash.orig, lines: r.items.map(x => ({ ...x })), edit: { idx, stash } };
  saveNow(); showTab('stock'); sMsg('');
}
function copyReceipt(idx) {
  const r = S.receipts[idx]; if (!r) return;
  if (S.stock.edit) { if (!confirm('Излизате от режим „редакция“ без запис. Продължавам?')) return; endStockEdit(); }
  else if (!stockEmpty() && !confirm('Текущата чернова ще бъде заменена. Продължавам?')) return;
  Object.assign(S.stock, { sup: r.sup, date: today(), comp: r.comp || S.stock.comp, lines: r.items.length ? r.items.map(x => ({ ...x })) : sBlank(RHALF) });
  saveNow(); showTab('stock'); sMsg(`Копирана е разписка от ${dmy(r.date)} (${r.sup}). Проверете количествата и цените.`, true);
}
function newReceipt() {
  if (S.stock.edit) endStockEdit();
  S.stock.sup = ''; S.stock.date = today(); S.stock.lines = sBlank(RHALF);
  saveNow(); showTab('stock'); sMsg(''); $('#s_sup').focus();
}
function lastStockDeal(sup, name) {
  let best = null;
  S.receipts.forEach(r => { if (r.sup !== sup) return; const hit = (r.items || []).find(x => x.name === name); if (hit && num(hit.price) && (!best || (r.date || '') >= best.date)) best = { date: r.date || '', ...hit }; });
  return best;
}

$('#tab-stock').addEventListener('input', e => {
  const t = e.target, d = S.stock;
  if (t.id === 's_sup') d.sup = t.value;
  else if (t.id === 's_date') d.date = t.value;
  else if (t.id === 's_comp') d.comp = t.value;
  else if (t.dataset.f) {
    if (t.tagName === 'TEXTAREA') { if (/\n/.test(t.value)) t.value = t.value.replace(/\s*\n+\s*/g, ' '); grow(t); }
    const tr = t.closest('tr'), l = d.lines[tr.dataset.i]; l[t.dataset.f] = t.value;
    if (t.dataset.f === 'name') {
      const a = findArt(t.value), deal = d.sup && t.value ? lastStockDeal(d.sup, t.value) : null, pe = $('[data-f=price]', tr);
      if (a && a.unit) { l.unit = a.unit; setUnit($('[data-f=unit]', tr), a.unit); }
      if (!num(l.price)) {
        if (deal) {
          l.price = String(deal.price); pe.value = l.price; pe.title = 'Последна цена от този доставчик: ' + dmy(deal.date);
          if (!l.qty) { l.qty = String(deal.qty || 1); $('[data-f=qty]', tr).value = l.qty; }
          if (deal.unit) { l.unit = deal.unit; setUnit($('[data-f=unit]', tr), deal.unit); }
        } else if (a && a.price !== '') { l.price = a.price; pe.value = a.price; }
      }
    }
  } else return;
  sMsg(''); updateStock(); save();
});
$('#s_items').addEventListener('click', e => {
  const b = e.target.closest('[data-del]'); if (!b) return;
  const L = S.stock.lines, i = +b.dataset.del;
  if (L.length > 1) L.splice(i, 1); else L[0] = { name: '', unit: 'БР.', qty: '', price: '' };
  save(); buildStockForm(); updateStock();
});
$('#s_items').addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.tagName === 'TEXTAREA') { e.preventDefault(); return; }
  if (e.key !== 'Enter' || e.defaultPrevented || e.target.dataset.f !== 'price') return;
  const tr = e.target.closest('tr'); e.preventDefault();
  if (tr.nextElementSibling) $('[data-f=name]', tr.nextElementSibling).focus(); else addStockRow(true);
});
$('#s_addRow').onclick = () => addStockRow(true);
$('#s_save').onclick = async () => {
  const m = validateStock(); if (m) return sMsg(m);
  const wasEdit = !!S.stock.edit;
  recordReceipt(); saveNow(); backupNow(); updateStock();
  await printSheet('p-stock');
  afterReceipt(); saveNow(); showTab('stock'); sMsg(wasEdit ? 'Разписката е обновена в регистъра.' : 'Разписката е записана в „Издадени разписки“.', true);
};
$('#s_saveonly').onclick = () => {
  const m = validateStock(); if (m) return sMsg(m);
  const wasEdit = !!S.stock.edit;
  recordReceipt(); afterReceipt(); saveNow(); backupNow(); showTab('stock'); sMsg(wasEdit ? 'Разписката е обновена в регистъра.' : 'Разписката е записана (без печат).', true);
};
$('#s_print').onclick = () => { const m = validateStock(); if (m) return sMsg(m); printSheet('p-stock'); };
$('#s_prev').onclick = () => showTab('sprev');
$('#sp_back').onclick = () => showTab('stock');
$('#sp_save').onclick = () => $('#s_save').onclick();
$('#sp_print').onclick = () => $('#s_print').onclick();
$('#s_clear').onclick = () => { if (S.stock.edit) endStockEdit(); else { S.stock.sup = ''; S.stock.lines = sBlank(RHALF); } save(); sMsg(''); syncStockForm(); updateStock(); };
$('#s_cancelEdit').onclick = () => { endStockEdit(); saveNow(); sMsg(''); syncStockForm(); updateStock(); };
$('#sp_pdf').onclick = async () => {
  const m = validateStock(); if (m) return sMsg(m);
  if (!HOST) { sMsg('Изберете „Save as PDF“ като принтер в прозореца за печат.', true); await printSheet('p-stock'); return; }
  const nm = `Разписка_${(S.stock.sup || '').replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 40)}_${S.stock.date}.pdf`.replace(/\s+/g, ' ');
  document.body.classList.add('p-stock'); await new Promise(r => setTimeout(r, 150));
  try { const r = await hostCall('pdf', { name: nm }); sMsg(r ? 'PDF е запазен: ' + r : '', !!r); }
  catch (e) { sMsg('Грешка при запис на PDF: ' + e.message); }
  finally { document.body.classList.remove('p-stock'); }
};
async function addSupDlg(prefill) {
  const wo = warnOnce();
  const o = await ask('Нов доставчик', [{ k: 'name', label: 'Име (ГЛАВНИ БУКВИ, без кавички, напр. АСЕТ - ПЕТЪР ПЕТРОВ ЕТ)', max: 51 }, { k: 'addr', label: 'Адрес' }, { k: 'vat', label: 'Регистриран по ЗДДС (№ по ДДС = BG + булстат)', type: 'check' }, { k: 'eik', label: 'Булстат' }, { k: 'mol', label: 'МОЛ' }], { name: prefill || '' },
    o => !o.name ? 'Въведете име.' : findSup(o.name) ? 'Доставчик с това име вече съществува.' : wo(o.eik, o.vat ? 'BG' + o.eik.replace(/\D/g, '') : ''));
  if (!o) return;
  const eik = o.eik.replace(/\D/g, '');
  S.suppliers.push({ name: o.name, vat: o.vat && eik ? 'BG' + eik : '', eik, mol: o.mol, addr: o.addr });
  S.stock.sup = o.name; saveNow(); fillDatalists(); syncStockForm(); updateStock();
}
$('#s_supX').onclick = () => { S.stock.sup = ''; $('#s_sup').value = ''; sMsg(''); updateStock(); save(); $('#s_sup').focus(); };
$('#s_editSup').onclick = e => { e.preventDefault(); editPartyDlg(S.suppliers, S.stock.sup, 'Редакция на доставчик', n => { S.stock.sup = n; syncStockForm(); updateStock(); }); };
$('#s_addSup').onclick = e => { e.preventDefault(); addSupDlg(S.stock.sup && !findSup(S.stock.sup) ? S.stock.sup : ''); };
$('#s_supInfo').addEventListener('click', e => { if (e.target.id === 's_nfAdd') { e.preventDefault(); addSupDlg(S.stock.sup); } });
$('#s_addArt').onclick = async () => {
  const o = await ask('Нов артикул', [{ k: 'name', label: 'Наименование' }, { k: 'unit', label: 'Мярка', type: 'select', opts: S.units }, { k: 'price', label: 'Ед. цена (по избор)' }], { unit: 'БР.' },
    o => !o.name ? 'Въведете наименование.' : findArt(o.name) ? 'Артикул с това име вече съществува.' : '');
  if (!o) return;
  S.articles.push({ name: o.name, unit: o.unit, price: o.price, added: today() }); saveNow(); fillDatalists();
};

/* ---- register of issued receipts ---- */
function renderSReg() {
  ensureStock();
  const ys = [...new Set(S.receipts.map(i => (i.date || '').slice(0, 4)).filter(Boolean))].sort().reverse();
  const ysel = $('#sr_y'), msel = $('#sr_m'), cy = ysel.value || 'Всички', cm = msel.value || '0';
  ysel.innerHTML = ['Всички', ...ys].map(y => `<option ${y === cy ? 'selected' : ''}>${y}</option>`).join('');
  msel.innerHTML = MONTHS.map((m, i) => `<option value="${i}" ${String(i) === cm ? 'selected' : ''}>${m}</option>`).join('');
  const q = $('#sr_q').value.trim().toLowerCase(), y = ysel.value, m = +msel.value;
  const list = S.receipts.map((r, idx) => ({ r, idx })).filter(({ r }) =>
    (y === 'Всички' || (r.date || '').startsWith(y)) && (!m || +(r.date || '').slice(5, 7) === m) &&
    (!q || [r.sup, r.eik, r.vat, ...r.lines].join(' ').toLowerCase().includes(q)));
  list.sort((a, b) => (b.r.date || '').localeCompare(a.r.date || '') || b.idx - a.idx);
  $('#sr_t tbody').innerHTML = list.map(({ r, idx }) => `<tr><td>${dmy(r.date)}</td><td>${esc(r.sup)}</td><td>${esc(r.eik)}</td><td>${esc(r.lines.filter(Boolean).join('; '))}</td><td class="n">${eur(r.total)}</td>
   <td style="white-space:nowrap"><button class="b g s" data-o="${idx}" title="Отвори във формата за преглед и повторен печат">Отвори</button> <button class="b g s" data-c="${idx}" title="Нова разписка на база тази">Копирай</button> <button class="b g s" data-d="${idx}">✕</button></td></tr>`).join('');
  $('#sr_sum').innerHTML = `${list.length} разписки · Общо: <b>${eur(list.reduce((a, { r }) => a + (+r.total || 0), 0))}</b>`;
}
$('#sr_t').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  if (b.dataset.o) openReceipt(+b.dataset.o);
  else if (b.dataset.c) copyReceipt(+b.dataset.c);
  else if (b.dataset.d && confirm('Да изтрия ли този запис?')) { S.receipts.splice(+b.dataset.d, 1); saveNow(); renderSReg(); }
});
['sr_q', 'sr_y', 'sr_m'].forEach(id => $('#' + id).addEventListener('input', renderSReg));
$('#sr_csv').onclick = () => {
  const q = v => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  const rows = $$('#sr_t tbody tr').map(tr => { const r = S.receipts[+tr.querySelector('[data-c]').dataset.c]; return [dmy(r.date), r.sup, r.vat, r.eik, r.lines.filter(Boolean).join('; '), csvDec(r.total)].map(q).join(';'); });
  download('izdadeni_razpiski.csv', '\ufeff' + [['Дата', 'Доставчик', 'Номер по ДДС', 'БУЛСТАТ', 'Артикули', 'Всичко'].map(q).join(';'), ...rows].join('\r\n'), 'text/csv');
};

/* ---------------- REGISTER ---------------- */
const MONTHS = ['Всички', 'Януари', 'Февруари', 'Март', 'Април', 'Май', 'Юни', 'Юли', 'Август', 'Септември', 'Октомври', 'Ноември', 'Декември'];
function renderReg() {
  const ys = [...new Set(S.invoices.map(i => (i.date || '').slice(0, 4)).filter(Boolean))].sort().reverse();
  const ysel = $('#r_y'), msel = $('#r_m'), cy = ysel.value || 'Всички', cm = msel.value || '0';
  ysel.innerHTML = ['Всички', ...ys].map(y => `<option ${y === cy ? 'selected' : ''}>${y}</option>`).join('');
  msel.innerHTML = MONTHS.map((m, i) => `<option value="${i}" ${String(i) === cm ? 'selected' : ''}>${m}</option>`).join('');
  const q = $('#r_q').value.trim().toLowerCase(), y = ysel.value, m = +msel.value;
  const list = S.invoices.map((inv, idx) => ({ inv, idx })).filter(({ inv }) =>
    (y === 'Всички' || (inv.date || '').startsWith(y)) && (!m || +(inv.date || '').slice(5, 7) === m) &&
    (!q || [inv.no, inv.client, inv.eik, inv.vat, ...inv.lines].join(' ').toLowerCase().includes(q)));
  list.sort((a, b) => (b.inv.date || '').localeCompare(a.inv.date || '') || (+b.inv.no || 0) - (+a.inv.no || 0));
  $('#r_t tbody').innerHTML = list.map(({ inv, idx }) => `<tr><td>${dmy(inv.date)}</td><td>${esc(inv.no)}</td><td>${esc(inv.client)}</td><td>${esc(inv.vat)}</td><td>${esc(inv.eik)}</td><td>${esc(inv.lines.filter(Boolean).join('; '))}</td><td class="n">${eur(inv.base)}</td><td class="n">${eur(inv.dds)}</td><td class="n">${eur(inv.total)}</td><td>${esc(inv.pay)}</td>
   <td style="white-space:nowrap">${inv.items ? `<button class="b g s" data-o="${idx}" title="Отвори във формата за преглед и повторен печат">Отвори</button> ` : ''}<button class="b g s" data-c="${idx}" title="Нова фактура на база тази">Копирай</button> <button class="b g s" data-e="${idx}" title="Редакция на данните в регистъра">Данни</button> <button class="b g s" data-d="${idx}">✕</button></td></tr>`).join('');
  const s = k => list.reduce((a, { inv }) => a + (+inv[k] || 0), 0);
  $('#r_sum').innerHTML = `${list.length} фактури · Ст-ст: <b>${eur(s('base'))}</b> · ДДС: <b>${eur(s('dds'))}</b> · Общо: <b>${eur(s('total'))}</b>`;
}
const invFields = (n = 8) => [{ k: 'date', label: 'Дата', type: 'date' }, { k: 'no', label: 'Номер' }, { k: 'client', label: 'Клиент', list: 'dl_clients' }, { k: 'vat', label: 'Номер по ДДС' }, { k: 'eik', label: 'Булстат' },
  ...Array.from({ length: n }, (_, i) => ({ k: 'l' + i, label: i ? `Основание ${i + 1} ред` : 'Основание / първи ред', list: 'dl_arts' })),
  { k: 'base', label: 'Ст-ст на сделката' }, { k: 'dds', label: 'ДДС' }, { k: 'total', label: 'Общо' }, { k: 'pay', label: 'Начин на плащане', type: 'select', opts: S.payments }];
async function editInv(idx) {
  fillDatalists();
  const cur = idx == null ? { date: today(), no: S.nextNo, client: '', vat: '', eik: '', lines: [], base: '', dds: '', total: '', pay: S.payments[0] } : S.invoices[idx];
  const v = { ...cur }; (cur.lines || []).forEach((l, i) => v['l' + i] = l);
  const o = await ask(idx == null ? 'Нов запис' : 'Редакция на запис', invFields(Math.max(8, (cur.lines || []).length + 1)), v, o => !o.client ? 'Въведете клиент.' : '');
  if (!o) return;
  const rec = { ...cur, date: o.date, no: /^\d+$/.test(o.no) ? +o.no : o.no, client: o.client, vat: o.vat, eik: o.eik, lines: Array.from({ length: Math.max(8, (cur.lines || []).length + 1) }, (_, i) => o['l' + i]).filter((x, i) => x || i < 8), base: num(o.base), dds: num(o.dds), total: num(o.total), pay: o.pay };
  if (idx == null) S.invoices.push(rec); else S.invoices[idx] = rec;
  saveNow(); renderReg();
}
$('#r_t').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  if (b.dataset.e) editInv(+b.dataset.e);
  else if (b.dataset.o) openInvoice(+b.dataset.o);
  else if (b.dataset.c) copyInvoice(+b.dataset.c);
  else if (b.dataset.d) deleteInvoice(+b.dataset.d);
});
function deleteInvoice(idx) {
  const r = S.invoices[idx]; if (!r) return;
  if (S.inv.edit) { alert('Първо излезте от режим „редакция“ на фактура (бутон „Отказ“), после изтрийте фактурата.'); return; }
  const no = Number(r.no), numbered = String(r.no).trim() !== '' && isFinite(no);
  const nos = S.invoices.map(i => Number(i.no)).filter(n => isFinite(n)), max = nos.length ? Math.max(...nos) : null;
  const next = parseInt(S.nextNo) || 0, who = `${r.no}${r.client ? ' — ' + r.client : ''}`;
  if (!numbered) { if (!confirm(`Да изтрия ли записа ${who}?`)) return; }
  else if (no === max) {                          // the latest invoice: its number becomes free again
    const roll = next === no + 1;
    if (!confirm(`Да изтрия ли фактура № ${who}?` + (roll ? `\n\nТова е последната издадена фактура — следващият номер ще се върне на ${no}.` : ''))) return;
    backupNow(); S.invoices.splice(idx, 1); if (roll) S.nextNo = no;
    saveNow(); renderReg(); return;
  }
  else if (!confirm(`ВНИМАНИЕ: фактура № ${r.no} НЕ е последната (последната е № ${max}).\n\nАко я изтриете, в номерацията ще остане пропуск — липсва № ${r.no}. Фактурите трябва да са с непрекъснати номера, така че това може да създаде проблем при проверка.\n\nСигурни ли сте, че искате да я изтриете?`)) return;
  backupNow(); S.invoices.splice(idx, 1); saveNow(); renderReg();
}
$('#r_add').onclick = () => editInv(null);
['r_q', 'r_y', 'r_m'].forEach(id => $('#' + id).addEventListener('input', renderReg));
$('#r_csv').onclick = () => {
  const q = v => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  const head = ['Дата', 'Номер', 'Клиент', 'Номер по ДДС', 'БУЛСТАТ', 'Основание', 'Ст-ст', 'ДДС', 'Общо', 'Плащане'];
  const rows = $$('#r_t tbody tr').map(tr => { const idx = +tr.querySelector('[data-c]').dataset.c, i = S.invoices[idx]; return [dmy(i.date), i.no, i.client, i.vat, i.eik, i.lines.filter(Boolean).join('; '), csvDec(i.base), csvDec(i.dds), csvDec(i.total), i.pay].map(q).join(';'); });
  download('izdadeni_fakturi.csv', '﻿' + [head.map(q).join(';'), ...rows].join('\r\n'), 'text/csv');
};
const csvDec = n => (Math.round((Number(n) || 0) * 100) / 100).toFixed(2).replace('.', ',');
function download(name, text, type) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type })); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000); }

/* ---------------- VAT summary ---------------- */
function vatRows(y) {
  const rows = Array.from({ length: 12 }, (_, i) => ({ m: i + 1, n: 0, base: 0, dds: 0, total: 0 }));
  S.invoices.forEach(r => { if ((r.date || '').slice(0, 4) !== y) return; const x = rows[+r.date.slice(5, 7) - 1]; if (!x) return; x.n++; x.base += +r.base || 0; x.dds += +r.dds || 0; x.total += +r.total || 0; });
  return rows;
}
function renderVat() {
  const ys = [...new Set(S.invoices.map(i => (i.date || '').slice(0, 4)).filter(Boolean))].sort().reverse();
  const ysel = $('#v_y'), cur = ysel.value || ys[0] || String(new Date().getFullYear());
  ysel.innerHTML = (ys.length ? ys : [cur]).map(y => `<option ${y === cur ? 'selected' : ''}>${y}</option>`).join('');
  const y = ysel.value, rows = vatRows(y), sum = k => rows.reduce((a, r) => a + r[k], 0);
  const line = (lbl, r, cls) => `<tr class="${cls || ''}"><td>${lbl}</td><td class="n">${r.n}</td><td class="n">${eur(r.base)}</td><td class="n">${eur(r.dds)}</td><td class="n">${eur(r.total)}</td></tr>`;
  let h = '';
  rows.forEach((r, i) => { h += line(`<a href="#" data-vm="${r.m}">${MONTHS[r.m]}</a>`, r); if (r.m % 3 === 0) { const q = rows.slice(r.m - 3, r.m), t = { n: 0, base: 0, dds: 0, total: 0 }; q.forEach(x => { t.n += x.n; t.base += x.base; t.dds += x.dds; t.total += x.total; }); h += line(`${r.m / 3}-во тримесечие`, t, 'q'); } });
  h += line(`Общо за ${y}`, { n: sum('n'), base: sum('base'), dds: sum('dds'), total: sum('total') }, 'tt');
  $('#v_t tbody').innerHTML = h;
}
$('#v_y').addEventListener('input', renderVat);
$('#v_t').addEventListener('click', e => { const a = e.target.closest('[data-vm]'); if (!a) return; e.preventDefault(); showTab('reg'); $('#r_y').value = $('#v_y').value; $('#r_m').value = a.dataset.vm; $('#r_q').value = ''; renderReg(); });
$('#v_csv').onclick = () => {
  const y = $('#v_y').value, q = v => '"' + String(v).replace(/"/g, '""') + '"';
  const rows = vatRows(y).map(r => [MONTHS[r.m], r.n, csvDec(r.base), csvDec(r.dds), csvDec(r.total)].map(q).join(';'));
  download(`dds_spravka_${y}.csv`, '\ufeff' + [['Месец', 'Фактури', 'Данъчна основа', 'ДДС', 'Общо'].map(q).join(';'), ...rows].join('\r\n'), 'text/csv');
};

/* ---------------- sales journal (Дневник на продажбите) ---------------- */
function packPages(items, wf, cap, lastCap) {
  const pages = []; let cur = [], w = 0;
  items.forEach(it => { const k = wf(it); if (cur.length && w + k > cap) { pages.push(cur); cur = []; w = 0; } cur.push(it); w += k; });
  pages.push(cur);
  const wt = p => p.reduce((a, it) => a + wf(it), 0), last = pages[pages.length - 1];
  if (wt(last) > lastCap) { let k = 1; while (k < last.length && wt(last.slice(k)) > lastCap) k++; pages.splice(pages.length - 1, 1, last.slice(0, k), last.slice(k)); if (!pages[pages.length - 1].length) pages.pop(); }
  return pages;
}
function journalData(y, m) {
  const pre = `${y}-${String(m).padStart(2, '0')}`;
  const rows = S.invoices.filter(i => (i.date || '').startsWith(pre)).slice()
    .sort((a, b) => (+a.no || 0) - (+b.no || 0) || String(a.no).localeCompare(String(b.no)) || (a.date || '').localeCompare(b.date || ''))
    .map((i, k) => ({ n: k + 1, date: i.date, no: i.no, client: i.client, eik: i.eik || '', vat: i.vat || '', lines: (i.lines || []).filter(Boolean).join('; '), base: +i.base || 0, dds: +i.dds || 0, total: +i.total || 0, rate: i.rate != null && i.rate !== '' ? +i.rate : (i.base ? Math.round(i.dds / i.base * 1000) / 10 : S.rate), pay: i.pay }));
  const tot = rows.reduce((a, r) => ({ base: a.base + r.base, dds: a.dds + r.dds, total: a.total + r.total }), { base: 0, dds: 0, total: 0 });
  const byRate = {}; rows.forEach(r => { const k = String(r.rate); (byRate[k] = byRate[k] || { rate: r.rate, base: 0, dds: 0 }); byRate[k].base += r.base; byRate[k].dds += r.dds; });
  return { rows, tot, byRate: Object.values(byRate).sort((a, b) => b.rate - a.rate), period: `${MONTHS[m]} ${y}` };
}
function journalHtml(y, m) {
  const J = journalData(y, m), co = S.company;
  if (!J.rows.length) return '';
  const pages = packPages(J.rows, r => Math.max(1, Math.ceil((r.client || '').length / 24)), 44, 34);
  const head = (first, n, total) => `<div class="jhead"><div><h1>ДНЕВНИК НА ПРОДАЖБИТЕ</h1></div><div class="jp"><b>Данъчен период: ${esc(J.period)}</b>${total > 1 ? `<br><span style="color:#555">стр. ${n} от ${total}</span>` : ''}</div></div>` +
    (first ? `<div class="jcomp"><b>${esc(co.name)}</b> · ЕИК ${esc(co.eik)} · ЗДДС № ${esc(co.vat)} · ${esc(co.addr)}</div>` : '');
  return pages.map((pg, pi) => {
    const lastPage = pi === pages.length - 1;
    const rows = pg.map(r => `<tr><td class="c">${r.n}</td><td class="c">${dmy(r.date)}</td><td class="c">${esc(padNo(r.no))}</td><td>${esc(r.client)}</td><td>${esc(r.vat || r.eik)}</td><td class="n">${eur(r.base)}</td><td class="c">${r.rate}%</td><td class="n">${eur(r.dds)}</td><td class="n">${eur(r.total)}</td></tr>`).join('');
    const tail = lastPage ? `<table class="jtot"><tr><td>Документи:</td><td>${J.rows.length}</td></tr>` +
      J.byRate.map(g => `<tr><td>Данъчна основа ${g.rate}%:</td><td>${eur(g.base)}</td></tr><tr><td>ДДС ${g.rate}%:</td><td>${eur(g.dds)}</td></tr>`).join('') +
      `<tr class="g"><td>Общо (с ДДС):</td><td>${eur(J.tot.total)}</td></tr></table>
      <div style="margin-top:auto;display:flex;justify-content:space-between;font-size:9px;color:#444;padding-top:8mm"><span>Съставил: ${esc(S.compilers[0] || '')}</span><span>Дата на отпечатване: ${todayDmy()}</span></div>` : '';
    return `<div class="page full"><div class="copy">${head(pi === 0, pi + 1, pages.length)}
      <table class="ib jt"><colgroup><col style="width:7mm"><col style="width:17mm"><col style="width:21mm"><col><col style="width:25mm"><col style="width:23mm"><col style="width:9mm"><col style="width:19mm"><col style="width:23mm"></colgroup>
       <thead><tr><th>№</th><th>дата</th><th>документ №</th><th>получател</th><th>ЗДДС № / ЕИК</th><th>данъчна основа</th><th>ст.</th><th>ДДС</th><th>общо</th></tr></thead><tbody>${rows}</tbody></table>${tail}</div></div>`;
  }).join('');
}
function renderJournal() {
  const ys = [...new Set(S.invoices.map(i => (i.date || '').slice(0, 4)).filter(Boolean))].sort().reverse();
  const ysel = $('#j_y'), msel = $('#j_m');
  const latest = S.invoices.map(i => i.date || '').sort().pop() || today();
  const cy = ysel.value || latest.slice(0, 4), cm = msel.value || String(+latest.slice(5, 7));
  ysel.innerHTML = (ys.length ? ys : [cy]).map(y => `<option ${y === cy ? 'selected' : ''}>${y}</option>`).join('');
  msel.innerHTML = MONTHS.slice(1).map((n, i) => `<option value="${i + 1}" ${String(i + 1) === cm ? 'selected' : ''}>${n}</option>`).join('');
  const y = ysel.value, m = +msel.value, J = journalData(y, m);
  $('#j_info').textContent = J.rows.length ? `${J.rows.length} документа · основа ${eur(J.tot.base)} · ДДС ${eur(J.tot.dds)} · общо ${eur(J.tot.total)}` : 'Няма фактури за този месец.';
  $('#journalSheet').innerHTML = journalHtml(y, m);
}
['j_y', 'j_m'].forEach(id => $('#' + id).addEventListener('input', renderJournal));
$('#j_print').onclick = async () => { if (!journalData($('#j_y').value, +$('#j_m').value).rows.length) return; await printSheet('p-journal'); };
$('#j_pdf').onclick = async () => {
  const y = $('#j_y').value, m = +$('#j_m').value, msg = (t, ok) => { $('#j_msg').innerHTML = t ? `<div class="${ok ? 'okmsg' : 'err'}">${t}</div>` : ''; };
  if (!journalData(y, m).rows.length) return msg('Няма фактури за този месец.');
  if (!HOST) { msg('Изберете „Save as PDF“ като принтер в прозореца за печат.', true); await printSheet('p-journal'); return; }
  document.body.classList.add('p-journal'); await new Promise(r => setTimeout(r, 150));
  try { const r = await hostCall('pdf', { name: `Дневник_продажби_${y}-${String(m).padStart(2, '0')}.pdf` }); msg(r ? 'PDF е запазен: ' + esc(r) : '', !!r); }
  catch (e) { msg('Грешка при запис на PDF: ' + esc(e.message)); }
  finally { document.body.classList.remove('p-journal'); }
};
$('#j_csv').onclick = () => {
  const y = $('#j_y').value, m = +$('#j_m').value, J = journalData(y, m), q = v => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  if (!J.rows.length) return;
  const head = ['Ред', 'Дата', 'Номер на документ', 'Получател', 'ЕИК/БУЛСТАТ', 'Номер по ЗДДС', 'Основание', 'Данъчна основа', 'Ставка %', 'ДДС', 'Общо', 'Плащане'];
  const rows = J.rows.map(r => [r.n, dmy(r.date), padNo(r.no), r.client, r.eik, r.vat, r.lines, csvDec(r.base), String(r.rate).replace('.', ','), csvDec(r.dds), csvDec(r.total), r.pay].map(q).join(';'));
  const foot = ['', '', '', 'ОБЩО', '', '', '', csvDec(J.tot.base), '', csvDec(J.tot.dds), csvDec(J.tot.total), ''].map(q).join(';');
  download(`dnevnik_prodazhbi_${y}-${String(m).padStart(2, '0')}.csv`, '\ufeff' + [head.map(q).join(';'), ...rows, foot].join('\r\n'), 'text/csv');
};

/* ---------------- generic list editors (search, sort, computed columns) ---------------- */
// column: { k, h, ml?, opts?, type?, kind?: 'text'|'num'|'date', ro?: true, v?: (item, stats) => value }
function usageStats(pairs) {   // pairs: [[name, date], ...] -> Map name -> { last, count }
  const m = new Map();
  pairs.forEach(([n, d]) => { if (!n) return; const e = m.get(n) || { last: '', count: 0 }; e.count++; if ((d || '') > e.last) e.last = d || ''; m.set(n, e); });
  return m;
}
const artStats = () => {
  const p = [];
  S.invoices.forEach(i => { const names = i.items && i.items.length ? i.items.map(x => x.name) : (i.lines || []); new Set(names.filter(Boolean)).forEach(n => p.push([n, i.date])); });
  (S.receipts || []).forEach(r => new Set((r.items || []).map(x => x.name).filter(Boolean)).forEach(n => p.push([n, r.date])));
  return usageStats(p);
};
const clientStats = () => usageStats(S.invoices.map(i => [i.client, i.date]));
const supStats = () => usageStats((S.receipts || []).map(r => [r.sup, r.date]));

function makeList(host, key, cols, maxName, statsFn) {
  const el = $(host); let q = '';
  const sortable = cols.filter(c => c.kind || !c.ro);
  const kindOf = c => c.kind || (c.k === 'price' ? 'num' : 'text');
  const sortOf = () => (S.sorts && S.sorts[key]) || null;
  const setSort = (k, d) => { S.sorts = S.sorts || {}; if (k) S.sorts[key] = { k, d }; else delete S.sorts[key]; save(); };
  const label = (c, d) => { const kd = kindOf(c); return `${c.h}: ` + (kd === 'text' ? (d === 1 ? 'А → Я' : 'Я → А') : kd === 'date' ? (d === -1 ? 'най-нови първо' : 'най-стари първо') : (d === -1 ? 'от най-голямо' : 'от най-малко')); };
  const api = {
    render() {
      const sel = sortOf();
      const opts = ['<option value="">По ред на въвеждане</option>'].concat(sortable.flatMap(c => [1, -1].map(d => ({ c, d })).sort((a, b) => (kindOf(c) === 'text' ? b.d - a.d : a.d - b.d)).map(({ c, d }) => `<option value="${c.k}|${d}" ${sel && sel.k === c.k && sel.d === d ? 'selected' : ''}>${esc(label(c, d))}</option>`))).join('');
      el.innerHTML = `<div class="row" style="margin-bottom:8px"><div style="width:280px"><label>Търсене</label><input class="lq" value="${esc(q)}"></div><div style="width:290px"><label>Подреди по</label><select class="lsort">${opts}</select></div><button class="b g add">+ Добави</button><span class="mut cnt"></span></div>
       <div class="scroll"><table class="t"><thead><tr>${cols.map(c => { const on = sel && sel.k === c.k; return `<th ${(c.kind || !c.ro) ? `class="srt" data-s="${c.k}"` : ''} title="${(c.kind || !c.ro) ? 'Кликнете за подреждане' : ''}">${c.h}${on ? (sel.d === 1 ? ' ▲' : ' ▼') : ''}</th>`; }).join('')}<th></th></tr></thead><tbody></tbody></table></div><div class="err hide"></div>`;
      $('.lq', el).oninput = e => { q = e.target.value; api.rows(); };
      $('.lsort', el).onchange = e => { const v = e.target.value; if (!v) setSort(null); else { const [k, d] = v.split('|'); setSort(k, +d); } api.render(); };
      $('.add', el).onclick = async () => {
        const fields = cols.filter(c => !c.ro).map(c => ({ k: c.k, label: c.h, max: c.k === 'name' ? maxName : 0, type: c.type, opts: c.opts && c.opts() }));
        const o = await ask('Нов запис', fields, Object.fromEntries(cols.filter(c => c.type === 'date').map(c => [c.k, today()])), o => !o.name ? 'Въведете име.' : S[key].some(x => x.name === o.name) ? 'Вече съществува.' : '');
        if (!o) return; S[key].push(o); saveNow(); fillDatalists(); q = o.name; api.render();
      };
      api.rows();
    },
    rows() {
      const ql = q.toLowerCase(), tb = $('tbody', el), st = statsFn ? statsFn() : null;
      const val = (c, x) => c.ro ? c.v(x, st) : x[c.k];
      let items = S[key].map((x, i) => [x, i]).filter(([x]) => !ql || cols.some(c => !c.ro && String(x[c.k] ?? '').toLowerCase().includes(ql)));
      const sel = sortOf(), sc = sel && cols.find(c => c.k === sel.k);
      if (sc) {
        const kd = kindOf(sc), empty = v => v == null || v === '' || (kd === 'num' && !isFinite(num(v)));
        items = items.map(it => [...it, val(sc, it[0])]).sort((a, b) => {
          const ea = empty(a[2]), eb = empty(b[2]); if (ea || eb) return ea === eb ? a[1] - b[1] : ea ? 1 : -1;   // blanks always last
          const c = kd === 'num' ? num(a[2]) - num(b[2]) : kd === 'date' ? String(a[2]).localeCompare(String(b[2])) : String(a[2]).localeCompare(String(b[2]), 'bg', { sensitivity: 'base' });
          return (c * sel.d) || a[1] - b[1];
        }).map(([x, i]) => [x, i]);
      }
      $('.cnt', el).textContent = `${items.length} от ${S[key].length}`;
      const cell = (c, x) => {
        if (c.ro) { const v = val(c, x); return `<td class="${c.kind === 'num' ? 'n' : ''} mut">${v === '' || v == null ? '—' : (c.kind === 'date' ? esc(dmy(v)) : esc(v))}</td>`; }
        if (c.opts) return `<td><select data-k="${c.k}">${['', ...c.opts()].map(o => `<option ${o === x[c.k] ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select></td>`;
        if (c.ml) return `<td><textarea class="ml" rows="1" data-k="${c.k}">${esc(x[c.k])}</textarea></td>`;
        if (c.type === 'date') return `<td><input type="date" data-k="${c.k}" value="${esc(x[c.k])}"></td>`;
        return `<td><input data-k="${c.k}" value="${esc(x[c.k])}"></td>`;
      };
      tb.innerHTML = items.slice(0, 600).map(([x, i]) => `<tr data-i="${i}">${cols.map(c => cell(c, x)).join('')}<td><button class="b g s del">✕</button></td></tr>`).join('');
      if (items.length > 600) tb.insertAdjacentHTML('beforeend', `<tr><td colspan="${cols.length + 1}" class="mut">Показани са първите 600 — използвайте търсенето.</td></tr>`);
    }
  };
  el.addEventListener('click', e => {
    const th = e.target.closest('th.srt'); if (!th) return;
    const c = cols.find(x => x.k === th.dataset.s), cur = sortOf(), first = kindOf(c) === 'text' ? 1 : -1;
    setSort(c.k, cur && cur.k === c.k ? -cur.d : first); api.render();
  });
  el.addEventListener('change', e => {
    const t = e.target, tr = t.closest('tr'); if (!tr || !t.dataset.k) return;
    const x = S[key][+tr.dataset.i], v = t.value.replace(/\s*\n+\s*/g, ' ').trim(), err = $('.err', el);
    if (t.dataset.k === 'name') {
      if (!v) { t.value = x.name; return; }
      if (maxName && v.length > maxName) { err.textContent = `Името е над ${maxName} знака.`; err.classList.remove('hide'); t.value = x.name; return; }
      if (S[key].some(o => o !== x && o.name === v)) { err.textContent = 'Вече съществува такова име.'; err.classList.remove('hide'); t.value = x.name; return; }
    }
    err.classList.add('hide'); x[t.dataset.k] = v; saveNow(); fillDatalists();
    if ((key === 'clients' || key === 'suppliers') && (t.dataset.k === 'eik' || t.dataset.k === 'vat')) { const w = partyWarn(x.eik, x.vat); if (w.length) { err.textContent = `Внимание (${x.name}): ${w.join('; ')}.`; err.classList.remove('hide'); } }
  });
  el.addEventListener('click', e => {
    const b = e.target.closest('.del'); if (!b) return;
    const i = +b.closest('tr').dataset.i; if (confirm(`Да изтрия ли „${S[key][i].name}“?`)) { S[key].splice(i, 1); saveNow(); fillDatalists(); api.rows(); }
  });
  return api;
}
let lists;
function initLists() {
  const last = (x, st) => (st.get(x.name) || {}).last || '', cnt = (x, st) => (st.get(x.name) || {}).count || '';
  lists = {
    clients: makeList('#l_clients', 'clients', [{ k: 'name', h: 'Клиент' }, { k: 'addr', h: 'Адрес' }, { k: 'vat', h: 'Номер по ЗДДС' }, { k: 'eik', h: 'Булстат' }, { k: 'mol', h: 'МОЛ' },
      { k: 'lastUsed', h: 'Последна фактура', ro: true, kind: 'date', v: last }, { k: 'uses', h: 'Фактури', ro: true, kind: 'num', v: cnt }], 51, clientStats),
    sups: makeList('#l_sups', 'suppliers', [{ k: 'name', h: 'Доставчик' }, { k: 'vat', h: '№ по ДДС' }, { k: 'eik', h: 'Булстат' }, { k: 'mol', h: 'МОЛ' }, { k: 'addr', h: 'Адрес' },
      { k: 'lastUsed', h: 'Последна разписка', ro: true, kind: 'date', v: last }, { k: 'uses', h: 'Разписки', ro: true, kind: 'num', v: cnt }], 51, supStats),
    arts: makeList('#l_arts', 'articles', [{ k: 'name', h: 'Артикул / основание', ml: true }, { k: 'unit', h: 'Мярка', opts: () => S.units }, { k: 'price', h: 'Ед. цена', kind: 'num' },
      { k: 'added', h: 'Добавен', type: 'date', kind: 'date' }, { k: 'lastUsed', h: 'Последно използван', ro: true, kind: 'date', v: last }, { k: 'uses', h: 'Използван (бр.)', ro: true, kind: 'num', v: cnt }], 0, artStats)
  };
}

/* ---------------- settings ---------------- */
function renderSettings() {
  renderAbout();
  applyPortableUi();
  $$('[data-c]').forEach(e => e.value = S.company[e.dataset.c] || '');
  $('#c_next').value = S.nextNo; $('#c_rate').value = S.rate; renderLogoCard();
  if (HOST) { $('#bk_host').classList.remove('hide'); hostCall('info').then(i => { $('#bk_info').innerHTML = `Всички данни са на едно място: <b>${esc(i.dataDir)}</b><br>data.json — документи, клиенти, доставчици, артикули и настройки · backups\\ — автоматични архиви (при всеки записан документ и веднъж дневно, последните ${i.keep}): ${i.count} файла${i.last ? ', последен ' + esc(i.last) : ''} · config.json — позиция на прозореца`; }).catch(() => { }); }
  $('#x_units').value = S.units.join('\n'); $('#x_pays').value = S.payments.join('\n'); $('#x_comps').value = S.compilers.join('\n');
}
$('#tab-set').addEventListener('input', e => {
  const t = e.target;
  if (t.dataset.c) S.company[t.dataset.c] = t.value;
  else if (t.id === 'c_next') S.nextNo = t.value;
  else if (t.id === 'c_rate') { S.rate = t.value; if (!S.inv.edit) S.inv.rate = t.value; }
  else if (t.id === 'c_logo') S.logo = t.checked;
  else if (t.id === 'x_units') S.units = t.value.split('\n').map(s => s.trim()).filter(Boolean);
  else if (t.id === 'x_pays') S.payments = t.value.split('\n').map(s => s.trim()).filter(Boolean);
  else if (t.id === 'x_comps') S.compilers = t.value.split('\n').map(s => s.trim()).filter(Boolean);
  else return;
  fillDatalists(); save();
});
$('#bk_now').onclick = async () => { const p = await backupNow(); alert(p ? 'Архивът е записан:\n' + p : 'Неуспешен архив.'); renderSettings(); };
$('#bk_open').onclick = () => hostCall('openDataDir').catch(() => { });
$('#bk_exp').onclick = () => { saveNow(); download(`faktura_backup_${today()}.json`, JSON.stringify(S), 'application/json'); };
$('#bk_imp').onchange = e => {
  const f = e.target.files[0]; if (!f) return;
  f.text().then(t => { const j = JSON.parse(t); if (!j.invoices || !j.clients) throw new Error('Невалиден файл'); if (!confirm('Да заменя ли всички текущи данни?')) return; S = j; saveNow(); setTimeout(() => location.reload(), 300); }).catch(err => alert('Грешка: ' + err.message));
};
$('#bk_reset').onclick = async () => { if (confirm('ВНИМАНИЕ: ще бъдат изтрити ВСИЧКИ фактури, разписки, клиенти, доставчици и артикули, а настройките ще се върнат към празни. Преди това ще се направи резервно копие. Продължавам ли?')) { await backupNow(); S = freshState(); saveNow(); setTimeout(() => location.reload(), 300); } };

/* ---------------- import from a newer Excel workbook (.xlsx / .xlsm) ---------------- */
// An .xlsx is a zip of XML files; the browser engine can inflate it natively, so no library is needed.
async function xlUnzip(buf) {
  const dv = new DataView(buf), u8 = new Uint8Array(buf);
  let e = u8.length - 22; while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('Файлът не е валиден Excel (.xlsx/.xlsm) документ.');
  const n = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true); const files = {};
  for (let i = 0; i < n && dv.getUint32(p, true) === 0x02014b50; i++) {
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    files[new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nl))] = { method: dv.getUint16(p + 10, true), csize: dv.getUint32(p + 20, true), lho: dv.getUint32(p + 42, true) };
    p += 46 + nl + el + cl;
  }
  return async name => {
    const f = files[name]; if (!f) return null;
    const start = f.lho + 30 + dv.getUint16(f.lho + 26, true) + dv.getUint16(f.lho + 28, true), data = u8.subarray(start, start + f.csize);
    if (f.method === 0) return data;
    return new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
  };
}
const xlCol = ref => { let n = 0; for (const ch of ref.replace(/\d/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
async function xlRead(file) {
  const read = await xlUnzip(await file.arrayBuffer());
  const txt = async n => { const d = await read(n); return d ? new TextDecoder('utf-8').decode(d) : null; };
  const xml = t => new DOMParser().parseFromString(t, 'application/xml');
  const wbx = await txt('xl/workbook.xml'); if (!wbx) throw new Error('Файлът не прилича на Excel работна книга.');
  const wb = xml(wbx), rels = xml(await txt('xl/_rels/workbook.xml.rels')), relMap = {};
  for (const r of rels.getElementsByTagName('Relationship')) relMap[r.getAttribute('Id')] = r.getAttribute('Target');
  const sst = [], ssx = await txt('xl/sharedStrings.xml');
  if (ssx) for (const si of xml(ssx).getElementsByTagName('si')) sst.push([...si.getElementsByTagName('t')].filter(t => t.parentNode.localName !== 'rPh').map(t => t.textContent).join(''));
  const sheets = {};
  for (const sh of wb.getElementsByTagName('sheet')) {
    let target = relMap[sh.getAttribute('r:id')]; if (!target) continue;
    target = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
    const data = await txt(target); if (!data) continue;
    const rows = [];
    for (const row of xml(data).getElementsByTagName('row')) {
      const arr = [];
      for (const c of row.getElementsByTagName('c')) {
        const t = c.getAttribute('t'), v = c.getElementsByTagName('v')[0]; let val = null;
        if (t === 's') val = v ? sst[+v.textContent] : null;
        else if (t === 'inlineStr') val = [...c.getElementsByTagName('t')].map(x => x.textContent).join('');
        else if (t === 'str') val = v ? v.textContent : null;
        else if (t === 'b' || t === 'e') val = null;
        else if (v) val = Number(v.textContent);
        arr[xlCol(c.getAttribute('r'))] = val;
      }
      rows[+row.getAttribute('r')] = arr;
    }
    sheets[sh.getAttribute('name')] = rows;
  }
  return sheets;
}
const xs = v => v == null ? '' : String(v).trim();
function xlDate(v) {
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5).toISOString().slice(0, 10);
  const m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})/.exec(xs(v)); return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : '';
}
function xlParse(sheets) {
  const miss = ['КЛИЕНТИ', 'ДОСТАВЧИЦИ', 'АРТИКУЛИ', 'ИЗДАДЕНИ ФАКТУРИ'].filter(n => !sheets[n]);
  if (miss.length) throw new Error('Липсват листове: ' + miss.join(', ') + '. Това не е Excel файлът за фактури.');
  const rows = (name, from = 2) => (sheets[name] || []).map((r, i) => [i, r]).filter(([i, r]) => i >= from && r).map(([, r]) => r);
  const num0 = v => typeof v === 'number' ? r2(v) : 0;
  const out = {
    clients: rows('КЛИЕНТИ').map(r => ({ name: xs(r[0]), addr: xs(r[1]), vat: xs(r[2]), eik: xs(r[3]), mol: xs(r[4]) })).filter(x => x.name),
    suppliers: rows('ДОСТАВЧИЦИ').map(r => ({ name: xs(r[0]), vat: xs(r[1]), eik: xs(r[2]), mol: xs(r[3]), addr: xs(r[4]) })).filter(x => x.name),
    articles: rows('АРТИКУЛИ').map(r => xs(r[1])).filter(Boolean),
    invoices: rows('ИЗДАДЕНИ ФАКТУРИ').filter(r => r[1] != null || r[2] != null).map(r => ({
      date: xlDate(r[0]), no: typeof r[1] === 'number' ? Math.round(r[1]) : xs(r[1]), client: xs(r[2]), vat: xs(r[3]), eik: xs(r[4]),
      lines: [5, 6, 7, 8, 9, 10, 11, 12].map(i => xs(r[i])), base: num0(r[13]), dds: num0(r[14]), total: num0(r[15]), pay: xs(r[16])
    })).filter(x => x.no !== ''),
    units: rows('РАЗНИ').map(r => xs(r[0])).filter(Boolean), payments: rows('РАЗНИ').map(r => xs(r[2])).filter(Boolean), compilers: rows('РАЗНИ').map(r => xs(r[4])).filter(Boolean),
    nextNo: sheets['ФАКТУРА'] && sheets['ФАКТУРА'][3] && typeof sheets['ФАКТУРА'][3][5] === 'number' ? Math.round(sheets['ФАКТУРА'][3][5]) : 0
  };
  return out;
}
function xlPlan(x) {
  const P = { clients: [], suppliers: [], articles: [], invoices: [], units: [], payments: [], compilers: [], fills: [], dupInv: 0 };
  const party = (src, list, out) => {
    const seen = new Set(list.map(c => c.name));
    src.forEach(c => {
      const ex = list.find(e => e.name === c.name);
      if (ex) { ['addr', 'vat', 'eik', 'mol'].forEach(k => { if (c[k] && !ex[k]) P.fills.push([ex, k, c[k]]); }); }
      else if (!seen.has(c.name)) { seen.add(c.name); out.push(c); }
    });
  };
  party(x.clients, S.clients, P.clients); party(x.suppliers, S.suppliers, P.suppliers);
  const art = new Set(S.articles.map(a => a.name)); x.articles.forEach(n => { if (!art.has(n)) { art.add(n); P.articles.push(n); } });
  const nos = new Set(S.invoices.map(i => String(i.no)));
  x.invoices.forEach(i => { if (nos.has(String(i.no))) P.dupInv++; else { nos.add(String(i.no)); P.invoices.push(i); } });
  const add = (src, have, out) => { const h = new Set(have); src.forEach(v => { if (!h.has(v)) { h.add(v); out.push(v); } }); };
  add(x.units, S.units, P.units); add(x.payments, S.payments, P.payments); add(x.compilers, S.compilers, P.compilers);
  const all = [...S.invoices, ...P.invoices].map(i => +i.no).filter(n => isFinite(n));
  P.nextNo = Math.max(parseInt(S.nextNo) || 0, x.nextNo || 0, all.length ? Math.max(...all) + 1 : 0);
  P.nextChanges = P.nextNo !== (parseInt(S.nextNo) || 0);
  P.total = P.clients.length + P.suppliers.length + P.articles.length + P.invoices.length + P.units.length + P.payments.length + P.compilers.length + P.fills.length;
  return P;
}
function xlApply(P) {
  P.fills.forEach(([ex, k, v]) => { ex[k] = v; });
  S.clients.push(...P.clients); S.suppliers.push(...P.suppliers);
  S.articles.push(...P.articles.map(name => ({ name, unit: '', price: '', added: today() })));
  S.invoices.push(...[...P.invoices].sort((a, b) => (a.date || '').localeCompare(b.date || '') || (+a.no || 0) - (+b.no || 0)));
  S.units.push(...P.units); S.payments.push(...P.payments); S.compilers.push(...P.compilers);
  if (S.inv.edit) S.inv.edit.prevNext = Math.max(parseInt(S.inv.edit.prevNext) || 0, P.nextNo); else S.nextNo = P.nextNo;
}
function confirmHtml(title, html, okText) {
  return new Promise(res => {
    const dlg = $('#dlg'); $('#dlg_t').textContent = title; $('#dlg_b').innerHTML = html;
    const ok = $('#dlg_ok'), old = ok.textContent, x = $('#dlg_x'), oldX = x.textContent;
    ok.textContent = okText || 'Добре'; ok.classList.toggle('hide', !okText); if (!okText) x.textContent = 'Затвори';
    const restore = () => { ok.textContent = old; ok.classList.remove('hide'); x.textContent = oldX; };
    const done = v => { restore(); dlg.close(); res(v); };
    ok.onclick = ev => { ev.preventDefault(); done(true); };
    $('#dlg_x').onclick = ev => { ev.preventDefault(); done(false); };
    dlg.oncancel = () => { restore(); res(false); };
    dlg.showModal();
  });
}
function xlPreviewHtml(P, name) {
  const sample = (a, f = x => x) => a.length ? `<div class="mut" style="margin:0 0 6px 0">${a.slice(0, 4).map(x => esc(f(x))).join('; ')}${a.length > 4 ? ' …' : ''}</div>` : '';
  const row = (lbl, n, extra = '') => `<tr><td>${lbl}</td><td class="n"><b>${n}</b></td></tr>${extra ? `<tr><td colspan="2">${extra}</td></tr>` : ''}`;
  const rng = P.invoices.length ? `№ ${Math.min(...P.invoices.map(i => +i.no || Infinity))} – ${Math.max(...P.invoices.map(i => +i.no || 0))}` : '';
  return `<div class="mut" style="margin-bottom:8px">Файл: <b>${esc(name)}</b></div>
  <table class="t"><tbody>
   ${row('Нови клиенти', P.clients.length, sample(P.clients, c => c.name))}
   ${row('Нови доставчици', P.suppliers.length, sample(P.suppliers, c => c.name))}
   ${row('Нови артикули', P.articles.length, sample(P.articles))}
   ${row('Нови издадени фактури', P.invoices.length, rng ? `<div class="mut">${rng}</div>` + sample(P.invoices, i => `${i.no} ${i.client}`) : '')}
   ${P.fills.length ? row('Допълнени празни полета в съществуващи клиенти/доставчици', P.fills.length) : ''}
   ${P.units.length + P.payments.length + P.compilers.length ? row('Нови мерни единици / плащания / съставители', P.units.length + P.payments.length + P.compilers.length) : ''}
   ${P.dupInv ? `<tr><td>Фактури, които вече има (пропуснати)</td><td class="n">${P.dupInv}</td></tr>` : ''}
   ${P.nextChanges ? row('Следващ номер на фактура', `${esc(String(S.nextNo))} → ${P.nextNo}`) : ''}
  </tbody></table>
  <div class="mut" style="margin-top:8px">Съществуващите записи не се променят (освен празни полета). Преди импорта се прави автоматичен архив.</div>`;
}
$('#xl_file').onchange = async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  const msg = (m, ok) => { $('#xl_msg').innerHTML = m ? `<div class="${ok ? 'okmsg' : 'err'}">${m}</div>` : ''; };
  msg('');
  try {
    const x = xlParse(await xlRead(f)), P = xlPlan(x);
    if (!P.total) { await confirmHtml('Импорт от Excel', `<p>В „${esc(f.name)}“ няма нищо ново спрямо данните в приложението.</p><div class="mut">Прочетени: ${x.clients.length} клиенти, ${x.suppliers.length} доставчици, ${x.articles.length} артикула, ${x.invoices.length} фактури.</div>`, ''); return; }
    if (!await confirmHtml('Импорт от Excel — преглед', xlPreviewHtml(P, f.name), 'Импортирай')) return;
    await backupNow();
    xlApply(P); saveNow(); fillDatalists();
    msg(`Импортирано: ${P.clients.length} клиенти, ${P.suppliers.length} доставчици, ${P.articles.length} артикула, ${P.invoices.length} издадени фактури${P.nextChanges ? `. Следващ номер на фактура: ${P.nextNo}` : ''}.`, true);
    renderSettings();
  } catch (err) { msg('Импортът не успя: ' + esc(err.message)); }
};

/* ---------------- portable edition (flash drive) + sync back ---------------- */
const mainNextNo = () => S.inv.edit ? S.inv.edit.prevNext : S.nextNo;
function buildPortableState() {   // always carries the client / supplier / article lists (not the issued documents)
  const P = JSON.parse(JSON.stringify(S));
  P.invoices = []; P.receipts = []; P.nextRcpt = 1; P.nextNo = mainNextNo();
  P.inv = { client: '', date: today(), tax: today(), taxManual: false, pay: S.payments[0] || '', comp: S.compilers[0] || '', rate: S.rate, orig: 'ОРИГИНАЛ', lines: blankLines(HALF) };
  P.stock = { sup: '', date: today(), comp: S.compilers[0] || '', lines: sBlank(8) };
  P.portable = { id: (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2)), created: today(), baseNextNo: P.nextNo, withLists: true };
  return P;
}
function applyPortableUi() {
  $('#pbadge').classList.toggle('hide', !S.portable);
  $('#pt_card').classList.toggle('hide', !!S.portable || !HOST);
  $('#sy_card').classList.toggle('hide', !!S.portable);
  $('#flashGrp').classList.toggle('hide', !!S.portable);
  $('#syncBtn').classList.toggle('hide', !!S.portable);
  $('#prepBtn').classList.toggle('hide', !!S.portable || !HOST);
  $('#xl_card').classList.toggle('hide', !!S.portable);
  $('#pt_info').classList.toggle('hide', !S.portable);
  if (S.portable) $('#pt_infotxt').innerHTML = `<p>Това е <b>преносима версия</b>, създадена на ${dmy(S.portable.created)}. Данните са във флашката (папка <b>invoice-data</b> до програмата).</p>
   <p>Номерата на фактурите започват от <b>№ ${esc(String(S.portable.baseNextNo))}</b>. Когато се върнете в офиса, включете флашката в основния компютър и от <b>Настройки → Синхронизация от преносима версия</b> изберете файла <b>invoice-data\\data.json</b>.</p>
   <p class="mut">Не издавайте фактури с номера от основното приложение, докато сте на път — така няма да има повтарящи се номера.</p>`;
}
async function prepareFlash() {
  const out = (m, ok) => { $('#pt_msg').innerHTML = m ? `<div class="${ok ? 'okmsg' : 'err'}">${m}</div>` : ''; };
  const tell = async (m, ok) => { out(m, ok); if (!onSettingsTab()) await confirmHtml('Преносима версия', `<p>${m}</p>`, ''); };
  out('');
  try {
    const drives = await hostCall('listDrives');
    if (!drives.length) { await tell('Не намерих флашка. Включете я и опитайте отново.'); return; }
    const labels = drives.map(d => `${d.root} — ${d.label || 'без име'} · ${d.freeGb} GB свободни от ${d.totalGb} GB${d.removable ? '' : ' (външен диск)'}`.replace(/\s+/g, ' '));
    const o = await ask('Подготви флашка', [{ k: 'd', label: 'Избери флашката (подготвя се целият диск — не може да се избира папка)', type: 'select', opts: labels }], { d: labels[0] });
    if (!o) return;
    saveNow();
    const dr = drives[labels.indexOf(o.d)], st = buildPortableState(), drive = dr.root;
    const r = await hostCall('makePortable', { state: JSON.stringify(st), drive });
    if (!r) return;                                   // cancelled in the overwrite question
    S.portables = (S.portables || []).filter(x => x.path !== r.path).concat([{ id: st.portable.id, created: st.portable.created, path: r.path }]); saveNow();
    await tell(`Готово: <b>${esc(r.path)}</b><br>${r.shortcut ? `Пряка връзка „invoice“: <b>${esc(r.shortcut)}</b><br>` : 'Не успях да създам пряка връзка в началото на флашката — стартирайте invoice-lite.exe от папката.<br>'}` +
      `Копирани са програмата за стари и нови компютри (invoice-lite) и ${st.clients.length} клиенти, ${st.suppliers.length} доставчици, ${st.articles.length} артикула. Номерата на фактурите започват от № ${esc(String(mainNextNo()))}. Не издавайте фактури от основното приложение, докато флашката е на път — иначе номерата ще се повторят.`, true);
  } catch (e) { await tell('Неуспешно: ' + esc(e.message)); }
}
$('#pt_make').onclick = prepareFlash;
$('#prepBtn').onclick = prepareFlash;
// ---- sync: merge what was issued on the flash drive into this app ----
function syncPlan(j) {
  if (!j || !Array.isArray(j.invoices) || !Array.isArray(j.clients)) throw new Error('Това не е файл с данни от приложението (data.json).');
  const P = { clients: [], suppliers: [], articles: [], invoices: [], receipts: [], fills: [], units: [], payments: [], compilers: [], dupInv: 0, dupRcp: 0, conflicts: [] };
  const party = (src, list, out) => {
    const seen = new Set(list.map(c => c.name));
    (src || []).forEach(c => {
      const ex = list.find(e => e.name === c.name);
      if (ex) ['addr', 'vat', 'eik', 'mol'].forEach(k => { if (c[k] && !ex[k]) P.fills.push([ex, k, c[k]]); });
      else if (!seen.has(c.name)) { seen.add(c.name); out.push(c); }
    });
  };
  party(j.clients, S.clients, P.clients); party(j.suppliers, S.suppliers, P.suppliers);
  const art = new Set(S.articles.map(a => a.name)); (j.articles || []).forEach(a => { if (a && a.name && !art.has(a.name)) { art.add(a.name); P.articles.push(a); } });
  const byNo = new Map(S.invoices.map(i => [String(i.no), i])), sig = i => `${i.date}|${i.client}|${Number(i.total).toFixed(2)}`;
  j.invoices.forEach(i => {
    const ex = byNo.get(String(i.no));
    if (!ex) { byNo.set(String(i.no), i); P.invoices.push(i); }
    else if (sig(ex) === sig(i)) P.dupInv++; else P.conflicts.push(i);
  });
  const rkey = r => `${r.date}|${r.sup}|${Number(r.total).toFixed(2)}|${(r.lines || []).join(',')}`, rk = new Set(S.receipts.map(rkey));
  (j.receipts || []).forEach(r => { if (rk.has(rkey(r))) P.dupRcp++; else { rk.add(rkey(r)); P.receipts.push(r); } });
  const add = (src, have, out) => { const hs = new Set(have); (src || []).forEach(v => { if (!hs.has(v)) { hs.add(v); out.push(v); } }); };
  add(j.units, S.units, P.units); add(j.payments, S.payments, P.payments); add(j.compilers, S.compilers, P.compilers);
  const nos = [...S.invoices, ...P.invoices].map(i => +i.no).filter(n => isFinite(n));
  P.nextNo = Math.max(parseInt(mainNextNo()) || 0, nos.length ? Math.max(...nos) + 1 : 0);
  P.nextChanges = P.nextNo !== (parseInt(mainNextNo()) || 0);
  P.total = P.clients.length + P.suppliers.length + P.articles.length + P.invoices.length + P.receipts.length + P.units.length + P.payments.length + P.compilers.length + P.fills.length;
  return P;
}
function syncApply(P) {
  P.fills.forEach(([ex, k, v]) => { ex[k] = v; });
  S.clients.push(...P.clients); S.suppliers.push(...P.suppliers); S.articles.push(...P.articles.map(a => ({ ...a, added: a.added || today() })));
  S.invoices.push(...[...P.invoices].sort((a, b) => (a.date || '').localeCompare(b.date || '') || (+a.no || 0) - (+b.no || 0)));
  S.receipts.push(...[...P.receipts].sort((a, b) => (a.date || '').localeCompare(b.date || '')));
  S.units.push(...P.units); S.payments.push(...P.payments); S.compilers.push(...P.compilers);
  if (S.inv.edit) S.inv.edit.prevNext = P.nextNo; else S.nextNo = P.nextNo;
}
function syncPreviewHtml(P, j) {
  const names = (a, f) => a.length ? `<div class="mut" style="margin:0 0 6px">${a.slice(0, 4).map(x => esc(f(x))).join('; ')}${a.length > 4 ? ' …' : ''}</div>` : '';
  const row = (l, n, ex = '') => `<tr><td>${l}</td><td class="n"><b>${n}</b></td></tr>${ex ? `<tr><td colspan="2">${ex}</td></tr>` : ''}`;
  const nos = P.invoices.map(i => +i.no).filter(n => isFinite(n));
  return `<div class="mut" style="margin-bottom:8px">Преносима версия от ${j.portable ? dmy(j.portable.created) : 'неизвестна дата'}</div>
  <table class="t"><tbody>
   ${row('Нови издадени фактури', P.invoices.length, nos.length ? `<div class="mut">№ ${Math.min(...nos)} – ${Math.max(...nos)}</div>` + names(P.invoices, i => `${i.no} ${i.client} · ${eur(i.total)}`) : '')}
   ${row('Нови складови разписки', P.receipts.length, names(P.receipts, r => `${dmy(r.date)} ${r.sup} · ${eur(r.total)}`))}
   ${row('Нови клиенти', P.clients.length, names(P.clients, c => c.name))}
   ${row('Нови доставчици', P.suppliers.length, names(P.suppliers, c => c.name))}
   ${row('Нови артикули', P.articles.length, names(P.articles, a => a.name))}
   ${P.fills.length ? row('Допълнени празни полета в съществуващи клиенти/доставчици', P.fills.length) : ''}
   ${P.dupInv + P.dupRcp ? `<tr><td>Вече синхронизирани (пропуснати)</td><td class="n">${P.dupInv + P.dupRcp}</td></tr>` : ''}
   ${P.nextChanges ? row('Следващ номер на фактура', `${esc(String(mainNextNo()))} → ${P.nextNo}`) : ''}
  </tbody></table>
  ${P.conflicts.length ? `<div class="err"><b>Конфликт на номера (${P.conflicts.length}) — няма да бъдат добавени:</b><br>${P.conflicts.slice(0, 6).map(i => `№ ${esc(String(i.no))} — във флашката: ${esc(i.client)} (${eur(i.total)}), в приложението е друга фактура`).join('<br>')}<br><span class="mut">Тези номера са издадени и на двете места. Проверете ги ръчно в „Издадени фактури“.</span></div>` : ''}
  <div class="mut" style="margin-top:8px">Преди добавянето се прави автоматичен архив.</div>`;
}
const syMsg = (m, ok) => { $('#sy_msg').innerHTML = m ? `<div class="${ok ? 'okmsg' : 'err'}">${m}</div>` : ''; };
async function runSync(j, label) {
  const P = syncPlan(j);
  if (!P.total && !P.conflicts.length) { await confirmHtml('Синхронизация', `<p>В „${esc(label)}“ няма нищо ново — всичко вече е в приложението.</p>`, ''); return; }
  if (!P.total) { await confirmHtml('Синхронизация', syncPreviewHtml(P, j), ''); return; }
  if (!await confirmHtml('Синхронизация — преглед', syncPreviewHtml(P, j), 'Добави')) return;
  await backupNow(); syncApply(P); saveNow(); fillDatalists();
  await syNotify(`Добавени: ${P.invoices.length} фактури, ${P.receipts.length} разписки, ${P.clients.length} клиенти, ${P.suppliers.length} доставчици, ${P.articles.length} артикула${P.nextChanges ? `. Следващ номер на фактура: ${P.nextNo}` : ''}.${HOST ? ' Подгответе флашката наново от „Преносима версия“, за да започне от новия номер.' : ''}`, true);
  renderSettings();
}
$('#sy_file').onchange = async e => {   // manual choice of the file
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  syMsg('');
  try { await runSync(JSON.parse(await f.text()), f.name); } catch (err) { syMsg('Синхронизацията не успя: ' + esc(err.message)); }
};
async function pickPortable(list) {
  if (list.length === 1) return list[0];
  const labels = list.map((c, i) => `${c.drive} ${c.folder.replace(/^.:\\/, '')} · ${c.docs} документа · ${c.modified}${c.match ? ' · ваша' : ''}`.replace(/\s+/g, ' '));
  const o = await ask('Намерени са няколко преносими версии', [{ k: 'p', label: 'Изберете откъде да синхронизирам', type: 'select', opts: labels }], { p: labels[0] });
  return o ? list[labels.indexOf(o.p)] : null;
}
function onSettingsTab() { return (document.querySelector('.tab:not(.hide)') || {}).id === 'tab-set'; }
async function syNotify(m, ok) {          // result line in Settings; a small dialog when the user is on another tab
  syMsg(m, ok);
  if (!onSettingsTab()) await confirmHtml('Синхронизация', `<p>${m}</p>`, '');
}
async function syncFromDrive() {          // look for the flash drive that was prepared earlier; fall back to a manual choice
  syMsg('');
  try {
    if (!HOST) { $('#sy_file').click(); return; }
    const known = S.portables || [];
    const found = await hostCall('findPortable', { paths: known.map(k => k.path), ids: known.map(k => k.id) });
    if (!found.length) {
      const m = 'Не намерих преносима версия автоматично (флашката не е включена или папката е преименувана). Изберете файла ръчно.';
      if (onSettingsTab()) { syMsg(m); $('#sy_file').click(); }
      else if (await confirmHtml('Синхронизация', `<p>${m}</p>`, 'Избери файла ръчно…')) $('#sy_file').click();
      return;
    }
    const pick = await pickPortable(found); if (!pick) return;
    await runSync(JSON.parse(await hostCall('readFile', pick.path)), pick.path);
    await liteRefreshOffer(pick.folder);
  } catch (err) { await syNotify('Синхронизацията не успя: ' + esc(err.message)); }
}
$('#sy_auto').onclick = syncFromDrive;
$('#syncBtn').onclick = syncFromDrive;

/* ---------------- about / version ---------------- */
const LITE_VERSION = '1.4.0';   // keep in sync with lite\Faktura-Lite.csproj
const CHANGELOG = [
  ['2.10.0', '14.10.2026', ['Печатни документи: логото на фирмата е най-отгоре в средата на фактурата; в Настройки може да се качи собствено лого (PNG, JPG, SVG…), копира се и в преносимата версия', 'Фактурата е с 7 реда на половин лист, по-големи полета горе и долу, второто копие е по-близо до пунктираната линия', 'По-удобен печат за лазерни принтери: черен текст и линии, по-едър дребен шрифт, по-плътни рамки, сив фон на заглавния ред и на общата сума', 'Lite 1.4.0: печатът е с точно същия вид като в основното приложение']],
  ['2.9.0', '13.10.2026', ['Lite 1.3.0: печатът е същият като в пълната версия (без смаляване), точно подравнен спрямо хартията — отместването на принтера се чете от самия принтер', 'Lite: собствен преглед, който показва листа точно както ще се отпечата', 'Lite: бутон ✕ във всеки ред на артикулите за изтриване на реда']],
  ['2.8.2', '13.10.2026', ['Lite 1.2.1: поправена грешка при щракване върху полето за клиент (Object reference not set to an instance of an object)']],
  ['2.8.1', '13.10.2026', ['Премахнато е форматирането на флашката — подготовката само копира файловете и прави пряка връзка (автоматичното откриване, синхронизиране и обновяване остават)', 'Lite: при щракване върху полето за клиент или артикул се показва списък с всички, а при писане се търси по част от името (и по ЕИК/адрес за клиенти)']],
  ['2.8.0', '13.10.2026', ['Цялото меню е в странична лента, с икони за всички раздели']],
  ['2.7.1', '13.10.2026', ['Бутоните за флашката са в странична лента — по-големи, с надписи']],
  ['2.7.0', '13.10.2026', ['Подготовка на флашка с форматиране и име INVOICE (само за тази програма)', 'Автоматично откриване на флашката при включване: предлага синхронизиране, ако има нови документи, и обновяване на Lite, ако е по-стар; опция за изключване в Настройки']],
  ['2.6.1', '13.10.2026', ['Проверка за обновления при всяко стартиране (може да се изключи в Настройки); при нова версия се показва предложение за обновяване']],
  ['2.6.0', '13.10.2026', ['Обновяване през интернет от GitHub (pachobg2/invoice): проверка, изтегляне с контролна сума и автоматичен рестарт; данните не се променят', 'Синхронизацията предлага обновяване на Lite във флашката', 'Програмата вече не носи вградени клиенти, артикули и фактури — стартира празна, данните са само в invoice-data', 'Lite: печатът е с по-големи полета (страницата е смалена до 93% и центрирана)']],
  ['2.5.4', '13.10.2026', ['Lite: печатът е с по-големи полета (страницата е смалена до 93% и центрирана) — повече разстояние от ляв и горен ръб']],
  ['2.5.3', '13.10.2026', ['Lite: поправено изместването на печата наляво — част от фактурата се отрязваше при принтери с непечатаема зона']],
  ['2.5.2', '11.10.2026', ['Всички файлове и папки с латински имена: invoice.exe, invoice-data, invoice-portable (старите имена се преименуват/откриват автоматично)']],
  ['2.5.1', '11.10.2026', ['Списъците за избор (клиент, доставчик, артикул) показват всички съвпадения с пълните имена']],
  ['2.5.0', '11.10.2026', ['Артикули, клиенти и доставчици: подреждане (по име, дата, използване), дата на добавяне, последно използване и брой']],
  ['2.4.0', '10.10.2026', ['Дневник на продажбите (печат, PDF, CSV за Excel)', 'Проверка на ЕИК/ЕГН и номер по ЗДДС при въвеждане', 'Раздел „За програмата“ с версия и информация за поддръжка']],
  ['2.3.0', '10.10.2026', ['Преносима версия само с Lite (Windows XP–11), пряка връзка във флашката', 'Подготовка на флашка с бутон в менюто — само цял диск', 'Изтриване на фактура: последната връща номера, по-стара — предупреждение']],
  ['2.2.0', '10.10.2026', ['Импорт от Excel и синхронизация от преносима версия (автоматично търсене на флашката)', 'Складова разписка като фактурата, една страница, без номер']],
  ['2.0.0', '10.10.2026', ['Самостоятелно Windows приложение, всички данни в папка invoice-data', 'Фактури с много редове, преглед, PDF, регистър, справка ДДС, тъмен режим']]
];
async function renderAbout() {
  let host = null; if (HOST) { try { host = await hostCall('about'); } catch (e) { } }
  const lines = [
    `Фактури ${host ? host.version : '(уеб)'}${host ? ' · WebView2 ' + host.webview2 + ' · ' + host.dotnet : ''}`,
    `Преносима версия (Lite): ${LITE_VERSION}`,
    host ? `Програма: ${host.exe}` : '',
    host ? `Данни: ${host.dataDir}` : '',
    `Документи: ${S.invoices.length} фактури, ${(S.receipts || []).length} разписки · ${S.clients.length} клиенти, ${S.suppliers.length} доставчици, ${S.articles.length} артикула`,
    `Следващ номер на фактура: ${S.nextNo} · Подготвени флашки: ${(S.portables || []).length}`
  ].filter(Boolean);
  window.__aboutText = lines.join('\n');
  $('#ab_txt').innerHTML = lines.map(esc).join('<br>');
  $('#ab_log').innerHTML = CHANGELOG.map(([v, d, items]) => `<div style="margin-top:6px"><b>${v}</b> <span class="mut">${d}</span><ul style="margin:2px 0 0 18px;padding:0">${items.map(i => `<li>${esc(i)}</li>`).join('')}</ul></div>`).join('');
}
$('#ab_copy').onclick = async () => { try { await navigator.clipboard.writeText(window.__aboutText || ''); $('#ab_msg').innerHTML = '<div class="okmsg">Информацията е копирана — поставете я в съобщение до поддръжката.</div>'; } catch (e) { $('#ab_msg').innerHTML = '<div class="err">Не успях да копирам.</div>'; } };

/* ---------------- keyboard shortcuts ---------------- */
document.addEventListener('keydown', e => {
  if (!(e.ctrlKey || e.metaKey) || e.altKey || $('#dlg').open) return;
  const tab = (document.querySelector('.tab:not(.hide)') || {}).id, k = e.key.toLowerCase();
  const inv = tab === 'tab-inv' || tab === 'tab-prev' || tab === 'tab-reg', stk = tab === 'tab-stock' || tab === 'tab-sprev' || tab === 'tab-sreg';
  if (k === 'p' && !e.shiftKey) { e.preventDefault(); if (inv) $('#i_save').onclick(); else if (stk) $('#s_save').onclick(); }
  else if (k === 's') { e.preventDefault(); if (inv) $('#i_saveonly').onclick(); else if (stk) $('#s_saveonly').onclick(); }
  else if (k === 'n') { e.preventDefault(); if (stk) newReceipt(); else newInvoice(); }
}, true);

/* ---------------- updates (GitHub releases of pachobg2/invoice) ---------------- */
const UPD_KEY = 'invoice_update';
function updPref() { try { return JSON.parse(localStorage.getItem(UPD_KEY) || '{}'); } catch (e) { return {}; } }
function updSet(o) { try { localStorage.setItem(UPD_KEY, JSON.stringify(Object.assign(updPref(), o))); } catch (e) { } }
let updInfo = null, updBusy = false;
function updShow(html) { const b = $('#upd_bar'); b.innerHTML = html || ''; b.classList.toggle('hide', !html); }
function updProgress(p) { const e = $('#upd_pct'); if (e) e.textContent = p + '%'; }
function updBarNew(info) {
  updShow(`<span>Има нова версия <b>${esc(info.latest)}</b> (вие имате ${esc(info.current)}).</span><button class="b" id="upd_go">Обнови сега</button><button class="b g" id="upd_info">Какво е ново</button><button class="b g" id="upd_x" title="По-късно">✕</button>`);
  $('#upd_go').onclick = updInstall; $('#upd_info').onclick = () => updDetails(false); $('#upd_x').onclick = () => updShow('');
}
async function updDetails(ask) {
  const i = updInfo; if (!i) return false;
  const html = `<p>Нова версия <b>${esc(i.latest)}</b> (вие имате ${esc(i.current)}), ${i.sizeMb} MB.</p><div style="white-space:pre-wrap;max-height:260px;overflow:auto" class="mut">${esc(i.notes || '')}</div>${ask ? '<p>Данните ви не се променят. Програмата ще се затвори и ще се отвори отново.</p>' : ''}`;
  return confirmHtml('Обновление', html, ask ? 'Обнови' : '');
}
async function updCheck(manual) {
  if (!HOST || S.portable || updBusy) return;
  const msg = t => { if (manual) $('#up_msg').innerHTML = t; };
  msg('<div class="mut">Проверявам…</div>');
  try {
    const info = await hostCall('updateCheck'); updSet({ last: Date.now() });
    if (info.newer && info.ready) { updInfo = info; updBarNew(info); if (!manual) { await updInstall(); return; } msg(`<div class="okmsg">Има нова версия <b>${esc(info.latest)}</b> — използвайте бутона „Обнови сега“ най-горе.</div>`); }
    else if (info.newer) msg('<div class="err">Има по-нова версия, но още не е напълно публикувана. Опитайте по-късно.</div>');
    else msg(`<div class="okmsg">Имате най-новата версия (${esc(info.current)}).</div>`);
  } catch (e) { msg('<div class="err">Не успях да проверя за обновления (няма интернет или GitHub не отговаря).</div>'); }
}
async function updInstall() {          // asks first (release notes + Обнови); the bar stays if the user says no
  if (updBusy || !updInfo) return;
  if (!await updDetails(true)) return;
  updBusy = true; saveNow(); await backupNow();
  updShow(`<span>Изтегляне на версия ${esc(updInfo.latest)}… <b id="upd_pct">0%</b> — не затваряйте програмата.</span>`);
  try { await hostCall('updateInstall'); updShow('<span>Рестартиране…</span>'); }
  catch (e) { updBusy = false; updShow(`<span class="err" style="margin:0">Обновяването не успя: ${esc(e.message)}. Старата версия остава.</span><button class="b g" id="upd_x">✕</button>`); $('#upd_x').onclick = () => updShow(''); }
}
async function liteRefreshOffer(folder) {   // after a sync: bring the Lite exe on the flash drive up to the version built into this program
  try {
    const s = await hostCall('liteRefresh', { folder, apply: false });
    if (!s.exists || !s.newer) return;
    if (!await confirmHtml('Преносима версия', `<p>На флашката е Lite ${esc(s.have)}, а тази програма съдържа по-нова — ${esc(s.embedded)}.</p><p>Да обновя ли „invoice-lite.exe“ във флашката? Данните там не се променят.</p>`, 'Обнови')) return;
    const r = await hostCall('liteRefresh', { folder, apply: true });
    await confirmHtml('Преносима версия', r.applied ? `<p>Lite на флашката е обновен до ${esc(r.embedded)}.</p>` : '<p>Не успях да обновя файла.</p>', '');
  } catch (e) { await confirmHtml('Преносима версия', `<p>Lite на флашката не беше обновен: ${esc(e.message)}</p><p>Затворете го, ако е отворен на друг компютър, и опитайте пак.</p>`, ''); }
}
$('#up_check').onclick = () => updCheck(true);
$('#up_auto').checked = updPref().auto !== false;
$('#up_auto').onchange = e => updSet({ auto: e.target.checked });
function updAutoCheck() {
  const p = updPref();
  if (!HOST || S.portable || p.auto === false) return;
  setTimeout(() => updCheck(false), 4000);
}

/* ---------------- flash drive: automatic detection when it is plugged in ---------------- */
const drvPref = () => updPref().drive !== false;
let drvBusy = false;
const drvSeen = new Set();          // drive states already handled (no repeated questions for an unchanged drive at start-up)
function drvNote(html) { updShow(`<span>${html}</span><button class="b g" id="upd_x">✕</button>`); $('#upd_x').onclick = () => updShow(''); setTimeout(() => { if ($('#upd_bar').innerText.includes('Флашка')) updShow(''); }, 9000); }
async function driveCheck(trigger) {
  if (!HOST || S.portable || !drvPref() || drvBusy || updBusy) return;
  drvBusy = true;
  try {
    for (let i = 0; i < 30 && $('#dlg').open; i++) await new Promise(r => setTimeout(r, 2000));   // another question is on screen
    const known = S.portables || [];
    const found = (await hostCall('findPortable', { paths: known.map(k => k.path), ids: known.map(k => k.id) })).filter(c => c.id);
    if (!found.length) return;
    const pick = found[0], name = pick.drive.replace(/\\$/, '');
    const j = JSON.parse(await hostCall('readFile', pick.path));
    const P = syncPlan(j), lite = await hostCall('liteRefresh', { folder: pick.folder, apply: false });
    const key = pick.path + '|' + pick.modified + '|' + P.total + '|' + (lite.have || '');
    if (trigger === 'start' && drvSeen.has(key)) return;
    drvSeen.add(key);
    if (P.total || P.conflicts.length) await runSync(j, pick.path);
    if (lite.exists && lite.newer) await liteRefreshOffer(pick.folder);
    else if (!P.total && !P.conflicts.length) drvNote(`Флашка ${esc(name)}: няма нови документи за синхронизиране${lite.exists ? `, Lite е актуален (${esc((lite.have || '').replace(/\.0$/, ''))})` : ''}.`);
  } catch (e) { /* a drive that cannot be read is not worth a message */ }
  finally { drvBusy = false; }
}
function drvSettings() {
  const c = $('#dv_auto'); if (!c) return;
  c.checked = drvPref(); c.onchange = () => updSet({ drive: c.checked });
}

/* ---------------- company logo: uploaded in Настройки, printed at the top of invoices and receipts ---------------- */
const logoTag = () => S.logo && S.logoImage ? `<img class="lg" src="${S.logoImage}" alt="">` : '';
const loadImg = src => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Картинката не може да се отвори.')); i.src = src; });
async function setLogo(file) {
  if (!/^image\//.test(file.type)) throw new Error('Изберете картинка (PNG, JPG, GIF, WebP или SVG).');
  const url = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(new Error('Файлът не може да се прочете.')); r.readAsDataURL(file); });
  const img = await loadImg(url);
  let w = img.naturalWidth, h = img.naturalHeight; if (!w || !h) { w = 400; h = 150; }      // an SVG without a size
  const k = Math.min(1, 900 / w, 300 / h), cw = Math.max(1, Math.round(w * k)), ch = Math.max(1, Math.round(h * k));
  const c = document.createElement('canvas'); c.width = cw; c.height = ch;
  const x = c.getContext('2d'); x.drawImage(img, 0, 0, cw, ch);
  let out = c.toDataURL('image/png');
  if (out.length > 450000) { x.globalCompositeOperation = 'destination-over'; x.fillStyle = '#fff'; x.fillRect(0, 0, cw, ch); out = c.toDataURL('image/jpeg', 0.88); }   // a photo: smaller as JPEG
  S.logoImage = out; S.logo = true; saveNow(); renderLogoCard();
}
async function migrateLogo() {                 // data from before the logo became uploadable: keep the logo that used to be built in
  if (S.logoImage !== undefined || S.portable) return;
  S.logoImage = '';
  if (!((S.company && S.company.name) || (S.invoices && S.invoices.length))) return;
  try {
    const [m, t] = await Promise.all([loadImg('logo-mark.png'), loadImg('logo-text.png')]);
    const H = 320, wm = m.width * H / m.height, th = H * 0.55, wt = t.width * th / t.height, gap = H * 0.18;
    const c = document.createElement('canvas'); c.width = Math.round(wm + gap + wt); c.height = H;
    const x = c.getContext('2d'); x.drawImage(m, 0, 0, wm, H); x.drawImage(t, wm + gap, (H - th) / 2, wt, th);
    S.logoImage = c.toDataURL('image/png'); S.logo = true; saveNow();
  } catch (e) { S.logoImage = ''; }
}
function renderLogoCard() {
  const has = !!S.logoImage;
  $('#c_logo').checked = !!S.logo;
  $('#lg_prev').src = has ? S.logoImage : '';
  $('#lg_prevwrap').classList.toggle('hide', !has); $('#lg_none').classList.toggle('hide', has); $('#lg_del').classList.toggle('hide', !has);
}
$('#lg_file').onchange = async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try { $('#lg_msg').innerHTML = ''; await setLogo(f); $('#lg_msg').innerHTML = '<div class="okmsg">Логото е запазено. Вижте го в „Преглед и печат“.</div>'; }
  catch (err) { $('#lg_msg').innerHTML = `<div class="err">${esc(err.message)}</div>`; }
};
$('#lg_del').onclick = () => { if (confirm('Да премахна ли логото?')) { S.logoImage = ''; saveNow(); renderLogoCard(); } };

/* ---------------- boot ---------------- */
(async () => {
  await load();
  await migrateLogo();
  S.rate = S.rate ?? 20;
  ensureStock();
  applyPortableUi();
  if (S.theme) applyTheme(S.theme); else S.theme = document.documentElement.dataset.theme;
  initLists(); fillDatalists();
  let t0 = 'inv'; try { t0 = sessionStorage.getItem('tab') || 'inv'; } catch (e) { }
  showTab(t0);
  window.__app = { slov, intWords, get S() { return S; }, HOST: !!HOST };
  updAutoCheck(); drvSettings(); if (HOST && !S.portable) setTimeout(() => driveCheck('start'), 7000);
  document.title = 'Фактури';
})();
