/* Halden: in-browser mini copy of the manager home page (fictional data, localStorage only) + screenshot lightbox. */
(function () {
  var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var en = document.documentElement.lang === 'en';
  /* ---- lightbox ---- */
  var dlg = document.getElementById('hx-lb');
  if (dlg && dlg.showModal) {
    var links = [].slice.call(document.querySelectorAll('.gal a.zoom')), cur = 0;
    var im = dlg.querySelector('img'), cap = dlg.querySelector('.hx-lbcap');
    function show(i) {
      cur = (i + links.length) % links.length;
      var a = links[cur], s = a.querySelector('img');
      im.src = a.getAttribute('href'); im.alt = s.alt; im.width = s.width; im.height = s.height;
      var fc = a.closest('figure').querySelector('figcaption'); cap.textContent = fc ? fc.textContent : '';
    }
    links.forEach(function (a, i) { a.addEventListener('click', function (e) { e.preventDefault(); show(i); dlg.showModal(); }); });
    dlg.querySelector('.hx-x').addEventListener('click', function () { dlg.close(); });
    dlg.querySelector('.hx-p').addEventListener('click', function () { show(cur - 1); });
    dlg.querySelector('.hx-n').addEventListener('click', function () { show(cur + 1); });
    dlg.addEventListener('click', function (e) { if (e.target === dlg) dlg.close(); });
    dlg.addEventListener('keydown', function (e) { if (e.key === 'ArrowLeft') show(cur - 1); if (e.key === 'ArrowRight') show(cur + 1); });
  }
  /* ---- demo ---- */
  var root = document.getElementById('hxdemo');
  if (!root) return;
  var T = en ? {
    hi: 'Good day, Demo Manager', wait: function (n) { return n ? ' — you have ' + n + ' item' + (n > 1 ? 's' : '') + ' waiting for a decision' : ' — nothing is waiting for you'; },
    tabs: ['Overview', 'My approvals', 'Projects', 'Documents'], act: 'Action required', pend: 'pending approvals', health: 'Portfolio health', crit: 'critical', ok: 'on track',
    docs: 'Key documents', files: 'files', need: 'Needs a decision', none: 'All decisions are made. Reset the demo to start again.', ap: 'Approve', rj: 'Reject', undo: 'Undo',
    did: { a: 'Approved: ', r: 'Rejected: ' }, undone: 'Decision undone: ', hist: 'Decided', g: 'On track', y: 'At risk', r: 'Critical', prog: 'progress',
    roles: { all: 'Team lead · all departments', HR: 'HR', Finance: 'Finance', PMO: 'PMO', IT: 'IT' }
  } : {
    hi: 'Добрый день, Демо Руководитель', wait: function (n) { return n ? ' — ждут вашего решения: ' + n : ' — решений не ждёт ничего'; },
    tabs: ['Обзор', 'Мои согласования', 'Проекты', 'Документы'], act: 'Требует действия', pend: 'заявок на согласование', health: 'Состояние портфеля', crit: 'критично', ok: 'в норме',
    docs: 'Ключевые документы', files: 'файлов', need: 'Нужно решение', none: 'Все решения приняты. Сбросьте демо, чтобы начать заново.', ap: 'Approve', rj: 'Reject', undo: 'Отменить',
    did: { a: 'Одобрено: ', r: 'Отклонено: ' }, undone: 'Решение отменено: ', hist: 'Решено', g: 'В норме', y: 'Под риском', r: 'Критично', prog: 'прогресс',
    roles: { all: 'Руководитель · все отделы', HR: 'HR', Finance: 'Финансы', PMO: 'PMO', IT: 'ИТ' }
  };
  var AP = [
    ['a1', 'HR', en ? 'Offer approval — platform lead' : 'Согласование оффера — ведущий инженер платформы', en ? 'Hiring' : 'Найм'],
    ['a2', 'HR', en ? 'Leave request — 10 days, December' : 'Отпуск — 10 дней, декабрь', en ? 'Leave' : 'Отпуск'],
    ['a3', 'Finance', en ? 'Vendor contract renewal — Northwind facilities' : 'Продление договора — Northwind facilities', en ? 'Contract' : 'Договор'],
    ['a4', 'Finance', en ? 'Budget request — Q1 marketing events' : 'Заявка на бюджет — мероприятия Q1', en ? 'Budget' : 'Бюджет'],
    ['a5', 'PMO', en ? 'Stage gate — CRM migration, phase 2' : 'Переход этапа — миграция CRM, фаза 2', en ? 'Stage gate' : 'Этап'],
    ['a6', 'IT', en ? 'Access request — Power BI Pro licence' : 'Запрос доступа — лицензия Power BI Pro', en ? 'Access' : 'Доступ']
  ];
  var PR = [
    ['HR', en ? 'Onboarding portal' : 'Портал адаптации', 'g', 80], ['Finance', en ? 'Expense automation' : 'Автоматизация авансовых', 'y', 55],
    ['PMO', en ? 'CRM migration' : 'Миграция CRM', 'r', 35], ['PMO', en ? 'Office move 2027' : 'Переезд офиса 2027', 'g', 20],
    ['IT', en ? 'Device refresh' : 'Обновление ноутбуков', 'y', 60], ['IT', en ? 'Helpdesk SLA' : 'SLA службы поддержки', 'g', 90]
  ];
  var DOC = [['HR', en ? 'Leave policy 2026.pdf' : 'Политика отпусков 2026.pdf'], ['Finance', en ? 'Budget Q1.xlsx' : 'Бюджет Q1.xlsx'], ['PMO', en ? 'Stage gate checklist.docx' : 'Чек-лист этапа.docx'], ['IT', en ? 'Security baseline.pdf' : 'Базовая безопасность.pdf'], ['Finance', en ? 'Vendor list.xlsx' : 'Поставщики.xlsx']];
  var KEY = 'halden-demo-v1', st;
  try { st = JSON.parse(localStorage.getItem(KEY)); } catch (e) {}
  if (!st || typeof st !== 'object') st = { role: 'all', tab: 0, done: {} };
  function save() { try { localStorage.setItem(KEY, JSON.stringify(st)); } catch (e) {} }
  var $ = function (id) { return document.getElementById(id); };
  var role = $('hx-role'), tabs = $('hx-tabs'), body = $('hx-body'), msg = $('hx-msg'), hi = $('hx-hi');
  var undoId = null;
  function mine(d) { return st.role === 'all' || d === st.role; }
  function el(t, c, x) { var e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; }
  function pending() { return AP.filter(function (a) { return mine(a[1]) && !st.done[a[0]]; }); }
  function say(t, cls, id) {
    msg.textContent = ''; msg.className = 'dd-msg ' + (cls || '');
    msg.appendChild(document.createTextNode(t));
    if (id) { var b = el('button', 'sec', T.undo); b.type = 'button'; b.addEventListener('click', function () { var x = AP.filter(function (a) { return a[0] === id; })[0]; delete st.done[id]; save(); say(T.undone + x[2]); render(); }); msg.appendChild(document.createTextNode(' ')); msg.appendChild(b); }
  }
  function decide(a, v) { st.done[a[0]] = v; save(); say(T.did[v] + a[2], 'ok', a[0]); render(); }
  function bump(n) { if (!reduce) { n.classList.remove('bump'); void n.offsetWidth; n.classList.add('bump'); } }
  function list(items, cap) {
    var wrap = el('div', 'hx-list');
    if (!items.length) { wrap.appendChild(el('p', 'hx-empty', T.none)); return wrap; }
    items.slice(0, cap || 99).forEach(function (a) {
      var r = el('div', 'hx-row'), t = el('div', 'hx-rt');
      t.appendChild(el('b', '', a[2])); t.appendChild(el('span', '', a[3] + ' · ' + (T.roles[a[1]] || a[1])));
      var ap = el('button', '', T.ap), rj = el('button', 'sec', T.rj); ap.type = rj.type = 'button';
      ap.addEventListener('click', function () { decide(a, 'a'); }); rj.addEventListener('click', function () { decide(a, 'r'); });
      r.appendChild(t); r.appendChild(ap); r.appendChild(rj); wrap.appendChild(r);
    });
    return wrap;
  }
  function kpi(lbl, val, sub, cls) { var k = el('div', 'hx-k ' + (cls || '')); k.appendChild(el('p', 'hx-kl', lbl)); var v = el('p', 'hx-kv'); v.innerHTML = val; k.appendChild(v); k.appendChild(el('p', 'hx-ks', sub)); return k; }
  var last = {};
  function render() {
    var p = pending(), pr = PR.filter(function (x) { return mine(x[0]); }), dc = DOC.filter(function (x) { return mine(x[0]); });
    var crit = pr.filter(function (x) { return x[2] === 'r'; }).length;
    hi.textContent = ''; hi.appendChild(document.createTextNode(T.hi)); hi.appendChild(el('em', '', T.wait(p.length)));
    var cnt = [null, p.length, pr.length, dc.length];
    [].forEach.call(tabs.children, function (b, i) {
      b.textContent = T.tabs[i] + (cnt[i] != null ? ' (' + cnt[i] + ')' : '');
      b.setAttribute('aria-selected', i === st.tab ? 'true' : 'false'); b.tabIndex = i === st.tab ? 0 : -1;
    });
    body.textContent = '';
    if (st.tab === 0) {
      var ks = el('div', 'hx-kpis');
      var k1 = kpi(T.act, String(p.length), T.pend, 'a'); ks.appendChild(k1);
      ks.appendChild(kpi(T.health, '<span class="r">' + crit + '</span> <small>' + T.crit + '</small> &nbsp;' + (pr.length - crit) + ' <small>' + T.ok + '</small>', ''));
      ks.appendChild(kpi(T.docs, String(dc.length), T.files));
      body.appendChild(ks);
      if (last.p != null && last.p !== p.length) bump(k1);
      var h = el('p', 'hx-h', T.need); h.appendChild(el('span', '', String(p.length))); body.appendChild(h);
      body.appendChild(list(p, 3));
    } else if (st.tab === 1) {
      body.appendChild(list(p));
      var d = AP.filter(function (a) { return mine(a[1]) && st.done[a[0]]; });
      if (d.length) { body.appendChild(el('p', 'hx-h', T.hist)); var ul = el('ul', 'hx-done'); d.forEach(function (a) { ul.appendChild(el('li', st.done[a[0]] === 'a' ? 'a' : 'r', (st.done[a[0]] === 'a' ? '✓ ' : '✕ ') + a[2])); }); body.appendChild(ul); }
    } else if (st.tab === 2) {
      var g = el('div', 'hx-proj');
      pr.forEach(function (x) { var c = el('div', 'hx-pc ' + x[2]); c.appendChild(el('b', '', x[1])); c.appendChild(el('span', 'hx-st', T[x[2]] + ' · ' + (T.roles[x[0]] || x[0])));
        var bar = el('div', 'hx-bar'); var f = el('i'); f.style.width = x[3] + '%'; bar.appendChild(f); c.appendChild(bar); c.appendChild(el('span', 'hx-ks', T.prog + ' ' + x[3] + '%')); g.appendChild(c); });
      body.appendChild(g);
    } else {
      var u = el('ul', 'hx-docs'); dc.forEach(function (x) { u.appendChild(el('li', '', x[1] + ' · ' + (T.roles[x[0]] || x[0]))); }); body.appendChild(u);
    }
    last.p = p.length;
  }
  Object.keys(T.roles).forEach(function (k) { var o = el('option', '', T.roles[k]); o.value = k; role.appendChild(o); });
  role.value = st.role;
  role.addEventListener('change', function () { st.role = role.value; save(); last = {}; say(''); render(); });
  T.tabs.forEach(function (n, i) { var b = el('button', '', n); b.type = 'button'; b.setAttribute('role', 'tab'); b.setAttribute('aria-controls', 'hx-body'); b.addEventListener('click', function () { st.tab = i; save(); render(); }); tabs.appendChild(b); });
  tabs.addEventListener('keydown', function (e) { var d = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0; if (!d) return; st.tab = (st.tab + d + 4) % 4; save(); render(); tabs.children[st.tab].focus(); });
  $('hx-reset').addEventListener('click', function () { st = { role: 'all', tab: 0, done: {} }; save(); role.value = 'all'; last = {}; say(''); render(); });
  render();
  root.classList.add('live');
})();
