/* Desk Reservation: in-browser demo of the SPFx web part logic (fictional data, localStorage only). */
(function () {
  var root = document.getElementById('deskdemo');
  if (!root) return;
  var en = document.documentElement.lang === 'en';
  var T = en ? {
    all: 'All zones', my: 'My reservations', find: 'Find a desk', pick: 'Select a date first.',
    booked: 'Desk %s is booked for %d.', cancelled: 'Booking for %s on %d is cancelled.', none: 'No free desks for this date and zone. Try other filters.',
    nomine: 'You have no reservations yet.', book: 'Book', cancel: 'Cancel booking', ready: 'Available', repair: 'Needs repair',
    busy: 'Booked by a colleague', mine: 'Your booking', date: 'Date', reset: 'Demo data reset.', plan: 'Office plan: free, booked and your desks for the selected date'
  } : {
    all: 'Все зоны', my: 'Мои бронирования', find: 'Найти место', pick: 'Сначала выберите дату.',
    booked: 'Место %s забронировано на %d.', cancelled: 'Бронь места %s на %d отменена.', none: 'На эту дату и зону свободных мест нет. Попробуйте другие фильтры.',
    nomine: 'У вас пока нет бронирований.', book: 'Забронировать', cancel: 'Отменить бронь', ready: 'Свободно', repair: 'Нужен ремонт',
    busy: 'Занято коллегой', mine: 'Ваша бронь', date: 'Дата', reset: 'Данные демо сброшены.', plan: 'План офиса: свободные, занятые и ваши места на выбранную дату'
  };
  // Same shape as the SharePoint list "Desks": Id, Title, Zone, EquipmentStatus
  var DESKS = [
    { Id: 1, Title: 'OS-01', Zone: 'Open Space', EquipmentStatus: 'Ready', x: 24, y: 40 },
    { Id: 2, Title: 'OS-02', Zone: 'Open Space', EquipmentStatus: 'Ready', x: 84, y: 40 },
    { Id: 3, Title: 'OS-03', Zone: 'Open Space', EquipmentStatus: 'Ready', x: 144, y: 40 },
    { Id: 4, Title: 'OS-04', Zone: 'Open Space', EquipmentStatus: 'Ready', x: 24, y: 96 },
    { Id: 5, Title: 'OS-05', Zone: 'Open Space', EquipmentStatus: 'Repair', x: 84, y: 96 },
    { Id: 6, Title: 'OS-06', Zone: 'Open Space', EquipmentStatus: 'Ready', x: 144, y: 96 },
    { Id: 7, Title: 'QZ-01', Zone: 'Quiet Zone', EquipmentStatus: 'Ready', x: 236, y: 40 },
    { Id: 8, Title: 'QZ-02', Zone: 'Quiet Zone', EquipmentStatus: 'Ready', x: 236, y: 96 },
    { Id: 9, Title: 'MR-01', Zone: 'Meeting Room', EquipmentStatus: 'Ready', x: 312, y: 40 },
    { Id: 10, Title: 'MR-02', Zone: 'Meeting Room', EquipmentStatus: 'Ready', x: 312, y: 96 }
  ];
  var KEY = 'deskDemo.v1';
  var mine = [];
  try { mine = JSON.parse(localStorage.getItem(KEY) || '[]') || []; } catch (e) { mine = []; }
  var save = function () { try { localStorage.setItem(KEY, JSON.stringify(mine)); } catch (e) {} };
  var $ = function (id) { return document.getElementById(id); };
  var dateI = $('dd-date'), zoneS = $('dd-zone'), viewB = $('dd-view'), msg = $('dd-msg'), cards = $('dd-cards'), plan = $('dd-plan'), resetB = $('dd-reset');
  var view = 'all';

  // "colleagues" bookings: stable pseudo-random per date, so the demo looks alive
  function others(date) {
    if (!date) return [];
    var h = 0; for (var i = 0; i < date.length; i++) h = (h * 31 + date.charCodeAt(i)) >>> 0;
    var out = [];
    DESKS.forEach(function (d) { h = (h * 1103515245 + 12345) >>> 0; if (d.EquipmentStatus === 'Ready' && (h >>> 16) % 10 < 3) out.push(d.Id); });
    return out;
  }
  function fmt(date) { var p = date.split('-'); return p[2] + '.' + p[1] + '.' + p[0]; }
  function say(kind, text) { msg.className = 'dd-msg ' + kind; msg.textContent = text; }
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function myOn(date) { return mine.filter(function (r) { return r.date === date; }).map(function (r) { return r.deskId; }); }

  function book(desk) {
    var date = dateI.value;
    if (!date) { say('warn', T.pick); dateI.focus(); return; }
    mine.push({ id: Date.now(), deskId: desk.Id, date: date }); save();
    say('ok', T.booked.replace('%s', desk.Title).replace('%d', fmt(date))); render();
  }
  function cancel(r) {
    var d = DESKS.filter(function (x) { return x.Id === r.deskId; })[0];
    mine = mine.filter(function (x) { return x.id !== r.id; }); save();
    say('ok', T.cancelled.replace('%s', d.Title).replace('%d', fmt(r.date))); render();
  }

  function drawPlan() {
    var NS = 'http://www.w3.org/2000/svg', date = dateI.value, busy = others(date), my = myOn(date), zone = zoneS.value;
    while (plan.firstChild) plan.removeChild(plan.firstChild);
    var add = function (tag, attrs, text) { var e = document.createElementNS(NS, tag); for (var k in attrs) e.setAttribute(k, attrs[k]); if (text) e.textContent = text; plan.appendChild(e); return e; };
    add('rect', { x: 8, y: 8, width: 376, height: 152, 'class': 'dw' });
    add('path', { d: 'M212 8v152M290 8v152M212 128h-12M290 128h-12', 'class': 'dw' });
    add('text', { x: 16, y: 24 }, 'OPEN SPACE'); add('text', { x: 218, y: 24 }, 'QUIET'); add('text', { x: 296, y: 24 }, 'MEETING');
    DESKS.forEach(function (d) {
      var st = d.EquipmentStatus !== 'Ready' ? 'rep' : my.indexOf(d.Id) > -1 ? 'mine' : busy.indexOf(d.Id) > -1 ? 'busy' : 'free';
      var dim = zone && d.Zone !== zone ? ' off' : '';
      var g = add('g', { 'class': 'desk ' + st + dim });
      var r = document.createElementNS(NS, 'rect'); r.setAttribute('x', d.x); r.setAttribute('y', d.y); r.setAttribute('width', 48); r.setAttribute('height', 40); g.appendChild(r);
      var t = document.createElementNS(NS, 'text'); t.setAttribute('x', d.x + 24); t.setAttribute('y', d.y + 24); t.setAttribute('text-anchor', 'middle'); t.textContent = d.Title; g.appendChild(t);
      var tt = document.createElementNS(NS, 'title');
      tt.textContent = d.Title + ' · ' + (st === 'rep' ? T.repair : st === 'mine' ? T.mine : st === 'busy' ? T.busy : T.ready); g.appendChild(tt);
      if (st === 'busy' || st === 'rep') {
        var hp = document.createElementNS(NS, 'path'); var p = '';
        for (var c = 8; c < 88; c += 8) { var xa = Math.max(0, c - 40), xb = Math.min(48, c); if (xb > xa) p += 'M' + (d.x + xa) + ' ' + (d.y + c - xa) + 'L' + (d.x + xb) + ' ' + (d.y + c - xb); }
        hp.setAttribute('d', p); hp.setAttribute('class', 'hz'); g.appendChild(hp);
      }
      if (st === 'free' && view === 'all' && !dim) {
        g.setAttribute('tabindex', '0'); g.setAttribute('role', 'button'); g.setAttribute('aria-label', T.book + ' ' + d.Title);
        g.addEventListener('click', function () { book(d); });
        g.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); book(d); } });
      }
    });
  }

  function render() {
    viewB.textContent = view === 'all' ? T.my : T.find;
    viewB.setAttribute('aria-pressed', view === 'my' ? 'true' : 'false');
    root.classList.toggle('mineview', view === 'my');
    while (cards.firstChild) cards.removeChild(cards.firstChild);
    var date = dateI.value;
    if (view === 'all') {
      var hidden = others(date).concat(myOn(date));
      var list = DESKS.filter(function (d) { return (!zoneS.value || d.Zone === zoneS.value) && hidden.indexOf(d.Id) === -1; });
      if (date && !list.length) cards.appendChild(el('p', 'dd-empty', T.none));
      list.forEach(function (d) {
        var c = el('div', 'dd-card'), ok = d.EquipmentStatus === 'Ready';
        c.appendChild(el('b', '', d.Title)); c.appendChild(el('span', 'z', d.Zone));
        c.appendChild(el('span', ok ? 's ok' : 's bad', ok ? T.ready : T.repair));
        var b = el('button', '', T.book); b.type = 'button'; b.disabled = !ok; b.addEventListener('click', function () { book(d); });
        c.appendChild(b); cards.appendChild(c);
      });
    } else {
      var my = mine.slice().sort(function (a, b) { return a.date < b.date ? -1 : 1; });
      if (!my.length) cards.appendChild(el('p', 'dd-empty', T.nomine));
      my.forEach(function (r) {
        var d = DESKS.filter(function (x) { return x.Id === r.deskId; })[0]; if (!d) return;
        var c = el('div', 'dd-card');
        c.appendChild(el('b', '', d.Title)); c.appendChild(el('span', 'z', d.Zone)); c.appendChild(el('span', 's', T.date + ': ' + fmt(r.date)));
        var b = el('button', 'sec', T.cancel); b.type = 'button'; b.addEventListener('click', function () { cancel(r); });
        c.appendChild(b); cards.appendChild(c);
      });
    }
    drawPlan();
  }

  var t = new Date(); t.setDate(t.getDate() + 1);
  dateI.value = t.toISOString().slice(0, 10);
  dateI.min = new Date().toISOString().slice(0, 10);
  plan.setAttribute('aria-label', T.plan);
  dateI.addEventListener('change', function () { msg.textContent = ''; msg.className = 'dd-msg'; render(); });
  zoneS.addEventListener('change', render);
  viewB.addEventListener('click', function () { view = view === 'all' ? 'my' : 'all'; msg.textContent = ''; msg.className = 'dd-msg'; render(); });
  resetB.addEventListener('click', function () { mine = []; save(); view = 'all'; say('ok', T.reset); render(); });
  root.classList.add('live');
  render();
})();
