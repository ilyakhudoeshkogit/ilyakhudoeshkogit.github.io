/* Teams Photo Bot: in-browser imitation of the bot dialog (fictional profile, nothing is sent anywhere). */
(function () {
  var root = document.getElementById('botdemo');
  if (!root) return;
  var en = document.documentElement.lang === 'en';
  var T = en ? {
    you: 'You', bot: 'Photo Bot', hello: 'Hi! I am Photo Bot. Type "photo" and I will show your Microsoft 365 profile card. "help" lists the commands.',
    help: 'Commands: "photo" shows your profile card with the photo from Microsoft Graph; "help" shows this list.',
    unknown: 'I do not know the command "%s". Try "photo" or "help".', empty: 'Type a message first.',
    title: 'Your profile', name: 'Demo User', job: 'Engineer · training tenant', mail: 'demo.user@contoso.example',
    nophoto: 'No photo in the profile: the bot shows initials instead.', refresh: 'Refresh', typing: 'Photo Bot is typing…',
    ini: 'DU', alt: 'Profile photo (drawn placeholder for the demo)', altIni: 'Initials DU instead of a photo'
  } : {
    you: 'Вы', bot: 'Photo Bot', hello: 'Привет! Я Photo Bot. Напишите «photo», и я покажу карточку вашего профиля Microsoft 365. «help» покажет команды.',
    help: 'Команды: «photo» показывает карточку профиля с фото из Microsoft Graph; «help» показывает этот список.',
    unknown: 'Команду «%s» я не знаю. Попробуйте «photo» или «help».', empty: 'Сначала напишите сообщение.',
    title: 'Ваш профиль', name: 'Демо Пользователь', job: 'Инженер · учебный тенант', mail: 'demo.user@contoso.example',
    nophoto: 'В профиле нет фото: бот показывает инициалы.', refresh: 'Обновить', typing: 'Photo Bot печатает…',
    ini: 'ДП', alt: 'Фото профиля (нарисованная заглушка для демо)', altIni: 'Инициалы ДП вместо фото'
  };
  var $ = function (id) { return document.getElementById(id); };
  var log = $('bd-log'), form = $('bd-form'), input = $('bd-in'), mode = $('bd-mode'), jsonB = $('bd-json'), pre = $('bd-pre'), typing = $('bd-typing');
  var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var lastCard = null, busy = false;

  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function time() { var d = new Date(); return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }
  function bubble(who, node) {
    var m = el('div', 'bd-m ' + (who === 'me' ? 'me' : 'bot'));
    var h = el('p', 'bd-h'); h.appendChild(el('b', '', who === 'me' ? T.you : T.bot)); h.appendChild(document.createTextNode(' · ' + time()));
    m.appendChild(h);
    if (typeof node === 'string') m.appendChild(el('p', 'bd-t', node)); else m.appendChild(node);
    log.appendChild(m); log.scrollTop = log.scrollHeight;
  }
  function avatar(withPhoto) {
    var NS = 'http://www.w3.org/2000/svg', s = document.createElementNS(NS, 'svg');
    s.setAttribute('viewBox', '0 0 96 96'); s.setAttribute('role', 'img'); s.setAttribute('class', 'bd-av' + (withPhoto ? '' : ' ini'));
    s.setAttribute('aria-label', withPhoto ? T.alt : T.altIni);
    var add = function (tag, a, txt) { var e = document.createElementNS(NS, tag); for (var k in a) e.setAttribute(k, a[k]); if (txt) e.textContent = txt; s.appendChild(e); return e; };
    add('rect', { x: 2, y: 2, width: 92, height: 92, 'class': 'fr' });
    if (withPhoto) {
      var p = ''; for (var c = 10; c < 184; c += 10) { var xa = Math.max(0, c - 92), xb = Math.min(92, c); p += 'M' + (2 + xa) + ' ' + (2 + c - xa) + 'L' + (2 + xb) + ' ' + (2 + c - xb); }
      add('path', { d: p, 'class': 'hz' });
      add('circle', { cx: 48, cy: 38, r: 17, 'class': 'fl' });
      add('path', { d: 'M16 94c3-20 16-30 32-30s29 10 32 30z', 'class': 'fl' });
    } else {
      add('text', { x: 48, y: 58, 'text-anchor': 'middle' }, T.ini);
    }
    return s;
  }
  function cardJson(withPhoto) {
    return {
      type: 'AdaptiveCard', version: '1.5', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      body: [
        { type: 'TextBlock', text: T.title, weight: 'Bolder', size: 'Medium' },
        { type: 'ColumnSet', columns: [
          { type: 'Column', width: 'auto', items: [ withPhoto
            ? { type: 'Image', url: 'data:image/jpeg;base64,/9j/4AAQ…', size: 'Large', altText: T.name }
            : { type: 'TextBlock', text: T.ini, size: 'ExtraLarge', weight: 'Bolder' } ] },
          { type: 'Column', width: 'stretch', items: [
            { type: 'TextBlock', text: T.name, weight: 'Bolder', wrap: true },
            { type: 'TextBlock', text: T.job, isSubtle: true, spacing: 'None', wrap: true },
            { type: 'TextBlock', text: T.mail, isSubtle: true, spacing: 'None' } ] } ] }
      ].concat(withPhoto ? [] : [{ type: 'TextBlock', text: T.nophoto, isSubtle: true, wrap: true }]),
      actions: [{ type: 'Action.Submit', title: T.refresh, data: { command: 'photo' } }]
    };
  }
  function card(withPhoto) {
    var c = el('div', 'bd-card');
    c.appendChild(el('p', 'bd-ct', T.title));
    var row = el('div', 'bd-row'); row.appendChild(avatar(withPhoto));
    var info = el('div', 'bd-info'); info.appendChild(el('b', '', T.name)); info.appendChild(el('span', '', T.job)); info.appendChild(el('span', '', T.mail));
    row.appendChild(info); c.appendChild(row);
    if (!withPhoto) c.appendChild(el('p', 'bd-note', T.nophoto));
    var b = el('button', 'sec', T.refresh); b.type = 'button'; b.addEventListener('click', function () { send('photo', true); });
    c.appendChild(b);
    c.appendChild(el('code', 'bd-trace', withPhoto ? 'GET /users/{id}/photo/$value → 200 image/jpeg' : 'GET /users/{id}/photo/$value → 404 ImageNotFound'));
    lastCard = cardJson(withPhoto); showJson();
    return c;
  }
  function showJson() { if (lastCard) pre.textContent = JSON.stringify(lastCard, null, 2); }
  function reply(text) {
    var cmd = text.trim().toLowerCase().replace(/^\//, '');
    if (cmd === 'photo' || cmd === 'фото') return card(mode.value === 'photo');
    if (cmd === 'help' || cmd === 'помощь') return T.help;
    return T.unknown.replace('%s', text.trim().slice(0, 40));
  }
  function send(text, silentUser) {
    if (busy) return;
    if (!text.trim()) { input.focus(); input.setAttribute('placeholder', T.empty); return; }
    if (!silentUser) bubble('me', text.trim());
    busy = true; typing.hidden = false; typing.textContent = T.typing;
    setTimeout(function () { typing.hidden = true; bubble('bot', reply(text)); busy = false; }, reduce ? 0 : 650);
  }
  form.addEventListener('submit', function (e) { e.preventDefault(); var v = input.value; input.value = ''; send(v); });
  Array.prototype.forEach.call(root.querySelectorAll('[data-cmd]'), function (b) { b.addEventListener('click', function () { send(b.getAttribute('data-cmd')); }); });
  jsonB.addEventListener('click', function () {
    var on = jsonB.getAttribute('aria-pressed') !== 'true';
    jsonB.setAttribute('aria-pressed', on ? 'true' : 'false'); pre.hidden = !on; showJson();
  });
  bubble('bot', T.hello);
  root.classList.add('live');
})();
