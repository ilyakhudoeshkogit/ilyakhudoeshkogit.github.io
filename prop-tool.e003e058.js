/* SPFx property pane builder: fields → props interface, getPropertyPaneConfiguration(), manifest properties, loc strings. Runs locally, sends nothing. */
(function () {
  var root = document.getElementById('spp'); if (!root) return;
  var en = document.documentElement.lang === 'en';
  var T = en ? {
    wpName: 'Web part name (PascalCase)', paneDesc: 'Pane description', nr: 'disableReactivePropertyChanges (Apply button)',
    fields: 'Fields', add: '+ field', type: 'Field type', name: 'Name (camelCase)', label: 'Label', group: 'Group', page: 'Page', def: 'Default value',
    multi: 'multiline', rows: 'rows', req: 'required', maxLen: 'max length', url: 'URL (https://)', opts: 'Options, one per line: key|Text', min: 'min', max: 'max', step: 'step', href: 'Link URL',
    up: 'Move up', del: 'Delete field', none: 'Add at least one field.', on: 'On', off: 'Off', apply: 'Apply', pageOf: 'Page %1 of %2', back: 'Back', next: 'Next',
    prev: 'Property pane preview', props: 'this.properties (what the web part gets)', pending: 'not applied yet: press Apply',
    b1: 'Props interface (%1WebPart.ts)', b2: 'getPropertyPaneConfiguration() and imports (%1WebPart.ts)', b3: 'preconfiguredEntries[0].properties (%1WebPart.manifest.json)',
    b4: 'loc/mystrings.d.ts', b5: 'loc/en-us.js', b6: 'Non-reactive pane (%1WebPart.ts)', copy: 'Copy', copied: 'Copied', warn: 'Things to remember', errs: 'Fix the errors in the fields: this code will not build as is.',
    eName: 'camelCase: a lowercase letter first, then letters and digits only.', eDup: 'This name is already used by another field.', eWp: 'PascalCase: a capital letter first, then letters and digits only.',
    eOpts: 'Add at least one option.', eKey: 'Option keys must be unique.', eRange: 'min must be less than max, step greater than 0, default within the range.', eHref: 'The link needs an https:// address.',
    vReq: 'Required field', vLen: 'No longer than %1 characters', vUrl: 'Enter an https:// address', page2: 'More settings', grp: 'Settings',
    h_json: 'Property values are saved together with the page and anyone who can edit it can read them. Never put keys, passwords or tokens there.',
    h_restart: 'The manifest is read when the dev server starts: after editing it, restart heft start (gulp serve in SPFx 1.21.1 and older) and add the web part again.',
    h_new: 'Values from preconfiguredEntries are defaults for NEW instances only. Web parts already on pages keep their saved values, so handle a missing property in code.',
    h_changed: 'To react to a change (reset a dependent field, reload data) override onPropertyPaneFieldChanged(propertyPath, oldValue, newValue) and call super.',
    h_async: 'Dropdown options from SharePoint (lists, items) are loaded asynchronously: in onPropertyPaneConfigurationStart() or loadPropertyPaneResources(), then this.context.propertyPane.refresh(); showLoadingIndicator shows a spinner meanwhile.',
    h_pnp: 'For picking a list or a colour there are ready-made controls in @pnp/spfx-property-controls: PropertyFieldListPicker, PropertyFieldColorPicker.',
    h_nr: 'In non-reactive mode changes reach the web part only after Apply, and while a field is invalid Apply is disabled.',
    h_valid: 'onGetErrorMessage returns an error string or an empty string (or a Promise<string> for a server check). Validation starts 200 ms after typing stops: deferredValidationTime changes that.',
    h_static: 'PropertyPaneLink and PropertyPaneLabel only show text: they need a name as the first argument but store nothing, so they are not in the interface or the manifest.',
    h_loc: 'Keep the App*Environment strings that Yeoman generated in mystrings.d.ts and en-us.js: only the lines below are new.'
  } : {
    wpName: 'Имя веб-части (PascalCase)', paneDesc: 'Описание панели', nr: 'disableReactivePropertyChanges (кнопка «Применить»)',
    fields: 'Поля', add: '+ поле', type: 'Тип поля', name: 'Имя (camelCase)', label: 'Подпись', group: 'Группа', page: 'Страница', def: 'По умолчанию',
    multi: 'многострочное', rows: 'строк', req: 'обязательное', maxLen: 'макс. длина', url: 'URL (https://)', opts: 'Варианты, по одному в строке: ключ|Текст', min: 'min', max: 'max', step: 'шаг', href: 'Адрес ссылки',
    up: 'Выше', del: 'Удалить поле', none: 'Добавьте хотя бы одно поле.', on: 'Вкл', off: 'Выкл', apply: 'Применить', pageOf: 'Страница %1 из %2', back: 'Назад', next: 'Далее',
    prev: 'Предпросмотр панели свойств', props: 'this.properties (что получит веб-часть)', pending: 'ещё не применено: нажмите «Применить»',
    b1: 'Интерфейс свойств (%1WebPart.ts)', b2: 'getPropertyPaneConfiguration() и импорты (%1WebPart.ts)', b3: 'preconfiguredEntries[0].properties (%1WebPart.manifest.json)',
    b4: 'loc/mystrings.d.ts', b5: 'loc/en-us.js', b6: 'Нереактивная панель (%1WebPart.ts)', copy: 'Копировать', copied: 'Скопировано', warn: 'О чём помнить', errs: 'Исправьте ошибки в полях: в таком виде код не соберётся.',
    eName: 'camelCase: первая строчная латинская буква, дальше только буквы и цифры.', eDup: 'Это имя уже занято другим полем.', eWp: 'PascalCase: первая заглавная латинская буква, дальше только буквы и цифры.',
    eOpts: 'Добавьте хотя бы один вариант.', eKey: 'Ключи вариантов должны быть разными.', eRange: 'min меньше max, шаг больше 0, значение по умолчанию внутри диапазона.', eHref: 'Ссылке нужен адрес https://.',
    vReq: 'Обязательное поле', vLen: 'Не длиннее %1 символов', vUrl: 'Укажите адрес https://', page2: 'Дополнительно', grp: 'Настройки',
    h_json: 'Значения свойств сохраняются вместе со страницей, и их видит любой, кто может её редактировать. Ключи, пароли и токены туда не кладут.',
    h_restart: 'Манифест читается при запуске сервера разработки: после правки перезапустите heft start (gulp serve в SPFx 1.21.1 и старше) и добавьте веб-часть заново.',
    h_new: 'Значения из preconfiguredEntries действуют только для НОВЫХ экземпляров. Веб-части, уже стоящие на страницах, сохраняют свои значения, поэтому отсутствие свойства надо обработать в коде.',
    h_changed: 'Чтобы отреагировать на изменение (сбросить зависимое поле, перезагрузить данные), переопределите onPropertyPaneFieldChanged(propertyPath, oldValue, newValue) и вызовите super.',
    h_async: 'Варианты выпадающего списка из SharePoint (списки, элементы) загружают асинхронно: в onPropertyPaneConfigurationStart() или loadPropertyPaneResources(), затем this.context.propertyPane.refresh(); на время загрузки showLoadingIndicator показывает индикатор.',
    h_pnp: 'Для выбора списка или цвета есть готовые контролы @pnp/spfx-property-controls: PropertyFieldListPicker, PropertyFieldColorPicker.',
    h_nr: 'В нереактивном режиме изменения попадают в веб-часть только по кнопке «Применить», а пока поле с ошибкой, кнопка недоступна.',
    h_valid: 'onGetErrorMessage возвращает текст ошибки или пустую строку (или Promise<string> для проверки на сервере). Проверка запускается через 200 мс после ввода, это меняет deferredValidationTime.',
    h_static: 'PropertyPaneLink и PropertyPaneLabel только показывают текст: имя первым аргументом им нужно, но значения они не хранят, поэтому их нет в интерфейсе и манифесте.',
    h_loc: 'Строки App*Environment, которые создал Yeoman, в mystrings.d.ts и en-us.js оставьте: новые только строки ниже.'
  };
  var TYPES = {text: 'PropertyPaneTextField', toggle: 'PropertyPaneToggle', dropdown: 'PropertyPaneDropdown', slider: 'PropertyPaneSlider',
    choice: 'PropertyPaneChoiceGroup', checkbox: 'PropertyPaneCheckbox', link: 'PropertyPaneLink', label: 'PropertyPaneLabel'};
  var ORDER = ['text', 'toggle', 'dropdown', 'slider', 'choice', 'checkbox', 'link', 'label'];
  var TS = {text: 'string', toggle: 'boolean', dropdown: 'string', slider: 'number', choice: 'string', checkbox: 'boolean'};
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  function fmt(s) { var a = arguments; return s.replace(/%(\d)/g, function (m, i) { return a[i]; }); }
  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
  function q(s) { return "'" + String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'"; }
  var uid = 0;
  function F(o) {
    var f = {id: ++uid, type: 'text', name: '', label: '', group: T.grp, page: 1, def: '', multiline: false, rows: 4, req: false, max: 0, url: false,
      opts: '', min: 1, mx: 10, step: 1, href: 'https://'};
    for (var k in o) f[k] = o[k];
    return f;
  }
  var PRESETS = {
    exec: function () { return {name: 'ExecutiveDashboard', desc: en ? 'Executive dashboard settings' : 'Настройки панели руководителя', nr: false, fields: [
      F({name: 'title', label: en ? 'Title' : 'Заголовок', def: en ? 'Executive dashboard' : 'Панель руководителя', req: true, max: 60}),
      F({name: 'approvalsList', label: en ? 'Approvals list' : 'Список согласований', def: 'Approvals', req: true}),
      F({type: 'slider', name: 'itemCount', label: en ? 'Items to show' : 'Сколько элементов показать', def: 10, min: 5, mx: 50, step: 5, group: en ? 'Display' : 'Отображение'}),
      F({type: 'toggle', name: 'showKpi', label: en ? 'Show KPI tiles' : 'Показывать плитки KPI', def: true, group: en ? 'Display' : 'Отображение'})]}; },
    desk: function () { return {name: 'DeskReservation', desc: en ? 'Desk booking settings' : 'Настройки бронирования мест', nr: false, fields: [
      F({name: 'desksList', label: en ? 'Desks list' : 'Список мест', def: 'Desks', req: true}),
      F({type: 'dropdown', name: 'defaultZone', label: en ? 'Default zone' : 'Зона по умолчанию', opts: en ? 'open|Open space\nquiet|Quiet zone\nmeet|Meeting rooms' : 'open|Опенспейс\nquiet|Тихая зона\nmeet|Переговорные', def: 'open'}),
      F({type: 'slider', name: 'daysAhead', label: en ? 'Days ahead to book' : 'На сколько дней вперёд', def: 14, min: 1, mx: 30, step: 1}),
      F({type: 'checkbox', name: 'allowCancel', label: en ? 'Allow cancelling a booking' : 'Разрешить отмену брони', def: true})]}; },
    news: function () { return {name: 'TeamNews', desc: en ? 'Team news settings' : 'Настройки новостей команды', nr: true, fields: [
      F({name: 'title', label: en ? 'Title' : 'Заголовок', def: en ? 'Team news' : 'Новости команды', req: true, max: 40}),
      F({name: 'feedUrl', label: en ? 'News site address' : 'Адрес сайта новостей', def: 'https://contoso.sharepoint.com/sites/news', req: true, url: true}),
      F({type: 'choice', name: 'layout', label: en ? 'Layout' : 'Вид', opts: en ? 'list|List\ncards|Cards\ncompact|Compact' : 'list|Список\ncards|Карточки\ncompact|Компактно', def: 'cards', group: en ? 'Display' : 'Отображение'}),
      F({type: 'slider', name: 'itemCount', label: en ? 'News items' : 'Сколько новостей', def: 6, min: 1, mx: 20, step: 1, group: en ? 'Display' : 'Отображение'}),
      F({type: 'toggle', name: 'showImages', label: en ? 'Show images' : 'Показывать картинки', def: true, group: en ? 'Display' : 'Отображение'}),
      F({type: 'text', name: 'footerNote', label: en ? 'Footer note' : 'Подпись внизу', def: '', multiline: true, rows: 3, max: 200, page: 2, group: T.page2}),
      F({type: 'link', name: 'helpLink', label: en ? 'How to publish news' : 'Как опубликовать новость', href: 'https://support.microsoft.com/', page: 2, group: T.page2}),
      F({type: 'label', name: 'aboutLabel', label: en ? 'Demo data, fictional site.' : 'Данные демо, сайт вымышленный.', page: 2, group: T.page2})]}; }
  };
  var S = PRESETS.exec();
  var $ = function (s) { return root.querySelector(s); };
  var ed = $('#p-fields'), pv = $('.p-prev'), out = $('.p-out');
  var vals = {}, pend = {}, curPage = 1;

  function opts(f) {
    return f.opts.split('\n').map(function (l) { l = l.trim(); if (!l) return null; var i = l.indexOf('|'); var k = (i < 0 ? l : l.slice(0, i)).trim(); return {key: k, text: (i < 0 ? l : l.slice(i + 1)).trim() || k}; }).filter(Boolean);
  }
  function stored(f) { return !!TS[f.type]; }
  function errors() {
    var E = {}, seen = {};
    S.fields.forEach(function (f) {
      var e = [];
      if (!/^[a-z][A-Za-z0-9]*$/.test(f.name)) e.push(T.eName);
      else if (seen[f.name]) e.push(T.eDup);
      seen[f.name] = 1;
      if (f.type === 'dropdown' || f.type === 'choice') {
        var o = opts(f), ks = o.map(function (x) { return x.key; });
        if (!o.length) e.push(T.eOpts);
        else if (ks.some(function (k, i) { return ks.indexOf(k) !== i; })) e.push(T.eKey);
      }
      if (f.type === 'slider' && !(+f.min < +f.mx && +f.step > 0 && +f.def >= +f.min && +f.def <= +f.mx)) e.push(T.eRange);
      if (f.type === 'link' && !/^https:\/\/\S+$/.test(f.href)) e.push(T.eHref);
      if (e.length) E[f.id] = e;
    });
    if (!/^[A-Z][A-Za-z0-9]*$/.test(S.name)) E.wp = [T.eWp];
    return E;
  }
  function defOf(f) {
    if (f.type === 'toggle' || f.type === 'checkbox') return !!f.def;
    if (f.type === 'slider') return +f.def;
    if (f.type === 'dropdown' || f.type === 'choice') { var o = opts(f); return o.some(function (x) { return x.key === f.def; }) ? f.def : (o[0] ? o[0].key : ''); }
    return String(f.def);
  }
  function resetVals() { vals = {}; S.fields.forEach(function (f) { if (stored(f)) vals[f.name] = defOf(f); }); pend = JSON.parse(JSON.stringify(vals)); }

  /* ---------- editor ---------- */
  function defCtl(f) {
    if (f.type === 'toggle' || f.type === 'checkbox') return '<label class="p-c"><input type="checkbox" data-k="def"' + (f.def ? ' checked' : '') + '> ' + T.def + '</label>';
    if (f.type === 'slider') return '<label>' + T.def + '<input type="number" data-k="def" value="' + esc(f.def) + '"></label>';
    if (f.type === 'dropdown' || f.type === 'choice') return '<label>' + T.def + '<select data-k="def">' + opts(f).map(function (o) { return '<option value="' + esc(o.key) + '"' + (o.key === defOf(f) ? ' selected' : '') + '>' + esc(o.text) + '</option>'; }).join('') + '</select></label>';
    if (f.type === 'link' || f.type === 'label') return '';
    return '<label>' + T.def + (f.multiline ? '<textarea data-k="def" rows="2">' + esc(f.def) + '</textarea>' : '<input type="text" data-k="def" value="' + esc(f.def) + '">') + '</label>';
  }
  function extra(f) {
    var n = function (k, l, v) { return '<label>' + l + '<input type="number" data-k="' + k + '" value="' + (k === 'max' && !v ? '' : esc(v)) + '"' + (k === 'max' ? ' min="0" placeholder="—"' : '') + '></label>'; };
    var c = function (k, l) { return '<label class="p-c"><input type="checkbox" data-k="' + k + '"' + (f[k] ? ' checked' : '') + '> ' + l + '</label>'; };
    if (f.type === 'text') return '<div class="p-x">' + c('multiline', T.multi) + (f.multiline ? n('rows', T.rows, f.rows) : '') + '<span class="p-sep">onGetErrorMessage:</span>' + c('req', T.req) + n('max', T.maxLen, f.max) + c('url', T.url) + '</div>';
    if (f.type === 'dropdown' || f.type === 'choice') return '<label class="p-wide">' + T.opts + '<textarea data-k="opts" rows="3" spellcheck="false">' + esc(f.opts) + '</textarea></label>';
    if (f.type === 'slider') return '<div class="p-x">' + n('min', T.min, f.min) + n('mx', T.max, f.mx) + n('step', T.step, f.step) + '</div>';
    if (f.type === 'link') return '<label class="p-wide">' + T.href + '<input type="url" data-k="href" value="' + esc(f.href) + '" spellcheck="false"></label>';
    return '';
  }
  function fieldHtml(f, i) {
    return '<div class="p-f" data-id="' + f.id + '"><div class="p-fh"><code>' + TYPES[f.type] + '</code><span>' +
      '<button type="button" class="q-x p-up" aria-label="' + T.up + '"' + (i ? '' : ' disabled') + '>↑</button> <button type="button" class="q-x p-del" aria-label="' + T.del + '">✕</button></span></div>' +
      '<div class="p-fi"><label>' + T.name + '<input type="text" data-k="name" value="' + esc(f.name) + '" spellcheck="false" autocomplete="off"></label>' +
      '<label>' + T.label + '<input type="text" data-k="label" value="' + esc(f.label) + '"></label>' +
      '<label>' + T.group + '<input type="text" data-k="group" value="' + esc(f.group) + '"></label>' +
      '<label>' + T.page + '<select data-k="page"><option value="1"' + (f.page == 1 ? ' selected' : '') + '>1</option><option value="2"' + (f.page == 2 ? ' selected' : '') + '>2</option></select></label>' +
      '<span class="p-def' + (f.type === 'text' ? ' w2' : '') + '">' + defCtl(f) + '</span>' + extra(f) + '</div><p class="p-err" aria-live="polite"></p></div>';
  }
  function drawEditor(newId) {
    $('#p-name').value = S.name; $('#p-desc').value = S.desc; $('#p-nr').checked = S.nr;
    ed.innerHTML = S.fields.length ? S.fields.map(fieldHtml).join('') : '<p class="g-none">' + T.none + '</p>';
    if (newId) { var n = ed.querySelector('[data-id="' + newId + '"]'); if (n) { n.classList.add('new'); n.querySelector('[data-k="name"]').focus(); } }
  }
  function byId(id) { for (var i = 0; i < S.fields.length; i++) if (S.fields[i].id == id) return i; return -1; }
  ed.addEventListener('input', function (e) {
    var t = e.target, k = t.getAttribute('data-k'), box = t.closest('.p-f'); if (!k || !box) return;
    var f = S.fields[byId(box.getAttribute('data-id'))];
    var old = f.name;
    f[k] = t.type === 'checkbox' ? t.checked : (t.type === 'number' ? (t.value === '' ? 0 : +t.value) : (k === 'page' ? +t.value : t.value));
    if (k === 'name' && old !== f.name && old in vals) { vals[f.name] = vals[old]; pend[f.name] = pend[old]; delete vals[old]; delete pend[old]; }
    if (k === 'def' || k === 'opts' || k === 'min' || k === 'mx') { vals[f.name] = pend[f.name] = defOf(f); }
    if (k === 'opts') box.querySelector('.p-def').innerHTML = defCtl(f);
    if (k === 'multiline') { box.querySelector('.p-def').innerHTML = defCtl(f); box.querySelector('.p-x').outerHTML = extra(f); }
    render();
  });
  ed.addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    var i = byId(b.closest('.p-f').getAttribute('data-id'));
    if (b.classList.contains('p-del')) { S.fields.splice(i, 1); }
    else if (b.classList.contains('p-up') && i > 0) { var x = S.fields[i]; S.fields[i] = S.fields[i - 1]; S.fields[i - 1] = x; }
    drawEditor(); resetVals(); render();
  });
  $('#p-name').addEventListener('input', function () { S.name = this.value.trim(); render(); });
  $('#p-desc').addEventListener('input', function () { S.desc = this.value; render(); });
  $('#p-nr').addEventListener('change', function () { S.nr = this.checked; pend = JSON.parse(JSON.stringify(vals)); render(); });
  var sel = $('#p-type');
  sel.innerHTML = ORDER.map(function (k) { return '<option value="' + k + '">' + TYPES[k] + '</option>'; }).join('');
  $('#p-addb').addEventListener('click', function () {
    var t = sel.value, base = {text: 'text', toggle: 'enabled', dropdown: 'option', slider: 'count', choice: 'mode', checkbox: 'flag', link: 'helpLink', label: 'note'}[t], nm = base, k = 2;
    var names = S.fields.map(function (f) { return f.name; });
    while (names.indexOf(nm) >= 0) nm = base + k++;
    var o = {type: t, name: nm, label: cap(nm)};
    if (t === 'dropdown' || t === 'choice') { o.opts = 'a|' + (en ? 'Option A' : 'Вариант A') + '\nb|' + (en ? 'Option B' : 'Вариант B'); o.def = 'a'; }
    if (t === 'slider') o.def = 5;
    if (t === 'toggle' || t === 'checkbox') o.def = false;
    if (t === 'link') o.href = 'https://learn.microsoft.com/';
    var f = F(o); S.fields.push(f); drawEditor(f.id); if (stored(f)) vals[f.name] = pend[f.name] = defOf(f); render();
  });

  /* ---------- preview ---------- */
  function vmsg(f, v) {
    if (f.type !== 'text' || !(f.req || f.max > 0 || f.url)) return '';
    v = v == null ? '' : String(v);
    if (f.req && !v.trim()) return T.vReq;
    if (f.max > 0 && v.length > f.max) return fmt(T.vLen, f.max);
    if (f.url && v && !/^https:\/\/\S+$/.test(v)) return T.vUrl;
    return '';
  }
  function pages() { var p = [1]; S.fields.forEach(function (f) { if (p.indexOf(f.page) < 0) p.push(f.page); }); return p.sort(); }
  function groups(pg) { var g = []; S.fields.forEach(function (f) { if (f.page == pg && g.indexOf(f.group) < 0) g.push(f.group); }); return g; }
  function ctl(f) {
    var v = pend[f.name], id = 'pp-' + f.id, m = vmsg(f, v);
    var lab = '<label for="' + id + '">' + esc(f.label || f.name) + '</label>';
    if (f.type === 'text') return lab + (f.multiline ? '<textarea id="' + id + '" data-n="' + esc(f.name) + '" rows="' + (f.rows || 3) + '">' + esc(v) + '</textarea>' : '<input id="' + id + '" type="text" data-n="' + esc(f.name) + '" value="' + esc(v) + '">') + '<small class="pp-e" role="alert">' + esc(m) + '</small>';
    if (f.type === 'toggle') return lab + '<span class="pp-tg"><input id="' + id + '" type="checkbox" role="switch" data-n="' + esc(f.name) + '"' + (v ? ' checked' : '') + '><i aria-hidden="true"></i><span>' + (v ? T.on : T.off) + '</span></span>';
    if (f.type === 'dropdown') return lab + '<select id="' + id + '" data-n="' + esc(f.name) + '">' + opts(f).map(function (o) { return '<option value="' + esc(o.key) + '"' + (o.key === v ? ' selected' : '') + '>' + esc(o.text) + '</option>'; }).join('') + '</select>';
    if (f.type === 'slider') return lab + '<span class="pp-sl"><input id="' + id + '" type="range" data-n="' + esc(f.name) + '" min="' + f.min + '" max="' + f.mx + '" step="' + f.step + '" value="' + v + '"><output>' + v + '</output></span>';
    if (f.type === 'choice') return '<fieldset class="pp-cg"><legend>' + esc(f.label || f.name) + '</legend>' + opts(f).map(function (o, i) { return '<label class="pp-r"><input type="radio" name="' + id + '" data-n="' + esc(f.name) + '" value="' + esc(o.key) + '"' + (o.key === v ? ' checked' : '') + '> ' + esc(o.text) + '</label>'; }).join('') + '</fieldset>';
    if (f.type === 'checkbox') return '<label class="pp-r"><input type="checkbox" data-n="' + esc(f.name) + '"' + (v ? ' checked' : '') + '> ' + esc(f.label || f.name) + '</label>';
    if (f.type === 'link') return '<a href="' + esc(f.href) + '" target="_blank" rel="noopener">' + esc(f.label || f.href) + '</a>';
    return '<p class="pp-lb">' + esc(f.label) + '</p>';
  }
  function drawPreview() {
    var a = document.activeElement, keep = a && pv.contains(a) ? (a.id ? '#' + a.id : (a.name ? 'input[name="' + a.name + '"][value="' + a.value + '"]' : (a.hasAttribute('data-go') ? '[data-go="' + a.getAttribute('data-go') + '"]' : null))) : null;
    drawPreviewInner();
    if (keep) { var n = pv.querySelector(keep); if (n && !n.disabled) n.focus(); }
  }
  function drawPreviewInner() {
    var P = pages(); if (P.indexOf(curPage) < 0) curPage = 1;
    var bad = S.fields.some(function (f) { return vmsg(f, pend[f.name]); });
    var h = '<div class="pp-h"><b>' + esc(S.name.replace(/([a-z])([A-Z])/g, '$1 $2') || 'Web part') + '</b><span aria-hidden="true">✕</span></div>';
    if (P.length > 1) h += '<p class="pp-pg">' + fmt(T.pageOf, P.indexOf(curPage) + 1, P.length) + '</p>';
    h += '<p class="pp-d">' + esc(curPage === 1 ? S.desc : T.page2) + '</p>';
    groups(curPage).forEach(function (g) {
      h += '<details class="pp-g" open><summary>' + esc(g) + '</summary>' + S.fields.filter(function (f) { return f.page == curPage && f.group === g; }).map(function (f) { return '<div class="pp-f">' + ctl(f) + '</div>'; }).join('') + '</details>';
    });
    if (P.length > 1) h += '<p class="pp-nav"><button type="button" class="q-pre" data-go="-1"' + (curPage === P[0] ? ' disabled' : '') + '>' + T.back + '</button><button type="button" class="q-pre" data-go="1"' + (curPage === P[P.length - 1] ? ' disabled' : '') + '>' + T.next + '</button></p>';
    if (S.nr) h += '<p class="pp-ap"><button type="button" class="q-pre pp-apply"' + (bad || JSON.stringify(pend) === JSON.stringify(vals) ? ' disabled' : '') + '>' + T.apply + '</button></p>';
    var diff = S.nr && JSON.stringify(pend) !== JSON.stringify(vals);
    pv.innerHTML = '<div class="pp" role="group" aria-label="' + T.prev + '">' + h + '</div><div class="q-blk pp-props"><div class="q-h"><small>' + T.props + '</small></div><pre><code>' + esc(JSON.stringify(vals, null, 2)) + '</code></pre>' + (diff ? '<p class="pp-pend">' + T.pending + '</p>' : '') + '</div>';
  }
  pv.addEventListener('input', function (e) {
    var t = e.target, n = t.getAttribute('data-n'); if (!n) return;
    var f = S.fields.filter(function (x) { return x.name === n; })[0]; if (!f) return;
    var v = t.type === 'checkbox' ? t.checked : (t.type === 'range' ? +t.value : t.value);
    pend[n] = v;
    if (!S.nr && !vmsg(f, v)) vals[n] = v;
    if (t.type === 'range') { t.nextSibling.textContent = v; var pr = pv.querySelector('.pp-props pre code'); if (pr) pr.textContent = JSON.stringify(vals, null, 2); return; }
    if (t.type === 'text' || t.tagName === 'TEXTAREA') {
      var sm = t.parentNode.querySelector('.pp-e'); if (sm) sm.textContent = vmsg(f, v);
      var pr2 = pv.querySelector('.pp-props pre code'); if (pr2) pr2.textContent = JSON.stringify(vals, null, 2);
      var ap = pv.querySelector('.pp-apply'); if (ap) ap.disabled = S.fields.some(function (x) { return vmsg(x, pend[x.name]); }) || JSON.stringify(pend) === JSON.stringify(vals);
      return;
    }
    drawPreview();
  });
  pv.addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    if (b.hasAttribute('data-go')) { var P = pages(); curPage = P[P.indexOf(curPage) + (+b.getAttribute('data-go'))] || curPage; drawPreview(); }
    else if (b.classList.contains('pp-apply')) { vals = JSON.parse(JSON.stringify(pend)); drawPreview(); }
  });

  /* ---------- code ---------- */
  function strKeys() {
    var K = [['PropertyPaneDescription', S.desc]], G = {}, n = 1;
    if (pages().indexOf(2) >= 0) K.push(['Page2Description', T.page2]);
    pages().forEach(function (p) { groups(p).forEach(function (g) { var key = p + '\u0001' + g; if (!G[key]) { G[key] = n === 1 ? 'BasicGroupName' : 'Group' + n + 'Name'; K.push([G[key], g]); n++; } }); });
    S.fields.forEach(function (f) { K.push([cap(f.name) + 'FieldLabel', f.label || f.name]); if (f.type === 'text' && f.max > 0) K.push([cap(f.name) + 'MaxLengthError', fmt(T.vLen, f.max)]); });
    if (S.fields.some(function (f) { return f.type === 'text' && f.req; })) K.push(['RequiredError', T.vReq]);
    if (S.fields.some(function (f) { return f.type === 'text' && f.url; })) K.push(['UrlError', T.vUrl]);
    return {K: K, G: G};
  }
  function fieldCode(f, ind) {
    var p = [], L = 'strings.' + cap(f.name) + 'FieldLabel', pad = ind + '  ';
    if (f.type === 'text') { p.push('label: ' + L); if (f.multiline) { p.push('multiline: true'); p.push('rows: ' + (f.rows || 3)); } if (f.req || f.max > 0 || f.url) p.push('onGetErrorMessage: this.validate' + cap(f.name) + '.bind(this)'); }
    else if (f.type === 'toggle') { p.push('label: ' + L); p.push('onText: strings.ToggleOnText'); p.push('offText: strings.ToggleOffText'); }
    else if (f.type === 'dropdown' || f.type === 'choice') {
      p.push('label: ' + L);
      p.push('options: [\n' + opts(f).map(function (o) { return pad + '  { key: ' + q(o.key) + ', text: ' + q(o.text) + ' }'; }).join(',\n') + '\n' + pad + ']');
    }
    else if (f.type === 'slider') { p.push('label: ' + L); p.push('min: ' + (+f.min)); p.push('max: ' + (+f.mx)); p.push('step: ' + (+f.step)); p.push('showValue: true'); }
    else if (f.type === 'checkbox') p.push('text: ' + L);
    else if (f.type === 'link') { p.push('href: ' + q(f.href)); p.push('text: ' + L); p.push("target: '_blank'"); }
    else p.push('text: ' + L);
    return ind + TYPES[f.type] + '(' + q(f.name) + ', {\n' + p.map(function (x) { return pad + x; }).join(',\n') + '\n' + ind + '})';
  }
  function gen() {
    var N = S.name || 'My', I = 'I' + N + 'WebPartProps', sk = strKeys();
    if (S.fields.some(function (f) { return f.type === 'toggle'; })) sk.K.push(['ToggleOnText', T.on], ['ToggleOffText', T.off]);
    var st = S.fields.filter(stored);
    var iface = 'export interface ' + I + ' {\n' + st.map(function (f) { return '  ' + f.name + ': ' + TS[f.type] + ';'; }).join('\n') + (st.length ? '\n' : '') + '}';
    var used = ORDER.filter(function (t) { return S.fields.some(function (f) { return f.type === t; }); }).map(function (t) { return TYPES[t]; });
    var imp = "import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';\nimport {\n  IPropertyPaneConfiguration" + used.map(function (u) { return ',\n  ' + u; }).join('') + "\n} from '@microsoft/sp-property-pane';\nimport * as strings from '" + N + "WebPartStrings';";
    var pgs = pages().map(function (p) {
      var gs = groups(p).map(function (g) {
        var fs = S.fields.filter(function (f) { return f.page == p && f.group === g; }).map(function (f) { return fieldCode(f, '                '); });
        return '          {\n            groupName: strings.' + sk.G[p + '\u0001' + g] + ',\n            groupFields: [\n' + fs.join(',\n') + '\n            ]\n          }';
      });
      return '      {\n        header: {\n          description: strings.' + (p === 1 ? 'PropertyPaneDescription' : 'Page2Description') + '\n        },\n        groups: [\n' + gs.join(',\n') + '\n        ]\n      }';
    });
    var vals2 = S.fields.filter(function (f) { return f.type === 'text' && (f.req || f.max > 0 || f.url); }).map(function (f) {
      var b = ['  private validate' + cap(f.name) + '(value: string): string {'];
      if (f.req) b.push('    if (!value || value.trim().length === 0) {', '      return strings.RequiredError;', '    }');
      if (f.max > 0) b.push('    if (value && value.length > ' + (+f.max) + ') {', '      return strings.' + cap(f.name) + 'MaxLengthError;', '    }');
      if (f.url) b.push('    if (value && !/^https:\\/\\/\\S+$/.test(value)) {', '      return strings.UrlError;', '    }');
      b.push("    return '';", '  }');
      return b.join('\n');
    });
    var cls = imp + '\n\nexport default class ' + N + 'WebPart extends BaseClientSideWebPart<' + I + '> {\n  // render(), onInit() ...\n\n' +
      '  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {\n    return {\n      pages: [\n' + pgs.join(',\n') + '\n      ]\n    };\n  }' +
      (vals2.length ? '\n\n' + vals2.join('\n\n') : '') + '\n}';
    var props = {}; st.forEach(function (f) { props[f.name] = defOf(f); });
    var man = '"properties": ' + JSON.stringify(props, null, 2);
    var dts = 'declare interface I' + N + 'WebPartStrings {\n' + sk.K.map(function (k) { return '  ' + k[0] + ': string;'; }).join('\n') + '\n}\n\ndeclare module \'' + N + 'WebPartStrings\' {\n  const strings: I' + N + 'WebPartStrings;\n  export = strings;\n}';
    var js = 'define([], function() {\n  return {\n' + sk.K.map(function (k) { return '    ' + JSON.stringify(k[0]) + ': ' + JSON.stringify(k[1]); }).join(',\n') + '\n  }\n});';
    var nr = 'protected get disableReactivePropertyChanges(): boolean {\n  return true;\n}';
    return {iface: iface, cls: cls, man: man, dts: dts, js: js, nr: nr, N: N};
  }
  function blk(title, text, k) {
    return '<div class="q-blk" data-b="' + k + '"><div class="q-h"><small>' + esc(title) + '</small><button type="button" class="q-cp">' + T.copy + '</button></div><pre><code>' + esc(text) + '</code></pre></div>';
  }
  function drawCode() {
    var E = errors(), c = gen();
    S.fields.forEach(function (f) {
      var box = ed.querySelector('[data-id="' + f.id + '"]'); if (!box) return;
      box.querySelector('.p-err').textContent = (E[f.id] || []).join(' ');
      box.querySelector('[data-k="name"]').setAttribute('aria-invalid', E[f.id] && (E[f.id].indexOf(T.eName) >= 0 || E[f.id].indexOf(T.eDup) >= 0) ? 'true' : 'false');
    });
    $('#p-name-err').textContent = E.wp ? E.wp[0] : '';
    if (!S.fields.length) { out.innerHTML = '<p class="g-none">' + T.none + '</p>'; return; }
    var H = [T.h_json, T.h_restart, T.h_new, T.h_changed];
    if (S.fields.some(function (f) { return f.type === 'dropdown'; })) H.push(T.h_async);
    if (S.fields.some(function (f) { return /list/i.test(f.name) || /colou?r/i.test(f.name); })) H.push(T.h_pnp);
    if (S.fields.some(function (f) { return f.type === 'text' && (f.req || f.max > 0 || f.url); })) H.push(T.h_valid);
    if (S.nr) H.push(T.h_nr);
    if (S.fields.some(function (f) { return !stored(f); })) H.push(T.h_static);
    H.push(T.h_loc);
    out.innerHTML = (Object.keys(E).length ? '<p class="p-bad" role="alert">' + T.errs + '</p>' : '') +
      blk(fmt(T.b1, c.N), c.iface, 'iface') + blk(fmt(T.b2, c.N), c.cls, 'cls') + blk(fmt(T.b3, c.N), c.man, 'man') + blk(T.b4, c.dts, 'dts') + blk(T.b5, c.js, 'js') +
      (S.nr ? blk(fmt(T.b6, c.N), c.nr, 'nr') : '') +
      '<div class="q-w"><small>' + T.warn + '</small><ul>' + H.map(function (h) { return '<li>' + esc(h) + '</li>'; }).join('') + '</ul></div>';
  }
  out.addEventListener('click', function (e) {
    var b = e.target.closest('.q-cp'); if (!b) return;
    var txt = b.parentNode.nextSibling.textContent;
    (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).then(function () { b.textContent = T.copied; setTimeout(function () { b.textContent = T.copy; }, 1400); }, function () {});
  });
  function render() { drawPreview(); drawCode(); }
  function load(k) { S = PRESETS[k](); curPage = 1; drawEditor(); resetVals(); render(); }
  [].forEach.call(root.querySelectorAll('[data-preset]'), function (b) { b.addEventListener('click', function () { load(b.getAttribute('data-preset')); }); });
  root.classList.add('on');
  load('exec');
})();
