/* SharePoint site script (site template) builder: JSON + PowerShell + schematic preview. Runs locally, sends nothing. */
(function () {
  var root = document.getElementById('sst'); if (!root) return;
  var en = document.documentElement.lang === 'en';
  var T = en ? {
    list: 'List', lname: 'List name', ltype: 'Template', ldesc: 'Description', nav: 'Add to navigation', fields: 'Columns',
    dn: 'Display name', inn: 'Internal name', ty: 'Type', ex: 'Choices (comma) / target list', req: 'Required', rq: 'Req.',
    addf: '+ column', dell: 'Remove list', delf: 'Remove column',
    tt: {100: 'Custom list', 101: 'Document library', 107: 'Tasks', 106: 'Events'},
    ft: {Text: 'Single line', Note: 'Multiple lines', Number: 'Number', Boolean: 'Yes/No', DateTime: 'Date and time', User: 'Person', Choice: 'Choice', Lookup: 'Lookup'},
    json: 'site-script.json', spo: 'SharePoint Online Management Shell', pnp: 'PnP PowerShell', copy: 'Copy', copied: 'Copied', warn: 'What to keep in mind',
    prev: 'Preview: what the site gets', seq: 'Order of actions', acts: 'actions', chars: 'characters', nonav: 'No navigation links', nolists: 'No lists', site: 'Site',
    tpl: {64: 'Team site', 1: 'Team site without a group', 68: 'Communication site'},
    v: {createSPList: 'create list', addSPField: 'column', addSPFieldXml: 'column (XML)', addSPLookupFieldXml: 'lookup (XML)', setDescription: 'description', addNavLink: 'navigation link', applyTheme: 'theme', setSiteBranding: 'header', setSiteLogo: 'logo', joinHubSite: 'join hub', setRegionalSettings: 'regional settings', setSiteExternalSharingCapability: 'guest access', installSolution: 'install solution'},
    h_paste: 'Save the JSON as UTF-8 and register it with one of the scripts below (tenant admin rights are needed). The template then appears under "Apply a site template" on sites of the chosen type.',
    h_ids: 'Every column gets a fixed GUID derived from the list and column name. Thanks to it, applying the template again updates the column instead of adding a duplicate.',
    h_30: 'More than 30 actions: Invoke-SPOSiteDesign (synchronous) will refuse it. Applying from the UI or with Add-SPOSiteDesignTask allows up to 300 actions or 100,000 characters.',
    h_300: 'More than 300 actions or 100,000 characters: split it into several site scripts and attach them to one template.',
    h_lk: 'Lookup "%1" points to list "%2", which this script does not create before it. Create the target list earlier in the script or make sure it already exists on the site.',
    h_lk0: 'Lookup "%1" has no target list.', h_ch: 'Choice column "%1" has no choices.',
    h_in: 'Column "%1": the internal name is generated from the display name and must use Latin letters and digits only, so I put "%2". Set your own, it cannot be changed after the column is created.',
    h_dup: 'List "%1" has two columns with the internal name "%2": the second one would update the first.',
    h_logo: 'setSiteLogo only works on communication sites (68); on team sites it is ignored.',
    h_theme: 'applyTheme only accepts a custom theme already added to the tenant (Add-SPOTheme). A built-in theme has to be copied as a custom one first.',
    h_hub: 'The hub site ID is not a GUID. Get it with Get-SPOHubSite or Get-PnPHubSite (the SiteId field).',
    h_sol: 'The solution ID is not a GUID. Take it from Get-PnPApp: the .sppkg must already be deployed in the tenant App Catalog.',
    h_lib: 'Library "%1" uses template 101: it is reached at /%1, not /Lists/%1.',
    h_ru: 'The list name becomes part of its URL. Cyrillic and spaces work but turn into long %D0%... sequences in links, so short Latin names are easier to live with.',
    h_reg: 'setRegionalSettings requires sortOrder (the collation): 7 is Cyrillic, 25 is General. The time zone is left as it is; add timeZone with a number from SPRegionalSettings.TimeZones if you need it.',
    h_rerun: 'Actions can be rerun: the same list name updates the existing list instead of creating a second one. Renaming a list later (setTitle) breaks that link.'
  } : {
    list: 'Список', lname: 'Имя списка', ltype: 'Шаблон', ldesc: 'Описание', nav: 'В меню сайта', fields: 'Столбцы',
    dn: 'Отображаемое имя', inn: 'Внутреннее имя', ty: 'Тип', ex: 'Варианты (через запятую) / целевой список', req: 'Обязательный', rq: 'Обяз.',
    addf: '+ столбец', dell: 'Убрать список', delf: 'Убрать столбец',
    tt: {100: 'Обычный список', 101: 'Библиотека документов', 107: 'Задачи', 106: 'События'},
    ft: {Text: 'Строка', Note: 'Несколько строк', Number: 'Число', Boolean: 'Да/Нет', DateTime: 'Дата и время', User: 'Пользователь', Choice: 'Выбор', Lookup: 'Подстановка'},
    json: 'site-script.json', spo: 'SharePoint Online Management Shell', pnp: 'PnP PowerShell', copy: 'Копировать', copied: 'Скопировано', warn: 'О чём помнить',
    prev: 'Предпросмотр: что получит сайт', seq: 'Порядок действий', acts: 'действий', chars: 'символов', nonav: 'Ссылок в меню нет', nolists: 'Списков нет', site: 'Сайт',
    tpl: {64: 'Сайт группы', 1: 'Сайт группы без группы M365', 68: 'Информационный сайт'},
    v: {createSPList: 'создать список', addSPField: 'столбец', addSPFieldXml: 'столбец (XML)', addSPLookupFieldXml: 'подстановка (XML)', setDescription: 'описание', addNavLink: 'ссылка в меню', applyTheme: 'тема', setSiteBranding: 'шапка', setSiteLogo: 'логотип', joinHubSite: 'присоединить к хабу', setRegionalSettings: 'региональные настройки', setSiteExternalSharingCapability: 'гостевой доступ', installSolution: 'установить решение'},
    h_paste: 'Сохраните JSON в UTF-8 и зарегистрируйте одним из скриптов ниже (нужны права администратора клиента). После этого шаблон появится в «Применить шаблон сайта» на сайтах выбранного типа.',
    h_ids: 'Каждому столбцу выдан постоянный GUID из имени списка и столбца. Благодаря ему повторное применение шаблона обновит столбец, а не создаст дубль.',
    h_30: 'Больше 30 действий: синхронный Invoke-SPOSiteDesign такой скрипт не примет. Из интерфейса и через Add-SPOSiteDesignTask можно до 300 действий или 100 000 символов.',
    h_300: 'Больше 300 действий или 100 000 символов: разбейте на несколько скриптов сайта и подключите их к одному шаблону.',
    h_lk: 'Подстановка «%1» ссылается на список «%2», который этот скрипт не создаёт раньше неё. Создайте целевой список выше по скрипту или убедитесь, что он уже есть на сайте.',
    h_lk0: 'У подстановки «%1» не указан целевой список.', h_ch: 'У столбца выбора «%1» нет вариантов.',
    h_in: 'Столбец «%1»: внутреннее имя строится из отображаемого и должно быть только из латиницы и цифр, поэтому поставил «%2». Задайте своё: после создания столбца его не поменять.',
    h_dup: 'В списке «%1» два столбца с внутренним именем «%2»: второй обновит первый.',
    h_logo: 'setSiteLogo работает только на информационных сайтах (68), на сайте группы он игнорируется.',
    h_theme: 'applyTheme принимает только свою тему, уже добавленную в клиент (Add-SPOTheme). Встроенную тему сначала копируют как собственную.',
    h_hub: 'ID хаба не похож на GUID. Возьмите его из Get-SPOHubSite или Get-PnPHubSite (поле SiteId).',
    h_sol: 'ID решения не похож на GUID. Его даёт Get-PnPApp, а сам .sppkg должен быть развёрнут в каталоге приложений клиента.',
    h_lib: 'Библиотека «%1» на шаблоне 101 открывается по адресу /%1, а не /Lists/%1.',
    h_ru: 'Имя списка становится частью его адреса. Кириллица и пробелы работают, но в ссылках превращаются в длинные %D0%..., поэтому короткие латинские имена удобнее.',
    h_reg: 'Для setRegionalSettings обязателен sortOrder (порядок сортировки): 7 это кириллица, 25 общий. Часовой пояс не трогаю; если нужен, добавьте timeZone с номером из SPRegionalSettings.TimeZones.',
    h_rerun: 'Действия можно применять повторно: то же имя списка обновит существующий список, а не создаст второй. Переименование списка потом (setTitle) эту связь ломает.'
  };
  var FT = ['Text', 'Note', 'Number', 'Boolean', 'DateTime', 'User', 'Choice', 'Lookup'];
  var GLYPH = {Text: 'Aa', Note: '¶', Number: '#', Boolean: '✓', DateTime: '◷', User: '☺', Choice: '◉', Lookup: '↗'};
  var DEPTS = 'Executive, HR, Finance, IT, Sales, PMO, Marketing';
  var P = {
    halden: {title: 'Halden department', tpl: '68', theme: 'Halden', hub: '', locale: '1033', header: 'Compact', bg: 'Soft', logo: '/SiteAssets/halden-logo.png', sharing: 'Disabled', sol: '', soln: '',
      lists: [
        {n: 'Halden Tickets', t: '100', d: 'IT requests from staff', nav: 1, f: [['Detail', 'HaldenDetail', 'Note', '', 0], ['Category', 'HaldenTicketCategory', 'Choice', 'Hardware, Software, Access, Network, Store systems, Other', 1], ['Priority', 'HaldenPriority', 'Choice', 'High, Medium, Low', 1], ['Status', 'HaldenTicketStatus', 'Choice', 'New, In progress, Resolved', 0], ['Requester', 'HaldenRequester', 'Text', '', 0], ['Ticket no.', 'HaldenTicketNo', 'Number', '', 0]]},
        {n: 'Halden FAQ', t: '100', d: 'Questions and answers', nav: 1, f: [['Answer', 'HaldenAnswer', 'Note', '', 1], ['Category', 'HaldenFaqCategory', 'Text', '', 0], ['Department', 'HaldenDepartment', 'Choice', DEPTS, 0]]},
        {n: 'Policies', t: '101', d: 'Policies and handbooks', nav: 1, f: [['Policy type', 'HaldenPolicyType', 'Choice', 'Policy, Handbook, Resource', 0], ['Review date', 'HaldenReviewDate', 'DateTime', '', 0]]}]},
    desk: {title: 'Desk booking', tpl: '64', theme: '', hub: '', locale: '1049', header: '', bg: '', logo: '', sharing: 'Disabled', sol: '', soln: 'desk-reservation-client-side-solution',
      lists: [
        {n: 'Desks', t: '100', d: 'Bookable desks', nav: 1, f: [['Zone', 'Zone', 'Choice', 'A, B, C', 1], ['Equipment status', 'EquipmentStatus', 'Choice', 'OK, Needs repair', 0]]},
        {n: 'Reservations', t: '100', d: 'One row per desk per day', nav: 0, f: [['Desk', 'Desk', 'Lookup', 'Desks', 1], ['Reserved date', 'ReservedDate', 'DateTime', '', 1], ['Reserved by', 'ReservedBy', 'User', '', 1]]}]},
    exec: {title: 'Executive dashboard', tpl: '64', theme: '', hub: '', locale: '1033', header: 'Standard', bg: 'Neutral', logo: '', sharing: 'ExistingExternalUserSharingOnly', sol: '', soln: 'Executive Dashboard (portfolio)',
      lists: [
        {n: 'Approvals', t: '100', d: 'Requests waiting for a decision', nav: 1, f: [['Status', 'Status', 'Choice', 'Pending, Approved, Rejected', 1], ['Amount', 'Amount', 'Number', '', 0], ['Requester', 'Requester', 'User', '', 1], ['Due date', 'DueDate', 'DateTime', '', 0]]},
        {n: 'Documents', t: '101', d: 'Board papers and reports', nav: 0, f: [['Report period', 'ReportPeriod', 'Text', '', 0]]}]}
  };
  var $ = function (s) { return root.querySelector(s); };
  var listsEl = $('#ss-lists'), out = $('.ss-out');
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  function xml(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  function fmt(s, a, b) { return s.split('%1').join(a).split('%2').join(b === undefined ? '' : b); }
  function isGuid(s) { return /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i.test(s); }
  function guid(seed) { // stable GUID from a string (FNV-1a over 4 lanes), version nibble 4, variant 8..b
    var h = [0x811c9dc5, 0x01000193, 0x2545f491, 0x9e3779b9];
    for (var i = 0; i < seed.length; i++) for (var k = 0; k < 4; k++) { h[k] ^= seed.charCodeAt(i) + k * 31; h[k] = Math.imul(h[k], 16777619) >>> 0; }
    var x = h.map(function (n) { return ('0000000' + n.toString(16)).slice(-8); }).join('');
    return x.slice(0, 8) + '-' + x.slice(8, 12) + '-4' + x.slice(13, 16) + '-' + '89ab'[parseInt(x[16], 16) & 3] + x.slice(17, 20) + '-' + x.slice(20, 32);
  }
  function autoName(d) { var s = d.replace(/[^A-Za-z0-9]/g, ''); if (/^[0-9]/.test(s)) s = 'F' + s; return s; }
  function opt(v, t, sel) { return '<option value="' + v + '"' + (v === sel ? ' selected' : '') + '>' + esc(t) + '</option>'; }
  function addField(box, f) {
    f = f || ['', '', 'Text', '', 0];
    var d = document.createElement('div'); d.className = 'ss-f';
    d.innerHTML = '<input class="f-dn" type="text" aria-label="' + T.dn + '" placeholder="' + T.dn + '" spellcheck="false">' +
      '<input class="f-in" type="text" aria-label="' + T.inn + '" placeholder="' + T.inn + '" spellcheck="false" autocomplete="off">' +
      '<select class="f-ty" aria-label="' + T.ty + '">' + FT.map(function (t) { return opt(t, T.ft[t], f[2]); }).join('') + '</select>' +
      '<input class="f-ex" type="text" aria-label="' + T.ex + '" placeholder="' + T.ex + '" spellcheck="false">' +
      '<label class="ss-c" title="' + T.req + '"><input class="f-rq" type="checkbox">' + T.rq + '</label>' +
      '<button type="button" class="q-x" aria-label="' + T.delf + '" title="' + T.delf + '">✕</button>';
    d.querySelector('.f-dn').value = f[0]; d.querySelector('.f-in').value = f[1]; d.querySelector('.f-ex').value = f[3]; d.querySelector('.f-rq').checked = !!f[4];
    d.querySelector('.q-x').addEventListener('click', function () { d.remove(); render(); });
    box.appendChild(d);
  }
  function addList(l) {
    l = l || {n: '', t: '100', d: '', nav: 0, f: [['', '', 'Text', '', 0]]};
    var d = document.createElement('fieldset'); d.className = 'q-fs ss-l';
    d.innerHTML = '<legend>' + T.list + '</legend><div class="ss-lh">' +
      '<label>' + T.lname + '<input class="l-n" type="text" spellcheck="false" autocomplete="off"></label>' +
      '<label>' + T.ltype + '<select class="l-t">' + ['100', '101', '107', '106'].map(function (t) { return opt(t, t + ' · ' + T.tt[t], l.t); }).join('') + '</select></label>' +
      '<label>' + T.ldesc + '<input class="l-d" type="text"></label>' +
      '<label class="ss-c"><input class="l-nav" type="checkbox">' + T.nav + '</label>' +
      '<button type="button" class="q-x" aria-label="' + T.dell + '" title="' + T.dell + '">✕</button></div>' +
      '<div class="ss-fh" aria-hidden="true"><span>' + T.dn + '</span><span>' + T.inn + '</span><span>' + T.ty + '</span><span>' + T.ex + '</span><span></span><span></span></div>' +
      '<div class="ss-fs"></div><button type="button" class="q-pre ss-addf">' + T.addf + '</button>';
    d.querySelector('.l-n').value = l.n; d.querySelector('.l-d').value = l.d; d.querySelector('.l-nav').checked = !!l.nav;
    var box = d.querySelector('.ss-fs'); l.f.forEach(function (f) { addField(box, f); });
    d.querySelector('.ss-addf').addEventListener('click', function () { addField(box); render(); });
    d.querySelector('.ss-lh .q-x').addEventListener('click', function () { d.remove(); render(); });
    listsEl.appendChild(d);
  }
  var S = ['title', 'tpl', 'theme', 'hub', 'locale', 'header', 'bg', 'logo', 'sharing', 'sol', 'soln'];
  function load(k) {
    var p = P[k];
    S.forEach(function (s) { $('#ss-' + s).value = p[s]; });
    listsEl.innerHTML = ''; p.lists.forEach(addList); render();
  }
  function read() {
    var st = {}; S.forEach(function (s) { st[s] = $('#ss-' + s).value.trim(); });
    st.lists = [].map.call(listsEl.querySelectorAll('.ss-l'), function (d) {
      return {n: d.querySelector('.l-n').value.trim(), t: d.querySelector('.l-t').value, d: d.querySelector('.l-d').value.trim(), nav: d.querySelector('.l-nav').checked,
        f: [].map.call(d.querySelectorAll('.ss-f'), function (r) {
          return {d: r.querySelector('.f-dn').value.trim(), i: r.querySelector('.f-in').value.trim(), t: r.querySelector('.f-ty').value, x: r.querySelector('.f-ex').value.trim(), r: r.querySelector('.f-rq').checked, el: r};
        }).filter(function (f) { return f.d; })};
    }).filter(function (l) { return l.n; });
    return st;
  }
  function build(st, H) {
    var A = [], created = {};
    if (st.theme) A.push({verb: 'applyTheme', themeName: st.theme});
    if (st.header || st.bg) { var b = {verb: 'setSiteBranding'}; if (st.header) b.headerLayout = st.header; if (st.bg) b.headerBackground = st.bg; A.push(b); }
    if (st.logo) A.push({verb: 'setSiteLogo', url: st.logo});
    if (st.locale) { A.push({verb: 'setRegionalSettings', locale: parseInt(st.locale, 10), sortOrder: /^(1049|1059)$/.test(st.locale) ? 7 : 25, hourFormat: st.locale === '1033' ? '12' : '24'}); H.push(T.h_reg); }
    if (st.sharing) A.push({verb: 'setSiteExternalSharingCapability', capability: st.sharing});
    st.lists.forEach(function (l) {
      var sub = [], seen = {};
      if (l.d) sub.push({verb: 'setDescription', description: l.d});
      l.f.forEach(function (f) {
        var inn = f.i || autoName(f.d);
        if (!f.i) { if (!inn) inn = 'Field' + (sub.length + 1); if (inn !== f.d) H.push(fmt(T.h_in, f.d, inn)); f.el.querySelector('.f-in').placeholder = inn; }
        if (seen[inn]) H.push(fmt(T.h_dup, l.n, inn)); seen[inn] = 1;
        var id = guid(l.n + '|' + inn), rq = f.r ? 'TRUE' : 'FALSE';
        if (f.t === 'Choice') {
          var ch = f.x.split(',').map(function (s) { return s.trim(); }).filter(Boolean);
          if (!ch.length) H.push(fmt(T.h_ch, f.d));
          sub.push({verb: 'addSPFieldXml', schemaXml: '<Field ID="{' + id + '}" Type="Choice" DisplayName="' + xml(f.d) + '" Required="' + rq + '" Format="Dropdown" StaticName="' + inn + '" Name="' + inn + '"><CHOICES>' + ch.map(function (c) { return '<CHOICE>' + xml(c) + '</CHOICE>'; }).join('') + '</CHOICES></Field>', addToDefaultView: true});
        } else if (f.t === 'Lookup') {
          if (!f.x) H.push(fmt(T.h_lk0, f.d)); else if (!created[f.x]) H.push(fmt(T.h_lk, f.d, f.x));
          sub.push({verb: 'addSPLookupFieldXml', schemaXml: '<Field ID="{' + id + '}" Type="Lookup" DisplayName="' + xml(f.d) + '" Required="' + rq + '" ShowField="Title" RelationshipDeleteBehavior="None" StaticName="' + inn + '" Name="' + inn + '" />', targetListName: f.x || 'TargetList', addToDefaultView: true});
        } else {
          sub.push({verb: 'addSPField', fieldType: f.t, displayName: f.d, internalName: inn, isRequired: f.r, id: id, addToDefaultView: true});
        }
      });
      A.push({verb: 'createSPList', listName: l.n, templateType: parseInt(l.t, 10), subactions: sub});
      created[l.n] = 1;
      if (/[^\x00-\x7f]/.test(l.n)) H.push(T.h_ru);
      if (l.nav) { A.push({verb: 'addNavLink', url: (l.t === '101' ? '/' : '/Lists/') + l.n, displayName: l.n, isWebRelative: true}); if (l.t === '101') H.push(fmt(T.h_lib, l.n)); }
    });
    if (st.hub) A.push({verb: 'joinHubSite', hubSiteId: st.hub.replace(/[{}]/g, '')});
    if (st.sol) A.push({verb: 'installSolution', id: st.sol.replace(/[{}]/g, ''), name: st.soln || 'SPFx solution'});
    return {$schema: 'https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json', actions: A};
  }
  function ps(st) {
    var t = (st.title || 'Site template').replace(/"/g, '`"'), tpl = st.tpl, pnpT = tpl === '68' ? 'CommunicationSite' : 'TeamSite';
    var spo = '# ' + (en ? 'once per tenant, as SharePoint admin' : 'один раз на клиент, от имени администратора SharePoint') + '\n' +
      'Connect-SPOService -Url https://contoso-admin.sharepoint.com\n' +
      '$json   = Get-Content ".\\site-script.json" -Raw -Encoding UTF8\n' +
      '$script = Add-SPOSiteScript -Title "' + t + '" -Content $json\n' +
      '$design = Add-SPOSiteDesign -Title "' + t + '" -WebTemplate "' + tpl + '" -SiteScripts $script.Id -Description "' + (en ? 'Lists, navigation and settings' : 'Списки, меню и настройки') + '"\n\n' +
      '# ' + (en ? 'apply to an existing site (asynchronous, up to 300 actions)' : 'применить к уже созданному сайту (асинхронно, до 300 действий)') + '\n' +
      'Add-SPOSiteDesignTask -SiteDesignId $design.Id -WebUrl https://contoso.sharepoint.com/sites/' + (tpl === '68' ? 'portal' : 'team');
    var pnp = 'Connect-PnPOnline -Url https://contoso-admin.sharepoint.com -Interactive -ClientId <app-id>\n' +
      '$script = Add-PnPSiteScript -Title "' + t + '" -Content (Get-Content ".\\site-script.json" -Raw -Encoding UTF8)\n' +
      '$design = Add-PnPSiteDesign -Title "' + t + '" -SiteScriptIds $script.Id -WebTemplate ' + pnpT + '\n' +
      'Invoke-PnPSiteDesign -Identity $design.Id -WebUrl https://contoso.sharepoint.com/sites/' + (tpl === '68' ? 'portal' : 'team');
    return [spo, pnp];
  }
  function count(A) { return A.reduce(function (n, a) { return n + 1 + (a.subactions ? a.subactions.length : 0); }, 0); }
  function preview(st, sc) {
    var A = sc.actions, navs = A.filter(function (a) { return a.verb === 'addNavLink'; });
    var site = '<div class="ss-site"><div class="ss-top"><b>' + esc(st.title || T.site) + '</b><span>' + esc(T.tpl[st.tpl]) + (st.theme ? ' · ' + esc(st.theme) : '') + '</span></div><div class="ss-body"><ul class="ss-nav">' +
      (navs.length ? navs.map(function (a) { return '<li>' + esc(a.displayName) + '</li>'; }).join('') : '<li class="ss-none">' + T.nonav + '</li>') + '</ul><div class="ss-cards">' +
      (st.lists.length ? st.lists.map(function (l) {
        return '<div class="ss-card"><b>' + (l.t === '101' ? '▤ ' : '☰ ') + esc(l.n) + '</b><small>' + esc(T.tt[l.t]) + '</small><ul>' + l.f.map(function (f) {
          return '<li><i aria-hidden="true">' + GLYPH[f.t] + '</i>' + esc(f.d) + (f.r ? ' *' : '') + (f.t === 'Lookup' && f.x ? ' → ' + esc(f.x) : '') + '</li>';
        }).join('') + '</ul></div>';
      }).join('') : '<p class="ss-none">' + T.nolists + '</p>') + '</div></div></div>';
    var seq = '<ol class="ss-seq">' + A.map(function (a) {
      var name = a.listName || a.displayName || a.themeName || a.capability || '';
      return '<li><code>' + a.verb + '</code> ' + esc(T.v[a.verb] || '') + (name ? ' · ' + esc(name) : '') + (a.subactions && a.subactions.length ? ' <small>+' + a.subactions.length + '</small>' : '') + '</li>';
    }).join('') + '</ol>';
    return '<div class="ss-prev"><small>' + T.prev + '</small>' + site + '<small>' + T.seq + '</small>' + seq + '</div>';
  }
  function blk(title, text) {
    return '<div class="q-blk"><div class="q-h"><small>' + title + '</small><button type="button" class="q-cp">' + T.copy + '</button></div><pre><code>' + esc(text) + '</code></pre></div>';
  }
  function render() {
    var st = read(), H = [];
    root.setAttribute('data-tpl', st.tpl);
    var sc = build(st, H), json = JSON.stringify(sc, null, 2), n = count(sc.actions), P2 = ps(st);
    var W = [T.h_paste, T.h_ids];
    if (n > 300 || json.length > 100000) W.push(T.h_300); else if (n > 30) W.push(T.h_30);
    if (st.logo && st.tpl !== '68') W.push(T.h_logo);
    if (st.theme) W.push(T.h_theme);
    if (st.hub && !isGuid(st.hub)) W.push(T.h_hub);
    if (st.sol && !isGuid(st.sol)) W.push(T.h_sol);
    H.forEach(function (h) { if (W.indexOf(h) < 0) W.push(h); });
    W.push(T.h_rerun);
    out.innerHTML = preview(st, sc) +
      '<p class="ss-cnt"><b>' + n + '</b> ' + T.acts + ' · <b>' + json.length.toLocaleString(en ? 'en-US' : 'ru-RU') + '</b> ' + T.chars + '</p>' +
      blk(T.json, json) + blk(T.spo, P2[0]) + blk(T.pnp, P2[1]) +
      '<div class="q-w"><small>' + T.warn + '</small><ul>' + W.map(function (h) { return '<li>' + esc(h) + '</li>'; }).join('') + '</ul></div>';
    var texts = [json, P2[0], P2[1]];
    [].forEach.call(out.querySelectorAll('.q-cp'), function (b, i) {
      b.addEventListener('click', function () {
        (navigator.clipboard ? navigator.clipboard.writeText(texts[i]) : Promise.reject()).then(function () { b.textContent = T.copied; setTimeout(function () { b.textContent = T.copy; }, 1400); }, function () {});
      });
    });
  }
  root.addEventListener('input', render); root.addEventListener('change', render);
  $('#ss-addl').addEventListener('click', function () { addList(); render(); });
  [].forEach.call(root.querySelectorAll('[data-preset]'), function (b) { b.addEventListener('click', function () { load(b.getAttribute('data-preset')); }); });
  load('halden');
})();
