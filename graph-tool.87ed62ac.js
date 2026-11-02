/* SPFx Microsoft Graph permissions picker: least-privileged delegated scopes + package-solution.json. Runs locally, sends nothing. */
(function () {
  var root = document.getElementById('sgp'); if (!root) return;
  var en = document.documentElement.lang === 'en';
  // id: [ru, en, method, endpoint, scope, note key]
  var A = [
    ['me', 'Имя, почта и должность текущего пользователя', 'Name, email and job title of the current user', 'GET', "/me?$select=displayName,mail,jobTitle", 'User.Read'],
    ['memberof', 'В каких группах состоит пользователь (показать блок только своим)', 'Which groups the user is in (show a block to members only)', 'GET', '/me/memberOf', 'User.Read'],
    ['photo', 'Фото коллег (аватары)', 'Colleagues\' photos (avatars)', 'GET', '/users/{id}/photo/$value', 'ProfilePhoto.Read.All'],
    ['users', 'Поиск коллег по имени (имя, почта)', 'Find colleagues by name (name, email)', 'GET', "/users?$filter=startswith(displayName,'Iv')&$select=displayName,mail", 'User.ReadBasic.All'],
    ['userfull', 'Отдел, должность и руководитель коллеги', 'A colleague\'s department, job title and manager', 'GET', '/users/{id}/manager', 'User.Read.All'],
    ['people', 'С кем пользователь чаще всего работает', 'People the user works with most', 'GET', '/me/people?$top=10', 'People.Read'],
    ['presence', 'Статус коллег в Teams (в сети, занят)', 'Colleagues\' Teams presence (available, busy)', 'POST', '/communications/getPresencesByUserId', 'Presence.Read.All'],
    ['calendar', 'Ближайшие встречи пользователя', 'The user\'s upcoming meetings', 'GET', '/me/calendarView?startDateTime={from}&endDateTime={to}', 'Calendars.ReadBasic'],
    ['mail', 'Непрочитанные письма (тема и отправитель)', 'Unread mail (subject and sender)', 'GET', "/me/mailFolders/inbox/messages?$filter=isRead eq false&$top=5", 'Mail.ReadBasic'],
    ['sendmail', 'Отправить письмо от имени пользователя', 'Send an email as the user', 'POST', '/me/sendMail', 'Mail.Send'],
    ['files', 'Недавние файлы пользователя в OneDrive', 'The user\'s recent OneDrive files', 'GET', '/me/drive/recent', 'Files.Read'],
    ['sites', 'Списки на ДРУГИХ сайтах и поиск сайтов', 'Lists on OTHER sites and site search', 'GET', '/sites/{site-id}/lists/{list-id}/items?$expand=fields', 'Sites.Read.All'],
    ['siteswrite', 'Создавать и менять элементы списков через Graph', 'Create and update list items through Graph', 'POST', '/sites/{site-id}/lists/{list-id}/items', 'Sites.ReadWrite.All'],
    ['teams', 'Команды Teams пользователя', 'The user\'s Teams teams', 'GET', '/me/joinedTeams', 'Team.ReadBasic.All'],
    ['chread', 'Читать сообщения канала Teams', 'Read Teams channel messages', 'GET', '/teams/{team-id}/channels/{channel-id}/messages', 'ChannelMessage.Read.All'],
    ['chsend', 'Написать в канал Teams', 'Post to a Teams channel', 'POST', '/teams/{team-id}/channels/{channel-id}/messages', 'ChannelMessage.Send'],
    ['planner', 'Задачи Planner пользователя', 'The user\'s Planner tasks', 'GET', '/me/planner/tasks', 'Tasks.Read'],
    ['members', 'Участники группы Microsoft 365', 'Members of a Microsoft 365 group', 'GET', '/groups/{group-id}/members', 'GroupMember.ReadBasic.All']
  ];
  // scope -> scopes from the same docs that already cover it (if one is requested, the narrower one is dropped)
  var COVER = {
    'ProfilePhoto.Read.All': ['User.ReadBasic.All', 'User.Read.All'],
    'User.ReadBasic.All': ['User.Read.All'],
    'Calendars.ReadBasic': [], 'Mail.ReadBasic': [],
    'Files.Read': ['Sites.Read.All', 'Sites.ReadWrite.All'],
    'Sites.Read.All': ['Sites.ReadWrite.All']
  };
  var SENS = {'Mail.ReadBasic': 1, 'Mail.Send': 1, 'Sites.ReadWrite.All': 1, 'ChannelMessage.Read.All': 1, 'ChannelMessage.Send': 1, 'User.Read.All': 1, 'Files.Read': 1, 'Calendars.ReadBasic': 1};
  var T = en ? {
    pick: 'What should the web part do?', none: 'Tick at least one task above.', tbl: 'Tasks, endpoints and scopes', task: 'Task', ep: 'Graph request', sc: 'Scope',
    pkg: 'config/package-solution.json (inside "solution")', code: 'Calling Graph from the web part (MSGraphClientV3)', ps: 'After deploying the .sppkg: approve (SharePoint Online Management Shell)',
    copy: 'Copy', copied: 'Copied', warn: 'What to keep in mind', covered: 'covered by ',
    h_shared: 'Grants go to one shared principal, SharePoint Online Client Extensibility Web Application Principal, not to your solution. Once approved, any SPFx code in the tenant can call Graph with these scopes on behalf of the user.',
    h_admin: 'Nobody can consent for themselves: after deploying to the app catalog the request waits on the API access page of the SharePoint admin center until an administrator approves it.',
    h_deleg: 'All scopes are delegated: the web part acts as the signed-in user and never sees more than that user can see.',
    h_local: 'Lists on the CURRENT site do not need Graph at all: SPHttpClient or PnPjs work with the user\'s own rights and need no approval.',
    h_sens: 'Sensitive scopes selected: ', h_sens2: '. Because the grant is shared, they open this data to every SPFx solution in the tenant. Ask for them only if the task really needs it.',
    h_basic: 'Calendars.ReadBasic and Mail.ReadBasic return no body or attachments. If you need the meeting agenda or message text, that is Calendars.Read or Mail.Read.',
    h_rw: 'Sites.ReadWrite.All lets the web part write to any site the user can edit. If writing is only needed on the current site, do it through SPHttpClient or PnPjs instead.',
    h_user: 'User.Read.All replaces User.ReadBasic.All and photo access: one scope instead of three.',
    h_dup: 'Approving a scope that is already granted gives an error: check the current grants first (Get-SPOTenantServicePrincipalPermissionGrants).'
  } : {
    pick: 'Что должна делать веб-часть?', none: 'Отметьте хотя бы одну задачу выше.', tbl: 'Задачи, запросы и разрешения', task: 'Задача', ep: 'Запрос к Graph', sc: 'Разрешение',
    pkg: 'config/package-solution.json (внутри "solution")', code: 'Вызов Graph из веб-части (MSGraphClientV3)', ps: 'После развёртывания .sppkg: одобрить (SharePoint Online Management Shell)',
    copy: 'Копировать', copied: 'Скопировано', warn: 'О чём помнить', covered: 'покрыто ',
    h_shared: 'Разрешения выдаются одному общему принципалу SharePoint Online Client Extensibility Web Application Principal, а не вашему решению. После одобрения любой код SPFx в тенанте сможет вызывать Graph с этими правами от имени пользователя.',
    h_admin: 'Сам пользователь согласие дать не может: после загрузки в каталог приложений запрос ждёт на странице «Доступ к API» в центре администрирования SharePoint, пока его не одобрит администратор.',
    h_deleg: 'Все разрешения делегированные: веб-часть действует от имени вошедшего пользователя и не увидит больше, чем видит он сам.',
    h_local: 'Для списков ТЕКУЩЕГО сайта Graph не нужен вовсе: SPHttpClient или PnPjs работают с правами пользователя и одобрения не требуют.',
    h_sens: 'Выбраны чувствительные разрешения: ', h_sens2: '. Раз выдача общая, они открывают эти данные всем решениям SPFx в тенанте. Просите их, только если задача без них никак.',
    h_basic: 'Calendars.ReadBasic и Mail.ReadBasic не отдают текст и вложения. Если нужна повестка встречи или текст письма, это уже Calendars.Read или Mail.Read.',
    h_rw: 'Sites.ReadWrite.All даёт писать на любой сайт, который пользователь может редактировать. Если писать нужно только на текущем сайте, делайте это через SPHttpClient или PnPjs.',
    h_user: 'User.Read.All заменяет User.ReadBasic.All и доступ к фото: одно разрешение вместо трёх.',
    h_dup: 'Одобрение уже выданного разрешения заканчивается ошибкой: сначала посмотрите текущие (Get-SPOTenantServicePrincipalPermissionGrants).'
  };
  var $ = function (s) { return root.querySelector(s); };
  var list = $('.g-list'), out = $('.g-out');
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  list.innerHTML = A.map(function (a) {
    return '<label class="g-it"><input type="checkbox" value="' + a[0] + '"><span>' + esc(en ? a[2] : a[1]) + ' <code>' + a[5] + '</code></span></label>';
  }).join('');
  function blk(title, text) {
    return '<div class="q-blk"><div class="q-h"><small>' + title + '</small><button type="button" class="q-cp">' + T.copy + '</button></div><pre><code>' + esc(text) + '</code></pre></div>';
  }
  function codeFor(a) {
    var ep = a[4].split('?')[0], q = a[4].indexOf('?') > 0 ? a[4].split('?')[1] : '';
    var lines = ["import { MSGraphClientV3 } from '@microsoft/sp-http';", '',
      "const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');"];
    var chain = "client.api('" + ep + "')";
    q.split('&').forEach(function (p) {
      var kv = p.split('='); if (!kv[0]) return;
      var k = kv[0].replace('$', ''), v = kv.slice(1).join('=');
      if (k === 'select') chain += ".select('" + v + "')";
      else if (k === 'filter') chain += '.filter("' + v + '")';
      else if (k === 'top') chain += '.top(' + v + ')';
      else if (k === 'expand') chain += ".expand('" + v + "')";
      else chain += ".query({ " + k + ": '" + v + "' })";
    });
    if (a[3] === 'POST') {
      var body = {sendmail: "{ message: { subject: 'Hi', body: { contentType: 'Text', content: '...' }, toRecipients: [{ emailAddress: { address: 'user@contoso.com' } }] } }",
        presence: "{ ids: [userId1, userId2] }", chsend: "{ body: { content: 'Hello from SPFx' } }", siteswrite: "{ fields: { Title: 'New item' } }"}[a[0]] || '{}';
      lines.push('const res = await ' + chain + '.post(' + body + ');');
    } else if (/\$value$/.test(ep)) {
      lines.push('const blob: Blob = await ' + chain + ".responseType('blob').get();", 'const src = URL.createObjectURL(blob);');
    } else lines.push('const res = await ' + chain + '.get();');
    return lines.join('\n');
  }
  function render() {
    var on = [].filter.call(list.querySelectorAll('input'), function (i) { return i.checked; }).map(function (i) { return i.value; });
    var sel = A.filter(function (a) { return on.indexOf(a[0]) >= 0; });
    if (!sel.length) { out.innerHTML = '<p class="g-none">' + T.none + '</p>'; return; }
    var scopes = []; sel.forEach(function (a) { if (scopes.indexOf(a[5]) < 0) scopes.push(a[5]); });
    var by = {};
    scopes.forEach(function (s) { (COVER[s] || []).forEach(function (c) { if (!by[s] && scopes.indexOf(c) >= 0) by[s] = c; }); });
    var final = scopes.filter(function (s) { return !by[s]; });
    var pkg = '"webApiPermissionRequests": ' + JSON.stringify(final.map(function (s) { return {resource: 'Microsoft Graph', scope: s}; }), null, 2);
    var ps = ['Connect-SPOService -Url https://contoso-admin.sharepoint.com', 'Get-SPOTenantServicePrincipalPermissionGrants',
      'Get-SPOTenantServicePrincipalPermissionRequests', '# ' + (en ? 'for each request of this solution:' : 'для каждого запроса этого решения:'),
      'Approve-SPOTenantServicePrincipalPermissionRequest -RequestId <Guid>'].join('\n');
    var rows = sel.map(function (a) {
      var s = a[5], c = by[s] ? ' <small class="g-cov">(' + T.covered + by[s] + ')</small>' : '';
      return '<tr><td>' + esc(en ? a[2] : a[1]) + '</td><td><code>' + a[3] + ' ' + esc(a[4]) + '</code></td><td><code>' + s + '</code>' + c + '</td></tr>';
    }).join('');
    var H = [T.h_shared, T.h_admin, T.h_deleg];
    var sens = final.filter(function (s) { return SENS[s]; });
    if (sens.length) H.push(T.h_sens + sens.join(', ') + T.h_sens2);
    if (final.indexOf('Calendars.ReadBasic') >= 0 || final.indexOf('Mail.ReadBasic') >= 0) H.push(T.h_basic);
    if (final.indexOf('Sites.ReadWrite.All') >= 0) H.push(T.h_rw);
    if (final.indexOf('User.Read.All') >= 0 && (by['User.ReadBasic.All'] || by['ProfilePhoto.Read.All'])) H.push(T.h_user);
    if (on.indexOf('sites') >= 0 || on.indexOf('siteswrite') >= 0) H.push(T.h_local);
    H.push(T.h_dup);
    out.innerHTML = '<div class="g-tw"><small>' + T.tbl + '</small><table class="g-t"><thead><tr><th>' + T.task + '</th><th>' + T.ep + '</th><th>' + T.sc + '</th></tr></thead><tbody>' + rows + '</tbody></table></div>' +
      blk(T.pkg, pkg) + blk(T.code, codeFor(sel[0])) + blk(T.ps, ps) +
      '<div class="q-w"><small>' + T.warn + '</small><ul>' + H.map(function (h) { return '<li>' + esc(h) + '</li>'; }).join('') + '</ul></div>';
    [].forEach.call(out.querySelectorAll('.q-cp'), function (b) {
      b.addEventListener('click', function () {
        var txt = b.parentNode.nextSibling.textContent;
        (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).then(function () { b.textContent = T.copied; setTimeout(function () { b.textContent = T.copy; }, 1400); }, function () {});
      });
    });
  }
  var P = {profile: ['me', 'photo'], day: ['calendar', 'mail', 'planner'], dir: ['users', 'photo', 'userfull', 'presence'], teams: ['teams', 'chsend']};
  function load(k) { [].forEach.call(list.querySelectorAll('input'), function (i) { i.checked = P[k].indexOf(i.value) >= 0; }); render(); }
  list.addEventListener('change', render);
  [].forEach.call(root.querySelectorAll('[data-preset]'), function (b) { b.addEventListener('click', function () { load(b.getAttribute('data-preset')); }); });
  load('profile');
})();
