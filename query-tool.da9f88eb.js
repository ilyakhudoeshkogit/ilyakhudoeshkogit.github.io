/* SharePoint list query builder: REST URL, PnPjs and Microsoft Graph from one form. Runs locally, sends nothing. */
(function () {
  var root = document.getElementById('spq'); if (!root) return;
  var en = document.documentElement.lang === 'en';
  var T = en ? {
    rest: 'SharePoint REST', pnp: 'PnPjs', graph: 'Microsoft Graph', copy: 'Copy', copied: 'Copied', and: 'and', or: 'or', del: 'Remove condition',
    w_exp: 'Auto-added to $expand: ', w_q: 'A quote inside a text value is doubled: \'\' (that is how OData escapes it).',
    w_top: '$top above 5000 does not work: one request returns at most 5000 items. Page through with @odata.nextLink or items.getAll() in PnPjs.',
    w_thr: 'If the list has more than 5000 items, the columns in $filter and $orderby must be indexed, otherwise the query fails with the list view threshold error.',
    w_lk: 'To filter by a lookup or person, it is cheaper to compare the id: ', w_sub: 'Graph has no substringof (contains) for list item fields: filter by startswith or on the client.',
    w_glk: 'In Graph a lookup or person column comes as ', w_glk2: ' and cannot be expanded to its Title as in REST: read the target list separately.',
    w_hdr: 'Graph filters and sorts by a non-indexed column only with the header Prefer: HonorNonIndexedQueriesWarningMayFailRandomly. Better to index the column.',
    w_none: 'No remarks.', warn: 'What to keep in mind', ph: 'Field'
  } : {
    rest: 'SharePoint REST', pnp: 'PnPjs', graph: 'Microsoft Graph', copy: 'Копировать', copied: 'Скопировано', and: 'и', or: 'или', del: 'Убрать условие',
    w_exp: 'Сам добавил в $expand: ', w_q: 'Кавычка внутри текстового значения удвоена: \'\' (так OData экранирует апостроф).',
    w_top: '$top больше 5000 не сработает: за один запрос приходит не больше 5000 элементов. Листайте по @odata.nextLink или берите items.getAll() в PnPjs.',
    w_thr: 'Если в списке больше 5000 элементов, столбцы из $filter и $orderby должны быть проиндексированы, иначе запрос упадёт с ошибкой порога представления списка.',
    w_lk: 'Фильтровать по подстановке или пользователю дешевле по id: ', w_sub: 'В Graph нет substringof (contains) для полей элементов списка: фильтруйте через startswith или уже на клиенте.',
    w_glk: 'В Graph поле подстановки или пользователя приходит как ', w_glk2: ' и не раскрывается до Title, как в REST: целевой список читается отдельно.',
    w_hdr: 'Graph фильтрует и сортирует по неиндексированному столбцу только с заголовком Prefer: HonorNonIndexedQueriesWarningMayFailRandomly. Лучше проиндексировать столбец.',
    w_none: 'Замечаний нет.', warn: 'О чём помнить', ph: 'Поле'
  };
  var P = {
    approvals: {list: 'Approvals', select: 'Id, Title, ApprovalType, Status, AssignedTo/Title, AssignedTo/EMail', conds: [['Status', 'eq', 'Pending', 's']], order: 'Created', desc: true, top: '50'},
    reservations: {list: 'Reservations', select: 'Id, DeskId, Desk/Title, ReservedDate, Author/Title', conds: [['ReservedDate', 'ge', '2026-11-02', 'd'], ['ReservedDate', 'lt', '2026-11-03', 'd']], order: 'ReservedDate', desc: false, top: '200'},
    projects: {list: 'Projects', select: 'Title, Status, Budget, Owner/Title', conds: [['Status', 'ne', 'Closed', 's'], ['Budget', 'gt', '10000', 'n']], order: 'Budget', desc: true, top: '20'}
  };
  var $ = function (s) { return root.querySelector(s); };
  var list = $('#q-list'), sel = $('#q-select'), conds = $('#q-conds'), join = $('#q-join'), order = $('#q-order'), dir = $('#q-dir'), top = $('#q-top'), out = $('.q-out');
  var OPS = ['eq', 'ne', 'gt', 'ge', 'lt', 'le', 'startswith', 'substringof'];
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  function addCond(c) {
    c = c || ['', 'eq', '', 's'];
    var d = document.createElement('div'); d.className = 'q-cond';
    d.innerHTML = '<input class="q-f" type="text" aria-label="' + T.ph + '" placeholder="' + T.ph + '" spellcheck="false">' +
      '<select class="q-op" aria-label="op">' + OPS.map(function (o) { return '<option>' + o + '</option>'; }).join('') + '</select>' +
      '<input class="q-v" type="text" aria-label="value" spellcheck="false">' +
      '<select class="q-t" aria-label="type"><option value="s">' + (en ? 'text' : 'текст') + '</option><option value="n">' + (en ? 'number' : 'число') + '</option><option value="d">' + (en ? 'date' : 'дата') + '</option></select>' +
      '<button type="button" class="q-x" aria-label="' + T.del + '" title="' + T.del + '">✕</button>';
    d.querySelector('.q-f').value = c[0]; d.querySelector('.q-op').value = c[1]; d.querySelector('.q-v').value = c[2]; d.querySelector('.q-t').value = c[3];
    d.querySelector('.q-x').addEventListener('click', function () { d.remove(); render(); });
    conds.appendChild(d);
  }
  function load(k) {
    var p = P[k]; list.value = p.list; sel.value = p.select; order.value = p.order; dir.value = p.desc ? 'desc' : 'asc'; top.value = p.top; join.value = 'and';
    conds.innerHTML = ''; p.conds.forEach(addCond); render();
  }
  function fields(s) { return s.split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
  function q(s) { return "'" + String(s).replace(/'/g, "''") + "'"; }
  function iso(v) { return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v + 'T00:00:00Z' : v; }
  function val(v, t, graph) {
    if (t === 'n') return String(parseFloat(v) || 0);
    if (t === 'd') return graph ? q(iso(v)) : "datetime'" + iso(v).replace(/'/g, '') + "'";
    return q(v);
  }
  function cond(c, graph) {
    var f = graph ? 'fields/' + c.f : c.f, v = val(c.v, c.t, graph);
    if (c.op === 'startswith') return 'startswith(' + f + ',' + v + ')';
    if (c.op === 'substringof') return 'substringof(' + v + ',' + f + ')';
    return f + ' ' + c.op + ' ' + v;
  }
  function render() {
    var L = list.value.trim() || 'List', F = fields(sel.value), W = [], j = ' ' + join.value + ' ';
    var C = [].map.call(conds.querySelectorAll('.q-cond'), function (d) {
      return {f: d.querySelector('.q-f').value.trim(), op: d.querySelector('.q-op').value, v: d.querySelector('.q-v').value, t: d.querySelector('.q-t').value};
    }).filter(function (c) { return c.f; });
    var ob = order.value.trim(), desc = dir.value === 'desc', n = parseInt(top.value, 10) || 0;
    var exp = [];
    F.concat(C.map(function (c) { return c.f; })).concat(ob ? [ob] : []).forEach(function (f) { var i = f.indexOf('/'); if (i > 0 && exp.indexOf(f.slice(0, i)) < 0) exp.push(f.slice(0, i)); });
    var filt = C.map(function (c) { return cond(c, false); }).join(j);
    // REST
    var qs = [];
    if (F.length) qs.push('$select=' + F.join(','));
    if (exp.length) qs.push('$expand=' + exp.join(','));
    if (filt) qs.push('$filter=' + filt);
    if (ob) qs.push('$orderby=' + ob + (desc ? ' desc' : ''));
    if (n) qs.push('$top=' + n);
    var rest = "/_api/web/lists/getbytitle(" + q(L) + ")/items" + (qs.length ? '?' + qs.join('&') : '');
    // PnPjs
    var pnp = 'const items = await sp.web.lists\n  .getByTitle(' + JSON.stringify(L) + ').items';
    if (F.length) pnp += '\n  .select(' + F.map(function (f) { return JSON.stringify(f); }).join(', ') + ')';
    if (exp.length) pnp += '\n  .expand(' + exp.map(function (f) { return JSON.stringify(f); }).join(', ') + ')';
    if (filt) pnp += '\n  .filter(' + JSON.stringify(filt) + ')';
    if (ob) pnp += '\n  .orderBy(' + JSON.stringify(ob) + (desc ? ', false' : '') + ')';
    if (n) pnp += '\n  .top(' + n + ')';
    pnp += '();';
    // Graph
    var gF = F.filter(function (f) { return f.indexOf('/') < 0 && f !== 'Id'; });
    var gs = ['$expand=fields' + (gF.length ? '($select=' + gF.join(',') + ')' : '')];
    var gC = C.filter(function (c) { return c.f.indexOf('/') < 0; });
    if (gC.length) gs.push('$filter=' + gC.map(function (c) { return cond(c, true); }).join(j));
    if (ob && ob.indexOf('/') < 0) gs.push('$orderby=fields/' + ob + (desc ? ' desc' : ''));
    if (n) gs.push('$top=' + n);
    var graph = 'GET /sites/{site-id}/lists/' + L + '/items?' + gs.join('&');
    if (gC.length || ob) graph += '\nPrefer: HonorNonIndexedQueriesWarningMayFailRandomly';
    // remarks
    if (exp.length) W.push(T.w_exp + exp.join(', ') + '.');
    if (C.some(function (c) { return c.t === 's' && c.v.indexOf("'") >= 0; }) || L.indexOf("'") >= 0) W.push(T.w_q);
    if (n > 5000) W.push(T.w_top);
    if (C.length || ob) W.push(T.w_thr);
    var lk = C.filter(function (c) { return c.f.indexOf('/') > 0; });
    if (lk.length) W.push(T.w_lk + lk.map(function (c) { return c.f.split('/')[0] + 'Id eq 5'; }).join(', ') + '.');
    if (C.some(function (c) { return c.op === 'substringof'; })) W.push(T.w_sub);
    if (exp.length) W.push(T.w_glk + exp.map(function (e) { return e + 'LookupId'; }).join(', ') + T.w_glk2);
    if (gC.length || ob) W.push(T.w_hdr);
    out.innerHTML = [[T.rest, rest], [T.pnp, pnp], [T.graph, graph]].map(function (b) {
      return '<div class="q-blk"><div class="q-h"><small>' + b[0] + '</small><button type="button" class="q-cp">' + T.copy + '</button></div><pre><code>' + esc(b[1]) + '</code></pre></div>';
    }).join('') + '<div class="q-w"><small>' + T.warn + '</small><ul>' + (W.length ? W : [T.w_none]).map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('') + '</ul></div>';
    [].forEach.call(out.querySelectorAll('.q-cp'), function (b) {
      b.addEventListener('click', function () {
        var t = b.parentNode.nextSibling.textContent;
        (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(function () { b.textContent = T.copied; setTimeout(function () { b.textContent = T.copy; }, 1400); }, function () {});
      });
    });
  }
  root.addEventListener('input', render); root.addEventListener('change', render);
  $('#q-add').addEventListener('click', function () { addCond(); render(); });
  [].forEach.call(root.querySelectorAll('[data-preset]'), function (b) { b.addEventListener('click', function () { load(b.getAttribute('data-preset')); }); });
  load('approvals');
})();
