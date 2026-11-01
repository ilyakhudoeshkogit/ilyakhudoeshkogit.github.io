/* SharePoint column formatting builder: JSON + live preview. Runs locally, sends nothing. */
(function () {
  var root = document.getElementById('scf'); if (!root) return;
  var en = document.documentElement.lang === 'en';
  var T = en ? {
    lv: {good: 'good', low: 'low', warning: 'warning', severeWarning: 'severe warning', blocked: 'blocked'},
    val: 'Value', del: 'Remove rule', prev: 'Preview (sample rows)', json: 'Column formatting JSON', copy: 'Copy', copied: 'Copied', warn: 'What to keep in mind',
    other: 'Other value', empty: '(empty)', days: ' d.',
    h_paste: 'Where to paste: column header menu, Column settings, Format this column, Advanced mode. Paste the JSON and press Save.',
    h_view: 'Formatting only changes how the value looks. It does not change the data and does not hide it: anyone with access still sees the value in the item form and through the API.',
    h_case: 'Values are compared exactly, case included: "approved" will not match "Approved". Copy them from the Choice column settings.',
    h_q: 'A value contains an apostrophe: it breaks the expression string. Rename the choice or use a value without it.',
    h_now: '@now is the time on the viewer\'s computer, so a row turns red at local midnight, not by server time.',
    h_num: 'The column must be a Number or Currency column. A Calculated column also works if its result type is Number.',
    h_neg: 'Negative values give a zero-width bar; values above the maximum are drawn at 100%.',
    h_dt: 'The column must be Date and Time. An empty date is left uncoloured thanks to the Number(@currentField) == 0 check.',
    h_dup: 'The same value appears twice: only the first rule applies.'
  } : {
    lv: {good: 'хорошо', low: 'низкий', warning: 'внимание', severeWarning: 'серьёзно', blocked: 'блок'},
    val: 'Значение', del: 'Убрать правило', prev: 'Предпросмотр (тестовые строки)', json: 'JSON для форматирования столбца', copy: 'Копировать', copied: 'Скопировано', warn: 'О чём помнить',
    other: 'Другое значение', empty: '(пусто)', days: ' дн.',
    h_paste: 'Куда вставить: меню заголовка столбца, «Параметры столбца», «Форматировать этот столбец», «Расширенный режим». Вставьте JSON и нажмите «Сохранить».',
    h_view: 'Форматирование меняет только вид значения. Данные оно не меняет и не прячет: у кого есть доступ, тот видит значение в форме элемента и через API.',
    h_case: 'Значения сравниваются точно, с учётом регистра: «approved» не совпадёт с «Approved». Копируйте их из настроек столбца «Выбор».',
    h_q: 'В значении есть апостроф: он ломает строку выражения. Переименуйте вариант или возьмите значение без него.',
    h_now: '@now это время на компьютере того, кто смотрит, поэтому строка краснеет в полночь по его часам, а не по серверу.',
    h_num: 'Столбец должен быть «Число» или «Денежный». Подойдёт и вычисляемый, если тип результата «Число».',
    h_neg: 'Отрицательное значение даёт полосу нулевой ширины, всё больше максимума рисуется на 100%.',
    h_dt: 'Столбец должен быть «Дата и время». Пустая дата не красится благодаря проверке Number(@currentField) == 0.',
    h_dup: 'Одно и то же значение встречается дважды: сработает только первое правило.'
  };
  var LV = ['good', 'low', 'warning', 'severeWarning', 'blocked'];
  var ICON = {good: 'CheckMark', low: 'Forward', warning: 'Warning', severeWarning: 'Error', blocked: 'ErrorBadge'};
  var GLYPH = {good: '✓', low: '→', warning: '!', severeWarning: '⚠', blocked: '✕'};
  var P = {
    approvals: {mode: 'status', field: 'Status', rules: [['Approved', 'good'], ['Pending', 'warning'], ['Rejected', 'blocked']]},
    projects: {mode: 'status', field: 'Health', rules: [['On track', 'good'], ['At risk', 'warning'], ['Late', 'severeWarning'], ['Closed', 'low']]},
    budget: {mode: 'bar', field: 'Budget', max: '50000'},
    due: {mode: 'due', field: 'DueDate', days: '3'}
  };
  var $ = function (s) { return root.querySelector(s); };
  var mode = $('#c-mode'), field = $('#c-field'), rules = $('#c-rules'), max = $('#c-max'), days = $('#c-days'), out = $('.c-out');
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  function addRule(r) {
    r = r || ['', 'good'];
    var d = document.createElement('div'); d.className = 'c-rule';
    d.innerHTML = '<input class="c-v" type="text" aria-label="' + T.val + '" placeholder="' + T.val + '" spellcheck="false">' +
      '<select class="c-l" aria-label="level">' + LV.map(function (l) { return '<option value="' + l + '">' + T.lv[l] + '</option>'; }).join('') + '</select>' +
      '<button type="button" class="q-x" aria-label="' + T.del + '" title="' + T.del + '">✕</button>';
    d.querySelector('.c-v').value = r[0]; d.querySelector('.c-l').value = r[1];
    d.querySelector('.q-x').addEventListener('click', function () { d.remove(); render(); });
    rules.appendChild(d);
  }
  function load(k) {
    var p = P[k]; mode.value = p.mode; field.value = p.field;
    rules.innerHTML = ''; (p.rules || [['Approved', 'good'], ['Pending', 'warning']]).forEach(addRule);
    max.value = p.max || '100'; days.value = p.days || '3'; render();
  }
  function sq(s) { return "'" + s + "'"; }
  function chain(R, f, last) { // nested if(...) over rules
    var e = sq(last);
    for (var i = R.length - 1; i >= 0; i--) e = 'if(@currentField == ' + sq(R[i].v) + ', ' + sq(f(R[i])) + ', ' + e + ')';
    return '=' + e;
  }
  var SCHEMA = 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json';
  function build(m, R, mx, nd) {
    if (m === 'status') return {
      $schema: SCHEMA, elmType: 'div',
      attributes: {'class': chain(R, function (r) { return 'sp-field-severity--' + r.l; }, '') + " + ' ms-fontColor-neutralSecondary'"},
      children: [
        {elmType: 'span', style: {display: 'inline-block', padding: '0 4px'}, attributes: {iconName: chain(R, function (r) { return ICON[r.l]; }, '')}},
        {elmType: 'span', txtContent: '@currentField'}
      ]
    };
    if (m === 'bar') return {
      $schema: SCHEMA, elmType: 'div', txtContent: '@currentField', attributes: {'class': 'sp-field-dataBars'},
      style: {width: '=if(@currentField > ' + mx + ", '100%', if(@currentField < 0, '0%', toString(@currentField * 100 / " + mx + ") + '%'))"}
    };
    return {
      $schema: SCHEMA, elmType: 'div', txtContent: '@currentField',
      attributes: {'class': "=if(Number(@currentField) == 0, '', if(@currentField <= @now, 'sp-field-severity--severeWarning', if(@currentField <= @now + " + (nd * 86400000) + ", 'sp-field-severity--warning', '')))"}
    };
  }
  function cell(cls, glyph, text) { return '<span class="c-cell ' + cls + '">' + (glyph ? '<i aria-hidden="true">' + glyph + '</i>' : '') + esc(text) + '</span>'; }
  function fmtDate(d) { return en ? d.toLocaleDateString('en-GB') : d.toLocaleDateString('ru-RU'); }
  function preview(m, R, mx, nd, f) {
    var rows = [];
    if (m === 'status') {
      R.forEach(function (r) { rows.push(cell('sev-' + r.l, GLYPH[r.l], r.v)); });
      rows.push(cell('sev-none', '', T.other));
    } else if (m === 'bar') {
      [0, .3, .75, 1.2].forEach(function (k) {
        var v = Math.round(mx * k), w = Math.max(0, Math.min(100, v * 100 / mx));
        rows.push('<span class="c-cell c-barw"><span class="c-bar" style="width:' + w + '%"></span><span class="c-bt">' + v.toLocaleString(en ? 'en-US' : 'ru-RU') + '</span></span>');
      });
    } else {
      var now = new Date();
      [-3, Math.max(0, Math.floor(nd / 2)), nd + 10].forEach(function (k) {
        var d = new Date(now.getTime() + k * 86400000 + (k > 0 ? 0 : -60000));
        var cls = d <= now ? 'sev-severeWarning' : (d <= new Date(now.getTime() + nd * 86400000) ? 'sev-warning' : 'sev-none');
        rows.push(cell(cls, '', fmtDate(d) + ' (' + (k > 0 ? '+' : '') + k + T.days + ')'));
      });
      rows.push(cell('sev-none', '', T.empty));
    }
    return '<div class="c-prev"><small>' + T.prev + '</small><table><thead><tr><th>Title</th><th>' + esc(f) + '</th></tr></thead><tbody>' +
      rows.map(function (r, i) { return '<tr><td>' + (en ? 'Item ' : 'Элемент ') + (i + 1) + '</td><td>' + r + '</td></tr>'; }).join('') + '</tbody></table></div>';
  }
  function render() {
    var m = mode.value, f = field.value.trim() || 'Status';
    root.setAttribute('data-mode', m);
    var R = [].map.call(rules.querySelectorAll('.c-rule'), function (d) { return {v: d.querySelector('.c-v').value.trim(), l: d.querySelector('.c-l').value}; }).filter(function (r) { return r.v; });
    var mx = parseFloat(max.value); if (!(mx > 0)) mx = 100;
    var nd = parseInt(days.value, 10); if (!(nd >= 0)) nd = 0;
    var json = JSON.stringify(build(m, R, mx, nd), null, 2);
    var H = [T.h_paste];
    if (m === 'status') {
      H.push(T.h_case);
      if (R.some(function (r) { return r.v.indexOf("'") >= 0; })) H.push(T.h_q);
      var seen = {}; if (R.some(function (r) { var d = seen[r.v]; seen[r.v] = 1; return d; })) H.push(T.h_dup);
    }
    if (m === 'bar') H.push(T.h_num, T.h_neg);
    if (m === 'due') H.push(T.h_dt, T.h_now);
    H.push(T.h_view);
    out.innerHTML = preview(m, R, mx, nd, f) +
      '<div class="q-blk"><div class="q-h"><small>' + T.json + '</small><button type="button" class="q-cp">' + T.copy + '</button></div><pre><code>' + esc(json) + '</code></pre></div>' +
      '<div class="q-w"><small>' + T.warn + '</small><ul>' + H.map(function (h) { return '<li>' + esc(h) + '</li>'; }).join('') + '</ul></div>';
    var b = out.querySelector('.q-cp');
    b.addEventListener('click', function () {
      (navigator.clipboard ? navigator.clipboard.writeText(json) : Promise.reject()).then(function () { b.textContent = T.copied; setTimeout(function () { b.textContent = T.copy; }, 1400); }, function () {});
    });
  }
  root.addEventListener('input', render); root.addEventListener('change', render);
  $('#c-add').addEventListener('click', function () { addRule(); render(); });
  [].forEach.call(root.querySelectorAll('[data-preset]'), function (b) { b.addEventListener('click', function () { load(b.getAttribute('data-preset')); }); });
  load('approvals');
})();
