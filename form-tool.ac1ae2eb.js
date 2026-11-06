/* SharePoint list form formatting builder: header / body / footer JSON + live preview. Runs locally, sends nothing. */
(function () {
  var root = document.getElementById('sfm'); if (!root) return;
  var en = document.documentElement.lang === 'en';
  var SCHEMA = 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json';
  var T = en ? {
    prev: 'Form preview (sample item)', copy: 'Copy', copied: 'Copied', warn: 'What to keep in mind',
    hdr: 'Header JSON', body: 'Body JSON', ftr: 'Footer JSON',
    icon: 'Fluent icon name', title: 'Title (text or expression)', sub: 'Subtitle',
    person: 'Show person', personF: 'Person field (internal name)',
    sec: 'Sections', addSec: '+ section', addFld: '+ field', del: 'Remove',
    secName: 'Section title', disp: 'Display name (in body JSON)', show: 'Show when',
    always: 'Always', custom: 'Custom formula',
    acts: 'Footer actions', addAct: '+ action', actTxt: 'Label', actStyle: 'Style',
    actKind: 'Action', url: 'URL (for link)', primary: 'Primary', def: 'Default',
    link: 'Open link', cancel: 'Cancel-looking',
    h_paste: 'Where to paste: open an item → Edit form → Configure layout → Apply formatting to Header / Body / Footer. Paste the matching JSON and Save. Re-open the form (or clear the cache) to see it.',
    h_tokens: 'Tokens use the internal name: [$Title], [$Status]. For a Person field use [$Author.displayName] and [$Author.email].',
    h_body: 'Body JSON only lays out sections. It does not hide fields. Conditional show/hide is a separate formula on each column (Edit columns → conditional formula). Old tenants may lack that UI.',
    h_schema: 'Header and footer use column-formatting JSON (elmType tree). Body uses a different shape: { "sections": [ { "displayname", "fields" } ] }. Field names in body are display names, not internal names.',
    h_view: 'Form formatting only changes how the form looks. It does not change permissions, validation or the data itself.',
    h_ct: 'Formatting is stored on the list content type. Content types from the hub are read-only until you temporarily switch them to Edit mode.',
    h_refresh: 'After Save, close and open the form again. A hard refresh sometimes helps if the old layout is cached.',
    sampleTitle: 'Laptop upgrade', sampleStatus: 'Pending', sampleAuthor: 'Ilya Khudoeshko',
    sampleAmount: '1 200 USD', sampleDesk: 'A-12', sampleZone: 'Open space', sampleDate: '12.11.2026',
    sampleDue: '18.11.2026', sampleAssigned: 'Anna K.',
    tabH: 'Header', tabB: 'Body', tabF: 'Footer'
  } : {
    prev: 'Предпросмотр формы (тестовый элемент)', copy: 'Копировать', copied: 'Скопировано', warn: 'О чём помнить',
    hdr: 'JSON заголовка', body: 'JSON тела', ftr: 'JSON подвала',
    icon: 'Имя иконки Fluent', title: 'Заголовок (текст или выражение)', sub: 'Подзаголовок',
    person: 'Показать человека', personF: 'Поле человека (внутр. имя)',
    sec: 'Секции', addSec: '+ секция', addFld: '+ поле', del: 'Убрать',
    secName: 'Название секции', disp: 'Отображаемое имя (в body JSON)', show: 'Показывать когда',
    always: 'Всегда', custom: 'Своя формула',
    acts: 'Действия в подвале', addAct: '+ действие', actTxt: 'Подпись', actStyle: 'Стиль',
    actKind: 'Действие', url: 'URL (для ссылки)', primary: 'Основная', def: 'Обычная',
    link: 'Открыть ссылку', cancel: 'Как «Отмена»',
    h_paste: 'Куда вставить: откройте элемент → «Изменить форму» → «Настроить макет» → «Применить форматирование к» Header / Body / Footer. Вставьте нужный JSON и нажмите «Сохранить». Чтобы увидеть результат, закройте и снова откройте форму (иногда помогает сброс кеша).',
    h_tokens: 'Токены берут внутреннее имя: [$Title], [$Status]. Для человека — [$Author.displayName] и [$Author.email].',
    h_body: 'Body JSON только раскладывает секции. Поля он не прячет. Условный показ — отдельная формула на столбце («Изменить столбцы» → условная формула). В старых тенантах этого UI может не быть.',
    h_schema: 'Header и footer — это column formatting JSON (дерево elmType). Body — другой формат: { "sections": [ { "displayname", "fields" } ] }. В fields указываются отображаемые имена столбцов, не внутренние.',
    h_view: 'Форматирование формы меняет только вид. Права, проверку данных и сами данные оно не трогает.',
    h_ct: 'JSON хранится на типе содержимого списка. Типы из хаба по умолчанию только для чтения: временно переключите в режим правки, примените JSON и верните обратно.',
    h_refresh: 'После «Сохранить» закройте и снова откройте форму. Иногда нужна жёсткая перезагрузка страницы, если старый макет закеширован.',
    sampleTitle: 'Апгрейд ноутбука', sampleStatus: 'Pending', sampleAuthor: 'Илья Худоешко',
    sampleAmount: '1 200 USD', sampleDesk: 'A-12', sampleZone: 'Опенспейс', sampleDate: '12.11.2026',
    sampleDue: '18.11.2026', sampleAssigned: 'Анна К.',
    tabH: 'Header', tabB: 'Body', tabF: 'Footer'
  };

  var P = {
    approvals: {
      icon: 'CheckboxComposite',
      title: "='Approval: ' + if([$Title] == '', 'New request', [$Title])",
      sub: "='Status: ' + [$Status]",
      person: true, personField: 'Author',
      sections: [
        { name: en ? 'Request' : 'Заявка', fields: [
          { d: 'Title', show: '' }, { d: 'Status', show: '' }, { d: 'Amount', show: '' }
        ]},
        { name: en ? 'People' : 'Люди', fields: [
          { d: 'Requestor', show: '' }, { d: 'Approver', show: "[$Status] == 'Pending'" }
        ]}
      ],
      actions: [
        { t: en ? 'Open in Dashboard' : 'Открыть в Dashboard', s: 'primary', k: 'link', u: 'https://contoso.sharepoint.com/sites/exec' },
        { t: en ? 'Close' : 'Закрыть', s: 'default', k: 'cancel', u: '' }
      ],
      sample: { Title: T.sampleTitle, Status: T.sampleStatus, Amount: T.sampleAmount, Author: T.sampleAuthor, Requestor: T.sampleAuthor, Approver: 'PM' }
    },
    desk: {
      icon: 'Locations',
      title: "='Desk ' + [$Desk]",
      sub: "='Zone: ' + [$Zone]",
      person: true, personField: 'ReservedBy',
      sections: [
        { name: en ? 'Place' : 'Место', fields: [
          { d: 'Desk', show: '' }, { d: 'Zone', show: '' }
        ]},
        { name: en ? 'Booking' : 'Бронь', fields: [
          { d: 'Date', show: '' }, { d: 'ReservedBy', show: '' }
        ]}
      ],
      actions: [
        { t: en ? 'Floor plan' : 'План этажа', s: 'primary', k: 'link', u: 'https://contoso.sharepoint.com/sites/office/SitePages/Floor.aspx' },
        { t: en ? 'Cancel' : 'Отмена', s: 'default', k: 'cancel', u: '' }
      ],
      sample: { Desk: T.sampleDesk, Zone: T.sampleZone, Date: T.sampleDate, ReservedBy: T.sampleAuthor, Title: T.sampleDesk }
    },
    task: {
      icon: 'TaskLogo',
      title: "=if([$Title] == '', 'New task', [$Title])",
      sub: "='Due: ' + toLocaleDateString([$DueDate])",
      person: true, personField: 'AssignedTo',
      sections: [
        { name: en ? 'Task' : 'Задача', fields: [
          { d: 'Title', show: '' }, { d: 'DueDate', show: '' }, { d: 'AssignedTo', show: '' }
        ]},
        { name: en ? 'Blocked details' : 'Почему блок', fields: [
          { d: 'BlockReason', show: "[$Status] == 'Blocked'" }, { d: 'Status', show: '' }
        ]}
      ],
      actions: [
        { t: en ? 'Related list' : 'Связанный список', s: 'default', k: 'link', u: 'https://contoso.sharepoint.com/lists/tasks' }
      ],
      sample: { Title: en ? 'Fix login' : 'Починить вход', DueDate: T.sampleDue, AssignedTo: T.sampleAssigned, Status: 'Blocked', BlockReason: en ? 'Waiting on IdP' : 'Ждём IdP' }
    }
  };

  var $ = function (s) { return root.querySelector(s); };
  var $$ = function (s) { return [].slice.call(root.querySelectorAll(s)); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
  function mode() { return root.getAttribute('data-mode') || 'header'; }
  function setMode(m) {
    root.setAttribute('data-mode', m);
    $$('.f-tab').forEach(function (b) { b.setAttribute('aria-selected', b.getAttribute('data-mode') === m ? 'true' : 'false'); });
    render();
  }

  /* ---- state readers ---- */
  function readHeader() {
    return {
      icon: ($('#f-icon') || {}).value || 'Group',
      title: ($('#f-title') || {}).value || '[$Title]',
      sub: ($('#f-sub') || {}).value || '',
      person: !!($('#f-person') && $('#f-person').checked),
      personField: (($('#f-person-f') || {}).value || 'Author').trim() || 'Author'
    };
  }
  function readSections() {
    return $$('#f-secs .f-sec').map(function (sec) {
      return {
        name: (sec.querySelector('.f-sec-name') || {}).value || '',
        fields: [].slice.call(sec.querySelectorAll('.f-fld')).map(function (row) {
          var showSel = row.querySelector('.f-show');
          var showCustom = row.querySelector('.f-show-c');
          var show = '';
          if (showSel && showSel.value === 'custom') show = (showCustom && showCustom.value || '').trim();
          else if (showSel && showSel.value && showSel.value !== 'always') show = showSel.value;
          return { d: (row.querySelector('.f-disp') || {}).value || '', show: show };
        }).filter(function (f) { return f.d.trim(); })
      };
    });
  }
  function readActions() {
    return $$('#f-acts .f-act').map(function (row) {
      return {
        t: (row.querySelector('.f-at') || {}).value || '',
        s: (row.querySelector('.f-as') || {}).value || 'default',
        k: (row.querySelector('.f-ak') || {}).value || 'link',
        u: (row.querySelector('.f-au') || {}).value || ''
      };
    }).filter(function (a) { return a.t.trim(); });
  }

  /* ---- JSON builders ---- */
  function exprOrText(v) {
    v = (v || '').trim();
    if (!v) return '';
    if (v.charAt(0) === '=') return v;
    if (v.indexOf('[$') >= 0 || v.indexOf('@') >= 0) return '=' + (v.charAt(0) === "'" || v.indexOf('+') >= 0 ? v : ("'" + v.replace(/'/g, "''") + "'"));
    /* plain text that already looks like an expression without = */
    if (/^(if|toLocaleDateString)\(/.test(v)) return '=' + v;
    return v;
  }
  function titleContent(v) {
    v = (v || '').trim() || '[$Title]';
    if (v.charAt(0) === '=') return v;
    if (v.indexOf('[$') >= 0 || v.indexOf('+') >= 0 || /^(if|toLocaleDateString)\(/.test(v)) return '=' + v.replace(/^=/, '');
    return v;
  }
  function buildHeader(h) {
    var kids = [];
    kids.push({
      elmType: 'div',
      style: { display: 'flex', 'box-sizing': 'border-box', 'align-items': 'center' },
      children: [{
        elmType: 'div',
        attributes: { iconName: h.icon || 'Group', class: 'ms-fontSize-42 ms-fontWeight-regular ms-fontColor-themePrimary', title: h.icon || 'Group' },
        style: { flex: 'none', padding: '0px', 'padding-left': '0px', height: '36px' }
      }]
    });
    var textKids = [{ elmType: 'div', txtContent: titleContent(h.title) }];
    if (h.sub && h.sub.trim()) {
      textKids.push({
        elmType: 'div',
        style: { 'margin-top': '4px', 'font-size': '14px', 'font-weight': '400' },
        attributes: { class: 'ms-fontColor-neutralSecondary' },
        txtContent: titleContent(h.sub)
      });
    }
    if (h.person) {
      textKids.push({
        elmType: 'div',
        style: { 'margin-top': '6px', 'font-size': '13px', 'font-weight': '400', display: 'flex', 'align-items': 'center' },
        children: [
          { elmType: 'span', attributes: { iconName: 'Contact', class: 'ms-fontSize-16 ms-fontColor-themePrimary' }, style: { 'margin-right': '6px' } },
          { elmType: 'span', txtContent: '=[' + '$' + h.personField + '.displayName]' }
        ]
      });
    }
    kids.push({
      elmType: 'div',
      attributes: { class: 'ms-fontColor-neutralSecondary ms-fontWeight-bold ms-fontSize-24' },
      style: { 'box-sizing': 'border-box', width: '100%', 'text-align': 'left', padding: '12px 12px 16px', overflow: 'hidden' },
      children: textKids
    });
    return {
      $schema: SCHEMA,
      elmType: 'div',
      attributes: { class: 'ms-borderColor-neutralTertiary' },
      style: {
        width: '99%', 'border-top-width': '0px', 'border-bottom-width': '1px',
        'border-left-width': '0px', 'border-right-width': '0px', 'border-style': 'solid', 'margin-bottom': '16px'
      },
      children: kids
    };
  }
  function buildBody(sections) {
    return {
      sections: sections.map(function (s) {
        return { displayname: s.name || '', fields: s.fields.map(function (f) { return f.d; }) };
      }).filter(function (s) { return s.fields.length; })
    };
  }
  function buildFooter(actions) {
    var children = actions.map(function (a) {
      var isPrimary = a.s === 'primary';
      var cls = isPrimary
        ? 'ms-bgColor-themePrimary ms-fontColor-white ms-fontWeight-semibold ms-fontSize-m'
        : 'ms-fontColor-themePrimary ms-borderColor-themePrimary ms-fontWeight-semibold ms-fontSize-m ms-bgColor-neutralLighter';
      var style = {
        display: 'inline-block', padding: '8px 16px', 'margin-right': '8px', 'margin-top': '8px',
        'text-decoration': 'none', 'border-radius': '2px', 'border-width': '1px', 'border-style': 'solid'
      };
      if (a.k === 'cancel') {
        return {
          elmType: 'div',
          txtContent: a.t,
          style: Object.assign({}, style, { cursor: 'default', opacity: '0.85' }),
          attributes: { class: 'ms-fontColor-neutralPrimary ms-bgColor-neutralLighter ms-borderColor-neutralTertiary' }
        };
      }
      return {
        elmType: 'a',
        txtContent: a.t,
        style: style,
        attributes: {
          class: cls,
          target: '_blank',
          href: a.u && a.u.trim() ? a.u.trim() : '#'
        }
      };
    });
    return {
      $schema: SCHEMA,
      elmType: 'div',
      style: { width: '100%', 'text-align': 'left', overflow: 'hidden', 'border-top-width': '1px', 'border-top-style': 'solid', 'border-top-color': '#edebe9', 'padding-top': '8px' },
      children: [{ elmType: 'div', style: { width: '100%', 'padding-top': '4px' }, children: children }]
    };
  }

  /* ---- preview helpers ---- */
  function resolveToken(expr, sample) {
    if (expr == null) return '';
    var s = String(expr);
    if (s.charAt(0) === '=') s = s.slice(1);
    /* very small evaluator for preview: replace [$Field] and [$Field.displayName], simple if(a=='',b,c) and concat with + */
    s = s.replace(/\[\$([A-Za-z0-9_]+)(?:\.(displayName|email))?\]/g, function (_, f, p) {
      var v = sample[f];
      if (v == null) v = '';
      if (p === 'email') return String(v).toLowerCase().replace(/\s+/g, '.') + '@contoso.com';
      return String(v);
    });
    /* if(x == '', a, b) */
    s = s.replace(/if\(\s*'([^']*)'\s*==\s*''\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*\)/g, function (_, x, a, b) { return x === '' ? a : b; });
    s = s.replace(/if\(\s*'([^']*)'\s*==\s*''\s*,\s*'([^']*)'\s*,\s*([^)]+)\)/g, function (_, x, a, b) { return x === '' ? a : b.replace(/^'|'$/g, ''); });
    /* strip simple string literals joined by + */
    var parts = s.split('+').map(function (p) { return p.trim().replace(/^'|'$/g, '').replace(/toLocaleDateString\(([^)]*)\)/g, '$1'); });
    return parts.join('').replace(/\s+/g, ' ').trim() || s;
  }
  function glyph(icon) {
    var M = { CheckboxComposite: '☑', Locations: '⌖', TaskLogo: '✓', Group: '👥', Contact: '👤', ShieldAlert: '⚠', DocumentSet: '▦' };
    return M[icon] || '◆';
  }
  function previewHTML(h, sections, actions, sample) {
    var title = resolveToken(h.title, sample) || sample.Title || '';
    var sub = h.sub ? resolveToken(h.sub, sample) : '';
    var person = h.person ? (sample[h.personField] || sample.Author || '') : '';
    var html = '<div class="f-form" aria-hidden="false"><small>' + T.prev + '</small>';
    html += '<div class="f-fh"><span class="f-ico" aria-hidden="true">' + glyph(h.icon) + '</span><div class="f-ft"><b>' + esc(title) + '</b>';
    if (sub) html += '<span class="f-sub">' + esc(sub) + '</span>';
    if (person) html += '<span class="f-per"><i aria-hidden="true">👤</i> ' + esc(person) + '</span>';
    html += '</div></div><div class="f-fb">';
    sections.forEach(function (sec) {
      if (!sec.fields.length) return;
      html += '<div class="f-secv">';
      if (sec.name) html += '<div class="f-sn">' + esc(sec.name) + '</div>';
      sec.fields.forEach(function (f) {
        var hidden = f.show && !evalShow(f.show, sample);
        var val = sample[f.d] != null ? sample[f.d] : (sample[f.d.replace(/\s/g, '')] != null ? sample[f.d.replace(/\s/g, '')] : '—');
        html += '<div class="f-row' + (hidden ? ' is-hid' : '') + '"><span class="f-lab">' + esc(f.d) + (f.show ? ' <em title="' + esc(f.show) + '">?</em>' : '') + '</span><span class="f-val">' + esc(val) + '</span></div>';
      });
      html += '</div>';
    });
    html += '</div>';
    if (actions.length) {
      html += '<div class="f-ff">';
      actions.forEach(function (a) {
        html += '<span class="f-btn' + (a.s === 'primary' ? ' prim' : '') + (a.k === 'cancel' ? ' can' : '') + '">' + esc(a.t) + '</span>';
      });
      html += '</div>';
    }
    html += '</div>';
    return html;
  }
  function evalShow(formula, sample) {
    /* support [$Field] == 'Value' and != */
    var m = /^\s*\[\$([A-Za-z0-9_]+)\]\s*(==|!=)\s*'([^']*)'\s*$/.exec(formula);
    if (!m) return true;
    var cur = String(sample[m[1]] == null ? '' : sample[m[1]]);
    return m[2] === '==' ? cur === m[3] : cur !== m[3];
  }

  /* ---- editors ---- */
  function addField(container, f) {
    f = f || { d: '', show: '' };
    var row = document.createElement('div'); row.className = 'f-fld';
    var showVal = f.show || 'always';
    var isCustom = showVal !== 'always' && showVal !== "[$Status] == 'Pending'" && showVal !== "[$Status] == 'Blocked'";
    var sel = isCustom ? 'custom' : (showVal === 'always' || !showVal ? 'always' : showVal);
    row.innerHTML =
      '<label>' + T.disp + '<input class="f-disp" type="text" spellcheck="false" autocomplete="off"></label>' +
      '<label>' + T.show + '<select class="f-show">' +
        '<option value="always">' + T.always + '</option>' +
        '<option value="[$Status] == \'Pending\'">[$Status] == \'Pending\'</option>' +
        '<option value="[$Status] == \'Blocked\'">[$Status] == \'Blocked\'</option>' +
        '<option value="custom">' + T.custom + '</option>' +
      '</select></label>' +
      '<label class="f-show-wrap">' + T.custom + '<input class="f-show-c" type="text" spellcheck="false" placeholder="[$Status] == \'Pending\'"></label>' +
      '<button type="button" class="q-x" aria-label="' + T.del + '" title="' + T.del + '">✕</button>';
    row.querySelector('.f-disp').value = f.d || '';
    row.querySelector('.f-show').value = sel;
    row.querySelector('.f-show-c').value = isCustom ? f.show : '';
    row.dataset.showMode = sel;
    row.querySelector('.f-show').addEventListener('change', function () {
      row.dataset.showMode = row.querySelector('.f-show').value;
      render();
    });
    row.querySelector('.q-x').addEventListener('click', function () { row.remove(); render(); });
    container.appendChild(row);
  }
  function addSection(s) {
    s = s || { name: '', fields: [{ d: '', show: '' }] };
    var box = document.createElement('div'); box.className = 'f-sec';
    box.innerHTML =
      '<div class="f-sec-h"><label>' + T.secName + '<input class="f-sec-name" type="text" autocomplete="off"></label>' +
      '<button type="button" class="q-x f-sec-x" aria-label="' + T.del + '" title="' + T.del + '">✕</button></div>' +
      '<div class="f-flds"></div>' +
      '<button type="button" class="q-pre f-add-fld">' + T.addFld + '</button>';
    box.querySelector('.f-sec-name').value = s.name || '';
    var flds = box.querySelector('.f-flds');
    (s.fields && s.fields.length ? s.fields : [{ d: '', show: '' }]).forEach(function (f) { addField(flds, f); });
    box.querySelector('.f-add-fld').addEventListener('click', function () { addField(flds); render(); });
    box.querySelector('.f-sec-x').addEventListener('click', function () { box.remove(); render(); });
    $('#f-secs').appendChild(box);
  }
  function addAction(a) {
    a = a || { t: '', s: 'default', k: 'link', u: '' };
    var row = document.createElement('div'); row.className = 'f-act';
    row.innerHTML =
      '<label>' + T.actTxt + '<input class="f-at" type="text" autocomplete="off"></label>' +
      '<label>' + T.actStyle + '<select class="f-as"><option value="primary">' + T.primary + '</option><option value="default">' + T.def + '</option></select></label>' +
      '<label>' + T.actKind + '<select class="f-ak"><option value="link">' + T.link + '</option><option value="cancel">' + T.cancel + '</option></select></label>' +
      '<label class="f-au-w">' + T.url + '<input class="f-au" type="text" spellcheck="false" autocomplete="off" placeholder="https://"></label>' +
      '<button type="button" class="q-x" aria-label="' + T.del + '" title="' + T.del + '">✕</button>';
    row.querySelector('.f-at').value = a.t || '';
    row.querySelector('.f-as').value = a.s || 'default';
    row.querySelector('.f-ak').value = a.k || 'link';
    row.querySelector('.f-au').value = a.u || '';
    row.querySelector('.q-x').addEventListener('click', function () { row.remove(); render(); });
    row.querySelector('.f-ak').addEventListener('change', function () { render(); });
    $('#f-acts').appendChild(row);
  }

  var sample = P.approvals.sample;

  function load(k) {
    var p = P[k]; if (!p) return;
    sample = p.sample || {};
    $('#f-icon').value = p.icon;
    $('#f-title').value = p.title;
    $('#f-sub').value = p.sub || '';
    $('#f-person').checked = !!p.person;
    $('#f-person-f').value = p.personField || 'Author';
    $('#f-secs').innerHTML = '';
    (p.sections || []).forEach(addSection);
    $('#f-acts').innerHTML = '';
    (p.actions || []).forEach(addAction);
    render();
  }

  function blk(label, json, key) {
    return '<div class="q-blk" data-b="' + key + '"><div class="q-h"><small>' + label + '</small><button type="button" class="q-cp" data-k="' + key + '">' + T.copy + '</button></div><pre><code>' + esc(json) + '</code></pre></div>';
  }
  function condHints(sections) {
    var list = [];
    sections.forEach(function (s) {
      s.fields.forEach(function (f) {
        if (f.show) list.push('=if(' + f.show + ", 'true', 'false')  /* " + f.d + ' */');
      });
    });
    return list;
  }

  function render() {
    var h = readHeader();
    var sections = readSections();
    var actions = readActions();
    root.dataset.person = h.person ? '1' : '0';
    $$('.f-fld').forEach(function (row) {
      row.dataset.showMode = (row.querySelector('.f-show') || {}).value || 'always';
    });
    $$('.f-act').forEach(function (row) {
      row.dataset.kind = (row.querySelector('.f-ak') || {}).value || 'link';
    });

    var headerJson = JSON.stringify(buildHeader(h), null, 2);
    var bodyJson = JSON.stringify(buildBody(sections), null, 2);
    var footerJson = JSON.stringify(buildFooter(actions), null, 2);
    var tips = [T.h_paste, T.h_tokens, T.h_schema, T.h_body, T.h_view, T.h_ct, T.h_refresh];
    var conds = condHints(sections);
    if (conds.length) {
      tips.push((en
        ? 'Conditional formulas for columns (paste under Edit columns → conditional formula), not into Body JSON: '
        : 'Условные формулы для столбцов (вставляются в «Изменить столбцы» → условная формула), не в Body JSON: ') + conds.join('; '));
    }

    var out = $('.f-out');
    out.innerHTML =
      previewHTML(h, sections, actions, sample) +
      blk(T.hdr, headerJson, 'header') +
      blk(T.body, bodyJson, 'body') +
      blk(T.ftr, footerJson, 'footer') +
      '<div class="q-w"><small>' + T.warn + '</small><ul>' + tips.map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('') + '</ul></div>';

    var map = { header: headerJson, body: bodyJson, footer: footerJson };
    $$.call ? null : null;
    [].forEach.call(out.querySelectorAll('.q-cp'), function (b) {
      b.addEventListener('click', function () {
        var json = map[b.getAttribute('data-k')];
        (navigator.clipboard ? navigator.clipboard.writeText(json) : Promise.reject()).then(function () {
          b.textContent = T.copied; setTimeout(function () { b.textContent = T.copy; }, 1400);
        }, function () {});
      });
    });
  }

  /* ---- wire ---- */
  $$('.f-tab').forEach(function (b) {
    b.addEventListener('click', function () { setMode(b.getAttribute('data-mode')); });
  });
  $('#f-add-sec').addEventListener('click', function () { addSection(); render(); });
  $('#f-add-act').addEventListener('click', function () { addAction(); render(); });
  root.addEventListener('input', render);
  root.addEventListener('change', render);
  [].forEach.call(root.querySelectorAll('[data-preset]'), function (b) {
    b.addEventListener('click', function () { load(b.getAttribute('data-preset')); });
  });
  load('approvals');
  setMode('header');
})();
