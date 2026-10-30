/* SPFx version helper. Data: learn.microsoft.com/sharepoint/dev/spfx/compatibility (checked 2026-10-30). Runs locally, sends nothing. */
(function () {
  var root = document.getElementById('spfxv'); if (!root) return;
  var en = document.documentElement.lang === 'en';
  var T = en ? {node:'Node.js', ts:'TypeScript', react:'React', tool:'Build toolchain', onprem:'SharePoint Server (on-prem)', online:'SharePoint Online only', cmd:'Pin React', ok:'fits', bad:'does not fit', none:'Paste the output of node -v', nf:'Could not find @microsoft/sp-core-library in this package.json', found:'Found SPFx ', notlisted:' is not in the official table, the closest lower version is selected', gulp:'gulp (legacy)', heft:'Heft', hint:'For a new project: Node.js 26 LTS (22.14+ and 24 also supported).', n22:'Node 22: version 22.14.0 or later.'}
              : {node:'Node.js', ts:'TypeScript', react:'React', tool:'Сборка', onprem:'SharePoint Server (локальный)', online:'только SharePoint Online', cmd:'Закрепить React', ok:'подходит для', bad:'не подходит для', none:'Вставьте вывод node -v', nf:'В этом package.json не нашёл @microsoft/sp-core-library', found:'Нашёл SPFx ', notlisted:' нет в официальной таблице, выбрана ближайшая младшая версия', gulp:'gulp (legacy)', heft:'Heft', hint:'Для нового проекта: Node.js 26 LTS (поддерживаются также 22.14+ и 24).', n22:'Node 22: версия 22.14.0 или новее.'};
  // [spfx, node, typescript, react]
  var D = [['1.24','22,24,26','2.9 - 5.8','18.3.1'],['1.23.2','22','2.9 - 5.8','17.0.1'],['1.23.0','22','2.9 - 5.8','17.0.1'],['1.22.2','22','2.9 - 5.8','17.0.1'],['1.22.1','22','2.9 - 5.8','17.0.1'],['1.22.0','22','2.9 - 5.8','17.0.1'],
    ['1.21.1','22','5.3','17.0.1'],['1.21.0','22','5.3','17.0.1'],['1.20.0','18','4.5, 4.7','17.0.1'],['1.19.0','18','4.5, 4.7','17.0.1'],['1.18.2','16,18','4.5, 4.7','17.0.1'],['1.18.1','16,18','4.5, 4.7','17.0.1'],['1.18.0','16,18','4.5, 4.7','17.0.1'],
    ['1.17.4','16.13+','4.5','17.0.1'],['1.17.3','16.13+','4.5','17.0.1'],['1.17.2','16.13+','4.5','17.0.1'],['1.17.1','16.13+','4.5','17.0.1'],['1.17.0','16.13+','4.5','17.0.1'],['1.16.1','16.13+','4.5','17.0.1'],['1.16.0','16.13+','4.5','17.0.1'],
    ['1.15.2','12,14,16','4.5','16.13.1'],['1.15.0','12,14,16','4.5','16.13.1'],['1.14.0','12,14','3.9','16.13.1'],['1.13.1','12,14','3.9','16.13.1'],['1.13.0','12,14','3.9','16.13.1'],['1.12.1','10,12,14','3.7','16.9.0'],['1.12.0','10,12','3.7','16.9.0'],
    ['1.11.0','10','3.3','16.8.5'],['1.10.0','8,10','3.3','16.8.5'],['1.9.1','8,10','2.9','16.8.5'],['1.8.2','8,10','2.9','16.7.0'],['1.8.1','8','2.7, 2.9, 3','16.7.0'],['1.8.0','8','2.7, 2.9, 3','16.7.0'],['1.7.1','8','2.4','16.3.2'],['1.7.0','8','2.4','16.3.2'],
    ['1.6.0','6,8','2.4','15'],['1.5.1','6,8','2.4','15'],['1.5.0','6,8','2.4','15'],['1.4.1','6,8','2.4','15'],['1.4.0','6','2.4','15'],['1.3.0','6','2.4','15'],['1.1.0','6','2.4','15'],['1.0.0','6','2.4','15']];
  function v(s) { return s.replace(/^[^\d]*/, '').split('.').map(function (x) { return parseInt(x, 10) || 0; }).concat([0, 0, 0]).slice(0, 3); }
  function cmp(a, b) { a = v(a); b = v(b); for (var i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i] - b[i]; return 0; }
  var sel = root.querySelector('select'), out = root.querySelector('.sv-out'), nodeIn = root.querySelector('#sv-node'), nodeRes = root.querySelector('.sv-noderes'), pj = root.querySelector('textarea'), pjRes = root.querySelector('.sv-pjres');
  D.forEach(function (r) { var o = document.createElement('option'); o.value = r[0]; o.textContent = 'SPFx ' + r[0]; sel.appendChild(o); });
  function row() { for (var i = 0; i < D.length; i++) if (D[i][0] === sel.value) return D[i]; return D[0]; }
  function onprem(s) {
    var l = [];
    if (cmp(s, '1.5') <= 0) l.push('Subscription Edition');
    if (cmp(s, '1.4.1') <= 0) l.push('2019');
    if (cmp(s, '1.1.0') <= 0) l.push('2016 FP2');
    return l.length ? l.join(', ') : T.online;
  }
  function nodeOk(r, ver) {
    var n = v(ver);
    return r[1].split(',').some(function (p) {
      if (/\+$/.test(p)) { var m = v(p); return n[0] === m[0] && n[1] >= m[1]; }
      var maj = parseInt(p, 10); if (n[0] !== maj) return false;
      if (r[0] === '1.24' && maj === 22) return cmp(ver, '22.14.0') >= 0;
      return true;
    });
  }
  function esc(s) { return String(s).replace(/[&<>]/g, function (c) { return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; }); }
  function render() {
    var r = row(), nodes = r[1].split(',').map(function (x) { return 'v' + x; }).join(', ');
    var cmd = r[3].indexOf('.') > 0 ? 'npm install react@' + r[3] + ' react-dom@' + r[3] + ' --save-exact' : 'npm install react@15 react-dom@15';
    out.innerHTML = '<div class="sv-c"><small>' + T.node + '</small><b>' + nodes + '</b></div>' +
      '<div class="sv-c"><small>' + T.ts + '</small><b>' + r[2] + '</b></div>' +
      '<div class="sv-c"><small>' + T.react + '</small><b>' + r[3] + '</b></div>' +
      '<div class="sv-c"><small>' + T.tool + '</small><b>' + (cmp(r[0], '1.22.0') >= 0 ? T.heft : T.gulp) + '</b></div>' +
      '<div class="sv-c w"><small>' + T.onprem + '</small><b>' + onprem(r[0]) + '</b></div>' +
      '<div class="sv-c w"><small>' + T.cmd + '</small><code>' + esc(cmd) + '</code></div>' +
      (r[0] === '1.24' ? '<p class="sv-hint">' + T.hint + ' ' + T.n22 + '</p>' : '');
    root.querySelectorAll('.sv-table tbody tr').forEach(function (tr) { tr.classList.toggle('on', tr.getAttribute('data-v') === r[0]); });
    checkNode();
  }
  function checkNode() {
    var s = (nodeIn.value || '').trim();
    if (!s) { nodeRes.textContent = ''; nodeRes.className = 'sv-noderes'; return; }
    if (!/^v?\d+(\.\d+){0,2}$/.test(s)) { nodeRes.textContent = T.none; nodeRes.className = 'sv-noderes'; return; }
    var ok = nodeOk(row(), s);
    nodeRes.textContent = 'Node ' + s.replace(/^v?/, 'v') + ': ' + (ok ? T.ok : T.bad) + ' SPFx ' + sel.value;
    nodeRes.className = 'sv-noderes ' + (ok ? 'ok' : 'bad');
  }
  function readPj() {
    var s = pj.value.trim(); if (!s) { pjRes.textContent = ''; return; }
    var m = s.match(/"@microsoft\/sp-core-library"\s*:\s*"[~^]?(\d+\.\d+\.\d+)/);
    if (!m) { pjRes.textContent = T.nf; return; }
    var found = m[1], pick = null;
    for (var i = 0; i < D.length; i++) if (cmp(D[i][0], found) <= 0) { pick = D[i][0]; break; }
    var exact = D.some(function (r) { return cmp(r[0], found) === 0; });
    if (!pick) pick = D[D.length - 1][0];
    sel.value = pick; render();
    pjRes.textContent = T.found + found + (exact ? '' : T.notlisted) + '.';
  }
  sel.addEventListener('change', render); nodeIn.addEventListener('input', checkNode); pj.addEventListener('input', readPj);
  root.querySelectorAll('.sv-table tbody tr').forEach(function (tr) { tr.addEventListener('click', function () { sel.value = tr.getAttribute('data-v'); render(); }); });
  render();
})();
