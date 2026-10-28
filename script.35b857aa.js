(function () {
  var root = document.documentElement;
  function save(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }

  // Theme: dark by default, choice is kept in localStorage and read before first paint (see inline script in <head>).
  var tb = document.getElementById('theme');
  if (tb) {
    tb.addEventListener('click', function () {
      var n = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      root.setAttribute('data-theme', n); save('theme', n);
    });
  }
  window.addEventListener('storage', function (e) {
    if (e.key === 'theme' && (e.newValue === 'light' || e.newValue === 'dark')) { root.setAttribute('data-theme', e.newValue); }
  });

  // Reveal sections and "draw" the schematics when they scroll into view.
  var rv = document.querySelectorAll('.rv');
  var figs = document.querySelectorAll('.proj svg, .hero-art svg');
  function show(el) { el.classList.add('in'); }
  if (!('IntersectionObserver' in window)) {
    for (var i = 0; i < rv.length; i++) { show(rv[i]); }
    for (var k = 0; k < figs.length; k++) { show(figs[k]); }
  } else {
    var io = new IntersectionObserver(function (es) {
      es.forEach(function (e) { if (e.isIntersecting) { show(e.target); io.unobserve(e.target); } });
    }, { threshold: 0.12, rootMargin: '0px 0px -4% 0px' });
    for (var j = 0; j < rv.length; j++) { io.observe(rv[j]); }
    for (var m = 0; m < figs.length; m++) { io.observe(figs[m]); }
  }
  root.classList.add('ok');

  // Service worker: caches hashed css/js/images, revalidates pages. Needs https (or localhost).
  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('/sw.js', { scope: '/', updateViaCache: 'none' }).catch(function () {});
    });
  }
})();
