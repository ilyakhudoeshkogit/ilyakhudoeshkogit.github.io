(function () {
  var root = document.documentElement;
  function save(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  document.getElementById('theme').addEventListener('click', function () {
    var n = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    root.setAttribute('data-theme', n); save('theme', n);
  });
  // Drawings are "drawn" when they scroll into view.
  var figs = document.querySelectorAll('.proj, .hero-art svg');
  if (!('IntersectionObserver' in window)) { for (var i = 0; i < figs.length; i++) { figs[i].classList.add('in'); } return; }
  var io = new IntersectionObserver(function (es) {
    es.forEach(function (e) { if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); } });
  }, { threshold: 0.15 });
  for (var j = 0; j < figs.length; j++) { io.observe(figs[j]); }
})();
