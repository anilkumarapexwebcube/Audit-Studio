/* GSC Suite — Navbar + view router (SPA shell)
 * Views are swapped purely in-memory. The URL is kept clean (no #hash fragments).
 */
(function () {
  const navLinks = document.getElementById('navLinks');
  const navToggle = document.getElementById('navToggle');

  function showView(name) {
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    const view = document.getElementById('view-' + name);
    if (view) view.classList.add('active');

    document.querySelectorAll('.nav-btn').forEach(b =>
      b.classList.toggle('active', b.dataset.nav === name));

    navLinks.classList.remove('open');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    document.dispatchEvent(new CustomEvent('viewchange', { detail: { name } }));
  }
  window.GSCNav = { showView };

  // Any element with data-nav navigates
  document.addEventListener('click', (e) => {
    const el = e.target.closest('[data-nav]');
    if (el) { e.preventDefault(); showView(el.dataset.nav); }
  });

  navToggle.addEventListener('click', () => navLinks.classList.toggle('open'));

  // Strip any leftover #fragment from the address bar so the URL stays clean.
  if (location.hash) history.replaceState({}, '', location.pathname + location.search);
})();
