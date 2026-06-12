/* GSC Suite — Admin gate + dashboard
 * Password-protects the Domain List, Setup tabs and the Admin Dashboard view.
 */
(function () {
  'use strict';
  const ADMIN_PASS = 'hello@Admin';
  const LS_ADMIN = 'gsc_admin';

  function isAdmin() { return sessionStorage.getItem(LS_ADMIN) === '1'; }

  function readJSON(store, key) {
    try { return JSON.parse(store.getItem(key) || '{}'); } catch (_) { return {}; }
  }

  function apply() {
    document.body.classList.toggle('admin-mode', isAdmin());
    const lbl = document.querySelector('#adminToggleBtn .atb-label');
    if (lbl) lbl.textContent = isAdmin() ? 'Exit Admin' : 'Admin';
    renderDashboard();
  }

  async function requestUnlock() {
    const pass = await GSCUI.prompt('Enter the admin password to unlock Domain List, Setup and the Admin Dashboard.', {
      title: 'Admin access', type: 'lock', password: true, placeholder: 'Password', okLabel: 'Unlock'
    });
    if (pass === null) return false;
    if (pass === ADMIN_PASS) {
      sessionStorage.setItem(LS_ADMIN, '1');
      apply();
      GSCUI.toast('Admin mode unlocked.', 'success');
      GSCNav.showView('admin');
      return true;
    }
    GSCUI.alert('That password is incorrect.', { title: 'Access denied', type: 'error' });
    return false;
  }

  async function lock() {
    const ok = await GSCUI.confirm('Exit admin mode? Domain List, Setup and the Admin Dashboard will be hidden again.', {
      title: 'Exit admin mode', okLabel: 'Exit', cancelLabel: 'Stay'
    });
    if (!ok) return;
    sessionStorage.removeItem(LS_ADMIN);
    apply();
    const runTab = document.querySelector('.subtab[data-sub="run"]');
    if (runTab) runTab.click();
    GSCNav.showView('home');
    GSCUI.toast('Admin mode locked.', 'info');
  }

  function renderDashboard() {
    const view = document.getElementById('view-admin');
    if (!view) return;
    const cfg = readJSON(localStorage, 'gsc_audit_config');
    const mapping = readJSON(localStorage, 'gsc_audit_mapping');
    const accounts = readJSON(sessionStorage, 'gsc_audit_accounts');

    const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    set('adm-accounts', Object.keys(accounts).length);
    set('adm-domains', Object.keys(mapping).length);

    const setBadge = (id, ok, okText, badText) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.textContent = ok ? okText : badText;
      el.className = 'status-pill ' + (ok ? 'ok' : 'warn');
    };
    setBadge('adm-oauth', !!cfg.clientId, 'Configured', 'Not set');
    setBadge('adm-script', !!cfg.appsScriptUrl, 'Configured', 'Not set');
  }

  // ---- wiring ----
  document.addEventListener('DOMContentLoaded', init);
  if (document.readyState !== 'loading') init();
  let inited = false;
  function init() {
    if (inited) return; inited = true;

    const toggle = document.getElementById('adminToggleBtn');
    if (toggle) toggle.addEventListener('click', () => (isAdmin() ? lock() : requestUnlock()));

    // Quick-action buttons inside the dashboard
    document.addEventListener('click', (e) => {
      const go = e.target.closest('[data-admin-go]');
      if (go) {
        const sub = go.dataset.adminGo;
        if (sub === 'lock') { lock(); return; }
        GSCNav.showView('audit');
        setTimeout(() => { const t = document.querySelector('.subtab[data-sub="' + sub + '"]'); if (t) t.click(); }, 60);
      }
    });

    document.addEventListener('viewchange', (e) => { if (e.detail.name === 'admin') renderDashboard(); });

    apply();
  }

  // Exposed so the audit studio can gate its admin-only subtabs and trigger the unlock flow.
  window.GSCAdmin = { isAdmin, requestUnlock };
})();
