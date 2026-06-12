/* GSC Suite — custom UI primitives (modals + toasts)
 * Drop-in, themed replacements for window.alert / confirm / prompt, plus toasts.
 * Exposed as window.GSCUI. All return Promises.
 */
window.GSCUI = (function () {
  const ICONS = {
    info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>',
    success: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
    warn: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    error: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
    question: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    lock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>'
  };

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

  function rootEl() {
    let r = document.getElementById('gscui-root');
    if (!r) { r = document.createElement('div'); r.id = 'gscui-root'; document.body.appendChild(r); }
    return r;
  }
  function toastRoot() {
    let r = document.getElementById('gscui-toasts');
    if (!r) { r = document.createElement('div'); r.id = 'gscui-toasts'; document.body.appendChild(r); }
    return r;
  }

  function modal({ title, message, type, buttons, input }) {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.className = 'gscui-overlay';
      const inputHtml = input
        ? `<input class="gscui-input" type="${input.password ? 'password' : 'text'}" placeholder="${esc(input.placeholder || '')}" value="${esc(input.value || '')}" spellcheck="false">`
        : '';
      overlay.innerHTML =
        `<div class="gscui-modal" role="dialog" aria-modal="true">
          <div class="gscui-modal-head">
            <span class="gscui-mi gscui-mi-${type || 'info'}">${ICONS[type] || ICONS.info}</span>
            <h3>${esc(title || '')}</h3>
          </div>
          <div class="gscui-modal-body">${esc(message || '').replace(/\n/g, '<br>')}</div>
          ${inputHtml}
          <div class="gscui-modal-actions"></div>
        </div>`;

      const actions = overlay.querySelector('.gscui-modal-actions');
      const inputEl = overlay.querySelector('.gscui-input');
      const cancelValue = input ? '__CANCEL__' : false;

      function close(val) {
        overlay.classList.remove('show');
        document.removeEventListener('keydown', onKey);
        setTimeout(() => overlay.remove(), 180);
        resolve(val);
      }
      function onKey(e) {
        if (e.key === 'Escape') close(cancelValue);
        else if (e.key === 'Enter' && inputEl) close(inputEl.value);
      }

      buttons.forEach((b) => {
        const btn = document.createElement('button');
        btn.className = 'btn ' + (b.variant || 'btn-ghost');
        btn.textContent = b.label;
        btn.addEventListener('click', () => close(b.value === '__INPUT__' ? (inputEl ? inputEl.value : null) : b.value));
        actions.appendChild(btn);
      });

      overlay.addEventListener('mousedown', (e) => { if (e.target === overlay) close(cancelValue); });
      document.addEventListener('keydown', onKey);
      rootEl().appendChild(overlay);
      requestAnimationFrame(() => overlay.classList.add('show'));
      if (inputEl) setTimeout(() => inputEl.focus(), 70);
      else { const p = actions.querySelector('.btn-primary, .btn-stop') || actions.lastChild; if (p) setTimeout(() => p.focus(), 70); }
    });
  }

  function alert(message, opts = {}) {
    return modal({
      title: opts.title || 'Notice', message, type: opts.type || 'info',
      buttons: [{ label: opts.okLabel || 'OK', value: true, variant: 'btn-primary' }]
    });
  }
  function confirm(message, opts = {}) {
    return modal({
      title: opts.title || 'Please confirm', message, type: opts.type || 'question',
      buttons: [
        { label: opts.cancelLabel || 'Cancel', value: false, variant: 'btn-ghost' },
        { label: opts.okLabel || 'Confirm', value: true, variant: opts.danger ? 'btn-stop' : 'btn-primary' }
      ]
    });
  }
  function prompt(message, opts = {}) {
    return modal({
      title: opts.title || 'Enter value', message, type: opts.type || 'lock',
      input: { password: opts.password, placeholder: opts.placeholder, value: opts.value },
      buttons: [
        { label: 'Cancel', value: '__CANCEL__', variant: 'btn-ghost' },
        { label: opts.okLabel || 'OK', value: '__INPUT__', variant: 'btn-primary' }
      ]
    }).then(v => v === '__CANCEL__' ? null : v);
  }

  function toast(message, type = 'info', ms = 3400) {
    const t = document.createElement('div');
    t.className = 'gscui-toast gscui-toast-' + type;
    t.innerHTML = `<span class="gscui-mi gscui-mi-${type}">${ICONS[type] || ICONS.info}</span><span>${esc(message)}</span>`;
    toastRoot().appendChild(t);
    requestAnimationFrame(() => t.classList.add('show'));
    setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 260); }, ms);
  }

  return { alert, confirm, prompt, toast };
})();
