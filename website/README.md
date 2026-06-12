# GSC Suite - Search Console Automation Studio

A single website that combines two Google Search Console automations (originally Chrome
extensions) into one clean, premium workspace with a home page and dynamic per-tool pages
switched from the navbar.

- **Crawl Date Tracker** - bulk-check the last Google crawl date + index status for up to
  100 URLs, with pause/resume/stop and CSV export.
- **GSC Audit Studio** - generate full PPTX audit decks (James / Omega / Neon) from the
  live Search Console API.

## Running it

Open it from any static web server (OAuth and module loading need an `http(s)` origin -
`file://` won't work for sign-in):

```bash
cd website
python -m http.server 8754
# then open http://localhost:8754
```

Any static host works (Netlify, GitHub Pages, Nginx, IIS, etc.).

## Crawl Date Tracker

Works out of the box - it talks to the same Google Apps Script webhook the original tool
used. No configuration required.

## GSC Audit Studio - one-time setup (Setup tab)

Because this runs as a website (not a Chrome extension), the Chrome-only pieces were
replaced with browser-native equivalents:

| Extension API | Web replacement |
|---|---|
| `chrome.identity` OAuth | Google Identity Services token client (no client secret) |
| `chrome.storage.local` | `localStorage` |
| `chrome.downloads` | data-URL anchor download |
| `chrome.debugger` screenshot capture | manual screenshot upload (Screenshots tab) |

To enable sign-in:

1. In the [Google Cloud Console](https://console.cloud.google.com/apis/credentials),
   create an **OAuth 2.0 Client ID** of type **Web application**.
2. Add this page's origin (shown on the Setup tab, e.g. `http://localhost:8754` or your
   deployed URL) under **Authorized JavaScript origins**.
3. Enable the **Google Search Console API** for the project.
4. Paste the Client ID into the Setup tab and Save.
5. In the **Accounts** tab, connect each Google account that has Search Console access.

The Apps Script URL (for the domain → account mapping) is pre-filled and can be changed in
the Setup tab.

### Admin vs. regular users

The **Domain List** and **Setup** tabs, plus the **Admin Dashboard** view, are hidden from
regular users. Click the **Admin** (lock) button in the navbar and enter the admin password
to reveal them; click **Exit Admin** to hide them again. Admin state is per browser tab
session.

All success / error / confirmation dialogs use the project's own themed modals and toasts
(no native browser `alert`/`confirm`/`prompt`).

### Screenshots

A website can't silently screenshot the Search Console UI like the extension's
`chrome.debugger` did. Upload screenshots manually in the **Screenshots** tab and they're
embedded into the relevant report slides. Any slide left without an image renders a tidy
placeholder, so reports always generate without errors.

## Project layout

```
website/
  index.html              SPA shell - navbar + Home / Crawl / Audit views
  css/style.css           design system (provided palette + JetBrains Mono / Plus Jakarta Sans)
  js/app.js               navbar + view router
  js/crawl-tracker.js     Crawl Date Tracker logic
  js/audit-studio.js      Audit Studio logic (web-adapted)
  js/report-builder.js    James PPTX builder (reused, unchanged)
  js/format-omega.js      Omega PPTX builder (reused, unchanged)
  js/format-neon.js       Neon PPTX builder (reused, unchanged)
  lib/pptxgen.bundle.js   PptxGenJS 3.12.0
```
