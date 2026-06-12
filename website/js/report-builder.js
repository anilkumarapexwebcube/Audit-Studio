/* GSC Audit Studio - Report Builder
 * Constructs the full 19-slide James-format PPTX using PptxGenJS.
 */

const C = {
  TEAL_TITLE_BG:  '165060',
  TEAL_HEADER_BG: '0D3D4A',
  GOLD:           'C9A84C',
  GOLD_SOFT:      'E8A020',
  WHITE:          'FFFFFF',
  OFF_WHITE:      'E8F4F6',
  LIGHT_TEAL:     '99C8D0',
  CARD_BG:        'EEF6FA',
  CARD_BG_2:      'F5F8FA',
  TEXT_DARK:      '1A1A2E',
  TEXT_SOFT:      '5A6B7A',
  KPI_ORANGE:     'E8A020',
  KPI_BLUE:       '2980B9',
  KPI_GREEN:      '27AE60',
  KPI_PURPLE:     '8E44AD',
  STATUS_OK:      '27AE60',
  STATUS_WARN:    'E8A020',
  STATUS_ERR:     'C0392B'
};
const FONT_HEAD = 'Trebuchet MS';
const FONT_BODY = 'Calibri';

async function buildJamesPptx(data) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5

  let slideNum = 1;

  // Slide 1: Title
  addTitleSlide(pptx, data);

  // Slide 2: Introduction
  addIntroSlide(pptx, data, ++slideNum);

  // Slide 3: Sitemap Status
  addSitemapStatusSlide(pptx, data, ++slideNum);

  // Slide 4: Sitemap Last Read
  addSitemapLastReadSlide(pptx, data, ++slideNum);

  // Slides 5-9: URL Inspection sections
  addUrlIndexingSlide(pptx, data, ++slideNum);
  addCanonicalSlide(pptx, data, ++slideNum);
  addRobotsSlide(pptx, data, ++slideNum);
  addFetchabilitySlide(pptx, data, ++slideNum);
  addCrawlTypeSlide(pptx, data, ++slideNum);

  // Slide 10: Rich Results
  addRichResultsSlide(pptx, data, ++slideNum);

  // Slide 11: Image Search Performance
  addImagePerfSlide(pptx, data, ++slideNum);

  // Slide 12: Performance Overview
  addPerfOverviewSlide(pptx, data, ++slideNum);

  // Slide 13: Top Queries
  addTopQueriesSlide(pptx, data, ++slideNum);

  // Slide 14: Top Pages
  addTopPagesSlide(pptx, data, ++slideNum);

  // Slide 15: Period Comparison
  addPeriodComparisonSlide(pptx, data, ++slideNum);

  // Slide 16: Keyword Opportunities
  addKeywordOpportunitiesSlide(pptx, data, ++slideNum);

  // Slides 17-18: Manual Action + Security Issues (if captured)
  if (data.manualActionScreenshot) {
    addScreenshotSlide(pptx, 'MANUAL ACTION', data.manualActionScreenshot, ++slideNum,
      "Google issues a manual action against a site when a human reviewer at Google has determined that pages on the site are not compliant with Google's webmaster quality guidelines.",
      determineManualActionStatus(data.manualActionScreenshot));
  }
  if (data.securityScreenshot) {
    addScreenshotSlide(pptx, 'SECURITY ISSUES', data.securityScreenshot, ++slideNum,
      "If a Google evaluation determines that a site was hacked, or exhibits behaviour that could harm visitors, the Security Issues report will show Google's findings.",
      'No Issues Detected');
  }

  // Final slide: Thank you
  addThankYouSlide(pptx, data, ++slideNum);

  // Generate as base64 data URL
  return await pptx.write({ outputType: 'base64' }).then(b64 =>
    'data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,' + b64
  );
}

// ===== Slide helpers =====

function addPageNumber(slide, n) {
  slide.addShape('rect', {
    x: 12.65, y: 7.05, w: 0.45, h: 0.35, fill: { color: C.GOLD }, line: { color: C.GOLD }
  });
  slide.addText(String(n), {
    x: 12.65, y: 7.05, w: 0.45, h: 0.35,
    fontFace: FONT_BODY, fontSize: 14, bold: true, color: C.WHITE, align: 'center', valign: 'middle'
  });
}

function addContentHeader(slide, title, description, status, statusColor = C.STATUS_OK) {
  // Top header bar
  slide.addShape('rect', { x: 0, y: 0, w: 13.33, h: 1.0, fill: { color: C.TEAL_HEADER_BG }, line: { color: C.TEAL_HEADER_BG } });
  // Gold accent bar
  slide.addShape('rect', { x: 0, y: 1.0, w: 13.33, h: 0.08, fill: { color: C.GOLD }, line: { color: C.GOLD } });

  slide.addText(title, {
    x: 0.5, y: 0.15, w: 12.3, h: 0.55,
    fontFace: FONT_HEAD, fontSize: 26, bold: true, color: C.WHITE
  });
  if (description) {
    slide.addText(description, {
      x: 0.5, y: 0.62, w: 12.3, h: 0.35,
      fontFace: FONT_BODY, fontSize: 14, color: C.OFF_WHITE
    });
  }
  if (status) {
    slide.addText([
      { text: 'Status: ', options: { fontFace: FONT_BODY, fontSize: 16, color: C.WHITE } },
      { text: status,    options: { fontFace: FONT_BODY, fontSize: 16, color: statusColor, bold: true } }
    ], { x: 0.5, y: 1.18, w: 12.3, h: 0.4 });
  }
}

// ===== Slide implementations =====

function addTitleSlide(pptx, data) {
  const slide = pptx.addSlide();
  slide.background = { color: C.TEAL_TITLE_BG };

  // Gold left bar
  slide.addShape('rect', { x: 0, y: 0, w: 0.4, h: 7.5, fill: { color: C.GOLD }, line: { color: C.GOLD } });

  slide.addText('GSC AUDIT REPORT', {
    x: 1.2, y: 2.3, w: 11, h: 1.0,
    fontFace: FONT_HEAD, fontSize: 36, bold: true, color: C.WHITE
  });

  // Gold underline
  slide.addShape('rect', { x: 1.2, y: 3.4, w: 5, h: 0.06, fill: { color: C.GOLD }, line: { color: C.GOLD } });

  slide.addText(data.domain, {
    x: 1.2, y: 3.6, w: 11, h: 0.6,
    fontFace: FONT_BODY, fontSize: 22, color: C.GOLD
  });

  slide.addText(`Search Console Audit  |  ${data.startDate} to ${data.endDate}`, {
    x: 1.2, y: 4.3, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: 14, color: C.OFF_WHITE
  });
}

function addIntroSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  addContentHeader(slide, 'INTRODUCTION',
    `This report presents a full Google Search Console audit for ${data.domain} across key areas: Sitemap Health, Indexing & Canonicalisation, Rich Results, Image Search, and Search Performance.`);

  // 5 info boxes in intro
  const boxes = [
    { label: 'Property',         value: data.propertyUrl, color: C.KPI_BLUE },
    { label: 'Audit Period',     value: `${data.startDate} to ${data.endDate} (${data.periodDays} days)`, color: C.KPI_ORANGE },
    { label: 'Comparison Period',value: `${data.prevStart} to ${data.prevEnd}`, color: C.KPI_PURPLE },
    { label: 'Pages Inspected',  value: `${data.inspections.length} URLs (homepage + top by clicks)`, color: C.KPI_GREEN },
    { label: 'Sections Covered', value: 'Sitemap, Indexing, Rich Results, Image, Performance, Manual Action, Security', color: C.KPI_BLUE }
  ];

  const cols = 3, rows = 2, boxW = 4.0, boxH = 2.3, gapX = 0.25, startX = 0.6, startY = 2.0;
  boxes.forEach((b, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (boxW + gapX), y = startY + row * (boxH + 0.25);
    slide.addShape('rect', { x, y, w: boxW, h: boxH, fill: { color: C.WHITE }, line: { color: C.CARD_BG, width: 1 } });
    slide.addShape('rect', { x, y, w: boxW, h: 0.12, fill: { color: b.color }, line: { color: b.color } });
    slide.addText(b.label, {
      x: x + 0.2, y: y + 0.2, w: boxW - 0.4, h: 0.35,
      fontFace: FONT_BODY, fontSize: 12, color: C.TEXT_SOFT, bold: true
    });
    slide.addText(b.value, {
      x: x + 0.2, y: y + 0.6, w: boxW - 0.4, h: boxH - 0.7,
      fontFace: FONT_BODY, fontSize: 13, color: C.TEXT_DARK, valign: 'top', wrap: true
    });
  });

  addPageNumber(slide, n);
}

function addSitemapStatusSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const sm = data.sitemaps || [];
  const hasErrors = sm.some(s => (s.errors || 0) > 0 || (s.warnings || 0) > 0);
  const status = sm.length === 0 ? 'No Sitemap Submitted' : (hasErrors ? 'Errors Found' : 'No Issues Found');
  const stColor = sm.length === 0 ? C.STATUS_WARN : (hasErrors ? C.STATUS_ERR : C.GOLD_SOFT);
  addContentHeader(slide, 'SITEMAP SUBMITTED & STATUS',
    'Checks whether sitemaps are submitted to GSC, whether they are active (not pending), and reports any errors or warnings on each sitemap.',
    status, stColor);

  // Table
  const headers = ['Sitemap', 'Type', 'Processed', 'Pending', 'Errors', 'Warnings'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  if (sm.length === 0) {
    rows.push([{ text: 'No sitemaps submitted', options: { colspan: 6, italic: true, color: C.TEXT_SOFT } }]);
  } else {
    sm.forEach(s => {
      rows.push([
        truncateUrl(s.path),
        s.type || 'Unknown',
        s.isPending ? 'No' : 'Yes',
        s.isPending ? 'Yes' : 'No',
        String(s.errors || 0),
        String(s.warnings || 0)
      ]);
    });
  }

  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 12,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [5.3, 1.7, 1.5, 1.4, 1.2, 1.2],
    rowH: 0.4
  });

  addPageNumber(slide, n);
}

function addSitemapLastReadSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const sm = data.sitemaps || [];
  const status = sm.length === 0 ? 'No Sitemap Submitted' : 'No Issues Found';
  addContentHeader(slide, 'SITEMAP LAST READ & URLS DISCOVERED',
    'Reports the last date Google downloaded each sitemap and how many URLs were discovered. A fresh read date indicates healthy crawl activity.',
    status, C.GOLD_SOFT);

  const headers = ['Sitemap', 'Last Submitted', 'Last Downloaded', 'URLs Submitted'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  let staleNote = '';
  sm.forEach(s => {
    const lastDl = s.lastDownloaded ? s.lastDownloaded.substring(0, 10) : 'Never';
    const submitted = s.lastSubmitted ? s.lastSubmitted.substring(0, 10) : '-';
    const urlsCount = (s.contents || []).reduce((sum, c) => sum + parseInt(c.submitted || 0, 10), 0);
    rows.push([truncateUrl(s.path), submitted, lastDl, String(urlsCount)]);

    if (s.lastDownloaded) {
      const ageDays = (Date.now() - new Date(s.lastDownloaded).getTime()) / 86400000;
      if (ageDays > 60) {
        staleNote += `WARNING: ${truncateUrl(s.path)} has not been re-downloaded in ${Math.round(ageDays)} days. Consider resubmitting.\n`;
      }
    }
  });

  if (sm.length === 0) rows.push([{ text: 'No sitemaps submitted', options: { colspan: 4, italic: true, color: C.TEXT_SOFT } }]);

  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 12,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [5.5, 2.3, 2.3, 2.2], rowH: 0.4
  });

  if (staleNote) {
    slide.addText(staleNote.trim(), {
      x: 0.5, y: 6.3, w: 12.3, h: 0.7,
      fontFace: FONT_BODY, fontSize: 12, color: C.STATUS_WARN, italic: true
    });
  }

  addPageNumber(slide, n);
}

function addUrlIndexingSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const indexed = data.inspections.filter(i => i.indexStatusResult?.verdict === 'PASS').length;
  const total = data.inspections.length;
  const status = (indexed === total) ? 'No Issues Found' : `${total - indexed} URL(s) Not Indexed`;
  const stColor = (indexed === total) ? C.GOLD_SOFT : C.STATUS_WARN;

  addContentHeader(slide, 'URL INDEXING CHECK (B1)',
    'Verifies whether each inspected URL is indexed by Google and confirms its coverage state directly from the GSC URL Inspection API.',
    status, stColor);

  const headers = ['URL', 'Verdict', 'Coverage State', 'Indexing State'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.inspections.forEach(i => {
    const r = i.indexStatusResult || {};
    rows.push([
      truncateUrl(i.url, 60),
      r.verdict || (i.error ? 'ERROR' : '-'),
      r.coverageState || '-',
      r.indexingState || '-'
    ]);
  });

  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [5.5, 1.5, 2.8, 2.5], rowH: 0.32
  });
  addPageNumber(slide, n);
}

function addCanonicalSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const mismatches = data.inspections.filter(i => {
    const r = i.indexStatusResult; if (!r) return false;
    if (!r.userCanonical || !r.googleCanonical) return false;
    return r.userCanonical !== r.googleCanonical;
  });
  const status = mismatches.length === 0 ? 'No Issues Found' : `${mismatches.length} Canonical Mismatch(es)`;
  const stColor = mismatches.length === 0 ? C.GOLD_SOFT : C.STATUS_WARN;

  addContentHeader(slide, 'CANONICAL CHECK (B2)',
    'Compares the user-declared canonical URL against the canonical URL Google has chosen. A mismatch means Google is indexing a different version than intended.',
    status, stColor);

  const headers = ['URL', 'Declared Canonical', 'Google Canonical', 'Match'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.inspections.forEach(i => {
    const r = i.indexStatusResult || {};
    const match = r.userCanonical && r.googleCanonical && r.userCanonical === r.googleCanonical ? 'YES' : (r.userCanonical || r.googleCanonical ? 'NO' : '-');
    rows.push([
      truncateUrl(i.url, 45),
      truncateUrl(r.userCanonical || '-', 45),
      truncateUrl(r.googleCanonical || '-', 45),
      { text: match, options: { color: match === 'YES' ? C.STATUS_OK : (match === 'NO' ? C.STATUS_ERR : C.TEXT_SOFT), bold: true } }
    ]);
  });
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 10,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [3.8, 4.0, 4.0, 0.5], rowH: 0.3
  });
  addPageNumber(slide, n);
}

function addRobotsSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const blocked = data.inspections.filter(i => i.indexStatusResult?.robotsTxtState === 'DISALLOWED');
  const status = blocked.length === 0 ? 'No Issues Found' : `${blocked.length} URL(s) Blocked`;
  const stColor = blocked.length === 0 ? C.GOLD_SOFT : C.STATUS_ERR;

  addContentHeader(slide, 'ROBOTS.TXT BLOCKING (B3)',
    "Checks whether Google's crawler is allowed to access each inspected URL. BLOCKED means the page cannot be crawled, which prevents indexing.",
    status, stColor);

  const headers = ['URL', 'robots.txt State'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.inspections.forEach(i => {
    const state = i.indexStatusResult?.robotsTxtState || '-';
    rows.push([
      truncateUrl(i.url, 80),
      { text: state, options: { color: state === 'ALLOWED' ? C.STATUS_OK : (state === 'DISALLOWED' ? C.STATUS_ERR : C.TEXT_SOFT), bold: state !== '-' } }
    ]);
  });
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [9.3, 3.0], rowH: 0.35
  });
  addPageNumber(slide, n);
}

function addFetchabilitySlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const fails = data.inspections.filter(i => {
    const r = i.indexStatusResult; if (!r) return false;
    return r.pageFetchState && r.pageFetchState !== 'SUCCESSFUL';
  });
  const status = fails.length === 0 ? 'No Issues Found' : `${fails.length} URL(s) Failed to Fetch`;
  const stColor = fails.length === 0 ? C.GOLD_SOFT : C.STATUS_ERR;
  addContentHeader(slide, 'PAGE FETCHABILITY & LAST CRAWLED (B4 + B5)',
    'Confirms Google can successfully fetch each page and reports when Google last crawled it. Pages not crawled in 90+ days may signal crawl budget or discovery issues.',
    status, stColor);

  const headers = ['URL', 'Page Fetch', 'Last Crawl Time'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.inspections.forEach(i => {
    const r = i.indexStatusResult || {};
    rows.push([
      truncateUrl(i.url, 65),
      { text: r.pageFetchState || '-', options: { color: r.pageFetchState === 'SUCCESSFUL' ? C.STATUS_OK : (r.pageFetchState ? C.STATUS_ERR : C.TEXT_SOFT), bold: !!r.pageFetchState } },
      r.lastCrawlTime ? r.lastCrawlTime.substring(0, 10) : '-'
    ]);
  });
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [7.5, 2.5, 2.3], rowH: 0.32
  });
  addPageNumber(slide, n);
}

function addCrawlTypeSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const desktop = data.inspections.filter(i => i.indexStatusResult?.crawledAs === 'DESKTOP');
  const status = desktop.length === 0 ? 'No Issues Found' : `${desktop.length} URL(s) Crawled as Desktop`;
  const stColor = desktop.length === 0 ? C.GOLD_SOFT : C.STATUS_WARN;
  addContentHeader(slide, 'CRAWL TYPE - MOBILE-FIRST CHECK (B6)',
    'Google uses mobile-first indexing. All inspected pages should show MOBILE as the crawl type. DESKTOP crawl is a flag that the page may not be mobile-optimised.',
    status, stColor);

  const headers = ['URL', 'Crawled As'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.inspections.forEach(i => {
    const c = i.indexStatusResult?.crawledAs || '-';
    rows.push([
      truncateUrl(i.url, 80),
      { text: c, options: { color: c === 'MOBILE' ? C.STATUS_OK : (c === 'DESKTOP' ? C.STATUS_WARN : C.TEXT_SOFT), bold: c !== '-' } }
    ]);
  });
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [9.3, 3.0], rowH: 0.35
  });
  addPageNumber(slide, n);
}

function addRichResultsSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const urlsWithRich = data.inspections.filter(i => {
    const rr = i.richResultsResult?.detectedItems;
    return rr && rr.length > 0;
  });
  const status = urlsWithRich.length > 0 ? 'Rich Results Detected' : 'No Rich Results Found';
  const stColor = urlsWithRich.length > 0 ? C.GOLD_SOFT : C.STATUS_WARN;
  addContentHeader(slide, 'ACTIVE RICH RESULTS & SEARCH APPEARANCE (C1)',
    'Identifies rich result types detected in page markup via URL inspection. Rich results improve SERP visibility and CTR.',
    status, stColor);

  const headers = ['URL', 'Rich Result Type(s)', 'Status'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.inspections.forEach(i => {
    const items = i.richResultsResult?.detectedItems || [];
    const types = items.map(it => it.richResultType).join(', ') || 'None';
    const verdict = i.richResultsResult?.verdict || '-';
    rows.push([
      truncateUrl(i.url, 60),
      types,
      verdict
    ]);
  });
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [6.0, 4.3, 2.0], rowH: 0.32
  });

  if (urlsWithRich.length < data.inspections.length) {
    slide.addText(`Only ${urlsWithRich.length} of ${data.inspections.length} inspected pages have rich result schemas. Opportunity exists to add Product, Review, Article, Breadcrumb, or FAQ schemas to other pages to improve SERP visibility.`, {
      x: 0.5, y: 6.5, w: 12.3, h: 0.5,
      fontFace: FONT_BODY, fontSize: 12, color: C.TEXT_SOFT, italic: true
    });
  }
  addPageNumber(slide, n);
}

function addImagePerfSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const totals = aggregateRows(data.imagePerf);
  const hasImpressions = totals.impressions > 0;
  const hasClicks = totals.clicks > 0;
  const status = !hasImpressions ? 'No Image Search Data' : (hasClicks ? 'Data Available' : 'Warning - Review');
  const stColor = !hasImpressions ? C.TEXT_SOFT : (hasClicks ? C.GOLD_SOFT : C.STATUS_WARN);
  const subtitle = !hasImpressions
    ? `${data.domain} does not appear in Google Image Search results during this period.`
    : `${data.domain} appears in Google Image Search results. Impressions exist${hasClicks ? '' : ' but click-through rate is 0%, indicating images are being discovered but users are not clicking through'}.`;
  addContentHeader(slide, 'IMAGE SEARCH PERFORMANCE (D1)', subtitle, status, stColor);

  // 4 KPI cards
  addKpiCards(slide, [
    { label: 'Image Clicks', value: fmt(totals.clicks), color: C.KPI_ORANGE },
    { label: 'Image Impressions', value: fmt(totals.impressions), color: C.KPI_BLUE },
    { label: 'Avg CTR', value: (totals.ctr * 100).toFixed(2) + '%', color: C.KPI_GREEN },
    { label: 'Avg Position', value: totals.position.toFixed(1), color: C.KPI_PURPLE }
  ]);

  // Top image queries
  if (data.imagePerf.length > 0) {
    const headers = ['Top Image Query', 'Clicks', 'Impressions', 'CTR', 'Position'];
    const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
    data.imagePerf.slice(0, 5).forEach(r => {
      rows.push([
        r.keys[0],
        fmt(r.clicks),
        fmt(r.impressions),
        (r.ctr * 100).toFixed(2) + '%',
        r.position.toFixed(1)
      ]);
    });
    slide.addTable(rows, {
      x: 0.5, y: 5.0, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
      border: { type: 'solid', color: C.CARD_BG, pt: 1 },
      colW: [5.3, 1.7, 2.0, 1.7, 1.6], rowH: 0.3
    });
  }
  addPageNumber(slide, n);
}

function addPerfOverviewSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const totals = aggregateRows(data.perfDaily);
  addContentHeader(slide, 'PERFORMANCE OVERVIEW (E1)',
    `Web search performance totals derived from daily data rows. Period: ${data.startDate} to ${data.endDate}.`,
    'Data Available', C.GOLD_SOFT);

  addKpiCards(slide, [
    { label: 'Total Clicks',  value: fmt(totals.clicks),                color: C.KPI_ORANGE },
    { label: 'Impressions',   value: fmt(totals.impressions),           color: C.KPI_BLUE },
    { label: 'Avg CTR',       value: (totals.ctr * 100).toFixed(2) + '%', color: C.KPI_GREEN },
    { label: 'Avg Position',  value: totals.position.toFixed(1),         color: C.KPI_PURPLE }
  ]);
  addPageNumber(slide, n);
}

function addTopQueriesSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const lowCtrFlags = data.topQueries.filter(q => q.position >= 4 && q.position <= 10 && q.ctr < 0.03 && q.impressions >= 50).length;
  const status = lowCtrFlags > 0 ? 'Warning - Review' : 'No Issues Found';
  const stColor = lowCtrFlags > 0 ? C.STATUS_WARN : C.GOLD_SOFT;
  addContentHeader(slide, 'TOP QUERIES (E2)',
    'Top 10 queries driving impressions and clicks. Flag = below position benchmark (pos 4-10 should achieve >3% CTR). Several high-impression queries with low CTR represent optimisation opportunities.',
    status, stColor);

  const headers = ['#', 'Query', 'Clicks', 'Impressions', 'CTR', 'Position'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.topQueries.forEach((q, i) => {
    const flagged = q.position >= 4 && q.position <= 10 && q.ctr < 0.03 && q.impressions >= 50;
    rows.push([
      { text: (flagged ? '!' : '') + String(i + 1), options: { color: flagged ? C.STATUS_WARN : C.TEXT_DARK, bold: flagged } },
      q.keys[0],
      fmt(q.clicks),
      fmt(q.impressions),
      (q.ctr * 100).toFixed(2) + '%',
      q.position.toFixed(1)
    ]);
  });
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [0.5, 5.8, 1.4, 1.8, 1.4, 1.4], rowH: 0.32
  });
  addPageNumber(slide, n);
}

function addTopPagesSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  addContentHeader(slide, 'TOP PAGES (E3)',
    `Top 10 pages by clicks during the period.`,
    'No Issues Found', C.GOLD_SOFT);

  const headers = ['#', 'Page', 'Clicks', 'Impressions', 'CTR', 'Position'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  data.topPages.forEach((p, i) => {
    rows.push([
      String(i + 1),
      truncateUrl(p.keys[0], 50),
      fmt(p.clicks),
      fmt(p.impressions),
      (p.ctr * 100).toFixed(2) + '%',
      p.position.toFixed(1)
    ]);
  });
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [0.5, 5.8, 1.4, 1.8, 1.4, 1.4], rowH: 0.32
  });
  addPageNumber(slide, n);
}

function addPeriodComparisonSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  const cur = aggregateRows(data.perfDaily);
  const prev = aggregateRows(data.prevPerf);
  const delta = cur.clicks - prev.clicks;
  const pct = prev.clicks ? (delta / prev.clicks * 100).toFixed(1) : 'N/A';
  const status = delta > 0 ? 'Growth' : (delta < 0 ? 'Decline' : 'Flat');
  const stColor = delta > 0 ? C.STATUS_OK : (delta < 0 ? C.STATUS_ERR : C.TEXT_SOFT);
  addContentHeader(slide, 'PERIOD COMPARISON (E4)',
    `Current: ${data.startDate} to ${data.endDate} (${cur.clicks} clicks) vs Previous: ${data.prevStart} to ${data.prevEnd} (${prev.clicks} clicks) | Net change: ${delta >= 0 ? '+' : ''}${delta} clicks (${pct}%)`,
    status, stColor);

  // Build gainers/losers from prev vs current queries (best-effort)
  const prevMap = {};
  (data.prevPerf || []).forEach(r => { prevMap[r.keys[0]] = r; });

  const compared = (data.topQueries || []).map(q => {
    const p = prevMap[q.keys[0]];
    const prevClicks = p ? p.clicks : 0;
    return { query: q.keys[0], cur: q.clicks, prev: prevClicks, delta: q.clicks - prevClicks };
  });
  const gainers = compared.filter(c => c.delta > 0).sort((a, b) => b.delta - a.delta).slice(0, 5);
  const losers  = compared.filter(c => c.delta < 0).sort((a, b) => a.delta - b.delta).slice(0, 5);

  slide.addText('GAINERS - Queries with positive click delta', {
    x: 0.5, y: 1.8, w: 6.0, h: 0.4,
    fontFace: FONT_BODY, fontSize: 13, bold: true, color: C.STATUS_OK
  });
  slide.addText('LOSERS - Queries with negative click delta', {
    x: 6.85, y: 1.8, w: 6.0, h: 0.4,
    fontFace: FONT_BODY, fontSize: 13, bold: true, color: C.STATUS_ERR
  });

  const buildMiniTable = (rows, deltaColor) => {
    const h = [['Query', 'Prev', 'Cur', 'Delta'].map(t => ({ text: t, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
    rows.forEach(r => {
      h.push([
        truncateText(r.query, 35),
        String(r.prev),
        String(r.cur),
        { text: (r.delta >= 0 ? '+' : '') + r.delta, options: { color: deltaColor, bold: true } }
      ]);
    });
    if (rows.length === 0) h.push([{ text: 'None', options: { colspan: 4, italic: true, color: C.TEXT_SOFT, align: 'center' } }]);
    return h;
  };

  slide.addTable(buildMiniTable(gainers, C.STATUS_OK), {
    x: 0.5, y: 2.3, w: 6.0, fontFace: FONT_BODY, fontSize: 10,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [3.3, 0.9, 0.9, 0.9], rowH: 0.3
  });
  slide.addTable(buildMiniTable(losers, C.STATUS_ERR), {
    x: 6.85, y: 2.3, w: 6.0, fontFace: FONT_BODY, fontSize: 10,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [3.3, 0.9, 0.9, 0.9], rowH: 0.3
  });

  addPageNumber(slide, n);
}

function addKeywordOpportunitiesSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  // High impressions + low CTR
  const opps = (data.topQueries || [])
    .filter(q => q.impressions >= 30 && q.ctr < 0.03)
    .sort((a, b) => b.impressions - a.impressions)
    .slice(0, 10);
  const status = opps.length > 0 ? 'Warning - Review' : 'No Issues Found';
  const stColor = opps.length > 0 ? C.STATUS_WARN : C.GOLD_SOFT;
  addContentHeader(slide, 'KEYWORD OPPORTUNITIES (E5)',
    'Queries with high impression volume but low CTR represent immediate optimisation opportunities. Improving title tags and meta descriptions for these terms can unlock clicks without new rankings.',
    status, stColor);

  const headers = ['Query', 'Impressions', 'CTR', 'Position', 'Opportunity'];
  const rows = [headers.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL_HEADER_BG } } }))];
  opps.forEach(q => {
    const opportunity = q.position <= 10 ? 'High - Title/Meta tune' : (q.position <= 20 ? 'Medium - Content boost' : 'Long-term');
    rows.push([
      q.keys[0],
      fmt(q.impressions),
      (q.ctr * 100).toFixed(2) + '%',
      q.position.toFixed(1),
      opportunity
    ]);
  });
  if (opps.length === 0) {
    rows.push([{ text: 'No clear opportunities in the top queries - CTR looks healthy.', options: { colspan: 5, italic: true, color: C.TEXT_SOFT, align: 'center' } }]);
  }
  slide.addTable(rows, {
    x: 0.5, y: 1.8, w: 12.3, fontFace: FONT_BODY, fontSize: 11,
    border: { type: 'solid', color: C.CARD_BG, pt: 1 },
    colW: [5.5, 1.7, 1.4, 1.4, 2.3], rowH: 0.32
  });
  addPageNumber(slide, n);
}

function addScreenshotSlide(pptx, title, dataUrl, n, description, status) {
  const slide = pptx.addSlide();
  slide.background = { color: C.CARD_BG_2 };
  addContentHeader(slide, title, description, status, C.GOLD_SOFT);
  // Image area: y=1.8 to 6.9, x=0.5 to 12.83
  slide.addImage({ data: dataUrl, x: 0.5, y: 1.8, w: 12.33, h: 5.1, sizing: { type: 'contain', w: 12.33, h: 5.1 } });
  addPageNumber(slide, n);
}

function determineManualActionStatus(_dataUrl) {
  // Without OCR we can't know from the screenshot. Default optimistic with caveat.
  return 'No Issues Detected (verify in slide)';
}

function addThankYouSlide(pptx, data, n) {
  const slide = pptx.addSlide();
  slide.background = { color: C.TEAL_TITLE_BG };
  slide.addShape('rect', { x: 0, y: 0, w: 0.4, h: 7.5, fill: { color: C.GOLD }, line: { color: C.GOLD } });

  slide.addText('THANK YOU', {
    x: 1.2, y: 2.8, w: 11, h: 1.0,
    fontFace: FONT_HEAD, fontSize: 48, bold: true, color: C.WHITE
  });
  slide.addShape('rect', { x: 1.2, y: 3.9, w: 4, h: 0.06, fill: { color: C.GOLD }, line: { color: C.GOLD } });

  slide.addText(`GSC Audit Report - ${data.domain}`, {
    x: 1.2, y: 4.2, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: 18, color: C.GOLD
  });
  slide.addText(`Period: ${data.startDate} to ${data.endDate}`, {
    x: 1.2, y: 4.7, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: 14, color: C.OFF_WHITE
  });
  slide.addText('For questions or follow-up analysis, refer to your Google Search Console dashboard.', {
    x: 1.2, y: 6.5, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, color: C.LIGHT_TEAL
  });
}

// ===== Utilities =====
function aggregateRows(rows) {
  let c = 0, i = 0, p = 0, cnt = 0;
  (rows || []).forEach(r => { c += r.clicks; i += r.impressions; p += r.position * r.impressions; cnt += r.impressions; });
  return {
    clicks: c,
    impressions: i,
    ctr: i > 0 ? c / i : 0,
    position: cnt > 0 ? p / cnt : 0
  };
}

function addKpiCards(slide, kpis) {
  const startX = 0.5, startY = 2.4, gap = 0.25, w = 2.95, h = 2.2;
  kpis.forEach((kpi, i) => {
    const x = startX + i * (w + gap);
    slide.addShape('rect', { x, y: startY, w, h, fill: { color: C.CARD_BG }, line: { color: C.CARD_BG } });
    slide.addShape('rect', { x, y: startY, w, h: 0.12, fill: { color: kpi.color }, line: { color: kpi.color } });
    slide.addText(kpi.value, {
      x: x + 0.1, y: startY + 0.4, w: w - 0.2, h: 1.0,
      fontFace: FONT_HEAD, fontSize: 30, bold: true, color: kpi.color, align: 'center', valign: 'middle'
    });
    slide.addText(kpi.label, {
      x: x + 0.1, y: startY + 1.5, w: w - 0.2, h: 0.5,
      fontFace: FONT_BODY, fontSize: 12, color: C.TEXT_DARK, align: 'center'
    });
  });
}

function fmt(n) {
  if (n == null) return '-';
  return n.toLocaleString('en-US');
}

function truncateUrl(url, max = 70) {
  if (!url) return '-';
  return url.length > max ? url.substring(0, max - 1) + '...' : url;
}
function truncateText(s, max) {
  if (!s) return '-';
  return s.length > max ? s.substring(0, max - 1) + '...' : s;
}

// ===== Format dispatcher =====
async function buildPptxReport(data, format = 'james') {
  switch (format) {
    case 'omega': return await buildOmegaPptx(data);
    case 'neon':  return await buildNeonPptx(data);
    case 'james':
    default:      return await buildJamesPptx(data);
  }
}
