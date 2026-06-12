/* GSC Audit Studio - Neon Format Builder
 * 7 slides, screenshot-driven, modern/vibrant styling with hybrid dark+bright palette.
 */

const NEON = {
  DARK:        '0A0A14',
  DARK_SOFT:   '1A1A2E',
  PINK:        'FF0099',
  PINK_SOFT:   'FF66C2',
  CYAN:        '00FFE1',
  CYAN_SOFT:   '66FFEF',
  WHITE:       'FFFFFF',
  OFF_WHITE:   'F0F0F8',
  CARD_BG:     'FAFAFC',
  TEXT_DARK:   '1A1A2E',
  TEXT_SOFT:   '4A4A6A',
  BORDER:      'E0E0F0',
  STATUS_OK:   '00CC88'
};
const NEON_HEAD = 'Trebuchet MS';
const NEON_BODY = 'Calibri';

async function buildNeonPptx(data) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  let n = 1;

  // Slide 1: Title (dark background, neon pink + cyan)
  {
    const slide = pptx.addSlide();
    slide.background = { color: NEON.DARK };
    // Pink accent left bar
    slide.addShape('rect', { x: 0, y: 0, w: 0.5, h: 7.5, fill: { color: NEON.PINK }, line: { color: NEON.PINK } });
    // Cyan accent bottom bar
    slide.addShape('rect', { x: 0.5, y: 7.3, w: 12.83, h: 0.2, fill: { color: NEON.CYAN }, line: { color: NEON.CYAN } });

    slide.addText('WEBMASTER', {
      x: 1.2, y: 2.0, w: 11, h: 0.9,
      fontFace: NEON_HEAD, fontSize: 48, bold: true, color: NEON.WHITE, charSpacing: 8
    });
    slide.addText('AUDIT REPORT', {
      x: 1.2, y: 2.9, w: 11, h: 0.9,
      fontFace: NEON_HEAD, fontSize: 48, bold: true, color: NEON.PINK, charSpacing: 8
    });
    // Cyan underline
    slide.addShape('rect', { x: 1.2, y: 4.0, w: 3.5, h: 0.08, fill: { color: NEON.CYAN }, line: { color: NEON.CYAN } });

    slide.addText(data.domain, {
      x: 1.2, y: 4.3, w: 11, h: 0.6,
      fontFace: NEON_BODY, fontSize: 24, color: NEON.CYAN
    });
    slide.addText(`Audit Date: ${data.endDate}`, {
      x: 1.2, y: 5.0, w: 11, h: 0.4,
      fontFace: NEON_BODY, fontSize: 14, color: NEON.OFF_WHITE
    });
  }

  // Slide 2: Introduction (white background, dark text, neon accents)
  {
    const slide = pptx.addSlide();
    slide.background = { color: NEON.OFF_WHITE };
    neonHeader(slide, 'INTRODUCTION');
    slide.addText('Reviewing the search console for the website is one of the major aspects of our work. We check the webmaster for all the important parameters, including sitemap, manual action, performance issues, and security issues, to confirm that the website is functioning correctly without errors.', {
      x: 1.0, y: 2.0, w: 11.3, h: 2.0,
      fontFace: NEON_BODY, fontSize: 16, color: NEON.TEXT_DARK, valign: 'top'
    });
    // Domain card
    slide.addShape('rect', { x: 1.0, y: 4.5, w: 11.3, h: 0.9, fill: { color: NEON.DARK }, line: { color: NEON.DARK } });
    slide.addShape('rect', { x: 1.0, y: 4.5, w: 0.15, h: 0.9, fill: { color: NEON.PINK }, line: { color: NEON.PINK } });
    slide.addText([
      { text: 'Domain: ', options: { fontFace: NEON_BODY, fontSize: 18, color: NEON.CYAN } },
      { text: data.domain, options: { fontFace: NEON_HEAD, fontSize: 22, bold: true, color: NEON.WHITE } }
    ], { x: 1.3, y: 4.5, w: 11.0, h: 0.9, valign: 'middle' });
    neonFooter(slide, n);
  }
  n++;

  // Slide 3: Sitemap
  neonScreenshotSlide(pptx, 'CHECK ERROR IN SITEMAP',
    'A sitemap is a record on your site that uncovers to Google which pages we should consider.',
    data.sitemapScreenshot, ++n);

  // Slide 4: Manual Action
  neonScreenshotSlide(pptx, 'CHECK MANUAL ACTION',
    'Manual action is a penalty imposed by Google for not following the webmaster guidelines.',
    data.manualActionScreenshot, ++n);

  // Slide 5: Performance Issue
  neonScreenshotSlide(pptx, 'CHECK PERFORMANCE ISSUE',
    'The Performance Report shows important metrics about how your site performs in Google Search Results.',
    data.performanceScreenshot, ++n);

  // Slide 6: Security Issue
  neonScreenshotSlide(pptx, 'CHECK SECURITY ISSUE',
    'The Security Issues report in Search Console alerts webmasters about malicious behaviour on their websites.',
    data.securityScreenshot, ++n);

  // Slide 7: Other Issues
  {
    const slide = pptx.addSlide();
    slide.background = { color: NEON.OFF_WHITE };
    neonHeader(slide, 'CHECK OTHER ISSUE IF IN WEBMASTER');
    slide.addText('Reviewed the Removals report and additional Search Console alerts.', {
      x: 1.0, y: 1.8, w: 11.3, h: 0.6,
      fontFace: NEON_BODY, fontSize: 14, color: NEON.TEXT_SOFT, italic: true
    });
    // Big "Not Found" badge
    slide.addShape('rect', { x: 4.5, y: 3.4, w: 4.3, h: 1.4, fill: { color: NEON.DARK }, line: { color: NEON.DARK } });
    slide.addShape('rect', { x: 4.5, y: 3.4, w: 4.3, h: 0.1, fill: { color: NEON.CYAN }, line: { color: NEON.CYAN } });
    slide.addText('NOT FOUND', {
      x: 4.5, y: 3.4, w: 4.3, h: 1.4,
      fontFace: NEON_HEAD, fontSize: 36, bold: true, color: NEON.CYAN, align: 'center', valign: 'middle', charSpacing: 6
    });
    slide.addText('No additional issues detected.', {
      x: 1.0, y: 5.1, w: 11.3, h: 0.5,
      fontFace: NEON_BODY, fontSize: 14, color: NEON.TEXT_DARK, align: 'center'
    });
    neonFooter(slide, n);
  }
  n++;

  // Slide 8: Thank You
  {
    const slide = pptx.addSlide();
    slide.background = { color: NEON.DARK };
    slide.addShape('rect', { x: 0, y: 0, w: 0.5, h: 7.5, fill: { color: NEON.PINK }, line: { color: NEON.PINK } });
    slide.addShape('rect', { x: 12.83, y: 0, w: 0.5, h: 7.5, fill: { color: NEON.CYAN }, line: { color: NEON.CYAN } });

    slide.addText('THANK', {
      x: 1.2, y: 2.0, w: 11, h: 1.2,
      fontFace: NEON_HEAD, fontSize: 64, bold: true, color: NEON.WHITE, align: 'center', charSpacing: 12
    });
    slide.addText('YOU', {
      x: 1.2, y: 3.3, w: 11, h: 1.2,
      fontFace: NEON_HEAD, fontSize: 64, bold: true, color: NEON.PINK, align: 'center', charSpacing: 12
    });
    slide.addShape('rect', { x: 5.66, y: 4.7, w: 2.0, h: 0.06, fill: { color: NEON.CYAN }, line: { color: NEON.CYAN } });

    slide.addText(data.domain, {
      x: 1.2, y: 5.0, w: 11, h: 0.5,
      fontFace: NEON_BODY, fontSize: 18, color: NEON.CYAN_SOFT, align: 'center'
    });
  }

  return await pptx.write({ outputType: 'base64' }).then(b64 =>
    'data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,' + b64
  );
}

function neonHeader(slide, title) {
  slide.addShape('rect', { x: 0, y: 0, w: 13.33, h: 1.2, fill: { color: NEON.DARK }, line: { color: NEON.DARK } });
  // Left pink + right cyan accent edges
  slide.addShape('rect', { x: 0, y: 0, w: 0.15, h: 1.2, fill: { color: NEON.PINK }, line: { color: NEON.PINK } });
  slide.addShape('rect', { x: 0, y: 1.2, w: 13.33, h: 0.08, fill: { color: NEON.CYAN }, line: { color: NEON.CYAN } });
  slide.addText(title, {
    x: 0.6, y: 0.2, w: 12.5, h: 0.8,
    fontFace: NEON_HEAD, fontSize: 26, bold: true, color: NEON.WHITE, charSpacing: 4
  });
}

function neonFooter(slide, n) {
  slide.addShape('rect', { x: 0, y: 7.1, w: 13.33, h: 0.4, fill: { color: NEON.DARK }, line: { color: NEON.DARK } });
  slide.addShape('rect', { x: 0, y: 7.1, w: 0.15, h: 0.4, fill: { color: NEON.PINK }, line: { color: NEON.PINK } });
  slide.addText(`Page ${n}`, {
    x: 0, y: 7.1, w: 13.33, h: 0.4,
    fontFace: NEON_BODY, fontSize: 10, color: NEON.CYAN_SOFT, align: 'center'
  });
}

function neonScreenshotSlide(pptx, title, description, dataUrl, n) {
  const slide = pptx.addSlide();
  slide.background = { color: NEON.OFF_WHITE };
  neonHeader(slide, title);
  slide.addText(description, {
    x: 1.0, y: 1.5, w: 11.3, h: 0.6,
    fontFace: NEON_BODY, fontSize: 14, color: NEON.TEXT_SOFT, italic: true
  });
  // Screenshot area with pink top edge accent
  if (dataUrl) {
    slide.addShape('rect', { x: 0.6, y: 2.15, w: 12.13, h: 4.85, fill: { color: NEON.CARD_BG }, line: { color: NEON.BORDER, width: 1 } });
    slide.addShape('rect', { x: 0.6, y: 2.15, w: 12.13, h: 0.06, fill: { color: NEON.PINK }, line: { color: NEON.PINK } });
    slide.addImage({ data: dataUrl, x: 0.7, y: 2.3, w: 11.93, h: 4.6, sizing: { type: 'contain', w: 11.93, h: 4.6 } });
  } else {
    slide.addShape('rect', { x: 0.6, y: 2.15, w: 12.13, h: 4.85, fill: { color: NEON.CARD_BG }, line: { color: NEON.BORDER, width: 1 } });
    slide.addText('Screenshot not captured.', {
      x: 0.6, y: 2.15, w: 12.13, h: 4.85,
      fontFace: NEON_BODY, fontSize: 18, color: NEON.TEXT_SOFT, align: 'center', valign: 'middle', italic: true
    });
  }
  neonFooter(slide, n);
}
