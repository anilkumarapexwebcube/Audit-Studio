/* GSC Audit Studio - Omega Format Builder
 * 7 slides, screenshot-driven, corporate/professional styling.
 * Mirrors the structure of the Omega DOCX webmaster audit report.
 */

const OMEGA = {
  NAVY:       '1B2A41',
  NAVY_DARK:  '0F1A2C',
  BRONZE:     'B08D57',
  BRONZE_SOFT:'D4B891',
  OFF_WHITE:  'F8F7F4',
  CARD_BG:    'FFFFFF',
  TEXT_DARK:  '2C2C2C',
  TEXT_SOFT:  '6B7280',
  BORDER:     'E5E0D5',
  STATUS_OK:  '4A7C59',
  STATUS_WARN:'B08D57',
  STATUS_ERR: '8B3A3A'
};
const OMEGA_HEAD = 'Georgia';
const OMEGA_BODY = 'Calibri';

async function buildOmegaPptx(data) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  let n = 1;

  // Slide 1: Title
  {
    const slide = pptx.addSlide();
    slide.background = { color: OMEGA.OFF_WHITE };
    // Top bronze bar
    slide.addShape('rect', { x: 0, y: 0, w: 13.33, h: 0.5, fill: { color: OMEGA.BRONZE }, line: { color: OMEGA.BRONZE } });
    // Bottom navy bar
    slide.addShape('rect', { x: 0, y: 7.0, w: 13.33, h: 0.5, fill: { color: OMEGA.NAVY }, line: { color: OMEGA.NAVY } });

    slide.addText('WEBMASTER AUDIT REPORT', {
      x: 1.5, y: 2.5, w: 10.3, h: 1.2,
      fontFace: OMEGA_HEAD, fontSize: 40, bold: true, color: OMEGA.NAVY, align: 'center'
    });
    // Decorative line
    slide.addShape('rect', { x: 5.66, y: 3.7, w: 2.0, h: 0.04, fill: { color: OMEGA.BRONZE }, line: { color: OMEGA.BRONZE } });

    slide.addText(data.domain, {
      x: 1.5, y: 4.0, w: 10.3, h: 0.6,
      fontFace: OMEGA_BODY, fontSize: 24, color: OMEGA.TEXT_DARK, align: 'center', italic: true
    });
    slide.addText(`Audit Date: ${data.endDate}`, {
      x: 1.5, y: 4.8, w: 10.3, h: 0.4,
      fontFace: OMEGA_BODY, fontSize: 14, color: OMEGA.TEXT_SOFT, align: 'center'
    });
  }

  // Slide 2: Introduction
  {
    const slide = pptx.addSlide();
    slide.background = { color: OMEGA.OFF_WHITE };
    omegaHeader(slide, 'INTRODUCTION');
    slide.addText('Reviewing the search console for the website is one of the major aspects of our work. We check the webmaster for all the important parameters, including sitemap health, manual actions, performance, and security issues, to confirm that the website is functioning correctly and free of errors.', {
      x: 1.0, y: 2.0, w: 11.3, h: 2.0,
      fontFace: OMEGA_BODY, fontSize: 16, color: OMEGA.TEXT_DARK, valign: 'top', paraSpaceAfter: 8
    });
    slide.addText(`Domain: ${data.domain}`, {
      x: 1.0, y: 4.5, w: 11.3, h: 0.5,
      fontFace: OMEGA_BODY, fontSize: 18, bold: true, color: OMEGA.NAVY
    });
    omegaFooter(slide, n);
  }
  n++;

  // Slide 3: Sitemap
  omegaScreenshotSlide(pptx, 'CHECK ERROR IN SITEMAP',
    'A sitemap is a record on your site that uncovers to Google which pages we should consider.',
    data.sitemapScreenshot, ++n);

  // Slide 4: Manual Action
  omegaScreenshotSlide(pptx, 'CHECK MANUAL ACTION',
    'Manual action is a penalty imposed by Google for not following the webmaster guidelines.',
    data.manualActionScreenshot, ++n);

  // Slide 5: Performance Issue
  omegaScreenshotSlide(pptx, 'CHECK PERFORMANCE ISSUE',
    'The Performance Report shows important metrics about how your site performs in Google Search Results.',
    data.performanceScreenshot, ++n);

  // Slide 6: Security Issue
  omegaScreenshotSlide(pptx, 'CHECK SECURITY ISSUE',
    'The Security Issues report in Search Console alerts webmasters about malicious behaviour on their websites.',
    data.securityScreenshot, ++n);

  // Slide 7: Other Issues
  {
    const slide = pptx.addSlide();
    slide.background = { color: OMEGA.OFF_WHITE };
    omegaHeader(slide, 'CHECK OTHER ISSUE IF IN WEBMASTER');
    slide.addText('Reviewed the Removals report and any additional Search Console alerts for the property.', {
      x: 1.0, y: 1.8, w: 11.3, h: 0.6,
      fontFace: OMEGA_BODY, fontSize: 14, color: OMEGA.TEXT_SOFT, italic: true
    });
    // Centred status statement
    slide.addShape('rect', { x: 3.5, y: 3.5, w: 6.3, h: 1.2, fill: { color: OMEGA.CARD_BG }, line: { color: OMEGA.BRONZE, width: 2 } });
    slide.addText('Not Found', {
      x: 3.5, y: 3.5, w: 6.3, h: 1.2,
      fontFace: OMEGA_HEAD, fontSize: 32, bold: true, color: OMEGA.STATUS_OK, align: 'center', valign: 'middle'
    });
    slide.addText('No additional issues detected in the Search Console webmaster reports for this property.', {
      x: 1.0, y: 5.0, w: 11.3, h: 0.6,
      fontFace: OMEGA_BODY, fontSize: 14, color: OMEGA.TEXT_DARK, align: 'center'
    });
    omegaFooter(slide, n);
  }
  n++;

  // Slide 8: Thank you
  {
    const slide = pptx.addSlide();
    slide.background = { color: OMEGA.NAVY };
    slide.addShape('rect', { x: 0, y: 3.4, w: 13.33, h: 0.04, fill: { color: OMEGA.BRONZE }, line: { color: OMEGA.BRONZE } });
    slide.addText('THANK YOU', {
      x: 1.5, y: 2.5, w: 10.3, h: 1.0,
      fontFace: OMEGA_HEAD, fontSize: 56, bold: true, color: OMEGA.BRONZE_SOFT, align: 'center', charSpacing: 8
    });
    slide.addText(data.domain, {
      x: 1.5, y: 3.7, w: 10.3, h: 0.5,
      fontFace: OMEGA_BODY, fontSize: 18, color: OMEGA.OFF_WHITE, align: 'center'
    });
  }

  return await pptx.write({ outputType: 'base64' }).then(b64 =>
    'data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,' + b64
  );
}

function omegaHeader(slide, title) {
  slide.addShape('rect', { x: 0, y: 0, w: 13.33, h: 1.2, fill: { color: OMEGA.NAVY }, line: { color: OMEGA.NAVY } });
  slide.addShape('rect', { x: 0, y: 1.2, w: 13.33, h: 0.06, fill: { color: OMEGA.BRONZE }, line: { color: OMEGA.BRONZE } });
  slide.addText(title, {
    x: 1.0, y: 0.25, w: 11.3, h: 0.7,
    fontFace: OMEGA_HEAD, fontSize: 28, bold: true, color: OMEGA.OFF_WHITE, charSpacing: 2
  });
}

function omegaFooter(slide, n) {
  slide.addShape('rect', { x: 0, y: 7.1, w: 13.33, h: 0.4, fill: { color: OMEGA.NAVY }, line: { color: OMEGA.NAVY } });
  slide.addText(`Page ${n}`, {
    x: 0, y: 7.1, w: 13.33, h: 0.4,
    fontFace: OMEGA_BODY, fontSize: 10, color: OMEGA.BRONZE_SOFT, align: 'center'
  });
}

function omegaScreenshotSlide(pptx, title, description, dataUrl, n) {
  const slide = pptx.addSlide();
  slide.background = { color: OMEGA.OFF_WHITE };
  omegaHeader(slide, title);
  slide.addText(description, {
    x: 1.0, y: 1.5, w: 11.3, h: 0.6,
    fontFace: OMEGA_BODY, fontSize: 14, color: OMEGA.TEXT_SOFT, italic: true
  });
  // Screenshot area with subtle border
  if (dataUrl) {
    slide.addShape('rect', { x: 0.7, y: 2.15, w: 11.93, h: 4.85, fill: { color: OMEGA.CARD_BG }, line: { color: OMEGA.BORDER, width: 1 } });
    slide.addImage({ data: dataUrl, x: 0.8, y: 2.25, w: 11.73, h: 4.65, sizing: { type: 'contain', w: 11.73, h: 4.65 } });
  } else {
    slide.addShape('rect', { x: 0.7, y: 2.15, w: 11.93, h: 4.85, fill: { color: OMEGA.CARD_BG }, line: { color: OMEGA.BORDER, width: 1 } });
    slide.addText('Screenshot not captured.', {
      x: 0.7, y: 2.15, w: 11.93, h: 4.85,
      fontFace: OMEGA_BODY, fontSize: 18, color: OMEGA.TEXT_SOFT, align: 'center', valign: 'middle', italic: true
    });
  }
  omegaFooter(slide, n);
}
