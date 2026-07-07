// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// Fully programmatic layout driven by design system reference.
// All positions/sizes expressed as proportions of actual slide dimensions
// so they scale correctly regardless of template size.
//
// Design system reference canvas: 1280×720px
// Brand: #2323CC electric blue / #070809 near-black / #F2F2F2 card bg
// Fonts: Oswald (headings/display) + Inter (body)
//
// Layout types (Gemini chooses per slide):
//   statement-trio  — 3 bold standalone statements, full width
//   item-list       — heading + card items, bold label before " — "
//   timeline-list   — heading + items with left blue border bar
//   cost-table      — heading + rows with right-aligned amounts + TOTAL
//   stat-callout    — problem/context left col + giant stat right col
//   two-stat        — two large numbers side by side (evidence)
//   two-column      — heading + two equal content columns
//   next-steps      — heading + numbered action items
//
// Deploy as: Web App / Execute as Me / Who has access: Anyone
// ═══════════════════════════════════════════════════════════════════════

var TEMPLATE_SLIDE_ID = '1SYcTXUmcg3ci2pg8kWrexzwiTgHdRrF0BCAwca5dBZ0';
var OUTPUT_FOLDER_ID  = '1kTvIOM06sQh5tk2cK0697xLjs9nPyERR';

var BRAND_BLUE  = '#2323CC';
var NEAR_BLACK  = '#070809';
var CARD_BG     = '#F2F2F2';
var WHITE       = '#FFFFFF';
var MUTED       = '#555555';

var FONT_DISPLAY = 'Oswald';
var FONT_BODY    = 'Inter';

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST body received.');
    }
    var data = JSON.parse(e.postData.contents);
    var result = buildDeck(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function buildDeck(data) {
  var date         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
  var clientName   = data.Client_Name   || 'Client';
  var projectTitle = data.Project_Title || 'Deck';
  var deckName     = clientName + ' — ' + projectTitle;

  var newFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID)
    .makeCopy(deckName, DriveApp.getFolderById(OUTPUT_FOLDER_ID));
  var deck = SlidesApp.openById(newFile.getId());

  // ── Dimensions from design system (proportional to 1280×720 canvas) ─
  var W = deck.getPageWidth();
  var H = deck.getPageHeight();

  var MARGIN      = W * 0.0422;
  var CONTENT_W   = W - MARGIN * 2;
  var HEADER_H    = H * 0.1431;
  var CONTENT_TOP = H * 0.1681;
  var COL_W       = W * 0.4063;
  var COL_R_LEFT  = W * 0.5;
  var ITEM_GAP    = H * 0.039;
  var SECT_GAP    = H * 0.056;
  var STMT_GAP    = H * 0.044;

  // ── Slide 1: title slide ─────────────────────────────────────────────
  var slides = deck.getSlides();
  var titleSlide  = slides[0];
  var titleShapes = titleSlide.getShapes();
  for (var j = 0; j < titleShapes.length; j++) {
    var s = titleShapes[j];
    if (!s.getText) continue;
    var raw = s.getText().asString().trim();
    if (raw === '[[title]]' || raw.indexOf('[[title]]') !== -1) {
      s.getText().clear();
      var ts = s.getText().appendText(projectTitle).getTextStyle();
      ts.setBold(true).setFontSize(60).setFontFamily(FONT_DISPLAY);
      try { s.setLeft(MARGIN); s.setTop(H * 0.25); s.setWidth(CONTENT_W); s.setHeight(H * 0.35); } catch(e) {}
    } else if (raw === '[[date]]' || raw.indexOf('[[date]]') !== -1) {
      s.getText().clear();
      var ts2 = s.getText().appendText(date).getTextStyle();
      ts2.setBold(false).setFontSize(20).setFontFamily(FONT_BODY);
    }
  }

  // ── Adjust slide count ───────────────────────────────────────────────
  var contentItems = data.slides || [];

  slides = deck.getSlides();
  while (slides.length - 2 < contentItems.length) {
    slides[1].duplicate();
    slides = deck.getSlides();
    slides[slides.length - 1].move(slides.length - 1);
    slides = deck.getSlides();
  }
  slides = deck.getSlides();
  while (slides.length - 2 > contentItems.length) {
    slides = deck.getSlides();
    slides[slides.length - 2].remove();
    slides = deck.getSlides();
  }
  slides = deck.getSlides();

  // ── Fill each content slide ─────────────────────────────────────────
  var dims = { W: W, H: H, MARGIN: MARGIN, CONTENT_W: CONTENT_W,
               HEADER_H: HEADER_H, CONTENT_TOP: CONTENT_TOP,
               COL_W: COL_W, COL_R_LEFT: COL_R_LEFT,
               ITEM_GAP: ITEM_GAP, SECT_GAP: SECT_GAP, STMT_GAP: STMT_GAP };

  for (var i = 0; i < contentItems.length; i++) {
    var slide = slides[i + 1];
    var item  = contentItems[i];

    var existing = slide.getShapes();
    for (var k = 0; k < existing.length; k++) {
      try { if (existing[k].getTop() >= HEADER_H) existing[k].remove(); } catch(e) {}
    }

    if (item.eyebrow) {
      addPill(slide, item.eyebrow, dims);
    }

    switch (item.layout || 'item-list') {
      case 'statement-trio': layoutStatementTrio(slide, item, dims); break;
      case 'item-list':      layoutItemList(slide, item, dims);      break;
      case 'timeline-list':  layoutTimelineList(slide, item, dims);  break;
      case 'cost-table':     layoutCostTable(slide, item, dims);     break;
      case 'stat-callout':   layoutStatCallout(slide, item, dims);   break;
      case 'two-stat':       layoutTwoStat(slide, item, dims);       break;
      case 'two-column':     layoutTwoColumn(slide, item, dims);     break;
      case 'next-steps':     layoutNextSteps(slide, item, dims);     break;
      default:               layoutItemList(slide, item, dims);      break;
    }

    if (item.speaker_note) {
      try { slide.getNotesPage().getSpeakerNotesShape().getText().setText(item.speaker_note); } catch(e) {}
    }
  }

  deck.saveAndClose();
  return {
    deckUrl:  'https://docs.google.com/presentation/d/' + newFile.getId() + '/edit',
    deckName: deckName
  };
}


// ── Pill / eyebrow label ──────────────────────────────────────────────
// Blue rounded pill, top-right corner, white uppercase label
function addPill(slide, text, d) {
  var pillH = d.H * 0.055;
  var pillW = d.W * 0.18;
  var pillX = d.W - d.MARGIN - pillW;
  var pillY = d.CONTENT_TOP;
  var rect  = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, pillX, pillY, pillW, pillH);
  rect.getFill().setSolidFill(BRAND_BLUE);
  rect.getBorder().setTransparent();
  var tf = rect.getText();
  tf.setText(text.toUpperCase());
  var ts = tf.getTextStyle();
  ts.setBold(true).setFontSize(10).setForegroundColor(WHITE).setFontFamily(FONT_BODY);
  tf.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}


// ── Shared: heading in brand blue (Oswald), returns new top ──────────
function addHeading(slide, text, d) {
  if (!text) return d.CONTENT_TOP;
  var headH = d.H * 0.067;
  var tb = slide.insertTextBox(text, d.MARGIN, d.CONTENT_TOP, d.CONTENT_W * 0.75, headH);
  var ts = tb.getText().getTextStyle();
  ts.setBold(true).setFontSize(32).setForegroundColor(BRAND_BLUE).setFontFamily(FONT_DISPLAY);
  return d.CONTENT_TOP + headH + d.SECT_GAP;
}

// ── Shared: divider line ──────────────────────────────────────────────
function addDivider(slide, y, d) {
  var line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, y, d.CONTENT_W, 1);
  line.getFill().setSolidFill(NEAR_BLACK);
  line.getBorder().setTransparent();
}


// ── statement-trio ────────────────────────────────────────────────────
// 3 bold statements — alternates black / blue / black for contrast
function layoutStatementTrio(slide, item, d) {
  var lines = item.lines || [];
  var top   = d.CONTENT_TOP;
  if (item.heading) {
    var htb = slide.insertTextBox(item.heading, d.MARGIN, top, d.CONTENT_W, d.H * 0.05);
    htb.getText().getTextStyle().setBold(false).setFontSize(14).setForegroundColor(MUTED).setFontFamily(FONT_BODY);
    top += d.H * 0.05 + d.ITEM_GAP;
  }
  var stmtH  = d.H * 0.1;
  var colors = [NEAR_BLACK, BRAND_BLUE, NEAR_BLACK];
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox(lines[i], d.MARGIN, top, d.CONTENT_W, stmtH);
    tb.getText().getTextStyle().setBold(true).setFontSize(32).setForegroundColor(colors[i] || NEAR_BLACK).setFontFamily(FONT_DISPLAY);
    top += stmtH + d.STMT_GAP;
  }
}


// ── item-list ─────────────────────────────────────────────────────────
// Heading + card items (light bg), bold label + regular description (Inter)
function layoutItemList(slide, item, d) {
  var top   = addHeading(slide, item.heading, d);
  var lines = item.lines || [];
  var cardH = d.H * 0.1;
  var padX  = d.MARGIN * 0.6;
  var padY  = d.H * 0.018;
  for (var i = 0; i < lines.length; i++) {
    var card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, top, d.CONTENT_W, cardH);
    card.getFill().setSolidFill(CARD_BG);
    card.getBorder().setTransparent();
    var tb = slide.insertTextBox('', d.MARGIN + padX, top + padY, d.CONTENT_W - padX * 2, cardH - padY * 2);
    var tf = tb.getText();
    var parts = lines[i].split(' — ');
    if (parts.length >= 2) {
      tf.appendText(parts[0] + ' — ').getTextStyle().setBold(true).setFontSize(17).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_BODY);
      tf.appendText(parts.slice(1).join(' — ')).getTextStyle().setBold(false).setFontSize(15).setForegroundColor('#444444').setFontFamily(FONT_BODY);
    } else {
      tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(15).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_BODY);
    }
    top += cardH + d.ITEM_GAP * 0.6;
  }
}


// ── timeline-list ─────────────────────────────────────────────────────
// Card items with a brand-blue 4pt left accent bar
function layoutTimelineList(slide, item, d) {
  var top   = addHeading(slide, item.heading, d);
  var lines = item.lines || [];
  var barW  = 4;
  var padL  = barW + d.MARGIN * 0.5;
  var cardH = d.H * 0.1;
  var padY  = d.H * 0.015;
  for (var i = 0; i < lines.length; i++) {
    var card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, top, d.CONTENT_W, cardH);
    card.getFill().setSolidFill(CARD_BG);
    card.getBorder().setTransparent();
    var bar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, top, barW, cardH);
    bar.getFill().setSolidFill(BRAND_BLUE);
    bar.getBorder().setTransparent();
    var colonIdx = lines[i].indexOf(':');
    var tb = slide.insertTextBox('', d.MARGIN + padL, top + padY, d.CONTENT_W - padL - d.MARGIN * 0.3, cardH - padY * 2);
    var tf = tb.getText();
    if (colonIdx !== -1) {
      tf.appendText(lines[i].substring(0, colonIdx + 1) + '  ').getTextStyle().setBold(true).setFontSize(14).setForegroundColor(BRAND_BLUE).setFontFamily(FONT_DISPLAY);
      tf.appendText(lines[i].substring(colonIdx + 1).trim()).getTextStyle().setBold(false).setFontSize(14).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_BODY);
    } else {
      tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(14).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_BODY);
    }
    top += cardH + d.ITEM_GAP * 0.6;
  }
}


// ── cost-table ────────────────────────────────────────────────────────
// Rows with right-aligned amount; blue divider + blue TOTAL row
function layoutCostTable(slide, item, d) {
  var top   = addHeading(slide, item.heading, d);
  var lines = item.lines || [];
  var rowH  = d.H * 0.072;
  for (var i = 0; i < lines.length; i++) {
    var isTotal = lines[i].toUpperCase().indexOf('TOTAL') === 0;
    if (isTotal) {
      var div = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, top - d.ITEM_GAP * 0.5, d.CONTENT_W, 2);
      div.getFill().setSolidFill(BRAND_BLUE);
      div.getBorder().setTransparent();
      top += d.ITEM_GAP * 0.5;
    }
    var parts  = lines[i].split(' — ');
    var label  = parts.length >= 2 ? parts[0] : lines[i];
    var amount = parts.length >= 2 ? parts.slice(1).join(' — ') : '';
    var fs     = isTotal ? 24 : 18;
    var color  = isTotal ? BRAND_BLUE : NEAR_BLACK;
    var font   = isTotal ? FONT_DISPLAY : FONT_BODY;
    var lblTb  = slide.insertTextBox(label, d.MARGIN, top, d.CONTENT_W * 0.65, rowH);
    lblTb.getText().getTextStyle().setBold(isTotal).setFontSize(fs).setForegroundColor(color).setFontFamily(font);
    if (amount) {
      var amtTb = slide.insertTextBox(amount, d.MARGIN + d.CONTENT_W * 0.65, top, d.CONTENT_W * 0.35, rowH);
      amtTb.getText().getTextStyle().setBold(isTotal).setFontSize(fs).setForegroundColor(color).setFontFamily(font);
    }
    top += rowH + (isTotal ? 0 : d.ITEM_GAP * 0.5);
  }
}


// ── stat-callout ──────────────────────────────────────────────────────
// Left col: heading (Oswald) + context body (Inter). Right col: giant blue stat.
function layoutStatCallout(slide, item, d) {
  var colTop = d.CONTENT_TOP;
  if (item.heading) {
    var htb = slide.insertTextBox(item.heading, d.MARGIN, colTop, d.COL_W, d.H * 0.13);
    htb.getText().getTextStyle().setBold(true).setFontSize(40).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_DISPLAY);
    colTop += d.H * 0.13 + d.SECT_GAP;
  }
  var lines = item.lines || [];
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox(lines[i], d.MARGIN, colTop, d.COL_W, d.H * 0.072);
    tb.getText().getTextStyle().setBold(false).setFontSize(15).setForegroundColor(MUTED).setFontFamily(FONT_BODY);
    colTop += d.H * 0.072 + d.ITEM_GAP;
  }
  var stat      = item.stat || '';
  var statLabel = item.stat_label || item.description || '';
  var statTop   = d.CONTENT_TOP + (d.H - d.CONTENT_TOP) * 0.05;
  var statTb    = slide.insertTextBox(stat, d.COL_R_LEFT, statTop, d.COL_W, d.H * 0.35);
  statTb.getText().getTextStyle().setBold(true).setFontSize(120).setForegroundColor(BRAND_BLUE).setFontFamily(FONT_DISPLAY);
  if (statLabel) {
    var labelTb = slide.insertTextBox(statLabel, d.COL_R_LEFT, statTop + d.H * 0.36, d.COL_W, d.H * 0.12);
    labelTb.getText().getTextStyle().setBold(false).setFontSize(15).setForegroundColor(MUTED).setFontFamily(FONT_BODY);
  }
}


// ── two-stat ──────────────────────────────────────────────────────────
// Two large blue numbers on light-card backgrounds (Oswald numbers, Inter labels)
function layoutTwoStat(slide, item, d) {
  var stats = item.stats || [];
  if (!stats.length && item.lines && item.lines.length >= 2) {
    for (var s = 0; s < item.lines.length; s += 2) {
      stats.push({ value: item.lines[s], label: item.lines[s + 1] || '' });
    }
  }
  if (item.heading) {
    var htb = slide.insertTextBox(item.heading, d.MARGIN, d.CONTENT_TOP, d.CONTENT_W, d.H * 0.067);
    htb.getText().getTextStyle().setBold(true).setFontSize(28).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_DISPLAY);
  }
  var statTop   = d.CONTENT_TOP + (item.heading ? d.H * 0.13 : d.H * 0.05);
  var positions = [d.MARGIN, d.COL_R_LEFT];
  for (var i = 0; i < Math.min(stats.length, 2); i++) {
    var xLeft = positions[i];
    var card  = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, xLeft, statTop - d.H * 0.02, d.COL_W, d.H * 0.42);
    card.getFill().setSolidFill(CARD_BG);
    card.getBorder().setTransparent();
    var numTb = slide.insertTextBox(stats[i].value || '', xLeft + d.MARGIN * 0.4, statTop, d.COL_W - d.MARGIN * 0.8, d.H * 0.28);
    numTb.getText().getTextStyle().setBold(true).setFontSize(90).setForegroundColor(BRAND_BLUE).setFontFamily(FONT_DISPLAY);
    if (stats[i].label) {
      var descTb = slide.insertTextBox(stats[i].label, xLeft + d.MARGIN * 0.4, statTop + d.H * 0.28, d.COL_W - d.MARGIN * 0.8, d.H * 0.1);
      descTb.getText().getTextStyle().setBold(false).setFontSize(15).setForegroundColor(MUTED).setFontFamily(FONT_BODY);
    }
  }
}


// ── two-column ────────────────────────────────────────────────────────
// Heading + two columns each with a blue Oswald title + card items
function layoutTwoColumn(slide, item, d) {
  var top  = addHeading(slide, item.heading, d);
  var cols = item.columns || [];
  if (!cols.length && item.lines) {
    var half = Math.ceil(item.lines.length / 2);
    cols = [{ title: '', items: item.lines.slice(0, half) },
            { title: '', items: item.lines.slice(half) }];
  }
  var colPositions = [d.MARGIN, d.COL_R_LEFT];
  for (var c = 0; c < Math.min(cols.length, 2); c++) {
    var cx = colPositions[c];
    var cy = top;
    if (cols[c].title) {
      var ctb = slide.insertTextBox(cols[c].title, cx, cy, d.COL_W, d.H * 0.067);
      ctb.getText().getTextStyle().setBold(true).setFontSize(20).setForegroundColor(BRAND_BLUE).setFontFamily(FONT_DISPLAY);
      cy += d.H * 0.067 + d.ITEM_GAP * 0.5;
    }
    var items     = cols[c].items || [];
    var itemCardH = d.H * 0.082;
    var padX      = d.MARGIN * 0.5;
    var padY      = d.H * 0.015;
    for (var r = 0; r < items.length; r++) {
      var bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx, cy, d.COL_W, itemCardH);
      bg.getFill().setSolidFill(CARD_BG);
      bg.getBorder().setTransparent();
      var itb = slide.insertTextBox(items[r], cx + padX, cy + padY, d.COL_W - padX * 2, itemCardH - padY * 2);
      itb.getText().getTextStyle().setBold(false).setFontSize(14).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_BODY);
      cy += itemCardH + d.ITEM_GAP * 0.5;
    }
  }
}


// ── next-steps ────────────────────────────────────────────────────────
// Large Oswald heading + numbered cards — blue number, Inter text
function layoutNextSteps(slide, item, d) {
  var headH = d.H * 0.1;
  var tb0   = slide.insertTextBox(item.heading || 'What happens next.', d.MARGIN, d.CONTENT_TOP, d.CONTENT_W, headH);
  tb0.getText().getTextStyle().setBold(true).setFontSize(48).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_DISPLAY);
  var top   = d.CONTENT_TOP + headH + d.SECT_GAP;
  var lines = item.lines || [];
  var stepH = d.H * 0.09;
  var padX  = d.MARGIN * 0.5;
  var padY  = d.H * 0.018;
  for (var i = 0; i < lines.length; i++) {
    var card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, top, d.CONTENT_W, stepH);
    card.getFill().setSolidFill(CARD_BG);
    card.getBorder().setTransparent();
    var tb = slide.insertTextBox('', d.MARGIN + padX, top + padY, d.CONTENT_W - padX * 2, stepH - padY * 2);
    var tf = tb.getText();
    tf.appendText((i + 1) + '.  ').getTextStyle().setBold(true).setFontSize(20).setForegroundColor(BRAND_BLUE).setFontFamily(FONT_DISPLAY);
    tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(17).setForegroundColor(NEAR_BLACK).setFontFamily(FONT_BODY);
    top += stepH + d.ITEM_GAP * 0.6;
  }
}
