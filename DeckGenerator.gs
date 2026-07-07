// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// Fully programmatic layout driven by design system reference.
// All positions/sizes expressed as proportions of actual slide dimensions
// so they scale correctly regardless of template size.
//
// Design system reference canvas: 1280×720px
// Proportion constants derived from that canvas — applied to W×H at runtime.
//
// Layout types (Gemini chooses per slide):
//   statement-trio  — 3 bold standalone statements, full width
//   item-list       — heading + line items, bold label before " — "
//   timeline-list   — heading + items with left border bar
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

  var MARGIN      = W * 0.0422;   // 54px / 1280
  var CONTENT_W   = W - MARGIN * 2;
  var HEADER_H    = H * 0.1431;   // 103px / 720
  var CONTENT_TOP = H * 0.1681;   // 121px / 720
  var COL_W       = W * 0.4063;   // 520px / 1280  (each column)
  var COL_R_LEFT  = W * 0.5;      // 640px / 1280  (right col start)
  var ITEM_GAP    = H * 0.039;    // 28px / 720
  var SECT_GAP    = H * 0.056;    // 40px / 720
  var STMT_GAP    = H * 0.044;    // 32px / 720

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
      s.getText().appendText(projectTitle).getTextStyle().setBold(true).setFontSize(60);
      try { s.setLeft(MARGIN); s.setTop(H * 0.25); s.setWidth(CONTENT_W); s.setHeight(H * 0.35); } catch(e) {}
    } else if (raw === '[[date]]' || raw.indexOf('[[date]]') !== -1) {
      s.getText().clear();
      s.getText().appendText(date).getTextStyle().setBold(false).setFontSize(20);
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

    // Remove all foreground shapes below the header zone
    var existing = slide.getShapes();
    for (var k = 0; k < existing.length; k++) {
      try { if (existing[k].getTop() >= HEADER_H) existing[k].remove(); } catch(e) {}
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


// ── Shared: insert heading, returns new top position ──────────────────
function addHeading(slide, text, d) {
  if (!text) return d.CONTENT_TOP;
  var headH = d.H * 0.067;
  var tb = slide.insertTextBox(text, d.MARGIN, d.CONTENT_TOP, d.CONTENT_W, headH);
  tb.getText().getTextStyle().setBold(true).setFontSize(32);
  return d.CONTENT_TOP + headH + d.SECT_GAP;
}

// ── Shared: insert a thin black divider line ──────────────────────────
function addDivider(slide, y, d) {
  var line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, y, d.CONTENT_W, 1);
  line.getFill().setSolidFill('#000000');
  line.getBorder().setTransparent();
}


// ── statement-trio ────────────────────────────────────────────────────
// 3 bold full-width statements, no heading — the statements fill the slide
function layoutStatementTrio(slide, item, d) {
  var lines = item.lines || [];
  // Optional small heading label above
  var top = item.heading ? d.CONTENT_TOP : d.CONTENT_TOP;
  if (item.heading) {
    var htb = slide.insertTextBox(item.heading, d.MARGIN, top, d.CONTENT_W, d.H * 0.05);
    htb.getText().getTextStyle().setBold(false).setFontSize(16);
    top += d.H * 0.05 + d.ITEM_GAP;
  }
  var stmtH = d.H * 0.1;
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox(lines[i], d.MARGIN, top, d.CONTENT_W, stmtH);
    tb.getText().getTextStyle().setBold(true).setFontSize(32);
    top += stmtH + d.STMT_GAP;
  }
}


// ── item-list ─────────────────────────────────────────────────────────
// Heading + line items; bold label before " — ", regular description after
function layoutItemList(slide, item, d) {
  var top   = addHeading(slide, item.heading, d);
  var lines = item.lines || [];
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox('', d.MARGIN, top, d.CONTENT_W, d.H * 0.072);
    var tf = tb.getText();
    var parts = lines[i].split(' — ');
    if (parts.length >= 2) {
      tf.appendText(parts[0] + ' — ').getTextStyle().setBold(true).setFontSize(20);
      tf.appendText(parts.slice(1).join(' — ')).getTextStyle().setBold(false).setFontSize(18);
    } else {
      tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(18);
    }
    top += d.H * 0.072 + d.ITEM_GAP;
  }
}


// ── timeline-list ─────────────────────────────────────────────────────
// Heading + items with 3pt left border bar, label bold, description below
function layoutTimelineList(slide, item, d) {
  var top   = addHeading(slide, item.heading, d);
  var lines = item.lines || [];
  var barW  = 3;
  var padL  = barW + d.MARGIN * 0.5;
  for (var i = 0; i < lines.length; i++) {
    var itemH = d.H * 0.095;
    // Left border bar
    var bar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, d.MARGIN, top, barW, itemH);
    bar.getFill().setSolidFill('#000000');
    bar.getBorder().setTransparent();
    // Text: period label bold + description regular
    var colonIdx = lines[i].indexOf(':');
    var tb = slide.insertTextBox('', d.MARGIN + padL, top, d.CONTENT_W - padL, itemH);
    var tf = tb.getText();
    if (colonIdx !== -1) {
      tf.appendText(lines[i].substring(0, colonIdx + 1) + '\n').getTextStyle().setBold(true).setFontSize(16);
      tf.appendText(lines[i].substring(colonIdx + 1).trim()).getTextStyle().setBold(false).setFontSize(18);
    } else {
      tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(18);
    }
    top += itemH + d.ITEM_GAP;
  }
}


// ── cost-table ────────────────────────────────────────────────────────
// Heading + rows with right-aligned amount, divider + TOTAL
function layoutCostTable(slide, item, d) {
  var top   = addHeading(slide, item.heading, d);
  var lines = item.lines || [];
  var rowH  = d.H * 0.072;
  for (var i = 0; i < lines.length; i++) {
    var isTotal = lines[i].toUpperCase().indexOf('TOTAL') === 0;
    if (isTotal) {
      addDivider(slide, top - d.ITEM_GAP * 0.5, d);
      top += d.ITEM_GAP * 0.5;
    }
    // Split on last " — " or "£" to separate label from amount
    var parts = lines[i].split(' — ');
    var label  = parts.length >= 2 ? parts[0] : lines[i];
    var amount = parts.length >= 2 ? parts.slice(1).join(' — ') : '';
    var fs = isTotal ? 24 : 18;
    var lblTb = slide.insertTextBox(label, d.MARGIN, top, d.CONTENT_W * 0.65, rowH);
    lblTb.getText().getTextStyle().setBold(isTotal).setFontSize(fs);
    if (amount) {
      var amtTb = slide.insertTextBox(amount, d.MARGIN + d.CONTENT_W * 0.65, top, d.CONTENT_W * 0.35, rowH);
      amtTb.getText().getTextStyle().setBold(isTotal).setFontSize(fs);
    }
    top += rowH + (isTotal ? 0 : d.ITEM_GAP);
  }
}


// ── stat-callout ──────────────────────────────────────────────────────
// Left col: context/problem heading + body items. Right col: giant stat + label.
// item.stat = the big number/stat string. item.stat_label = description below it.
// item.lines = left column body items (18pt)
function layoutStatCallout(slide, item, d) {
  var colTop = d.CONTENT_TOP;
  // Left column — heading + body items
  if (item.heading) {
    var htb = slide.insertTextBox(item.heading, d.MARGIN, colTop, d.COL_W, d.H * 0.13);
    htb.getText().getTextStyle().setBold(true).setFontSize(48);
    colTop += d.H * 0.13 + d.SECT_GAP;
  }
  var lines = item.lines || [];
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox(lines[i], d.MARGIN, colTop, d.COL_W, d.H * 0.072);
    tb.getText().getTextStyle().setBold(false).setFontSize(18);
    colTop += d.H * 0.072 + d.ITEM_GAP;
  }
  // Right column — giant stat
  var stat      = item.stat || '';
  var statLabel = item.stat_label || item.description || '';
  var statTop   = d.CONTENT_TOP + (d.H - d.CONTENT_TOP) * 0.1;
  var statTb = slide.insertTextBox(stat, d.COL_R_LEFT, statTop, d.COL_W, d.H * 0.35);
  statTb.getText().getTextStyle().setBold(true).setFontSize(120);
  if (statLabel) {
    var labelTb = slide.insertTextBox(statLabel, d.COL_R_LEFT, statTop + d.H * 0.38, d.COL_W, d.H * 0.12);
    labelTb.getText().getTextStyle().setBold(false).setFontSize(18);
  }
}


// ── two-stat ──────────────────────────────────────────────────────────
// Two large numbers side by side — evidence/proof slide.
// item.stats = [{value: "3x", label: "description"}, {value: "73%", label: "..."}]
// Fallback: split item.lines into two halves (first line = stat, second = label)
function layoutTwoStat(slide, item, d) {
  var stats = item.stats || [];
  // Fallback: build from lines[]
  if (!stats.length && item.lines && item.lines.length >= 2) {
    for (var s = 0; s < item.lines.length; s += 2) {
      stats.push({ value: item.lines[s], label: item.lines[s + 1] || '' });
    }
  }
  if (item.heading) {
    var htb = slide.insertTextBox(item.heading, d.MARGIN, d.CONTENT_TOP, d.CONTENT_W, d.H * 0.067);
    htb.getText().getTextStyle().setBold(true).setFontSize(28);
  }
  var statTop = d.CONTENT_TOP + (item.heading ? d.H * 0.13 : 0);
  var positions = [d.MARGIN, d.COL_R_LEFT];
  for (var i = 0; i < Math.min(stats.length, 2); i++) {
    var xLeft = positions[i];
    var numTb = slide.insertTextBox(stats[i].value || '', xLeft, statTop, d.COL_W, d.H * 0.28);
    numTb.getText().getTextStyle().setBold(true).setFontSize(90);
    if (stats[i].label) {
      var descTb = slide.insertTextBox(stats[i].label, xLeft, statTop + d.H * 0.3, d.COL_W, d.H * 0.1);
      descTb.getText().getTextStyle().setBold(false).setFontSize(20);
    }
  }
}


// ── two-column ────────────────────────────────────────────────────────
// Heading + two equal columns, each with a title and list of items.
// item.columns = [{title: "For Kids", items: [...]}, {title: "For Parents", items: [...]}]
// Fallback: split item.lines in half
function layoutTwoColumn(slide, item, d) {
  var top = addHeading(slide, item.heading, d);
  var cols = item.columns || [];
  if (!cols.length && item.lines) {
    var half = Math.ceil(item.lines.length / 2);
    cols = [{ title: '', items: item.lines.slice(0, half) },
            { title: '', items: item.lines.slice(half) }];
  }
  var colPositions = [d.MARGIN, d.COL_R_LEFT];
  for (var c = 0; c < Math.min(cols.length, 2); c++) {
    var cx  = colPositions[c];
    var cy  = top;
    if (cols[c].title) {
      var ctb = slide.insertTextBox(cols[c].title, cx, cy, d.COL_W, d.H * 0.067);
      ctb.getText().getTextStyle().setBold(true).setFontSize(26);
      cy += d.H * 0.067 + d.ITEM_GAP;
    }
    var items = cols[c].items || [];
    for (var r = 0; r < items.length; r++) {
      var itb = slide.insertTextBox(items[r], cx, cy, d.COL_W, d.H * 0.06);
      itb.getText().getTextStyle().setBold(false).setFontSize(18);
      cy += d.H * 0.06 + d.ITEM_GAP;
    }
  }
}


// ── next-steps ────────────────────────────────────────────────────────
// Heading (large) + numbered action items
function layoutNextSteps(slide, item, d) {
  var headH = d.H * 0.1;
  var tb0 = slide.insertTextBox(item.heading || 'What happens next.', d.MARGIN, d.CONTENT_TOP, d.CONTENT_W, headH);
  tb0.getText().getTextStyle().setBold(true).setFontSize(48);
  var top  = d.CONTENT_TOP + headH + d.SECT_GAP;
  var lines = item.lines || [];
  var stepH = d.H * 0.085;
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox('', d.MARGIN, top, d.CONTENT_W, stepH);
    var tf = tb.getText();
    tf.appendText((i + 1) + '.  ').getTextStyle().setBold(true).setFontSize(22);
    tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(20);
    top += stepH + d.ITEM_GAP * 1.5;
  }
}
