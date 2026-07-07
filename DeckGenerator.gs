// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// Fully programmatic layout — no template markers on content slides.
// Apps Script reads actual page dimensions at runtime and places text
// boxes proportionally. The template contributes only its visual layer
// (background, header branding, header rule) which survives because
// getShapes().remove() only removes foreground shapes, not master elements.
//
// Slide 1 (title) fills [[title]] and [[date]] markers.
// Slides 2–N are rebuilt from the slides[] array in the JSON.
// Last slide (closing) is left as-is from the template.
//
// JSON input shape:
// {
//   Client_Name, Project_Title, Sector,
//   slides: [{ heading, layout, lines[], speaker_note, stat?, description?, total? }],
//   Next_Steps
// }
//
// Layout types: statement-trio, item-list, timeline-list, cost-table,
//               stat-callout, two-column, next-steps
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

  // Read actual page dimensions — never hardcode
  var W = deck.getPageWidth();
  var H = deck.getPageHeight();
  var MARGIN      = Math.round(W * 0.035);       // ~25pt on 720pt slide
  var CONTENT_W   = W - MARGIN * 2;
  var HEADER_H    = Math.round(H * 0.143);       // ~58pt header zone
  var CONTENT_TOP = HEADER_H + Math.round(H * 0.025); // gap below header rule

  // ── Slide 1: title ──────────────────────────────────────────────────
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
      try { s.setLeft(MARGIN); s.setTop(H * 0.2); s.setWidth(CONTENT_W); s.setHeight(H * 0.4); } catch(e) {}
    } else if (raw === '[[date]]' || raw.indexOf('[[date]]') !== -1) {
      s.getText().clear();
      s.getText().appendText(date).getTextStyle().setBold(false).setFontSize(20);
    }
  }

  // ── Adjust slide count to match content slides ──────────────────────
  // Template structure: [0]=title, [1..N-2]=content, [N-1]=closing
  // We preserve title (0) and closing (last), rebuild everything between.
  var contentItems = data.slides || [];

  // Add extra content slides by duplicating slide index 1 (first content slide)
  slides = deck.getSlides();
  while (slides.length - 2 < contentItems.length) {
    slides[1].duplicate();
    slides = deck.getSlides();
    // Move the duplicate (appended at end) to just before closing
    slides[slides.length - 1].move(slides.length - 2);
    slides = deck.getSlides();
  }

  // Remove excess content slides (working backwards to avoid index shifts)
  slides = deck.getSlides();
  while (slides.length - 2 > contentItems.length) {
    slides = deck.getSlides();
    slides[slides.length - 2].remove();
    slides = deck.getSlides();
  }

  slides = deck.getSlides();

  // ── Fill each content slide ─────────────────────────────────────────
  for (var i = 0; i < contentItems.length; i++) {
    var slide = slides[i + 1];
    var item  = contentItems[i];

    // Remove all foreground shapes in the content area
    var existing = slide.getShapes();
    for (var k = 0; k < existing.length; k++) {
      try {
        if (existing[k].getTop() >= HEADER_H) {
          existing[k].remove();
        }
      } catch(e) {}
    }

    // Dispatch to layout function
    switch (item.layout || 'item-list') {
      case 'statement-trio': layoutStatementTrio(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W); break;
      case 'item-list':      layoutItemList(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W);      break;
      case 'timeline-list':  layoutTimelineList(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W);  break;
      case 'cost-table':     layoutCostTable(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W);     break;
      case 'stat-callout':   layoutStatCallout(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W);   break;
      case 'two-column':     layoutTwoColumn(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W);     break;
      case 'next-steps':     layoutNextSteps(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W);     break;
      default:               layoutItemList(slide, item, W, H, MARGIN, CONTENT_TOP, CONTENT_W);      break;
    }

    // Speaker notes
    if (item.speaker_note) {
      try {
        slide.getNotesPage().getSpeakerNotesShape().getText().setText(item.speaker_note);
      } catch(e) {}
    }
  }

  deck.saveAndClose();

  return {
    deckUrl:  'https://docs.google.com/presentation/d/' + newFile.getId() + '/edit',
    deckName: deckName
  };
}


// ── Shared helper: insert heading, return updated top ─────────────────
function addHeading(slide, text, margin, top, contentW) {
  if (!text) return top;
  var tb = slide.insertTextBox(text, margin, top, contentW, 44);
  tb.getText().getTextStyle().setBold(true).setFontSize(26);
  return top + 52;
}


// ── statement-trio ────────────────────────────────────────────────────
// Heading + 3 bold standalone statements, evenly spaced
function layoutStatementTrio(slide, item, W, H, margin, contentTop, contentW) {
  var top   = addHeading(slide, item.heading, margin, contentTop, contentW);
  var lines = item.lines || [];
  var available = H - top - margin;
  var lineH = Math.floor(available / Math.max(lines.length, 1));
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox(lines[i], margin, top + i * lineH, contentW, lineH - 8);
    tb.getText().getTextStyle().setBold(true).setFontSize(28);
  }
}


// ── item-list ─────────────────────────────────────────────────────────
// Heading + line items; bold phase name before " — "
function layoutItemList(slide, item, W, H, margin, contentTop, contentW) {
  var top   = addHeading(slide, item.heading, margin, contentTop, contentW);
  var lines = item.lines || [];
  var available = H - top - margin;
  var lineH = Math.min(Math.floor(available / Math.max(lines.length, 1)), 52);
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox('', margin, top + i * lineH, contentW, lineH - 6);
    var tf = tb.getText();
    var parts = lines[i].split(' — ');
    if (parts.length >= 2) {
      tf.appendText(parts[0] + ' — ').getTextStyle().setBold(true).setFontSize(20);
      tf.appendText(parts.slice(1).join(' — ')).getTextStyle().setBold(false).setFontSize(18);
    } else {
      tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(20);
    }
  }
}


// ── timeline-list ─────────────────────────────────────────────────────
// Heading + dated entries; bold the period label before ":"
function layoutTimelineList(slide, item, W, H, margin, contentTop, contentW) {
  var top   = addHeading(slide, item.heading, margin, contentTop, contentW);
  var lines = item.lines || [];
  var available = H - top - margin;
  var lineH = Math.min(Math.floor(available / Math.max(lines.length, 1)), 52);
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox('', margin, top + i * lineH, contentW, lineH - 6);
    var tf = tb.getText();
    var colonIdx = lines[i].indexOf(':');
    if (colonIdx !== -1) {
      tf.appendText(lines[i].substring(0, colonIdx + 1)).getTextStyle().setBold(true).setFontSize(18);
      tf.appendText(lines[i].substring(colonIdx + 1)).getTextStyle().setBold(false).setFontSize(18);
    } else {
      tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(18);
    }
  }
}


// ── cost-table ────────────────────────────────────────────────────────
// Heading + line items; TOTAL line is bold and slightly larger
function layoutCostTable(slide, item, W, H, margin, contentTop, contentW) {
  var top   = addHeading(slide, item.heading, margin, contentTop, contentW);
  var lines = item.lines || [];
  var available = H - top - margin;
  var lineH = Math.min(Math.floor(available / Math.max(lines.length, 1)), 44);
  for (var i = 0; i < lines.length; i++) {
    var isTotal = lines[i].toUpperCase().indexOf('TOTAL') === 0;
    var tb = slide.insertTextBox(lines[i], margin, top + i * lineH, contentW, lineH - 4);
    tb.getText().getTextStyle()
      .setBold(isTotal)
      .setFontSize(isTotal ? 24 : 18);
  }
}


// ── stat-callout ──────────────────────────────────────────────────────
// Giant stat/number on the left, short description on the right
function layoutStatCallout(slide, item, W, H, margin, contentTop, contentW) {
  var available = H - contentTop - margin;
  var halfW     = contentW * 0.45;
  var stat      = item.stat || (item.lines && item.lines[0]) || '';
  var desc      = item.description || (item.lines && item.lines.slice(1).join('\n')) || '';

  var statTb = slide.insertTextBox(stat, margin, contentTop, halfW, available);
  statTb.getText().getTextStyle().setBold(true).setFontSize(80);

  if (desc) {
    var descTop = contentTop + available * 0.25;
    var descTb  = slide.insertTextBox(desc, margin + halfW + margin, descTop, halfW, available * 0.6);
    descTb.getText().getTextStyle().setBold(false).setFontSize(18);
  }
}


// ── two-column ────────────────────────────────────────────────────────
// Heading + lines split evenly into left and right columns
function layoutTwoColumn(slide, item, W, H, margin, contentTop, contentW) {
  var top   = addHeading(slide, item.heading, margin, contentTop, contentW);
  var lines = item.lines || [];
  var colW  = (contentW - margin) / 2;
  var available = H - top - margin;
  var half  = Math.ceil(lines.length / 2);
  var lineH = Math.min(Math.floor(available / Math.max(half, 1)), 44);

  var left  = lines.slice(0, half);
  var right = lines.slice(half);

  for (var i = 0; i < left.length; i++) {
    var tb = slide.insertTextBox(left[i], margin, top + i * lineH, colW, lineH - 4);
    tb.getText().getTextStyle().setBold(false).setFontSize(18);
  }
  for (var j = 0; j < right.length; j++) {
    var tb2 = slide.insertTextBox(right[j], margin + colW + margin, top + j * lineH, colW, lineH - 4);
    tb2.getText().getTextStyle().setBold(false).setFontSize(18);
  }
}


// ── next-steps ────────────────────────────────────────────────────────
// Heading + numbered action items; number is bold
function layoutNextSteps(slide, item, W, H, margin, contentTop, contentW) {
  var top   = addHeading(slide, item.heading || 'What happens next.', margin, contentTop, contentW);
  var lines = item.lines || [];
  var available = H - top - margin;
  var lineH = Math.min(Math.floor(available / Math.max(lines.length, 1)), 56);
  for (var i = 0; i < lines.length; i++) {
    var tb = slide.insertTextBox('', margin, top + i * lineH, contentW, lineH - 8);
    var tf = tb.getText();
    tf.appendText((i + 1) + '.  ').getTextStyle().setBold(true).setFontSize(22);
    tf.appendText(lines[i]).getTextStyle().setBold(false).setFontSize(22);
  }
}
