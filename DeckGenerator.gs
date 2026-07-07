// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// Template setup: each content text box contains a tiny invisible marker
// in 1pt white text. The script finds each marker, clears the box, and
// inserts real content at the correct font size.
//
// Markers → slots:
//   [[title]]      Slide 1 — project title
//   [[date]]       Slide 1 — month/year date
//   [[sector]]     Slide 2 — sector / industry
//   [[takeaway]]   Slide 3 — 3 key statements (one per line)
//   [[milestones]] Slide 4 — 4–5 milestones
//   [[timeline]]   Slide 5 — timeline paragraph
//   [[cost]]       Slide 6 — cost breakdown line items
//
// Deploy as: Web App / Execute as Me / Who has access: Anyone
// ═══════════════════════════════════════════════════════════════════════

var TEMPLATE_SLIDE_ID = '1SYcTXUmcg3ci2pg8kWrexzwiTgHdRrF0BCAwca5dBZ0';
var OUTPUT_FOLDER_ID  = '1kTvIOM06sQh5tk2cK0697xLjs9nPyERR';

// Font sizes per slot. Font family inherits from the text box style in the template.
var SLOT_STYLES = {
  '[[title]]':      { fontSize: 60, bold: true  },
  '[[date]]':       { fontSize: 20, bold: false },
  '{{sector}}':     { fontSize: 60, bold: true  },
  '{{Takeaway}}':   { fontSize: 34, bold: false },
  '{{Milestones}}': { fontSize: 22, bold: false },
  '{{Timeine}}':    { fontSize: 22, bold: false },
  '{{Cost}}':       { fontSize: 22, bold: false },
};


function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST body received.');
    }
    const data = JSON.parse(e.postData.contents);
    const result = buildDeck(data);
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
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');

  // Map markers to their content — matches exact text in template (after trim)
  var content = {
    '[[title]]':       data.Project_Title       || '',
    '[[date]]':        date,
    '{{sector}}':      data.Sector              || '',
    '{{Takeaway}}':    data.Key_Takeaway        || '',
    '{{Milestones}}':  data.Project_Milestones  || '',
    '{{Timeine}}':     data.Project_Timeline    || '',  // typo in template — do not fix here
    '{{Cost}}':        data.Cost_Breakdown      || '',
  };

  var clientName   = data.Client_Name   || 'Client';
  var projectTitle = data.Project_Title || 'Deck';
  var deckName     = clientName + ' — ' + projectTitle; // em dash

  var newFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID)
    .makeCopy(deckName, DriveApp.getFolderById(OUTPUT_FOLDER_ID));

  var deck   = SlidesApp.openById(newFile.getId());
  var slides = deck.getSlides();

  for (var i = 0; i < slides.length; i++) {
    var shapes = slides[i].getShapes();
    for (var j = 0; j < shapes.length; j++) {
      var shape = shapes[j];
      if (!shape.getText) continue;
      var raw = shape.getText().asString().trim();

      if (content.hasOwnProperty(raw)) {
        var text    = content[raw];
        var style   = SLOT_STYLES[raw];
        var tf      = shape.getText();

        tf.clear();

        if (raw === '{{Milestones}}') {
          // Bold phase name, regular description — split on " — "
          var lines = text.split('\n');
          for (var l = 0; l < lines.length; l++) {
            if (l > 0) tf.appendText('\n');
            var parts = lines[l].split(' — ');
            if (parts.length >= 2) {
              var boldPart = tf.appendText(parts[0] + ' — ');
              boldPart.getTextStyle().setBold(true).setFontSize(style.fontSize);
              var regularPart = tf.appendText(parts.slice(1).join(' — '));
              regularPart.getTextStyle().setBold(false).setFontSize(style.fontSize - 2);
            } else {
              var plain = tf.appendText(lines[l]);
              plain.getTextStyle().setBold(false).setFontSize(style.fontSize);
            }
          }
        } else if (raw === '{{Cost}}') {
          // Regular line items, bold TOTAL line
          var costLines = text.split('\n');
          for (var c = 0; c < costLines.length; c++) {
            if (c > 0) tf.appendText('\n');
            var isTotal = costLines[c].toUpperCase().indexOf('TOTAL') === 0;
            var costRange = tf.appendText(costLines[c]);
            costRange.getTextStyle()
              .setBold(isTotal)
              .setFontSize(isTotal ? style.fontSize + 4 : style.fontSize);
          }
        } else if (raw === '{{Takeaway}}') {
          // Each statement on its own line with extra spacing
          var statements = text.split('\n');
          for (var s = 0; s < statements.length; s++) {
            if (s > 0) tf.appendText('\n\n');
            var stRange = tf.appendText(statements[s]);
            stRange.getTextStyle().setBold(false).setFontSize(style.fontSize);
          }
        } else {
          // Default: insert as-is
          var range = tf.appendText(text);
          range.getTextStyle().setBold(style.bold).setFontSize(style.fontSize);
        }
      }
    }
  }

  deck.saveAndClose();

  return {
    deckUrl:  'https://docs.google.com/presentation/d/' + newFile.getId() + '/edit',
    deckName: deckName
  };
}
