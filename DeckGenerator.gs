// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// Template setup: each content text box contains a tiny invisible marker
// in 1pt white text. The script finds each marker, clears the box, and
// inserts real content at the correct font size.
//
// Slide dimensions: 720pt wide × 405pt tall (standard 16:9)
// Content area (below header rule): top=58, left=25, width=670, height=332
//
// Markers → slots:
//   [[title]]      Slide 1 — project title
//   [[date]]       Slide 1 — month/year date
//   {{sector}}     Slide 2 — sector / industry
//   {{Takeaway}}   Slide 3 — 3 key statements (one per line)
//   {{Milestones}} Slide 4 — 4–5 milestones
//   {{Timeine}}    Slide 5 — timeline (typo in template, do not fix)
//   {{Cost }}      Slide 6 — cost breakdown line items
//
// Deploy as: Web App / Execute as Me / Who has access: Anyone
// ═══════════════════════════════════════════════════════════════════════

var TEMPLATE_SLIDE_ID = '1SYcTXUmcg3ci2pg8kWrexzwiTgHdRrF0BCAwca5dBZ0';
var OUTPUT_FOLDER_ID  = '1kTvIOM06sQh5tk2cK0697xLjs9nPyERR';

// Font sizes per slot
var SLOT_STYLES = {
  '[[title]]':           { fontSize: 60, bold: true  },
  '[[date]]':            { fontSize: 20, bold: false },
  '{{sector}}':          { fontSize: 60, bold: true  },
  '{{Heading_Insight}}': { fontSize: 28, bold: true  },
  '{{Takeaway}}':        { fontSize: 34, bold: false },
  '{{Heading_Milestones}}': { fontSize: 28, bold: true },
  '{{Milestones}}':      { fontSize: 22, bold: false },
  '{{Heading_Timeline}}':   { fontSize: 28, bold: true },
  '{{Timeine}}':         { fontSize: 22, bold: false },
  '{{Heading_Cost}}':    { fontSize: 28, bold: true  },
  '{{Cost }}':           { fontSize: 22, bold: false },
};

// Layout per slot — repositions and resizes the shape to fill the content area.
// Content area on a 720×405pt slide below the header rule: top≈58, left=25, width=670, height=332
var SLOT_LAYOUT = {
  '[[title]]':           { left: 25, top: 80,  width: 670, height: 200 },
  '[[date]]':            null,
  '{{sector}}':          { left: 25, top: 120, width: 670, height: 220 },
  '{{Heading_Insight}}': { left: 25, top: 58,  width: 670, height: 45  },
  '{{Takeaway}}':        { left: 25, top: 110, width: 670, height: 280 },
  '{{Heading_Milestones}}': { left: 25, top: 58, width: 670, height: 45 },
  '{{Milestones}}':      { left: 25, top: 110, width: 670, height: 280 },
  '{{Heading_Timeline}}':   { left: 25, top: 58, width: 670, height: 45 },
  '{{Timeine}}':         { left: 25, top: 110, width: 670, height: 280 },
  '{{Heading_Cost}}':    { left: 25, top: 58,  width: 670, height: 45  },
  '{{Cost }}':           { left: 25, top: 110, width: 670, height: 280 },
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


function smartSplit(text) {
  if (text.indexOf('\n') !== -1) {
    return text.split('\n').filter(function(line) { return line.trim() !== ''; });
  }
  var delimiters = ['. ', '; ', ' | '];
  for (var i = 0; i < delimiters.length; i++) {
    if (text.indexOf(delimiters[i]) !== -1) {
      return text.split(delimiters[i]).filter(function(line) { return line.trim() !== ''; });
    }
  }
  return [text];
}

function buildDeck(data) {
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');

  var content = {
    '[[title]]':              data.Project_Title       || '',
    '[[date]]':               date,
    '{{sector}}':             data.Sector              || '',
    '{{Heading_Insight}}':    data.Heading_Insight     || '',
    '{{Takeaway}}':           data.Key_Takeaway        || '',
    '{{Heading_Milestones}}': data.Heading_Milestones  || '',
    '{{Milestones}}':         data.Project_Milestones  || '',
    '{{Heading_Timeline}}':   data.Heading_Timeline    || '',
    '{{Timeine}}':            data.Project_Timeline    || '',  // typo in template — do not fix here
    '{{Heading_Cost}}':       data.Heading_Cost        || '',
    '{{Cost }}':              data.Cost_Breakdown      || '',
  };

  var clientName   = data.Client_Name   || 'Client';
  var projectTitle = data.Project_Title || 'Deck';
  var deckName     = clientName + ' — ' + projectTitle;

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

      var matchedKey = null;
      for (var k in content) {
        if (raw === k) { matchedKey = k; break; }
        if (k.indexOf('[[') === 0 && raw.indexOf(k) !== -1) { matchedKey = k; break; }
      }

      if (matchedKey !== null) {
        var text  = content[matchedKey];
        var style = SLOT_STYLES[matchedKey];
        var tf    = shape.getText();

        tf.clear();

        if (matchedKey === '{{Milestones}}') {
          var lines = smartSplit(text);
          for (var l = 0; l < lines.length; l++) {
            if (l > 0) tf.appendText('\n');
            var parts = lines[l].trim().split(' — ');
            if (parts.length >= 2) {
              var boldPart = tf.appendText(parts[0] + ' — ');
              boldPart.getTextStyle().setBold(true).setFontSize(style.fontSize);
              var regularPart = tf.appendText(parts.slice(1).join(' — '));
              regularPart.getTextStyle().setBold(false).setFontSize(style.fontSize - 2);
            } else {
              var plain = tf.appendText(lines[l].trim());
              plain.getTextStyle().setBold(false).setFontSize(style.fontSize);
            }
          }
        } else if (matchedKey === '{{Cost }}') {
          var costLines = smartSplit(text);
          for (var c = 0; c < costLines.length; c++) {
            if (c > 0) tf.appendText('\n');
            var costLine = costLines[c].trim();
            var isTotal = costLine.toUpperCase().indexOf('TOTAL') === 0;
            var costRange = tf.appendText(costLine);
            costRange.getTextStyle()
              .setBold(isTotal)
              .setFontSize(isTotal ? style.fontSize + 4 : style.fontSize);
          }
        } else if (matchedKey === '{{Takeaway}}') {
          var statements = smartSplit(text);
          for (var s = 0; s < statements.length; s++) {
            if (s > 0) tf.appendText('\n\n');
            var stRange = tf.appendText(statements[s].trim());
            stRange.getTextStyle().setBold(false).setFontSize(style.fontSize);
          }
        } else {
          var range = tf.appendText(text);
          range.getTextStyle().setBold(style.bold).setFontSize(style.fontSize);
        }

        // Resize and reposition the shape to fill the content area
        var layout = SLOT_LAYOUT[matchedKey];
        if (layout) {
          try {
            shape.setLeft(layout.left);
            shape.setTop(layout.top);
            shape.setWidth(layout.width);
            shape.setHeight(layout.height);
          } catch(e) { /* shape locked or grouped — skip resize */ }
        }
      }
    }
  }

  // Write speaker notes to slide notes panels
  var speakerNotes = {
    2: data.Speaker_Notes_Insight    || '',
    3: data.Speaker_Notes_Milestones || '',
    4: data.Speaker_Notes_Timeline   || '',
    5: data.Speaker_Notes_Cost       || '',
  };
  for (var idx in speakerNotes) {
    var noteText = speakerNotes[idx];
    if (!noteText) continue;
    var slideIdx = parseInt(idx, 10);
    if (slideIdx >= slides.length) continue;
    try {
      slides[slideIdx].getNotesPage().getSpeakerNotesShape().getText().setText(noteText);
    } catch(e) { /* notes panel unavailable on this slide — skip */ }
  }

  deck.saveAndClose();

  return {
    deckUrl:  'https://docs.google.com/presentation/d/' + newFile.getId() + '/edit',
    deckName: deckName
  };
}
