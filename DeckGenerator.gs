// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// HOW TO DEPLOY:
//   1. Open https://script.google.com and create a new project.
//   2. Paste this entire file into the editor (replace any existing code).
//   3. Set the two constants below:
//        TEMPLATE_SLIDE_ID — the file ID of your Cyphr Slides template
//        OUTPUT_FOLDER_ID  — the Google Drive folder ID where decks are saved
//   4. Click Deploy > New Deployment.
//        - Type: Web App
//        - Execute as: Me
//        - Who has access: Anyone
//   5. Copy the Web App URL — paste it into Cyphr Flow Settings as the Deck Generator URL.
//
// NOTE: No Gemini API key needed — the caller (Cyphr Flow) sends pre-processed
//       slide copy. This script only does Slides manipulation.
//
// INCOMING POST BODY (JSON) — all values are flat strings:
//   {
//     Client_Name, Project_Title, Sector, Date,
//     Key_Takeaway,          (3 lines, one per line)
//     Project_Milestones,    (3–5 lines formatted "Phase: description")
//     Project_Timeline,      (short paragraph)
//     Cost_Breakdown,        (line items "Item — £X,XXX" + "TOTAL — £XX,XXX")
//     Next_Steps,
//     Stakeholders
//   }
//
// TEMPLATE SLIDE TAGS (must match exactly):
//   Slide 1 (Cover):    {{Client_Name}}, {{Project_Title}}, {{Sector}}, {{Date}}
//   Slide 2:            {{Key_Takeaway}}
//   Slide 3:            {{Project_Milestones}}
//   Slide 4:            {{Project_Timeline}}
//   Slide 5:            {{Cost_Breakdown}}
//   Slide 6 (Closing):  {{Next_Steps}}, {{Stakeholders}}
//
// RESPONSE (JSON):
//   Success: { deckUrl, deckName }
//   Error:   { error: "message" }
// ═══════════════════════════════════════════════════════════════════════


var TEMPLATE_SLIDE_ID = '1SYcTXUmcg3ci2pg8kWrexzwiTgHdRrF0BCAwca5dBZ0';
var OUTPUT_FOLDER_ID  = '1kTvIOM06sQh5tk2cK0697xLjs9nPyERR';


/**
 * Entry point for the Web App.
 * Expects a JSON body with pre-processed slide copy — no Gemini call here.
 */
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


/**
 * Copies the Slides template, replaces all {{tags}}, saves, returns the URL.
 */
function buildDeck(data) {
  if (!TEMPLATE_SLIDE_ID) throw new Error('TEMPLATE_SLIDE_ID is not set.');
  if (!OUTPUT_FOLDER_ID)  throw new Error('OUTPUT_FOLDER_ID is not set.');

  // Inject current month/year if not provided
  if (!data.Date) {
    data.Date = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'MMMM yyyy'
    );
  }

  const clientName   = data.Client_Name   || 'Client';
  const projectTitle = data.Project_Title || 'Deck';
  const deckName     = clientName + ' — ' + projectTitle;

  const newFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID)
    .makeCopy(deckName, DriveApp.getFolderById(OUTPUT_FOLDER_ID));

  const deck = SlidesApp.openById(newFile.getId());

  for (var key in data) {
    if (Object.prototype.hasOwnProperty.call(data, key)) {
      var value = (data[key] !== null && data[key] !== undefined)
        ? String(data[key])
        : '—';
      deck.replaceAllText('{{' + key + '}}', value);
    }
  }

  deck.saveAndClose();

  return {
    deckUrl:  'https://docs.google.com/presentation/d/' + newFile.getId() + '/edit',
    deckName: deckName
  };
}
