// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// HOW TO DEPLOY:
//   1. Open https://script.google.com and create a new project.
//   2. Paste this entire file into the editor (replace any existing code).
//   3. Set Script Properties (Project Settings > Script Properties):
//        GEMINI_API_KEY    — your Google AI Studio API key
//        TEMPLATE_SLIDE_ID — the file ID of your Cyphr Slides template
//                            (the long ID in the Google Slides URL)
//        OUTPUT_FOLDER_ID  — the Google Drive folder ID where decks are saved
//   4. Click Deploy > New Deployment.
//        - Type: Web App
//        - Execute as: Me
//        - Who has access: Anyone
//   5. Copy the Web App URL — paste it into Cyphr Flow Settings as the Deck Generator URL.
//
// TEMPLATE SLIDE TAGS (must match exactly — no spaces, no leading dots):
//   Slide 1 (Cover):    {{Project_Title}}, {{Date}}
//   Slide 2:            {{Sector}}
//   Slide 3:            {{Key_Takeaway_1}}
//   Slide 4:            {{Key_Takeaway_2}}
//   Slide 5:            {{Key_Takeaway_3}}
//   Slide 6:            {{Project_Milestones}}
//   Slide 7:            {{Project_Timeline}}
//   Slide 8:            {{Cost_Breakdown}}
//   Slide 9 (Closing):  No tags — fixed "THANK YOU" slide
//
// INCOMING POST BODY (JSON):
//   { clientName, projectName, sector, budget, timeline,
//     requirements, bgNotes, stakeholders, nextSteps,
//     briefOutput, estimateOutput }
//
// RESPONSE (JSON):
//   Success: { deckUrl, deckName }
//   Error:   { error: "message" }
// ═══════════════════════════════════════════════════════════════════════


// ── FILL THESE IN ──────────────────────────────────────────────────────
var GEMINI_API_KEY    = 'PASTE_YOUR_GEMINI_KEY_HERE';
var TEMPLATE_SLIDE_ID = '1SYcTXUmcg3ci2pg8kWrexzwiTgHdRrF0BCAwca5dBZ0';
var OUTPUT_FOLDER_ID  = '1kTvIOM06sQh5tk2cK0697xLjs9nPyERR';
// ───────────────────────────────────────────────────────────────────────


/**
 * Entry point for the Web App.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST body received.');
    }

    const projectData = JSON.parse(e.postData.contents);
    const compressed  = compressForSlides(projectData);
    const result      = buildDeck(compressed);

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
 * Calls Gemini to distil raw project data into structured slide copy.
 * Returns a flat object whose keys match the {{tags}} in the Slides template.
 */
function compressForSlides(projectData) {
  const apiKey = GEMINI_API_KEY;
  if (!apiKey || apiKey === 'PASTE_YOUR_GEMINI_KEY_HERE') throw new Error('GEMINI_API_KEY is not set at the top of the script.');

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' +
    apiKey;

  const rawContext = [
    'Client Name: '      + (projectData.clientName    || ''),
    'Project Name: '     + (projectData.projectName   || ''),
    'Sector: '           + (projectData.sector         || ''),
    'Budget: '           + (projectData.budget         || ''),
    'Timeline: '         + (projectData.timeline       || ''),
    'Requirements: '     + (projectData.requirements   || ''),
    'Background Notes: ' + (projectData.bgNotes        || ''),
    'Stakeholders: '     + (projectData.stakeholders   || ''),
    'Next Steps: '       + (projectData.nextSteps      || ''),
    'AI Brief Output: '  + (projectData.briefOutput    || ''),
    'Estimate Output: '  + (projectData.estimateOutput || ''),
  ].join('\n');

  const systemPrompt = `You are a professional copywriter for Cyphr, a creative strategy and innovation studio.
Distil the raw project data below into polished, concise slide copy for a client-facing deck.
Return ONLY a single valid JSON object — no markdown, no code fences, no commentary.

The JSON must contain exactly these keys:

  Project_Title        — string, a punchy 4–8 word project title
  Sector               — string, the industry / sector in 2–4 words
  Key_Takeaway         — string, exactly 3 punchy insights or statements, each on its own line,
                         each max 20 words, no leading dashes or bullets
  Project_Milestones   — string, 3–5 key project milestones, each on its own line formatted as
                         "Phase Name: one-sentence description"
  Project_Timeline     — string, a clear narrative of the project schedule — phases, durations,
                         and key dates formatted as a short paragraph (3–5 sentences)
  Cost_Breakdown       — string, a cost breakdown formatted as line items, one per line:
                         "Item name — £X,XXX" (if no currency given, assume GBP).
                         End with a total line: "TOTAL — £XX,XXX"

Rules:
- Write in Cyphr's voice: confident, direct, intelligent, no fluff.
- If a value is unknown or not in the source data, write "TBC" for short fields or a brief
  "To be confirmed following scoping." for longer fields.
- Do not invent specific numbers, names, or dates not present in the source data.
- Keep all values as flat strings (no nested objects or arrays).`;

  const requestBody = {
    system_instruction: { parts: [{ text: systemPrompt }] },
    contents: [{ role: 'user', parts: [{ text: 'Raw project data:\n\n' + rawContext }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      maxOutputTokens: 1500,
      temperature: 0.3
    }
  };

  const response     = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error('Gemini API error ' + responseCode + ': ' + responseText);
  }

  let rawJson;
  try {
    rawJson = JSON.parse(responseText).candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('Unexpected Gemini response shape: ' + responseText);
  }

  try {
    return JSON.parse(rawJson);
  } catch (e) {
    throw new Error('Gemini did not return valid JSON. Raw: ' + rawJson);
  }
}


/**
 * Copies the Slides template, replaces all {{tags}}, saves, returns the URL.
 */
function buildDeck(data) {
  const templateId = TEMPLATE_SLIDE_ID;
  const folderId   = OUTPUT_FOLDER_ID;
  if (!folderId || folderId === 'PASTE_YOUR_DRIVE_FOLDER_ID_HERE') throw new Error('OUTPUT_FOLDER_ID is not set at the top of the script.');

  // Inject the current month/year as {{Date}}
  data.Date = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );

  const clientName   = data.Client_Name   || 'Client';
  const projectTitle = data.Project_Title || 'Deck';
  const deckName     = clientName + ' — ' + projectTitle; // em dash

  // Copy template into output folder
  const newFile = DriveApp.getFileById(templateId)
    .makeCopy(deckName, DriveApp.getFolderById(folderId));

  const deck = SlidesApp.openById(newFile.getId());

  // Replace every {{tag}} in the deck
  for (var key in data) {
    if (Object.prototype.hasOwnProperty.call(data, key)) {
      var value = (data[key] !== null && data[key] !== undefined)
        ? String(data[key])
        : '—'; // em dash fallback
      deck.replaceAllText('{{' + key + '}}', value);
    }
  }

  deck.saveAndClose();

  return {
    deckUrl:  'https://docs.google.com/presentation/d/' + newFile.getId() + '/edit',
    deckName: deckName
  };
}
