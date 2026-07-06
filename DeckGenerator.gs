// ═══════════════════════════════════════════════════════════════════════
// CYPHR DECK GENERATOR — Google Apps Script Web App
// ═══════════════════════════════════════════════════════════════════════
//
// HOW TO DEPLOY:
//   1. Open https://script.google.com and create a new project.
//   2. Paste this entire file into the editor (replace any existing code).
//   3. Set Script Properties (Project Settings > Script Properties):
//        GEMINI_API_KEY   — your Google AI Studio API key
//        TEMPLATE_SLIDE_ID — the file ID of your Cyphr Slides template
//                            (the long ID in the Google Slides URL)
//        OUTPUT_FOLDER_ID  — the Google Drive folder ID where decks are saved
//   4. Click Deploy > New Deployment.
//        - Type: Web App
//        - Execute as: Me
//        - Who has access: Anyone
//   5. Copy the Web App URL — paste it into cyphr-flow as VITE_DECK_GEN_URL
//      (or wherever the front-end reads its endpoint).
//
// INCOMING POST BODY (JSON):
//   { clientName, projectName, sector, budget, timeline,
//     requirements, bgNotes, stakeholders, nextSteps, briefOutput }
//
// RESPONSE (JSON):
//   Success: { deckUrl, deckName }
//   Error:   { error: "message" }
// ═══════════════════════════════════════════════════════════════════════


/**
 * Entry point for the Web App.
 * Parses the POST body, calls compressForSlides, then buildDeck.
 */
function doPost(e) {
  // Apps Script Web Apps cannot respond to preflight OPTIONS requests —
  // CORS headers are set on the text output object below.
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST body received.');
    }

    const projectData = JSON.parse(e.postData.contents);

    // Step 1 — compress raw project data into structured slide copy
    const compressed = compressForSlides(projectData);

    // Step 2 — create the deck from the template
    const result = buildDeck(compressed);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errBody = JSON.stringify({ error: err.message || String(err) });
    return ContentService
      .createTextOutput(errBody)
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Calls the Gemini API to compress raw project data into structured
 * slide-ready copy.  Returns a plain JS object with the keys the
 * template placeholders expect.
 *
 * @param {Object} projectData  Raw fields from cyphr-flow.
 * @returns {Object}            Structured data object for buildDeck().
 */
function compressForSlides(projectData) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Script Property GEMINI_API_KEY is not set.');

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' +
    apiKey;

  // Build a readable summary of what was passed in
  const rawContext = [
    'Client Name: '      + (projectData.clientName   || ''),
    'Project Name: '     + (projectData.projectName  || ''),
    'Sector: '           + (projectData.sector        || ''),
    'Budget: '           + (projectData.budget        || ''),
    'Timeline: '         + (projectData.timeline      || ''),
    'Requirements: '     + (projectData.requirements  || ''),
    'Background Notes: ' + (projectData.bgNotes       || ''),
    'Stakeholders: '     + (projectData.stakeholders  || ''),
    'Next Steps: '       + (projectData.nextSteps     || ''),
    'AI Brief Output: '  + (projectData.briefOutput   || ''),
  ].join('\n');

  const systemPrompt = `You are a professional copywriter for Cyphr, a creative strategy and innovation studio.
Your job is to distil raw project intake data into polished, concise slide copy for a client-facing deck.
Return ONLY a single valid JSON object — no markdown, no code fences, no commentary.

The JSON must contain exactly these keys:
  Client_Name          — string, the client's company or individual name
  Project_Title        — string, a punchy 4-8 word project title
  Sector               — string, the industry / sector
  Budget               — string, budget formatted as £X,XXX or £XX,XXX (if no currency symbol given, assume GBP)
  Timeline             — string, project timeline / deadline
  Executive_Summary    — string, 2-3 polished sentences summarising the whole engagement
  Background           — string, 3-4 sentences giving context about the client and their challenge
  Key_Takeaway_1       — string, max 15 words, a punchy bullet (no leading dash)
  Key_Takeaway_2       — string, max 15 words, a punchy bullet (no leading dash)
  Key_Takeaway_3       — string, max 15 words, a punchy bullet (no leading dash)
  Proposed_Approach    — string, 3-4 sentences describing how Cyphr would tackle this project
  Next_Steps           — string, a short bullet list separated by " | " (pipe character), max 5 items
  Stakeholders         — string, each person as "Name — Role", separated by newline, max 5 people

Rules:
- Write in Cyphr's voice: confident, direct, intelligent, no fluff.
- If a value is unknown or not provided, write "TBC" rather than leaving it blank.
- Do not invent specific numbers, dates, or names that were not in the source data.
- Keep all values as flat strings (no nested objects or arrays).`;

  const requestBody = {
    system_instruction: {
      parts: [{ text: systemPrompt }]
    },
    contents: [
      {
        role: 'user',
        parts: [{ text: 'Here is the raw project data:\n\n' + rawContext }]
      }
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      maxOutputTokens: 1200,
      temperature: 0.3
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error('Gemini API error ' + responseCode + ': ' + responseText);
  }

  const geminiResponse = JSON.parse(responseText);

  // Extract the model's text output
  let rawJson;
  try {
    rawJson = geminiResponse.candidates[0].content.parts[0].text;
  } catch (parseErr) {
    throw new Error('Unexpected Gemini response shape: ' + responseText);
  }

  // Parse the JSON the model returned
  let structured;
  try {
    structured = JSON.parse(rawJson);
  } catch (jsonErr) {
    throw new Error('Gemini did not return valid JSON. Raw output: ' + rawJson);
  }

  return structured;
}


/**
 * Copies the Slides template into the output folder, replaces all
 * {{placeholder}} tokens with values from data, and returns the deck URL.
 *
 * @param {Object} data  Structured data from compressForSlides().
 * @returns {{ deckUrl: string, deckName: string }}
 */
function buildDeck(data) {
  const props = PropertiesService.getScriptProperties();

  const templateId = props.getProperty('TEMPLATE_SLIDE_ID');
  if (!templateId) throw new Error('Script Property TEMPLATE_SLIDE_ID is not set.');

  const folderId = props.getProperty('OUTPUT_FOLDER_ID');
  if (!folderId) throw new Error('Script Property OUTPUT_FOLDER_ID is not set.');

  // Add a formatted date so the template can show {{Date}}
  data.Date = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );

  // Build a safe deck name from client + project title
  const clientName   = data.Client_Name   || 'Client';
  const projectTitle = data.Project_Title || 'Deck';
  const deckName     = clientName + ' — ' + projectTitle; // em dash

  // Copy the template into the output folder
  const templateFile = DriveApp.getFileById(templateId);
  const outputFolder = DriveApp.getFolderById(folderId);
  const newFile      = templateFile.makeCopy(deckName, outputFolder);
  const newFileId    = newFile.getId();

  // Open as a Presentation and replace all placeholders
  const deck = SlidesApp.openById(newFileId);

  for (const key in data) {
    if (Object.prototype.hasOwnProperty.call(data, key)) {
      const placeholder = '{{' + key + '}}';
      const value       = (data[key] !== null && data[key] !== undefined)
        ? String(data[key])
        : '—'; // em dash fallback for empty values
      deck.replaceAllText(placeholder, value);
    }
  }

  deck.saveAndClose();

  const deckUrl = 'https://docs.google.com/presentation/d/' + newFileId + '/edit';

  return { deckUrl: deckUrl, deckName: deckName };
}
