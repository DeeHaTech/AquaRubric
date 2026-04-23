/**
 * Swimming Continuum - Full Code.gs
 * Model: Gemini Flash Latest
 * Persona: Educational Swim Coach (Gemma 3 Style)
 * Tone: Natural, Blended, Flowing Storytelling
 * Update: Removed Grammar Checker
 */

// --- 1. CONFIGURATION ---
// API key is stored in Script Properties, not source code.
// Set it via: Apps Script editor > Project Settings (gear) > Script Properties > Add
//   property: GEMINI_API_KEY
//   value:    <your key>
const AI_MODEL = "gemini-flash-latest";

function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error("GEMINI_API_KEY is not set. Open the Apps Script editor > Project Settings > Script Properties and add it.");
  }
  return key;
}

// --- STANDALONE MULTI-SPREADSHEET CONFIG ---
// SPREADSHEET_CONFIG is stored as a JSON string in Script Properties, mapping
// friendly labels to spreadsheet IDs, e.g.:
//   {"2025-2026":"1AbC...xyz","2024-2025":"1QrS...abc"}
// The allow-list prevents the deployed web app from opening any arbitrary
// spreadsheet the executing user happens to have access to.
function getSpreadsheetConfig() {
  const raw = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_CONFIG') || '{}';
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

// Returns the configured spreadsheets in insertion order (newest first, since
// addSpreadsheetToConfig prepends). Entries whose Drive file is trashed or
// inaccessible are HIDDEN from the dropdown but remain in SPREADSHEET_CONFIG —
// that way if a teacher accidentally trashes a file and restores it, the entry
// reappears automatically without needing to re-register.
function listSpreadsheets() {
  const config = getSpreadsheetConfig();
  const result = [];
  Object.keys(config).forEach(function(label) {
    const id = config[label];
    try {
      const file = DriveApp.getFileById(id);
      if (file.isTrashed()) return; // hide from UI this load; keep in config
      result.push({ label: label, id: id });
    } catch (e) {
      // File hard-deleted or temporarily inaccessible — hide but keep in config.
    }
  });
  return result;
}

// Prepend a new entry to SPREADSHEET_CONFIG so newest-created appears first
// in the dropdown (and becomes the default selection on load).
// Throws if a LIVE entry with the same label already exists. If the colliding
// entry's file is trashed or hard-deleted (a "ghost" left behind by the
// non-destructive listSpreadsheets filter), the new entry silently replaces it.
function addSpreadsheetToConfig(label, spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  const config = getSpreadsheetConfig();

  if (config[label]) {
    // Same label exists — only block if the existing file is still live.
    let existingIsLive = false;
    try {
      const existing = DriveApp.getFileById(config[label]);
      existingIsLive = !existing.isTrashed();
    } catch (e) {
      existingIsLive = false; // hard-deleted / inaccessible → treat as ghost
    }
    if (existingIsLive) {
      throw new Error('A spreadsheet named "' + label + '" already exists in the config.');
    }
    // Fall through and overwrite the ghost entry with the new ID.
  }

  // Rebuild with the new/updated key first. JSON object key order is preserved
  // in V8; listSpreadsheets() iterates via Object.keys() which respects it.
  const reordered = {};
  reordered[label] = spreadsheetId;
  Object.keys(config).forEach(function(k) {
    if (k !== label) reordered[k] = config[k];
  });
  props.setProperty('SPREADSHEET_CONFIG', JSON.stringify(reordered));
}

// Clone the template spreadsheet into the target Drive folder, wipe any
// existing student rows (keep header row 1 on every sheet), register it in
// SPREADSHEET_CONFIG, and return its metadata.
//
// Requires Script Properties:
//   TEMPLATE_SPREADSHEET_ID — ID of the master template spreadsheet
//   TARGET_FOLDER_ID       — Drive folder ID new copies go into
// `label`    — friendly name used in SPREADSHEET_CONFIG / dropdown (e.g. "2028-2029")
// `fileName` — actual Drive file name of the copy     (e.g. "2028-29 VS Swim Rubric")
//              falls back to `label` if omitted (keeps older callers working).
function createNewSpreadsheet(label, fileName) {
  const cleanLabel    = (label || '').toString().trim();
  const cleanFileName = (fileName || cleanLabel).toString().trim();
  if (!cleanLabel)    throw new Error('Please provide a name for the new spreadsheet.');
  if (!cleanFileName) throw new Error('Missing Drive file name.');
  if (cleanLabel.length > 80)    throw new Error('Name is too long (80 char max).');
  if (cleanFileName.length > 120) throw new Error('File name is too long.');

  const props = PropertiesService.getScriptProperties();
  const templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID');
  const folderId   = props.getProperty('TARGET_FOLDER_ID');
  if (!templateId) throw new Error('TEMPLATE_SPREADSHEET_ID is not set in Script Properties.');
  if (!folderId)   throw new Error('TARGET_FOLDER_ID is not set in Script Properties.');

  // Duplicate-label handling is delegated to addSpreadsheetToConfig(), which
  // permits overwriting a ghost entry (trashed/deleted file) but blocks
  // collisions with a live one.

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const folder = DriveApp.getFolderById(folderId);
    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(cleanFileName, folder);
    const newId = newFile.getId();

    // Wipe student rows on every tab so the new spreadsheet starts empty but
    // keeps the exact column layout from the template.
    const ss = SpreadsheetApp.openById(newId);
    ss.getSheets().forEach(function(sh) {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        const lastCol = sh.getLastColumn();
        sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    });

    addSpreadsheetToConfig(cleanLabel, newId);
    return { id: newId, label: cleanLabel, fileName: cleanFileName, url: ss.getUrl() };
  } finally {
    lock.releaseLock();
  }
}

function resolveSpreadsheet(spreadsheetId) {
  if (!spreadsheetId) throw new Error('Missing spreadsheet ID.');
  const config = getSpreadsheetConfig();
  const allowed = Object.keys(config).map(function(k) { return config[k]; });
  if (allowed.length && allowed.indexOf(spreadsheetId) === -1) {
    throw new Error('Spreadsheet ID is not in SPREADSHEET_CONFIG allow-list.');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

// Fetch the dolphin logo from Drive and return a base64 data URI.
// Used by the client-side PDF exporter (html2canvas cannot render
// cross-origin drive.google.com thumbnails directly). Cached 1 hour.
const LOGO_FILE_ID = '1k1hToRtNT9TWfhJKMbRrPCmDj4XfrFdW';
function getLogoDataUri() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('logoDataUri');
  if (cached) return cached;
  try {
    const blob = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
    const b64 = Utilities.base64Encode(blob.getBytes());
    const dataUri = 'data:' + blob.getContentType() + ';base64,' + b64;
    // Cache up to ~90KB (CacheService limit is 100KB per entry)
    if (dataUri.length < 95000) cache.put('logoDataUri', dataUri, 3600);
    return dataUri;
  } catch (e) {
    return '';
  }
}

// --- 2. GRADE LEVEL STANDARDS & RUBRIC TEXT ---
const STANDARDS = {
  1: { free: 2, back: 2, breast: 1, stamina: 10 },
  2: { free: 2, back: 2, breast: 2, stamina: 25 },
  3: { free: 3, back: 3, breast: 2, stamina: 50 },
  4: { free: 4, back: 4, breast: 3, stamina: 75 },
  5: { free: 4, back: 4, breast: 3, stamina: 100 }
};

const RUBRIC_TEXT = {
  free: {
    1: "enter water, blow bubbles, float face down, and exit the water safely",
    2: "swim in a prone position with alternating arm and leg action for a short time",
    3: "swim in a prone position with a body position approaching horizontal",
    4: "swim efficient freestyle with a horizontal body position, consistent kick, and proper side breathing",
    5: "swim exceptional freestyle with a streamlined horizontal body position, efficient pull, and exceptional timing"
  },
  back: {
    1: "float in a supine position",
    2: "swim in a supine position with alternating arm and leg action for a short time",
    3: "swim in a supine position with a body position approaching horizontal",
    4: "swim efficient backstroke with a horizontal body position, consistent kick, and proper timing",
    5: "swim exceptional backstroke with an efficient pull and exceptional timing"
  },
  breast: {
    1: "explore simultaneous pulls underwater",
    2: "demonstrate breaststroke movements on land and begin exploring them in the water",
    3: "swim in a prone position with a synchronized arm pull recovering underwater and whip kick",
    4: "swim legal breaststroke with proper pull, breathe, kick, glide sequence",
    5: "swim exceptional breaststroke demonstrating high efficiency and proper timing"
  }
};

// onOpen() removed: standalone Apps Script has no container spreadsheet to
// attach a menu to. The old draftAIComments batch action relied on the active
// selection in the bound sheet; the per-student AI button in the web UI now
// covers that workflow.

function extractScoreAndNotes(val) {
  if (!val) return { score: "", notes: "" };
  const str = String(val);
  const match = str.match(/(\d+(\.\d+)?)/);
  const score = match ? parseFloat(match[0]) : "";
  const notes = str.replace(score, "").trim();
  return { score: score, notes: notes };
}

function getRubricContext(stroke, score) {
  if (!score) return "is continuing to develop foundational skills";
  const level = Math.max(1, Math.min(5, Math.floor(score)));
  return RUBRIC_TEXT[stroke][level];
}

// Ability label for a score relative to the grade's expected level.
// Used by the AI prompt so the model receives MEANING, not rubric sentences.
function getAbilityLabel(score, expected) {
  if (!score) return "not yet assessed";
  const s = Math.floor(score);
  const diff = s - expected;
  if (diff <= -2) return "well below grade level (early learner)";
  if (diff === -1) return "approaching grade level (still developing)";
  if (diff === 0)  return "meeting grade level (proficient)";
  if (diff === 1)  return "above grade level (strong/capable)";
  return "well above grade level (exceptional)";
}

// Short descriptor of the SKILL FOCUS at this level — for technique hints,
// not for copy/paste into the paragraph.
function getSkillFocus(stroke, score) {
  if (!score) return "foundational water confidence";
  const level = Math.max(1, Math.min(5, Math.floor(score)));
  const FOCUS = {
    free:   { 1: "water entry and face-down floating", 2: "coordinating arms and legs over short distances", 3: "building a horizontal body position over 25m", 4: "efficient freestyle with side breathing and a steady kick", 5: "streamlined technique and refined timing" },
    back:   { 1: "supine floating", 2: "coordinating arms and legs on the back over short distances", 3: "building a horizontal body position on the back over 25m", 4: "efficient backstroke with a steady kick and proper timing", 5: "refined pull and timing" },
    breast: { 1: "simultaneous underwater pulls", 2: "learning breaststroke arms and legs in the water", 3: "synchronizing an underwater arm pull with a whip kick", 4: "the pull-breathe-kick-glide sequence", 5: "high efficiency and timing" }
  };
  return FOCUS[stroke][level];
}

// Absolute-skill descriptor — what the swimmer ACTUALLY looks like at this level,
// independent of grade expectations. Prevents the AI from calling a "still exploring"
// swimmer "proficient" just because they're above grade level.
function getAbsoluteSkillDescription(stroke, score) {
  if (!score) return "has not yet been assessed on this stroke";
  const level = Math.max(1, Math.min(5, Math.floor(score)));
  const DESC = {
    free: {
      1: "is still in the water-comfort stage — getting face in the water, floating, and exiting safely. They are NOT yet swimming.",
      2: "is beginning to swim short distances on their front but the stroke is still rough and tires quickly. Body position is not yet flat.",
      3: "can cover 25m on their front with a body position that is getting close to flat, but the stroke is not yet efficient and side breathing is still inconsistent.",
      4: "swims a proper, efficient freestyle — flat body, steady kick, proper side breathing and timing. This is competent freestyle.",
      5: "swims an exceptional, streamlined freestyle with refined pull, kick, and breathing timing. This is advanced."
    },
    back: {
      1: "can only float on their back — they are NOT yet swimming backstroke.",
      2: "is beginning to swim short distances on their back but the stroke is still rough. Body position is not yet flat.",
      3: "can cover 25m on their back with a body position getting close to flat, but the stroke is not yet efficient and timing is still inconsistent.",
      4: "swims a proper, efficient backstroke — flat body, steady kick, proper timing. This is competent backstroke.",
      5: "swims an exceptional backstroke with refined pull and timing. This is advanced."
    },
    breast: {
      1: "can only do simple underwater pulls — they are NOT yet swimming breaststroke in any form.",
      2: "is EXPLORING breaststroke movements in the water. The arm action and kick are NOT yet correct — they are experimenting with the shape of the stroke, not performing it. Do NOT describe them as 'coordinating' or 'in harmony'.",
      3: "can perform the correct breaststroke movements in the water (arm pull recovering underwater, whip-style kick), but the stroke is NOT yet legal competitive breaststroke — timing and finer form are still off.",
      4: "swims a legal breaststroke with the proper pull-breathe-kick-glide sequence. This is competent breaststroke.",
      5: "swims an exceptional, highly efficient breaststroke with refined timing. This is advanced."
    }
  };
  return DESC[stroke][level];
}

// --- CACHED HEADERS (5 min TTL, per spreadsheet+sheet) ---
// Cache key is scoped by spreadsheet ID so two different spreadsheets with
// different column orders don't collide.
function getCachedHeaders(sheet, sheetName) {
  const cache = CacheService.getScriptCache();
  const key = 'headers:' + sheet.getParent().getId() + ':' + sheetName;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => h.toString().toLowerCase().trim());
  cache.put(key, JSON.stringify(headers), 300);
  return headers;
}

function invalidateHeaderCache(spreadsheetId, sheetName) {
  CacheService.getScriptCache().remove('headers:' + spreadsheetId + ':' + sheetName);
}

// --- HEADER INDEX MAP (used by AI prompt + batch) ---
function getPromptHeaderIdx(headers) {
  return {
    firstName: headers.findIndex(h => h.includes("first")),
    fullName:  headers.findIndex(h => h.includes("full") || h === "name"),
    gender:    headers.findIndex(h => h.includes("gender") || h.includes("sex")),
    classVal:  headers.findIndex(h => h.includes("class")),
    free:      headers.findIndex(h => h.includes("free")),
    back:      headers.findIndex(h => h.includes("back")),
    breast:    headers.findIndex(h => h.includes("breast")),
    stamina:   headers.findIndex(h => h.includes("stamina") && !h.includes("kickboard")),
    comment:   headers.findIndex(h => h.includes("comment") && !h.includes("low"))
  };
}

// --- PARSE STUDENT ROW INTO PROMPT INPUTS ---
function parseStudentForPrompt(row, idx) {
  let studentName = "";
  if (idx.fullName > -1 && String(row[idx.fullName]).trim() !== "") {
    studentName = String(row[idx.fullName]).split(" ")[0];
  } else if (idx.firstName > -1 && String(row[idx.firstName]).trim() !== "") {
    studentName = String(row[idx.firstName]).trim();
  }

  let pronounContext = "Use the student's name or gender-neutral pronouns (they/them).";
  if (idx.gender > -1 && row[idx.gender]) {
    const g = String(row[idx.gender]).trim().toLowerCase();
    if (g === "m" || g === "male" || g === "boy") pronounContext = "Use male pronouns (he/him/his).";
    else if (g === "f" || g === "female" || g === "girl") pronounContext = "Use female pronouns (she/her/hers).";
  }

  const gradeMatch = String(row[idx.classVal]).match(/\d+/);
  const gradeLevel = gradeMatch ? parseInt(gradeMatch[0]) : 1;
  const expected = STANDARDS[gradeLevel] || STANDARDS[1];

  const freeData = extractScoreAndNotes(row[idx.free]);
  const backData = extractScoreAndNotes(row[idx.back]);
  const breastData = extractScoreAndNotes(row[idx.breast]);

  let staminaData = "0";
  if (idx.stamina > -1 && row[idx.stamina]) {
    staminaData = String(row[idx.stamina]).replace(/[^0-9.+]/g, '');
  }

  return { studentName, pronounContext, expected, freeData, backData, breastData, staminaData };
}

// --- SHARED AI PROMPT BUILDER ---
function buildAIPrompt(p) {
  const freeAbility   = getAbilityLabel(p.freeData.score,   p.expected.free);
  const backAbility   = getAbilityLabel(p.backData.score,   p.expected.back);
  const breastAbility = getAbilityLabel(p.breastData.score, p.expected.breast);

  const freeDesc   = getAbsoluteSkillDescription('free',   p.freeData.score);
  const backDesc   = getAbsoluteSkillDescription('back',   p.backData.score);
  const breastDesc = getAbsoluteSkillDescription('breast', p.breastData.score);

  const staminaGap = Number(p.staminaData) < Number(p.expected.stamina);

  return `
    You are an experienced Swim Coach writing a warm, natural end-of-unit report for a parent to read. Your job is to describe the swimmer as a person with a skill profile — NOT to list rubric bullet points.

    PRONOUN INSTRUCTION: ${p.pronounContext}

    ==========================================
    STUDENT CONTEXT (for YOUR understanding — do NOT quote this text back)
    ==========================================
    Name: ${p.studentName}

    WHAT THE SWIMMER CAN ACTUALLY DO (describe THIS, not grade level):
    - Freestyle (score ${p.freeData.score || 'N/A'}): ${freeDesc}
    - Backstroke (score ${p.backData.score || 'N/A'}): ${backDesc}
    - Breaststroke (score ${p.breastData.score || 'N/A'}): ${breastDesc}

    How this compares to grade expectations (for tone calibration only):
    - Freestyle: ${freeAbility}
    - Backstroke: ${backAbility}
    - Breaststroke: ${breastAbility}
    - Stamina: swam ${p.staminaData}m vs target ${p.expected.stamina}m (${staminaGap ? 'below target' : 'at or above target'})

    CRITICAL CALIBRATION RULE:
    Describe what the swimmer ACTUALLY does (the absolute skill), NOT how they compare to their grade.
    A young student who is "above grade level" may still only be EXPLORING a stroke — do not promote them to "proficient" or "coordinated" or "in harmony" just because they are ahead of expectations. If the absolute description says "still exploring" or "NOT yet correct", your comment MUST reflect that honestly (e.g., "is beginning to explore breaststroke movements", "is experimenting with the arm and leg shapes"). You may note that they are ahead of expectations, but accuracy about what they can do comes first.

    Teacher's notes on current errors (may be empty):
    - Freestyle: "${p.freeData.notes || 'none'}"
    - Backstroke: "${p.backData.notes || 'none'}"
    - Breaststroke: "${p.breastData.notes || 'none'}"

    ==========================================
    HOW TO WRITE THE COMMENT
    ==========================================

    OPENING (required): Pick ONE randomly and use it verbatim:
    - "It's been a joy coaching ${p.studentName} this unit!"
    - "It's been a pleasure coaching ${p.studentName} this unit!"
    - "It's been wonderful coaching ${p.studentName} this unit!"

    VOICE: Warm, confident, specific. Write as a coach who actually watched this child swim — not as someone summarizing a spreadsheet.

    ANTI-PARROTING RULE (MOST IMPORTANT):
    Do NOT copy or lightly paraphrase rubric phrases. Forbidden phrases include but are not limited to: "blow bubbles", "float face down", "enter the water safely", "supine position", "prone position", "alternating arm and leg action", "approaching horizontal", "synchronized arm pull", "whip kick", "pull-breathe-kick-glide".
    Instead, translate each skill into natural coach-speak. Examples of good translation:
      - "float face down" → "is comfortable stretching out on the surface" or "has found her balance in the water"
      - "alternating arm and leg action" → "is starting to coordinate her arms and legs"
      - "body position approaching horizontal" → "is flattening out nicely as she swims"
      - "legal breaststroke with pull-breathe-kick-glide" → "her breaststroke rhythm is coming together"
    You may name the STROKE ("freestyle", "backstroke", "breaststroke") — that is not parroting.

    WORD CHOICE BY ABSOLUTE SKILL LEVEL (NOT grade-relative):
    - Score 1: "is building water confidence", "is learning to...", "is comfortable in the water" — do NOT imply they are swimming the stroke.
    - Score 2 freestyle/backstroke: "is beginning to swim short distances on his front/back", "is starting to connect arms and legs" — stroke is rough.
    - Score 2 breaststroke SPECIFICALLY: "is exploring the shapes of breaststroke", "is experimenting with the arm and leg movements" — NEVER say "coordinated", "in harmony", "working together", "proficient", or imply the stroke is correct.
    - Score 3: "is flattening out", "is putting the pieces together", "the stroke is taking shape" — correct-ish but not yet efficient/legal.
    - Score 4: "proficient", "capable", "efficient", "swims a proper [stroke]" — confident language, no hedging.
    - Score 5: "exceptional", "refined", "stands out", "advanced" — highlight what makes them strong.

    Avoid "developing" unless the student is actually below grade level.

    TEACHER NOTES = MISTAKES:
    If a teacher note exists, it describes something the student is CURRENTLY doing wrong. Flip it into positive practice advice (e.g., "arms wide" → "practice reaching straight back"). Never tell them to keep doing the mistake. If notes are "none", do not invent a correction — just describe proficiency.

    STAMINA:
    ${staminaGap
      ? `Stamina is below target — weave in ONE brief line about how building endurance will help ${p.studentName} hold technique longer. No numbers, no "meters".`
      : `Stamina is at or above target — do NOT mention stamina, endurance, or distance at all.`}

    BANNED WORDS: "unassisted", "independently", "beautiful", "lovely", "will aim to", "technical goals", "solid progress", "shows promise", "Level", "Grade", "Standard", "Score", "Meters", and any number followed by "m".

    FORMAT:
    - One flowing paragraph, 3–5 sentences.
    - Cover all three strokes, but combine them naturally where it reads better (e.g., "Her freestyle and backstroke are both proficient...").
    - End with exactly: "Great work this unit!"

    Write the comment now. Output only the paragraph — no headings, no preamble.
  `;
}

// draftAIComments() removed: it depended on getActiveSheet()/getActiveRange(),
// which only exist in a container-bound script. The per-student "AI Comment"
// button (generateCommentForStudent) replaces it in the standalone web app.

// --- 4. HELPER: CALL GEMINI API WITH RETRY LOGIC ---
function callGeminiAI(promptText, isJson = false, maxRetries = 3) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${AI_MODEL}:generateContent?key=${getApiKey()}`;
  const payload = { 
    "contents": [{ "parts": [{ "text": promptText }] }],
    "generationConfig": { "temperature": 0.8 }
  };

  if (isJson) {
    payload.generationConfig.responseMimeType = "application/json";
    payload.generationConfig.temperature = 0.2; 
  }
  
  const options = { 
    "method": "post", 
    "contentType": "application/json", 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  
  let attempt = 0;
  let delay = 2000; 

  while (attempt < maxRetries) {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (!json.error) {
      return json.candidates[0].content.parts[0].text.trim();
    }
    
    if (json.error.code === 503 || json.error.code === 429 || json.error.message.toLowerCase().includes("high demand")) {
      attempt++;
      if (attempt >= maxRetries) {
        throw new Error(`Failed after ${maxRetries} attempts. Google servers are experiencing high demand.`);
      }
      console.log('Gemini server busy. Retrying in ' + (delay/1000) + 's (attempt ' + attempt + '/' + maxRetries + ')');
      Utilities.sleep(delay);
      delay *= 2; 
    } else {
      throw new Error(json.error.message);
    }
  }
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.spreadsheets = listSpreadsheets();                        // [{label,id},...]
  template.initialSpreadsheetId = (e && e.parameter && e.parameter.ssid) || '';
  return template.evaluate()
      .setTitle('Swim Rubric Generator')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getData(spreadsheetId, sheetName) {
  const ss = resolveSpreadsheet(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: "Tab not found!" };

  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(h => h.toString().toLowerCase().trim()); 

  const idx = {
    firstName: headers.findIndex(h => h.includes("first")),
    lastName:  headers.findIndex(h => h.includes("last")),
    fullName:  headers.findIndex(h => h.includes("full") || h === "name"),
    gender:    headers.findIndex(h => h.includes("gender")),
    classVal:  headers.findIndex(h => h.includes("class")),
    teacher:   headers.findIndex(h => h.includes("teacher")),
    free:      headers.findIndex(h => h.includes("free")),
    back:      headers.findIndex(h => h.includes("back")),
    breast:    headers.findIndex(h => h.includes("breast")),
    stamina:   headers.findIndex(h => h.includes("stamina") && !h.includes("kickboard")),
    staminaKb: headers.findIndex(h => h.includes("stamina") && h.includes("kickboard")),
    comments:  headers.findIndex(h => h.includes("comment") && !h.includes("low")),
    commentsLow: headers.findIndex(h => h.includes("comment") && h.includes("low")),
    spec1:     headers.findIndex(h => h.includes("face")),
    spec2:     headers.findIndex(h => h.includes("bubbles")),
    spec3:     headers.findIndex(h => h.includes("bob")),
    spec4:     headers.findIndex(h => h.includes("float")),
    spec5:     headers.findIndex(h => h.includes("flutter")),
    rubricType: headers.findIndex(h => h.includes("use low level") || (h.includes("low level") && h.includes("rubric")))
  };

  const extractNumber = (val) => {
    if (!val) return "";
    const match = val.toString().match(/(\d+(\.\d+)?)/); 
    return match ? parseFloat(match[0]) : ""; 
  };

  const extractStamina = (val) => {
    if (!val) return "";
    return val.toString().replace(/[^0-9.+]/g, ''); 
  };

  const extractNotes = (val) => {
    if (!val) return "";
    const str = val.toString();
    const match = str.match(/^\s*\d+(\.\d+)?\s*(.*)$/);
    return match ? match[2].trim() : str.trim();
  };

  const extractSpecNotes = (val) => {
    if (!val) return "";
    const str = val.toString().trim();
    if (str.toLowerCase().startsWith('o')) return str.substring(1).trim();
    return "";
  };

  return values.map((row, i) => ({
    rowIndex:  i + 2, // 1-based, accounting for header row
    firstName: idx.firstName > -1 ? row[idx.firstName] : "",
    lastName:  idx.lastName > -1 ? row[idx.lastName] : "",
    fullName:  idx.fullName > -1 ? row[idx.fullName] : "",
    gender:    idx.gender > -1 ? row[idx.gender] : "",
    class:     idx.classVal > -1 ? row[idx.classVal] : "",
    teacher:   idx.teacher > -1 ? row[idx.teacher] : "",
    free:      idx.free > -1 ? extractNumber(row[idx.free]) : "",
    back:      idx.back > -1 ? extractNumber(row[idx.back]) : "",
    breast:    idx.breast > -1 ? extractNumber(row[idx.breast]) : "",
    freeNotes: idx.free > -1 ? extractNotes(row[idx.free]) : "",
    backNotes: idx.back > -1 ? extractNotes(row[idx.back]) : "",
    breastNotes: idx.breast > -1 ? extractNotes(row[idx.breast]) : "",
    stamina:   idx.stamina > -1 ? extractStamina(row[idx.stamina]) : "",
    staminaKb: idx.staminaKb > -1 ? extractStamina(row[idx.staminaKb]) : "",
    comments:  idx.comments > -1 ? row[idx.comments] : "",
    commentsLow: idx.commentsLow > -1 ? row[idx.commentsLow] : "",
    spec1:     idx.spec1 > -1 ? row[idx.spec1].toString().trim() : "",
    spec2:     idx.spec2 > -1 ? row[idx.spec2].toString().trim() : "",
    spec3:     idx.spec3 > -1 ? row[idx.spec3].toString().trim() : "",
    spec4:     idx.spec4 > -1 ? row[idx.spec4].toString().trim() : "",
    spec5:     idx.spec5 > -1 ? row[idx.spec5].toString().trim() : "",
    spec1Notes: idx.spec1 > -1 ? extractSpecNotes(row[idx.spec1]) : "",
    spec2Notes: idx.spec2 > -1 ? extractSpecNotes(row[idx.spec2]) : "",
    spec3Notes: idx.spec3 > -1 ? extractSpecNotes(row[idx.spec3]) : "",
    spec4Notes: idx.spec4 > -1 ? extractSpecNotes(row[idx.spec4]) : "",
    spec5Notes: idx.spec5 > -1 ? extractSpecNotes(row[idx.spec5]) : "",
    rubricType: idx.rubricType > -1 ? row[idx.rubricType].toString().trim() : ""
  })).filter(s => s.firstName || s.lastName || s.fullName);
}

// --- 5. EDIT MODE: FIELD KEY -> HEADER MATCHER ---
function getFieldColumnIndex(headers, fieldKey) {
  const h = headers;
  switch (fieldKey) {
    case 'free':     return h.findIndex(x => x.includes("free"));
    case 'back':     return h.findIndex(x => x.includes("back"));
    case 'breast':   return h.findIndex(x => x.includes("breast"));
    case 'stamina':  return h.findIndex(x => x.includes("stamina") && !x.includes("kickboard"));
    case 'staminaKb':return h.findIndex(x => x.includes("stamina") && x.includes("kickboard"));
    case 'comments': return h.findIndex(x => x.includes("comment") && !x.includes("low"));
    case 'commentsLow': return h.findIndex(x => x.includes("comment") && x.includes("low"));
    case 'spec1':    return h.findIndex(x => x.includes("face"));
    case 'spec2':    return h.findIndex(x => x.includes("bubbles"));
    case 'spec3':    return h.findIndex(x => x.includes("bob"));
    case 'spec4':    return h.findIndex(x => x.includes("float"));
    case 'spec5':    return h.findIndex(x => x.includes("flutter"));
    case 'rubricType': return h.findIndex(x => x.includes("use low level") || (x.includes("low level") && x.includes("rubric")));
    default: return -1;
  }
}

// --- 6. EDIT MODE: SAVE SINGLE FIELD ---
// LockService prevents concurrent write collisions when 3-4 teachers save simultaneously.
// Each write holds the lock for ~200-500ms; at most 3-4 teachers queued = <2s worst-case wait.
function saveStudentField(spreadsheetId, sheetName, rowIndex, fieldKey, value) {
  const ss = resolveSpreadsheet(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const headers = getCachedHeaders(sheet, sheetName);
  const colIdx = getFieldColumnIndex(headers, fieldKey);
  if (colIdx === -1) throw new Error("Column not found for field: " + fieldKey);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.getRange(rowIndex, colIdx + 1).setValue(value);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// --- 7. EDIT MODE: SWITCH RUBRIC TYPE (write column only; do not clear data) ---
function switchStudentRubric(spreadsheetId, sheetName, rowIndex, newType) {
  const ss = resolveSpreadsheet(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const headers = getCachedHeaders(sheet, sheetName);
  const typeColIdx = getFieldColumnIndex(headers, 'rubricType');
  if (typeColIdx === -1) throw new Error("Column 'Use Low Level Rubric?' not found.");

  const label = newType === 'lowLevel' ? 'Low-Level Rubric' : 'Standard Rubric';

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.getRange(rowIndex, typeColIdx + 1).setValue(label);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// --- 8. EDIT MODE: AI COMMENT FOR A SINGLE STUDENT ---
function generateCommentForStudent(spreadsheetId, sheetName, rowIndex) {
  const ss = resolveSpreadsheet(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const headers = getCachedHeaders(sheet, sheetName);
  const idx = getPromptHeaderIdx(headers);
  const row = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  const parsed = parseStudentForPrompt(row, idx);
  if (!parsed.studentName) throw new Error("Student name not found.");

  let aiComment = callGeminiAI(buildAIPrompt(parsed), false);
  aiComment = aiComment.replace(/\r?\n|\r/g, " ").replace(/\s+/g, " ").trim();

  if (idx.comment > -1) {
    sheet.getRange(rowIndex, idx.comment + 1).setValue(aiComment);
  }
  return { ok: true, comment: aiComment };
}