// =======================
// ON OPEN MENU
// =======================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔑 API Key Tools")
    .addItem("Reset API Key Now", "manualResetApiKey")
    .addItem("Refresh Status Panel", "updateStatusPanel")
    .addSeparator()
    .addItem("Re-initialize Setup", "firstTimeUserSetup") // optional, safe to rerun
    .addItem("Data & Query 1.1 release","initializeDashboard")
    .addToUi();

  initializeDashboard();
  insertWelcomeNote();
  test();
}





// =======================
// 1. FIRST-TIME / RE-INITIALIZE SETUP
// =======================
function firstTimeUserSetup() {
  const ui = SpreadsheetApp.getUi();

  // Authorize script (dummy call to trigger auth)
  authorizeScript();

  // Install required triggers
  ScriptApp.newTrigger("onEditHandler")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  ScriptApp.newTrigger("onEditSettings")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  ScriptApp.newTrigger("onOpen")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();

  // Insert welcome/instructions note
  insertWelcomeNote();


  ui.alert("✅ Setup completed!\nAll triggers installed, protections applied, and instructions added.");
  setApiKeyAndFetchUrl();
  updateStatusPanel();

}




// =======================
// 2. AUTHORIZE SCRIPT (dummy call)
// =======================
function authorizeScript() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var file = DriveApp.getFilesByName("dummy"); // touches Drive scope
  SpreadsheetApp.getUi().alert("✅ Script authorized successfully.");

}


// =======================
// 3. REMOVE EXISTING TRIGGERS (internal use)
// =======================
function removeExistingTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}




// =======================
// 4. MANUAL RESET API KEY
// =======================
function manualResetApiKey() {
  PropertiesService.getScriptProperties().setProperty("API_KEY", "");
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Settings")   // 👈 directly refer to the "Settings" sheet
    .getRange("B6")
    .setValue("No API Key Stored");

  PropertiesService.getScriptProperties().setProperty("FETCH_URL", "");
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Settings")   // 👈 directly refer to the "Settings" sheet
    .getRange("B8")
    .setValue("No Fetch URL Stored");

  PropertiesService.getScriptProperties().setProperty("AI_MODEL", "");
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Settings")   // 👈 directly refer to the "Settings" sheet
    .getRange("B10")
    .setValue("No AI Model Stored");


  SpreadsheetApp.getUi().alert("⚠️ API Key, Fetch URL and AI Model reset via menu!");
  updateStatusPanel();

}




// SETUP API KEY AND FETCH URL
function setApiKeyAndFetchUrl() {
  // 🔑 Read entered information at B7 and B9
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var e = ss.getSheetByName("Settings");

  // const myApiKey = e.getRange("B7").getDisplayValue().toString().trim();
  const myApiKey = String(e.getRange("B7").getDisplayValue()).trim().replace(/[^\x20-\x7E]/g, '').normalize("NFKC");


  // const myFetchUrl = e.getRange("B9").getDisplayValue().toString().trim();
  const myFetchUrl = String(e.getRange("B9").getDisplayValue()).trim().replace(/[^\x20-\x7E]/g, '').normalize("NFKC");


  // const myAIModel = e.getRange("B11").getDisplayValue().toString().trim();
  const myAIModel = String(e.getRange("B11").getDisplayValue()).trim().replace(/[^\x20-\x7E]/g, '').normalize("NFKC");


  // Save it in Script Properties
  if (myApiKey) {
    PropertiesService.getScriptProperties().setProperty("API_KEY", myApiKey);
    e.getRange("B7").setValue(""); // clear B7 after saving
    SpreadsheetApp.getUi().alert("✅ API key saved successfully.");

  //  const checkKey = PropertiesService.getScriptProperties().getProperty("API_KEY");
  //  Logger.log(`Saved key: ${checkKey}`);
  //  Logger.log("API Key saved. Length: " + checkKey.length);


  } else {
    SpreadsheetApp.getUi().alert("⚠️ No API key entered at Settings!B7.");
  }
 



  if (myFetchUrl) {
    PropertiesService.getScriptProperties().setProperty("FETCH_URL", myFetchUrl);
    e.getRange("B9").setValue(""); // clear B9 after saving
    SpreadsheetApp.getUi().alert("✅ Fetch URL saved successfully.");

  //  const checkUrl = PropertiesService.getScriptProperties().getProperty("FETCH_URL");
  //  Logger.log(`Saved Url: ${checkUrl}`);
  //  Logger.log("Fetch Url saved. Length: " + checkUrl.length);


  } else {
    SpreadsheetApp.getUi().alert("⚠️ No Fetch URL entered at Settings!B9.");
  }



    if (myAIModel) {
    PropertiesService.getScriptProperties().setProperty("AI_MODEL", myAIModel);
    e.getRange("B11").setValue(""); // clear B11 after saving
    SpreadsheetApp.getUi().alert("✅ AI Model saved successfully.");

  //  const checkModel = PropertiesService.getScriptProperties().getProperty("AI_MODEL");
  //  Logger.log(`Saved AI Model: ${checkModel}`);
  //  Logger.log("AI Model saved. Length: " + checkModel.length);


  } else {
    SpreadsheetApp.getUi().alert("⚠️ No AI Model entered at Settings!B11.");
  }

}








// Check for a valid API Key saved in Script Properties and Return
function getApiKey() {
  const rawKey = PropertiesService.getScriptProperties().getProperty("API_KEY");
  const API_KEY = rawKey ? String(rawKey).trim() : "";
  // Logger.log(`Saved key: ${API_KEY}`);
  // Logger.log("API Key saved. Length: " + API_KEY.length);

  if (!API_KEY) {
    throw new Error("⚠️ No API key found. Please enter one in Settings!B7.");
  }
  return API_KEY;
}




// Check for a valid Fetch URL saved in Script Properties and Return
function getFetchUrl() {
  const rawUrl = PropertiesService.getScriptProperties().getProperty("FETCH_URL");
  const FETCH_URL = rawUrl ? String(rawUrl).trim() : "";
  // Logger.log(`Saved key: ${FETCH_URL}`);
  // Logger.log("API Key saved. Length: " + FETCH_URL.length);

  if (!FETCH_URL) {
    throw new Error("⚠️ No Fetch URL found. Please enter one in Settings!B9.");
  }
  return FETCH_URL;
}



// Check for a valid AI Model saved in Script Properties and Return
function getAIModel() {
  const rawModel = PropertiesService.getScriptProperties().getProperty("AI_MODEL");
  const AI_MODEL = rawModel ? String(rawModel).trim() : "";
  // Logger.log(`Saved AI Model: ${AI_MODEL}`);
  // Logger.log("AI Model saved. Length: " + AI_MODEL.length);

  if (!AI_MODEL) {
    throw new Error("⚠️ No AI Model found. Please enter one in Settings!B11.");
  }
  return AI_MODEL;
}



// GEMINI API CONNECTION
function callGeminiDynamic(prompt,apiKey, fetchUrl, AIModel) {
  const baseUrl = fetchUrl;
  const finalUrl = `${baseUrl}${AIModel}:generateContent?key=${apiKey}`;
  const payload = {
    contents: [
      { parts: [{ text: prompt }] }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(finalUrl, options);
  const json = JSON.parse(response.getContentText());
  return json;
}





// 🔧 Smart cleaner to fix spacing & punctuation
function cleanOutput(text) {
  return text
    .replace(/\s+([.,!?;:])/g, "$1")              // spaces before punctuation
    .replace(/\s+’\s+/g, "’")                     // spaces around curly apostrophes
    .replace(/\s+'(\w)/g, "’$1")                  // straight apostrophe into curly
    .replace(/(\w)\s*-\s*(\w)/g, "$1-$2")         // hyphenated words
    .replace(/(\d+)\s*[–—-]\s*(\d+)/g, "$1–$2")   // number ranges (10–15)
    .replace(/–\s*(\d+)/g, "–$1")                 // en dash cleanup
    .replace(/-\s*(minute|hour|day|week|month|year|priority|back)/gi, "-$1") // common compounds
    .replace(/\s{2,}/g, " ")                      // collapse multiple spaces
    .trim();
}








function runReplicateModel(promptText, apiKey, fetchUrl, AIModel, maxRetries) {
  var apiKey = getApiKey();
  var fetchUrl = getFetchUrl();
  var AIModel = getAIModel();
  var Retries = maxRetries || 20;

  // 🔹 Prevent overlapping runs
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // 🔹 Check cache first
    var cache = CacheService.getUserCache();
    var cacheKey = "replicate:" + Utilities.base64Encode(promptText).substring(0, 100);
    var cachedResponse = cache.get(cacheKey);
    if (cachedResponse) {
      Logger.log("Cache hit for prompt: " + promptText);
      return cachedResponse;
    }

    // 🔹 Send initial request
    const payload = {
      version: AIModel,
      input: { prompt: promptText }
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${apiKey}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const initialResp = UrlFetchApp.fetch(fetchUrl, options);
    if (initialResp.getResponseCode() !== 201) {
      throw new Error(`Request failed (HTTP ${initialResp.getResponseCode()}): ${initialResp.getContentText()}`);
    }

    const initialJson = JSON.parse(initialResp.getContentText());
    const statusUrl = initialJson.urls.get;

    let statusObj;
    let attempts = 0;

    // 🔹 Poll until finished AND output is present
    do {
      Utilities.sleep(3000);
      attempts++;

      const statusResp = UrlFetchApp.fetch(statusUrl, {
        headers: { Authorization: `Bearer ${apiKey}` },
        muteHttpExceptions: true
      });

      if (statusResp.getResponseCode() < 200 || statusResp.getResponseCode() >= 300) {
        throw new Error(`Status check failed (HTTP ${statusResp.getResponseCode()}): ${statusResp.getContentText()}`);
      }

      statusObj = JSON.parse(statusResp.getContentText());

      if (statusObj.status === 'failed') {
        const errMsg = statusObj.error ? JSON.stringify(statusObj.error) : 'Unknown error';
        throw new Error(`Prediction failed: ${errMsg}`);
      }

    } while ((statusObj.status !== 'succeeded' || !statusObj.output || statusObj.output.length === 0) 
             && attempts < Retries);

    if (statusObj.status !== 'succeeded' || !statusObj.output) {
      throw new Error(`Prediction did not succeed after ${Retries} attempts.`);
    }

    // 🔹 Normalize output → always return a single string
    const rawOutput = statusObj.output;
    let finalText = "";

    if (Array.isArray(rawOutput)) {
      // Join all array elements into one string, filter empty values
      finalText = cleanOutput(rawOutput.filter(t => t).join(""));
    } else if (typeof rawOutput === 'string') {
      finalText = cleanOutput(rawOutput);
    } else {
      // Fallback for unexpected format
      finalText = cleanOutput(String(rawOutput));
    }

    // 🔹 Cache final result (TTL = 60s)
    cache.put(cacheKey, finalText, 60);

    // 🔹 Return a single string
    return finalText;

  } finally {
    lock.releaseLock();
  }
}








// =======================
// 6. STATUS PANEL (SETTINGS TAB)
// =======================
function updateStatusPanel() {
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const props = PropertiesService.getScriptProperties();

  const key = props.getProperty("API_KEY");
  const url = props.getProperty("FETCH_URL");
  const model = props.getProperty("AI_MODEL");
  const triggers = ScriptApp.getProjectTriggers();
  const hasEdit1 = triggers.some(t => t.getHandlerFunction() === "onEditHandler");
  const hasEdit2 = triggers.some(t => t.getHandlerFunction() === "onEditSettings");
  const hasOpen = triggers.some(t => t.getHandlerFunction() === "onOpen");

 // 🔑 Always restore B6 from Script Properties
  if (key) {
    const b6 = sheet.getRange("B6");
    if (b6.getValue() !== key) {
      b6.setValue(key);

     // Logger.log(`Saved key: ${key}`);
     // Logger.log("API Key saved. Length: " + key.length);
    }
  }

 // 🔑 Always restore B8 from Script Properties
  if (url) {
    const b8 = sheet.getRange("B8");
    if (b8.getValue() !== url) {
      b8.setValue(url);

      // Logger.log(`Saved Url: ${url}`);
      // Logger.log("Fetch Url saved. Length: " + url.length);

    }
  }

 // 🔑 Always restore B10 from Script Properties
  if (url) {
    const b10 = sheet.getRange("B10");
    if (b10.getValue() !== model) {
      b10.setValue(model);

      // Logger.log(`Saved Model: ${model}`);
      // Logger.log("AI Model saved. Length: " + model.length);

    }
  }




  const status = [
    ["📊 STATUS", ""],
    ["Authorization:", "✅ Done (if you ran 'Re-initialize Setup')"],
    ["Triggers:", (hasEdit1 && hasEdit2 && hasOpen) ? "✅ Installed" : "❌ Not Installed"],
    ["API Key:", key ? "🔑 Stored" : "❌ Missing"]
  ];

  const range = sheet.getRange("A15:B18");
  range.clearContent();
  range.setValues(status);

  // Style
  range.setFontWeight("normal");
  range.setFontSize(11);


    // Header row
  const header = sheet.getRange("A15:B15");
  if (header.isPartOfMerge()) {
    header.breakApart();
  }
  header.merge();
  header.setFontWeight("bold")
        .setFontSize(12)
        .setBackground("#f4f4f4")
        .setHorizontalAlignment("center");


  // Labels column
  const labels = sheet.getRange("A16:A18");
  labels.setFontWeight("bold");

  // Values column
  const values = sheet.getRange("A16:B18");
  values.setHorizontalAlignment("left");
}





// =======================
// 8. WELCOME / IN-SHEET INSTRUCTIONS
// =======================
function insertWelcomeNote() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");

  // Merge A1:B5 for the welcome note
  const range = sheet.getRange("A1:B5");
  range.merge();
  range.setValue(
    "📌 Welcome! Mobile/Desktop Instructions: Data&Query 1.1 \n\n" +
    "1️⃣ Enter your API Key in B7 → will appear in B6.\n" +
    "2️⃣ Enter your FETCH URL in B9 → will appear in B8.\n" +
    "3️⃣ Enter your Selected AI Model in B11 → will appear in B10.\n" +
    "4️⃣ Type RESET in B13 → clears the API Key.\n" +
    "5️⃣ Check the box in B14 → runs setup (Re-initialize).\n" +
    "6️⃣ Status updates appear in B16:B18. \n"
  );

  range.setFontWeight("bold");
  range.setFontSize(10);
  range.setWrap(true);
}





function initializeDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");

  dash.getRange("A3").setValue("Enter data here...").setBackground(null);
  dash.getRange("A11").setValue("Enter question here...").setBackground(null);

  dash.getRange("A8").clearContent().setBackground(null);
  dash.getRange("A16").clearContent().setBackground(null);

  var answerArea = dash.getRange("A19");
  answerArea.setValue("Answer to Question will appear here...");
  answerArea.setWrap(true);
  answerArea.setVerticalAlignment("top");
  answerArea.setFontColor("black");
  answerArea.setBackground(null);
}







// =======================
// 4. ON EDIT HANDLER
// =======================
function onEditHandler(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'Dashboard') return;

  var cell = e.range;
  var val = e.value;




  // --- Submit checkbox at F5 ---
  if (cell.getA1Notation() === "F5" && val === "TRUE") {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(0)) return;  

    try {
      // Highlight question area (A3) and flash checkbox block (F5)
      flashCell(sheet.getRange("F5"), "#a5d6a7");              // green flash
      highlightEntryAndFlash(sheet.getRange("A3"), "#c8e6c9"); // light green highlight

      // Delay reset so the user-trigger fully registers
      Utilities.sleep(200);  
      cell.setValue(false);
      sheet.getRange("F5").setBackground(null);

      // Continue with AI flow
      flashProcessing(sheet.getRange("A8"), function () {
        submitData();                                  // pass mode
      }, sheet.getRange("A3"));

    } finally {
      lock.releaseLock();
    }
  }






  if (cell.getA1Notation() === "F12" && val === "TRUE") {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(0)) return;  

    try {
      // Highlight question area (A11) and flash checkbox block (F12)
      flashCell(sheet.getRange("F12"), "#90caf9");              // blue flash
      highlightEntryAndFlash(sheet.getRange("A11"), "#90caf9"); // light blue highlight

      // Delay reset so the user-trigger fully registers
      Utilities.sleep(200);  
      cell.setValue(false);
      sheet.getRange("F12").setBackground(null);

      // Continue with AI flow
      flashProcessing(sheet.getRange("A16"), function () {
        askQuestion("Offline");                                  // pass mode
      }, sheet.getRange("A11"));

    } finally {
      lock.releaseLock();
    }
  }






  if (cell.getA1Notation() === "F15" && val === "TRUE") {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(0)) return;  

    try {
      // Highlight question area (A11) and flash checkbox block (F15)
      flashCell(sheet.getRange("F15"), "#00FA9A");              // flash
      highlightEntryAndFlash(sheet.getRange("A11"), "#90caf9"); // highlight

      // Delay reset so the user-trigger fully registers
      Utilities.sleep(200);  
      cell.setValue(false);
      sheet.getRange("F15").setBackground(null);

      // Continue with AI flow
      flashProcessing(sheet.getRange("A16"), function () {
        askQuestion("AI-assisted");
      }, sheet.getRange("A11"));

    } finally {
      lock.releaseLock();
    }
  }






      // --- REFRESH UI at D1 checkbox ---
  if (cell.getA1Notation() === "D1" && val === "TRUE") {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(0)) return;  

    try {
      flashCell(sheet.getRange("A1:F19"), "#00FFFF");              // blue flash
      Utilities.sleep(200);
      cell.setValue(false);
      sheet.getRange("A1:F19").setBackground(null);
      initializeDashboard();  
      showAnswerSidebar("","");

    } finally {
      lock.releaseLock();
    }
  }

}








function onEditSettings(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Settings") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // --- B7: save new API key
  if (row === 7 && col === 2) {

    const newKey = String(e.range.getDisplayValue()).trim().replace(/[^\x20-\x7E]/g, '').normalize("NFKC");
    if (newKey) {
      PropertiesService.getScriptProperties().setProperty("API_KEY", newKey);
      const checkKey = PropertiesService.getScriptProperties().getProperty("API_KEY");
      Logger.log("Saved key:", checkKey);

      const b6 = sheet.getRange("B6");
    
      // Write new key into B6
      b6.setValue(newKey);

      // Clear B7 input
      e.range.setValue(""); // clear B7 after saving
    }
  }



  // --- B9: save new Fetch URL
  if (row === 9 && col === 2) {

      const newUrl = String(e.range.getDisplayValue()).trim().replace(/[^\x20-\x7E]/g, '').normalize("NFKC");
      if (newUrl) {
        PropertiesService.getScriptProperties().setProperty("FETCH_URL", newUrl);
        const checkUrl = PropertiesService.getScriptProperties().getProperty("FETCH_URL");
        Logger.log("Saved Url:", checkUrl);

        const b8 = sheet.getRange("B8");

      // Write new key into B8
      b8.setValue(newUrl);

      // Clear B9 input
      e.range.setValue(""); // clear B9 after saving

    }
  }


  // --- B11: save new AI Model
  if (row === 11 && col === 2) {

      const newModel = String(e.range.getDisplayValue()).trim().replace(/[^\x20-\x7E]/g, '').normalize("NFKC");
      if (newModel) {
        PropertiesService.getScriptProperties().setProperty("AI_MODEL", newModel);
        const checkModel = PropertiesService.getScriptProperties().getProperty("AI_MODEL");
        Logger.log("Saved Model:", checkModel);

        const b10 = sheet.getRange("B10");

      // Write new key into B10
      b10.setValue(newModel);

      // Clear B11 input
      e.range.setValue(""); // clear B11 after saving

    }
  }




  // --- B13: RESET API key
  if (row === 13 && col === 2 && e.range.getValue() === "RESET") {
      PropertiesService.getScriptProperties().setProperty("API_KEY", "");
      sheet.getRange("B6").setValue("No API Key Stored");
      e.range.setValue("");

      PropertiesService.getScriptProperties().setProperty("FETCH_URL", "");
      sheet.getRange("B8").setValue("No Fetch URL Stored");
      e.range.setValue("");

      PropertiesService.getScriptProperties().setProperty("AI_MODEL", "");
      sheet.getRange("B10").setValue("No AI Model Stored");
      e.range.setValue("");


  }

  // --- B14: Checkbox for Re-initialize Setup (mobile-friendly)
  if ((row === 14) && col === 2 && e.range.isChecked()) {
    firstTimeUserSetup();
    e.range.setValue(false); // reset checkbox after running
  }

  updateStatusPanel(); // now updates B16:B18
  
}






function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== "Dashboard") return;

  if (range.getA1Notation() === "A3") {
    if (range.getValue() === "Enter data here...") range.setValue("");
    sheet.getRange("A8").clearContent().setBackground(null);
  }

  if (range.getA1Notation() === "A11") {
    if (range.getValue() === "Enter question here...") range.setValue("");
    sheet.getRange("A16").clearContent().setBackground(null);
  }
}





function highlightEntryAndFlash(entryRange, color) {
  entryRange.setBackground(color);
}

function flashCell(range, color) {
  var original = range.getBackgrounds();
  for (var i = 0; i < 2; i++) {
    range.setBackground(color);
    SpreadsheetApp.flush();
    Utilities.sleep(50);
    range.setBackgrounds(original);
    SpreadsheetApp.flush();
    Utilities.sleep(50);
  }
}

function flashProcessing(statusCell, callback, entryRange) {
  var originalBg = statusCell.getBackground();
  statusCell.setBackground("#fff176").setValue("Processing...");
  SpreadsheetApp.flush();

  // Flash twice while keeping text
  for (var i = 0; i < 2; i++) {
    Utilities.sleep(50);
    statusCell.setBackground(originalBg);
    SpreadsheetApp.flush();
    Utilities.sleep(50);
    statusCell.setBackground("#fff176");
    SpreadsheetApp.flush();
  }

  // Run actual function
  callback();

  // Gradual fade for status
  fadeBackground(statusCell, "#fff176", originalBg, 3);

  // Gradual fade for entry highlight, then clear
  if (entryRange) {
    fadeBackground(entryRange, entryRange.getBackground(), "#ffffff", 3);
    entryRange.setBackground(null);
  }
}

function fadeBackground(cell, startColor, endColor, steps) {
  var startRGB = hexToRgb(startColor);
  var endRGB = hexToRgb(endColor);
  for (var i = 1; i <= steps; i++) {
    var r = Math.round(startRGB.r + (endRGB.r - startRGB.r) * (i / steps));
    var g = Math.round(startRGB.g + (endRGB.g - startRGB.g) * (i / steps));
    var b = Math.round(startRGB.b + (endRGB.b - startRGB.b) * (i / steps));
    cell.setBackground(rgbToHex(r, g, b));
    SpreadsheetApp.flush();
    Utilities.sleep(50);
  }
}

function hexToRgb(hex) {
  hex = hex.replace("#", "");
  var bigint = parseInt(hex, 16);
  return { r: (bigint >> 16) & 255, g: (bigint >> 8) & 255, b: bigint & 255 };
}

function rgbToHex(r, g, b) {
  return "#" + [r, g, b].map(x => {
    var hex = x.toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  }).join('');
}






function askQuestion(mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");

  var statusCell = dash.getRange("A16");
  statusCell.clearContent();

  var question = (dash.getRange("A11").getValue() + "").trim();

  if (!question || question === "Enter question here...") {
    statusCell.setValue("⚠ Please type something before asking.").setBackground(null);
    return;
  }

  // --- bridge to getAnswers ---
  if (mode === "AI-assisted" || mode === "Offline") {

    getAnswers(mode,question);

  } else {
    statusCell.setValue("⚠ Unknown mode.").setBackground(null);
  }

}






function getAnswers(mode, question) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");

  if (mode === "Offline") {

    getOfflineAnswers(question);  // removed return
  } else if (mode === "AI-assisted") {

    getAiAssistedAnswers(question);
  } else {
    dash.getRange("A19").setValue("⚠ Unknown mode: " + mode);
  }

}







function getOfflineAnswers(question) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");
  var source = ss.getSheetByName("Source Table");

  var entryRange = dash.getRange("A11:E15");
  var statusCell = dash.getRange("A16");
  var answerArea = dash.getRange("A19");

    // === AUTO-CLEAR PREVIOUS ENTRIES ===
  entryRange.setBackground(null);              // Only clear background, keep text
  statusCell.clearContent().setBackground(null); // Status line
  answerArea.clearContent().setBackground(null); // Answer area

    // Show temporary AI processing message
  answerArea.setValue("📊 Offline query response for: " + question)
            .setWrap(true)
            .setBackground(null);
  statusCell.setValue("⏳ Processing Offline query...").setBackground("#fff176");



  // helper: escape regex special chars
  function escapeRegExp(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }




  // --- Extract top 10 keywords (excluding stopwords) ---
  var cleanQ = question.replace(/[^\w\s]/g, "");
  var words = cleanQ.split(/\s+/).filter(w => w.length > 0);

  // Load stopwords from StopWords sheet into a Set
  var stopSheet = ss.getSheetByName("StopWords");
  var stopSet = new Set();
  if (stopSheet) {
    var stopVals = stopSheet.getRange(1, 1, stopSheet.getLastRow(), 1).getValues();
    stopVals.forEach(function(row) {
      if (row[0]) stopSet.add((row[0] + "").toLowerCase());
    });
  }

  // Filter out stopwords
  var filtered = words.filter(function(w) {
    return !stopSet.has(w.toLowerCase());
  });


  // --- Load irregular verbs into a map (base -> [variants]) ---
  var irregSheet = ss.getSheetByName("IrregularVerbs");
  var irregMap = {};
  if (irregSheet) {
    var irregData = irregSheet.getDataRange().getValues();
    for (var r = 0; r < irregData.length; r++) {     // use r to avoid confusion with i
      var base = (irregData[r][0] + "").toLowerCase().trim();
      if (!base) continue;
      var variants = [];
      for (var c = 1; c < irregData[r].length; c++) {
        if (!irregData[r][c]) continue;
        // split in case of multiple variants in one cell e.g. "was/were"
        var parts = (irregData[r][c] + "").toLowerCase().split(/[\/\s]+/);
        parts.forEach(p => { if (p) variants.push(p); });
        }
      // ensure the base itself is included (helps lookup when question has base)
      if (!variants.includes(base)) variants.unshift(base);
      irregMap[base] = variants;
    }
  }

  // --- Build reverse map for irregular verbs (variant -> base) ---
  var irregularReverse = {};
  for (var base in irregMap) {
    (irregMap[base] || []).forEach(v => {
      irregularReverse[v.toLowerCase()] = base;
    });
  }




  // Sort by length (longer words first), fallback = order of appearance
  filtered.sort(function(a, b) {
    return b.length - a.length || words.indexOf(a) - words.indexOf(b);
  });


  // Get keyword limit from Settings!B22 (default 10)
  var rawKwLimit = settingsSheet ? settingsSheet.getRange("B22").getValue() : null;
  var keywordLimit = parseInt(rawKwLimit, 10);
  if (isNaN(keywordLimit) || keywordLimit <= 0) keywordLimit = 10;

  // Keep top keywords up to keywordLimit
  var keywords = filtered.slice(0, keywordLimit);


  var data = source.getDataRange().getValues();
  var results = [];

  for (var i = 0; i < data.length; i++) {
    var rowText = (data[i][1] + "").toLowerCase();
    var matchedSet = new Set();
    var debugHits = []; // collect debug info for this row

    // Check each top keyword
    for (var k = 0; k < keywords.length; k++) {
      var kw = (keywords[k] + "").toLowerCase();
      

      // --- Build family of keyword variants (modular, consistent with irregMap) ---
      var family = new Set();
      family.add(kw); // always include the keyword itself

      var baseVerb = kw.toLowerCase();

      // Build validBaseWords: tokens from Source Table entries + irregular base verbs
      var validBaseWords = new Set();

      // tokenize Source Table column B entries (remove punctuation)
      for (var r = 0; r < data.length; r++) {
        var text = (data[r][1] || "").toString().toLowerCase();
        text = text.replace(/[^\w\s]/g, ""); // remove punctuation
        if (!text) continue;
        var toks = text.split(/\s+/);
        toks.forEach(function(w) {
          if (w) validBaseWords.add(w);
        });
      }

      // include all irregular base verbs (from irregMap)
      for (var b in irregMap) {
        if (b) validBaseWords.add(b);
      }


      // 1) If kw is exactly a base irregular (e.g., "eat")
      if (irregMap[baseVerb]) {
      (irregMap[baseVerb] || []).forEach(f => family.add(f));

      // 2) If kw is a known irregular variant (e.g., "eaten", "ate")
      } else if (irregularReverse[baseVerb]) {
      var base = irregularReverse[baseVerb];
      family.add(base);               // include base
      (irregMap[base] || []).forEach(f => family.add(f)); // include all variants

      // 3) Otherwise treat as regular verb and generate plausible variants
      } else {
        // regular verb -> generate family using dictionary of valid bases
        family = generateRegularVerbFamily(kw, validBaseWords);
      }

      // --- Test each variant in the family ---
      // add main keyword once to matchedSet if any variant matches; log which variant(s) hit
      var familyMatched = false;
      var matchedTerms = [];

      // 🔹 Collect debug info for all rows
      var ENABLE_DEBUG = false;  // toggle here
      var debugRows = [];


      // iterate deterministically using Array.from() (avoid Set iterator surprises)
      Array.from(family).forEach(function(term) {
      var re = new RegExp("\\b" + escapeRegExp(term) + "\\b", "i");
      var hit = re.test(rowText); // safe: no 'g' flag
      
      // record attempt for debugging
      debugHits.push("TRY:" + term + " => " + (hit ? "MATCH" : "no")); // Let ENABLE_DEBUG = TRUE to activate function getOrCreateDebugSheet_ at the bottom of the main function to write info into tab "DebugMatches"
        if (hit) {
          familyMatched = true;
          matchedTerms.push(term);
        }
      });

      if (familyMatched) {
        matchedSet.add(kw); // add base keyword once
        debugHits.push("MATCHED_FAMILY:" + kw + " via: " + matchedTerms.join(", ")); // Let ENABLE_DEBUG = TRUE to activate function getOrCreateDebugSheet_ at the bottom of the main function to write info into tab "DebugMatches"
      } else {
      // helpful: record the whole family when nothing matched
        debugHits.push("NO_MATCH_FAMILY:" + kw + " family=[" + Array.from(family).join(", ") + "]"); // Let ENABLE_DEBUG = TRUE to activate function getOrCreateDebugSheet_ at the bottom of the main function to write info into tab "DebugMatches"
      }

    }    
      

    
    var matchCount = matchedSet.size;
    var matchedFamilies = Array.from(matchedSet);

    // --- ALWAYS record debug for troubleshooting (temporary) ---
    debugRows.push([
    data[i][0],
    i + 1,
    data[i][1],
    matchCount,
    matchedFamilies.join(", "),
    debugHits.join(" | ")
    ]);

    if (matchCount > 0) {
      results.push({
      timestamp: data[i][0],
      rowNum: i + 1,
      entry: data[i][1],
      matches: matchCount,
      matched: matchedFamilies
      });
    }
  }


  // Sort results: first by matches desc, then by timestamp desc
  results.sort(function(a, b) {
    if (b.matches !== a.matches) return b.matches - a.matches;
    return new Date(b.timestamp) - new Date(a.timestamp);
  });

  // Get max results from Settings!B21 (default 5)
  var settingsSheet = ss.getSheetByName("Settings");
  var rawMax = settingsSheet ? settingsSheet.getRange("B21").getValue() : null;
  var maxResults = parseInt(rawMax, 10);
  if (isNaN(maxResults) || maxResults <= 0) maxResults = 5;

  // Build display text using top N results
  var displayResults = results.slice(0, maxResults);
  var displayText = "";
  var segments = [];

  for (var i = 0; i < displayResults.length; i++) {
    var r = displayResults[i];
    var ts = r.timestamp;
    var mm = ts.getMonth() + 1, dd = ts.getDate(), yy = ts.getFullYear().toString().slice(-2);
    var dateStr = (mm < 10 ? "0"+mm:mm) + (dd<10?"0"+dd:dd) + yy;

    var numText = (i + 1) + ".";
    var line = numText + " " + r.entry + ".  " + dateStr + "-R" + r.rowNum;

    segments.push({start: displayText.length, end: displayText.length + numText.length});
    displayText += line + (i < displayResults.length - 1 ? "\n" : "");
  }

  var builder = SpreadsheetApp.newRichTextValue().setText(displayText);
  segments.forEach(s => {
    builder.setTextStyle(s.start, s.end, SpreadsheetApp.newTextStyle().setBold(true).build());
  });
  var richText = builder.build();
  answerArea.setRichTextValue(richText);
  answerArea.setWrap(true);

  // Auto row height
  var lineCount = displayText ? displayText.split("\n").length : 1;
  dash.setRowHeight(19, Math.max(25, lineCount * 25));

  // Write full results to ResultsOFFline sheet
  function getOrCreateResultsSheet_() {
    var sh = ss.getSheetByName("ResultsOFFline");
    if (!sh) sh = ss.insertSheet("ResultsOFFline");
    return sh;
  }
  var resultsSheet = getOrCreateResultsSheet_();
  resultsSheet.clearContents();
  resultsSheet.getRange(1, 1, 1, 7).setValues([[
    "Date", "Row", "Answer: Source Table Entry Match",
    "Question", "Matches", "Matched Keywords", "Keywords"
  ]]);

  if (results.length > 0) {
    var allRows = results.map(function(res) {
      var ts = res.timestamp;
      var mm = ts.getMonth() + 1, dd = ts.getDate(), yy = ts.getFullYear().toString().slice(-2);
      var dateStr = (mm < 10 ? "0"+mm:mm) + (dd<10?"0"+dd:dd) + yy;
      return [
        dateStr,
        res.rowNum,
        res.entry,
        question,
        res.matches,
        (res.matched || []).join(", "),
        keywords.join(", ")
      ];
    });
    resultsSheet.getRange(2, 1, allRows.length, 7).setValues(allRows);
  }

  // Wrap-up UI
  entryRange.setBackground(null);
  statusCell.setValue("✅ Question recorded successfully!").setBackground("#c8e6c9");
  answerArea.setBackground("#c8e6c9");
  entryRange.setBackground(null);

  // Show sidebar preview
  showAnswerSidebar(displayResults, "Offline");
  addReturnLink("ResultsOFFline");


  // Debug logging sheet 
  if (ENABLE_DEBUG) {
    function getOrCreateDebugSheet_() {
      var sh = ss.getSheetByName("DebugMatches");
      if (!sh) sh = ss.insertSheet("DebugMatches");
      return sh;
    }

    var debugSheet = getOrCreateDebugSheet_();
    debugSheet.clearContents();
    debugSheet.getRange(1, 1, 1, 6).setValues([[
  "Date", "Row", "Entry", "Matches", "Matched Families", "Debug Hits"
    ]]);
    if (debugRows.length > 0) {
      debugSheet.getRange(2, 1, debugRows.length, 6).setValues(debugRows);
    }
  }
}






/**
 * Generate plausible variants for a regular verb keyword
 * @param {string} kw - the keyword
 * @param {Set} validBaseWords - a set of known valid base verbs (from Source Table + irregulars)
 * @returns {Set} family - set of keyword variants
 */
function generateRegularVerbFamily(kw, validBaseWords) {
  kw = (kw || "").toLowerCase();
  var family = new Set();
  family.add(kw);

  var root = kw;

  // Recover candidate root from suffixes using dictionary checks
  if (kw.endsWith("ed")) {
    var candidate = kw.slice(0, -2); // strip ed
    if (validBaseWords.has(candidate)) root = candidate;
    else if (validBaseWords.has(candidate + "e")) root = candidate + "e";
    else root = candidate; // fallback

    // --- double consonant handling for -ed ---
    var doubleMatch = root.match(/([b-df-hj-np-tv-z])([aeiou])\1$/);
    if (doubleMatch) {
      var ccRoot = root.slice(0, -1); // drop one consonant
      if (validBaseWords.has(ccRoot)) root = ccRoot;
    }

  } else if (kw.endsWith("ing")) {
    var candidate = kw.slice(0, -3); // strip ing
    if (validBaseWords.has(candidate)) root = candidate;
    else if (validBaseWords.has(candidate + "e")) root = candidate + "e";
    else root = candidate;

    // --- double consonant handling for -ing ---
    var doubleMatch = root.match(/([b-df-hj-np-tv-z])([aeiou])\1$/);
    if (doubleMatch) {
      var ccRoot = root.slice(0, -1); // drop one consonant
      if (validBaseWords.has(ccRoot)) root = ccRoot;
    }

  } else if (kw.endsWith("es")) {
    var candidate = kw.slice(0, -2); // strip es
    if (validBaseWords.has(candidate)) root = candidate;
    else if (validBaseWords.has(candidate + "e")) root = candidate + "e";
    else root = candidate;
  } else {
    // leave root = kw (base form)
    root = kw;
  }

  if (root && root !== kw) family.add(root);

  if (root) {
    // 1) plural / third-person: apply -es rule for certain endings
    var last1 = root.slice(-1);

    if (/(s|sh|ch|x|z|o)$/.test(root)) {
      // ends with s/sh/ch/x/z/o -> use -es
      family.add(root + "es");
    } else if (root.length > 1 && last1 === "y" && !/[aeiou]/.test(root.charAt(root.length - 2))) {
      // consonant + y -> replace y with ies (study -> studies)
      family.add(root.slice(0, -1) + "ies");
    } else {
      // normal plural
      family.add(root + "s");
    }

    // 2) past (-ed) and gerund (-ing)
    if (root.endsWith("e")) {
      // keep 'e' for past; drop it for -ing
      family.add(root + "d");                    // schedule -> scheduled
      family.add(root.slice(0, -1) + "ing");     // schedule -> scheduling
    } else {
      family.add(root + "ed");
      family.add(root + "ing");

      // --- double consonant variant forms for past/gerund ---
      var cvcMatch = root.match(/([b-df-hj-np-tv-z])([aeiou])([b-df-hj-np-tv-z])$/);
      if (cvcMatch) {
        var lastConsonant = cvcMatch[3];
        family.add(root + lastConsonant + "ed");   // wrap -> wrapped
        family.add(root + lastConsonant + "ing");  // wrap -> wrapping
      }
    }
  }

  return family;
}









// === MAIN AI-ASSISTED ANSWER FUNCTION ===
function buildAiContext(question) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName("Source Table");
  var data = source.getDataRange().getValues(); // assuming column 0 = timestamp, column 1 = entry

  // --- Step 1: extract top 5 keywords ---
  var words = question.split(/\s+/).filter(w => w.length > 1);
  words.sort((a, b) => b.length - a.length || words.indexOf(a) - words.indexOf(b));
  var keywords = words.slice(0, 5);

  // --- Step 2: filter relevant rows ---
  var relevantRows = [];
  for (var i = 0; i < data.length; i++) {
    var rowText = (data[i][1] + "").toLowerCase();
    var matchCount = keywords.reduce((acc, kw) => rowText.includes(kw.toLowerCase()) ? acc + 1 : acc, 0);
    if (matchCount > 0) {
      relevantRows.push({
        timestamp: new Date(data[i][0]), // keep timestamp for future use
        entry: data[i][1],
        rowNumber: i + 1 // row number in Source Table
      });
    }
  }

  // --- Step 3: sort by timestamp descending ---
  relevantRows.sort((a, b) => b.timestamp - a.timestamp);

  // --- Step 4: build AI context string ---
  var context = relevantRows.map(r => {
    var ts = r.timestamp;
    var mm = ts.getMonth() + 1, dd = ts.getDate(), yy = ts.getFullYear().toString().slice(-2);
    var dateStr = (mm < 10 ? "0" + mm : mm) + (dd < 10 ? "0" + dd : dd) + yy;
    return dateStr + " (R" + r.rowNumber + "): " + r.entry;
  }).join("\n");

  // return both context string and full relevantRows for future use
  return { context: context, contextRows: relevantRows };
}







function getAiAssistedAnswers(question) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");
  var answerArea = dash.getRange("A19");
  var statusCell = dash.getRange("A16");
  var questionArea = dash.getRange("A11:E15"); 
  
    // === AUTO-CLEAR PREVIOUS ENTRIES ===
  questionArea.setBackground(null);              // Only clear background, keep text
  statusCell.clearContent().setBackground(null); // Status line
  answerArea.clearContent().setBackground(null); // Answer area
  
  
  // Show temporary AI processing message
  answerArea.setValue("🤖 AI response for: " + question)
            .setWrap(true)
            .setBackground(null);
  statusCell.setValue("⏳ Processing AI-assisted query...").setBackground("#fff176");

  // Ensure ResultsAI sheet exists
  var resultsAI = (function() {
    var sh = ss.getSheetByName("ResultsAI");
    if (!sh) {
      sh = ss.insertSheet("ResultsAI");
      sh.appendRow(["Timestamp", "Question", "Answer"]); // header row
    }
    return sh;
  })();

  // Retrieve and Check API Key from Script Properties
  const apiKey = getApiKey();

  // Retrieve and Check AI Model from Script Properties
  const AIModel = getAIModel();

  // Retrieve and Check Fetch Url from Script Properties
  const fetchUrl = getFetchUrl();




  try {
    // --- Get AI context from Source Table ---
    var contextData = buildAiContext(question); // { context, contextRows }
    var contextText = contextData.context;
    
   
    if (!apiKey) throw new Error("AI API key not found.");
   

   // Instruct AI to give a concise answer 50-100 words, shorter if possible
    var prompt = "Use the following Source Table data for context, then answer the user question concisely (50-100 words, shorter is fine). Only    answer if context is clear:\n\n"
             + "Source Table Data:\n" + contextText + "\n\nUser question:\n" + question;

   
    // Prepare request payload for AI
    var payload = {
      model: AIModel,
      messages: [{ role: "user", content: prompt }],
      temperature: 0.7,
      max_tokens: 300
    };

    var options = {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(payload)
    };


    var aiText = "";

    // GOOGLE AI/GEMINI CONNECTION
    if (fetchUrl.includes("google")) {
    
      // Condition is met → call another function
      var json = callGeminiDynamic(prompt,apiKey,fetchUrl,AIModel);
      aiText = json.candidates?.[0]?.content?.parts?.[0]?.text || '';
      // SpreadsheetApp.getUi().alert("line 1186" + " " + fetchUrl + " " +aiText);
    } 




    // REPLICATE CONNECTION

  if (fetchUrl.includes("replicate")) {
    try {
    // 🔹 One clean call — returns a stable string
      aiText = runReplicateModel(prompt, apiKey, fetchUrl, AIModel, 20) || "";

    // Debug/logging (optional)
      Logger.log("Replicate output: " + aiText);
    // SpreadsheetApp.getUi().alert("line 1201: " + fetchUrl + " " + aiText);

    // Write to sheet or UI
      answerArea.setValue(aiText).setBackground("#e8f5e9"); // success = light green
    } catch (err) {
      Logger.log("Replicate error: " + err);
      answerArea
        .setValue(`Error: ${err.message}`)
        .setBackground("#ffcdd2"); // error = light red
    // SpreadsheetApp.getUi().alert("line 1210: " + aiText);
      return;
    }
    // SpreadsheetApp.getUi().alert("line 1213" + " " + aiText);
  }

    // SpreadsheetApp.getUi().alert("line 1216" + " " + aiText);


    // OPENAI-FORMAT CONNECTION
    if (fetchUrl.includes("v1/chat/completions")) {
      var response = UrlFetchApp.fetch(fetchUrl, options);
      var json = JSON.parse(response.getContentText());
      aiText = String(json?.choices?.[0]?.message?.content || "").trim();
      var apiWarning = "";

    // SpreadsheetApp.getUi().alert("line 1226" + " " + fetchUrl + " " + aiText); 
      if (!aiText && apiWarning) {
    // SpreadsheetApp.getUi().alert("line 1228" + " " + fetchUrl + " " + aiText);
      // Only show error if there is no valid AI text
      throw new Error("API error: " + apiWarning);
      }

    // --- Capture any API error without throwing if we still have content ---
    // SpreadsheetApp.getUi().alert("line 1234" + " " + fetchUrl + " " + aiText);    

      if (json.error) {
      //  SpreadsheetApp.getUi().alert("line 1237" + " " + fetchUrl + " " + aiText);
        apiWarning = json.error.message || JSON.stringify(json.error);
        Logger.log("API warning: " + apiWarning);
      // SpreadsheetApp.getUi().alert("line 1240" + " " + fetchUrl + " " + aiText);
      }
    } 


    // SpreadsheetApp.getUi().alert("line 1245" + " " + aiText);
    // --- Update ResultsAI tab ---
    updateResultsAI(question, aiText); // keeps max 20 latest

    // SpreadsheetApp.getUi().alert("line 1249" + " " + aiText);
    statusCell.setValue("✅ AI-assisted answer generated successfully!").setBackground("#c8e6c9");

    // SpreadsheetApp.getUi().alert("line 1252" + " " + aiText);


 
    try {
    // 🔹 Ensure aiText is a string
    aiText = typeof aiText === "string" ? aiText : "";

    // --- Estimate number of wrapped lines ---
    var avgCharsPerLine = 80; // adjust based on merged cell width + font size
    var lineCount = Math.ceil(aiText.length / avgCharsPerLine);
    // SpreadsheetApp.getUi().alert("line 1263: " + aiText);

    // --- Each line ~25px, set row height safely ---
    if (dash && dash.getRowHeight) {
      var newHeight = Math.max(25, lineCount * 25);
      try { dash.setRowHeight(19, newHeight); }
      catch(err) { Logger.log("Warning: cannot set row height: " + err); }
    } else {
      Logger.log("Warning: 'dash' range not found or invalid.");
    }
    // SpreadsheetApp.getUi().alert("line 1273: " + aiText);

    // --- Display in dashboard ---
    if (answerArea && answerArea.setValue) {
    try {
      answerArea.setValue(aiText).setWrap(true).setBackground("#c8e6c9");
      } catch(err) { Logger.log("Warning: cannot set answerArea value: " + err); }
    } else {
      Logger.log("Warning: 'answerArea' range not found or invalid.");
    }
    // SpreadsheetApp.getUi().alert("line 1283: " + aiText);

    // --- Show to sidebar ---
    try {
      var aiResults = [{ entry: aiText }]; // for sidebar, no timestamp needed
      // SpreadsheetApp.getUi().alert("line 1288: " + aiText);
      showAnswerSidebar(aiResults, "AI-assisted");
      } catch(err) {
        Logger.log("Warning: cannot show sidebar: " + err);
    }

    // --- Clear question input area background (A11) ---
    try { dash.getRange("A11").setBackground(null); } 
      catch(err) { Logger.log("Warning: cannot clear A11 background: " + err); }

    } catch(mainErr) {
      Logger.log("Unexpected error in post-1286 block: " + mainErr);
      // SpreadsheetApp.getUi().alert("Error in post-1286 block: " + mainErr.message);
    }

  // SpreadsheetApp.getUi().alert("line 1303: " + aiText);



  } catch (e) {


      var errorMsg = e.message || "";
      var detailedMsg = "";

      if (errorMsg.includes("fetch") || errorMsg.includes("network") || errorMsg.includes("Timeout")) {
          detailedMsg = "❌ Connection issue. Check your internet.";
      } 
      else if (errorMsg.includes("400")) {
          detailedMsg = "❌ 400 - Invalid Request/Format. Misconfigured request. Verify your key, url, and model settings.";
      } 
      else if (errorMsg.includes("401")) {
          detailedMsg = "❌ 401 - Authentication error/failure. Unauthorized, Missing or Invalid API Key. Verify your key, url, and model settings.";
      } 
      else if (errorMsg.includes("402")) {
          detailedMsg = "❌ 402 - Payment required/Insufficient Balance. The account associated with the API Key has reached its maximum allowed limit. Verify your credit limits, key, url, and model settings.";
      } 
      else if (errorMsg.includes("403")) {
          detailedMsg = "❌ 403 - Bad Request/Forbidden. Verify your credit limits, key, url, and model settings.";
      } 
      else if (errorMsg.includes("404")) {
          detailedMsg = "❌ 404 - Not Found. Invalid URL or Model name. Verify your url, and model settings.";
      }
      else if (errorMsg.includes("permission")) {
          detailedMsg = "❌ an API Key permission issue. Verify your key or provider settings.";
      }  
      else if (errorMsg.includes("422") || errorMsg.includes("parameter")) {
          detailedMsg = "❌ 422 - Invalid Parameters. Verify your model settings.";
      } 
      else if (errorMsg.includes("429") || errorMsg.includes("quota")) {
          detailedMsg = "❌ 429 - Quota exceeded/Rate Limit Reached. Too many requests sent in a short period of time. Check your plan or credits or Try again a few moments later.";
      } 
      else if (errorMsg.includes("500") || errorMsg.includes("server error")) {
          detailedMsg = "❌ 500 - Server error. Unknown server error. Please try again later.";
      } 
      else if (errorMsg.includes("503") || errorMsg.includes("unavailable")) {
          detailedMsg = "❌ 503 - Server overload/Service unavailable. Our servers are seeing high amounts of traffic. Please try again later.";
      } 
      else {
          detailedMsg = "❌ Unexpected error: " + errorMsg;
      }

  // SpreadsheetApp.getUi().alert("line 1350" + " " + aiText);

    // Only show error in dashboard if no valid AI text exists
    if (aiText === "") {
      statusCell.setValue("❌ Error generating AI-assisted answer - EMPTY").setBackground("#fff176");
      answerArea.setValue(detailedMsg).setWrap(true).setBackground("#fff176");
      dash.getRange("A11:E15").clearContent().setBackground(null);
      
      // Show error in SideBar
      var html = "<div style='font-family:Arial; font-size:14px;'>";
      html += "<p><i>" + detailedMsg + "</i></p>";  // Use concatenation with + here
      html += "</div>";

      var ui = HtmlService.createHtmlOutput(html)
        .setTitle("Answer Sidebar");
      SpreadsheetApp.getUi().showSidebar(ui);

    } else {
    // Log warning but keep the valid AI text displayed
      Logger.log("Warning ignored (valid aiText exists): " + detailedMsg);      statusCell.setValue("❌ Error generating AI-assisted answer - non-EMPTY").setBackground("#fff176");
      answerArea.setValue(detailedMsg).setWrap(true).setBackground("#fff176");
      dash.getRange("A11:E15").clearContent().setBackground(null);

    }
  }
    Logger.log("Warning ignored (valid aiText exists): " + detailedMsg);
}








// === RESULTS AI TAB HANDLER ===
function updateResultsAI(question, aiText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("ResultsAI");

  // Create sheet if missing, add header
  if (!sh) {
    sh = ss.insertSheet("ResultsAI");
    sh.appendRow(["Timestamp", "Question", "Answer"]);
  }

  // Insert new answer at row 2 (below header)
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 3).setValues([
    [new Date(), question, aiText]
  ]);

  // Keep only last 20 answers (delete row 22 if exists)
  if (sh.getLastRow() > 21) {
    sh.deleteRow(22);
  }
  addReturnLink("ResultsAI");
}







// === SIDEBAR DISPLAY HANDLER ===
function showAnswerSidebar(results, mode) {
  var html = "<div style='font-family:Arial; font-size:14px;'>";

  if (mode === "AI-assisted") {
    html += "<h3>🤖 AI-assisted Answer</h3>";
    if (results.length > 0) {
      html += "<p style='white-space:pre-wrap;'>" + results[0].entry + "</p>";
    } else {
      html += "<p><i>No AI-assisted answer available.</i></p>";
    }

  } else if (mode === "Offline") {
    html += "<h3>📊 Offline Match Results</h3>";
    if (results.length > 0) {
      results.forEach(function(r, i) {
        var ts = r.timestamp;
        var mm = ts.getMonth() + 1;
        var dd = ts.getDate();
        var yy = ts.getFullYear().toString().slice(-2);
        var dateStr = (mm < 10 ? "0"+mm:mm) + (dd<10?"0"+dd:dd) + yy;

        html += "<p><b>" + (i + 1) + ".</b> " + r.entry + ". "  + dateStr + "-R" + r.rowNum + "</p>";
      });
    } else {
      html += "<p><i>No offline matches found.</i></p>";
    }

  } else {
    html += "<p><i>Answer to Question will appear here.</i></p>";


  }

  html += "</div>";

  var ui = HtmlService.createHtmlOutput(html)
    .setTitle("Answer Sidebar");
  SpreadsheetApp.getUi().showSidebar(ui);
}














function submitData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Dashboard");
  var source = ss.getSheetByName("Source Table");

  var status = dash.getRange("A8");
  status.clearContent();

  var val = (dash.getRange("A3").getValue() + "").trim();
  if (!val || val === "Enter data here...") {
    status.setValue("⚠ Please type something before submitting.").setBackground(null);
    return;
  }

  var timestamp = new Date();
  var lastRow = source.getLastRow();
  var lastRowColA = source.getRange(lastRow, 1).getValue().toString().trim();
  

  var insertRow;
  if (lastRow > 0 && lastRowColA.toString().indexOf("Back to Dashboard") !== -1) {
  insertRow = lastRow - 1;
  
  } else {
    insertRow = lastRow + 1; // normal case
  }

  // --- Write new entry ---
  source.getRange(insertRow, 1).setValue(
    Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")
  );
  source.getRange(insertRow, 2).setValue(val);

  // Reset status & dashboard input
  status.setValue("✅ Data submitted successfully!").setBackground(null);
  dash.getRange("A3").setValue("Enter data here...");

  // Always re-add backlink at bottom
  addReturnLink("Source Table");
}








function addDashboardLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheets().find(s =>
    s.getName().toLowerCase().includes("dashboard")
  );
  if (!dashboard) return;

  var resultsSheet = ss.getSheetByName("ResultsOFFline");
  var resultsAiSheet = ss.getSheetByName("ResultsAI");
  var sourceTableSheet = ss.getSheetByName("Source Table");

  // Place the two links (adjust row numbers as desired)
  dashboard.getRange("B21").setFormula(
    `=HYPERLINK("#gid=${resultsSheet.getSheetId()}", "📊 OFFLINE Full Results")`
  );
  
  dashboard.getRange("B23").setFormula(
    `=HYPERLINK("#gid=${resultsAiSheet.getSheetId()}", "🤖📊 AI-ASSISTED Results Log")`
  );
  
  dashboard.getRange("B25").setFormula(
    `=HYPERLINK("#gid=${sourceTableSheet.getSheetId()}", "📋 SOURCE Table")`
  );
}
addDashboardLinks()






function addReturnLink(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var dashboardSheet = ss.getSheets().find(s =>
    s.getName().toLowerCase().includes("dashboard")
  );
  if (!dashboardSheet) return;

  // Remove any old "Back to Dashboard" link
  var lastRow = sheet.getLastRow();
  var backrows = (lastRow >= 12) ? 10 : lastRow - 2;
  var range = sheet.getRange(lastRow - backrows, 1, backrows + 2, 1); // look back up to 12 rows
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {

    if (values[i][0] && values[i][0].toString().includes("Back to Dashboard")) {
      range.getCell(i + 1, 1).clearContent();

    }
  }

  var lastRow = sheet.getLastRow();

  // Add link 1 row below last row of data
  var linkRow = sheet.getLastRow() + 2;
  sheet.getRange(linkRow, 1).setFormula(
    `=HYPERLINK("#gid=${dashboardSheet.getSheetId()}", "← Back to ${dashboardSheet.getName()}")`
  );

} 






