/**
 * Syncs page counts with Super Fuzzy logic and Error Logging
 */
function syncEbookPages() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName("Goodreads_Cleaned"); // Updated to match your new sheet name!
    const destSheet = ss.getSheetByName("eBook Page Counts for Book Tracker"); 
    const errorLog = ss.getSheetByName("Error Log");

    // =====================================================================
    // ⚙️ CONFIGURATION ZONE
    // =====================================================================
    const DEST_START_ROW = 67;   
    const DEST_TITLE_COL = 2;    // B = 2
    const DEST_PAGES_COL = 3;    // C = 3
    // =====================================================================

    if (!sourceSheet || !destSheet) return;

    const lastSourceRow = sourceSheet.getLastRow();
    if (lastSourceRow < 2) return;

    // Dynamically find columns
    const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    const titleIdx = headers.indexOf("Clean Title");
    const pageIdx = headers.indexOf("Number of Pages"); 

    const sourceData = sourceSheet.getRange(2, 1, lastSourceRow - 1, sourceSheet.getLastColumn()).getValues();

    // 1. Map existing titles
    const destLastRow = destSheet.getLastRow();
    let existingTitlesFuzzy = [];
    if (destLastRow >= DEST_START_ROW) {
        existingTitlesFuzzy = destSheet.getRange(DEST_START_ROW, DEST_TITLE_COL, Math.max(1, destLastRow - (DEST_START_ROW - 1)), 1)
            .getValues()
            .flat()
            .map(t => normalize(t));
    }

    let newEntries = [];
    let errors = [];

    // 2. Compare and find missing pages
    sourceData.forEach(row => {
        let title = titleIdx > -1 ? row[titleIdx] : "";
        let pages = pageIdx > -1 ? row[pageIdx] : "";

        if (!title || !pages) {
            if (title || pages) {
                errors.push([new Date(), "Page Sync", "Missing data", `Title: ${title || 'N/A'}, Pages: ${pages || 'N/A'}`]);
            }
        } else if (!existingTitlesFuzzy.includes(normalize(title))) {
            newEntries.push([title, pages]);
        }
    });

    // 3. Write to Destination
    if (newEntries.length > 0) {
        let startRow = findFirstEmptyRow(destSheet, "B", DEST_START_ROW);
        let targetRange = destSheet.getRange(startRow, DEST_TITLE_COL, newEntries.length, 2);

        targetRange.setValues(newEntries);

        // Format and Sort
        destSheet.getRange(Math.max(1, startRow - 1), DEST_TITLE_COL, 1, 2).copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        destSheet.getRange(DEST_START_ROW, DEST_TITLE_COL, destSheet.getLastRow() - (DEST_START_ROW - 1), 2).sort({column: DEST_TITLE_COL, ascending: true});
    }

    // 4. Log errors 
    if (errors.length > 0 && errorLog) {
        errorLog.getRange(errorLog.getLastRow() + 1, 1, errors.length, 4).setValues(errors);
    }

    // 🌟 THE FIX: This is exactly where the puller trigger goes!
    pullPagesToTracker();
}


/**
 * FETCH TOTAL PAGES (Book tracker 🧸 EDITION)
 * Scans your Reading list and fills in missing page counts!
 */
function pullPagesToTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName("Reading list");
  const pagesSheet = ss.getSheetByName("eBook Page Counts for Book Tracker"); 

  // =====================================================================
  // ⚙️ CONFIGURATION ZONE
  // =====================================================================
  const TRACKER_START_ROW = 4;
  const TRACKER_TITLE_COL = 3;    // C = 3
  const TRACKER_PAGES_COL = 13;   // M = 13 (TOTAL PAGE column)

  const PAGES_START_ROW = 67;     // Where your page database starts
  const PAGES_TITLE_COL = 2;      // B = 2
  const PAGES_COUNT_COL = 3;      // C = 3
  // =====================================================================

  if (!trackerSheet || !pagesSheet) return;

  // 1. Build the Page Dictionary
  const lastPageRow = pagesSheet.getLastRow();
  let pageMap = {};
  if (lastPageRow >= PAGES_START_ROW) {
    const pageData = pagesSheet.getRange(PAGES_START_ROW, PAGES_TITLE_COL, lastPageRow - PAGES_START_ROW + 1, 2).getValues();
    pageData.forEach(row => {
      if (row[0]) pageMap[superFuzzy(row[0])] = row[1];
    });
  }

  // 2. Scan the Reading list
  const lastTrackerRow = trackerSheet.getLastRow();
  if (lastTrackerRow < TRACKER_START_ROW) return;

  const trackerRange = trackerSheet.getRange(TRACKER_START_ROW, TRACKER_TITLE_COL, lastTrackerRow - TRACKER_START_ROW + 1, (TRACKER_PAGES_COL - TRACKER_TITLE_COL) + 1);
  const trackerData = trackerRange.getValues();

  let chunks = [];
  let currentChunk = null;
  let updateCount = 0;

  for (let i = 0; i < trackerData.length; i++) {
    let title = trackerData[i][0];
    let currentPages = trackerData[i][TRACKER_PAGES_COL - TRACKER_TITLE_COL];

    if (!title || (currentPages !== "" && currentPages !== null)) continue;

    let cleanTitle = superFuzzy(title);
    let foundPages = pageMap[cleanTitle];

    if (foundPages) {
      if (!currentChunk || currentChunk.endRow !== i + TRACKER_START_ROW - 1) {
        if (currentChunk) chunks.push(currentChunk);
        currentChunk = { startRow: i + TRACKER_START_ROW, endRow: i + TRACKER_START_ROW, values: [[foundPages]] };
      } else {
        currentChunk.endRow = i + TRACKER_START_ROW;
        currentChunk.values.push([foundPages]);
      }
      updateCount++;
    }
  }

  if (currentChunk) chunks.push(currentChunk);

  // 4. Paste the values instantly
  if (chunks.length > 0) {
    for (let c of chunks) {
      trackerSheet.getRange(c.startRow, TRACKER_PAGES_COL, c.values.length, 1).setValues(c.values);
    }
  }

  if (updateCount > 0) {
    ss.toast(`✅ Successfully fetched page counts for ${updateCount} books!`, "Page Sync");
  } else {
    ss.toast(`✅ Scan complete. No missing page counts found!`, "Page Sync");
  }
}

// --- HELPER FUNCTIONS ---

function normalize(str) {
    if (!str) return "";
    return str.toString().toLowerCase().replace(/^(the |a |an )/i, "").replace(/[^a-z0-9]/g, "").trim();
}

function superFuzzy(str) {
    if (!str) return "";
    return str.toString().toLowerCase().split(/[(\（]/)[0].replace(/&amp;/g, " ").replace(/&/g, " ").replace(/\bamp\b/g, " ").replace(/\band\b/g, " ").replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();
}

function findFirstEmptyRow(sheet, colLetter, startRow) {
    const vals = sheet.getRange(colLetter + startRow + ":" + colLetter).getValues();
    for (let i = 0; i < vals.length; i++) {
        if (vals[i][0] === "") return startRow + i;
    }
    return startRow + vals.length;
}