/**
 * MASTER GOODREADS SYNC (Book tracker 🧸 EDITION - V4)
 * 🎯 FIXED: Maps correctly to Wishlist Col G, fixes Gap Overwrite, and fixes Deduplication!
 */
function masterGoodreadsSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName("Goodreads_Cleaned"); 
  const trackerSheet = ss.getSheetByName("Reading list"); 
  const wishlistSheet = ss.getSheetByName("Wishlist"); 
  
  // =====================================================================
  // ⚙️ CONFIGURATION ZONE 
  // =====================================================================
  // TRACKER SETTINGS (Reading list)
  const TRACKER_START_ROW = 4;             
  const TRACKER_TITLE_COL = 3;      // C = 3
  const TRACKER_SERIES_COL = 17;    // Q = 17
  const TRACKER_AUTHOR_COL = 5;     // E = 5
  const TRACKER_STATUS_COL = 8;     // H = 8 
  const TRACKER_START_DATE = 9;     // I = 9
  const TRACKER_END_DATE = 10;      // J = 10
  
  // WISHLIST SETTINGS (Wishlist tab)
  const WISH_START_ROW = 13;        // Start at Row 13
  const WISH_TITLE_COL = 7;         // G = 7
  const WISH_AUTHOR_COL = 19;       // S = 19
  // =====================================================================

  const lastRow = importSheet.getLastRow();
  if (lastRow < 2) return ss.toast("No data found in Goodreads_Cleaned.", "Error");

  const headers = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues()[0];
  const titleIdx = headers.indexOf("Clean Title") > -1 ? headers.indexOf("Clean Title") : headers.indexOf("Title");
  const authorIdx = headers.indexOf("Author");
  const seriesIdx = headers.indexOf("Series"); 
  const coAuthorIdx = headers.indexOf("Additional Authors"); 
  const shelfIdx = headers.indexOf("Exclusive Shelf"); 
  const dateAddedIdx = headers.indexOf("Date Added"); 
  const dateReadIdx = headers.indexOf("Date Read");   

  if (titleIdx === -1) return ss.toast("Error: Title header not found!", "Error");

  // Smarter normalization so books in the same series aren't falsely skipped
  const normalize = (str) => {
    if (!str) return "";
    return str.toString().toLowerCase().replace(/[^a-z0-9]/g, ""); 
  };

  // Build maps of existing books so we NEVER duplicate them
  let trackerMap = {};
  const tLast = trackerSheet.getLastRow();
  if (tLast >= TRACKER_START_ROW) {
    trackerSheet.getRange(TRACKER_START_ROW, TRACKER_TITLE_COL, tLast - TRACKER_START_ROW + 1, 1).getValues().forEach((row, i) => {
      let norm = normalize(row[0]);
      if (norm) trackerMap[norm] = i + TRACKER_START_ROW;
    });
  }

  let wishlistMap = {};
  const wLast = wishlistSheet.getLastRow();
  if (wLast >= WISH_START_ROW) {
    wishlistSheet.getRange(WISH_START_ROW, WISH_TITLE_COL, wLast - WISH_START_ROW + 1, 1).getValues().forEach((row, i) => {
      let norm = normalize(row[0]);
      if (norm) wishlistMap[norm] = i + WISH_START_ROW;
    });
  }

  const filterData = importSheet.getRange(2, 1, lastRow - 1, importSheet.getLastColumn()).getValues();
  
  let trackerBatch = [];
  let wishlistBatch = [];
  let updateCount = 0;

  filterData.forEach(row => {
    let rawTitle = row[titleIdx] ? row[titleIdx].toString().trim() : "";
    if (!rawTitle) return;

    let searchTitle = normalize(rawTitle);
    let status = shelfIdx > -1 ? (row[shelfIdx] || "").toString().toLowerCase().trim() : "";
    let series = seriesIdx !== -1 ? row[seriesIdx] : ""; 
    let author = authorIdx !== -1 ? row[authorIdx] : "";
    let coAuthor = coAuthorIdx !== -1 ? row[coAuthorIdx] : "";
    let fullAuthor = coAuthor ? `${author} & ${coAuthor}` : author;

    // 1. Update Existing (Prevents duplicates across tabs)
    if (trackerMap[searchTitle]) {
      trackerSheet.getRange(trackerMap[searchTitle], TRACKER_TITLE_COL).setValue(rawTitle);
      trackerSheet.getRange(trackerMap[searchTitle], TRACKER_SERIES_COL).setValue(series);
      updateCount++;
    } 
    else if (wishlistMap[searchTitle]) {
      wishlistSheet.getRange(wishlistMap[searchTitle], WISH_TITLE_COL).setValue(rawTitle);
      updateCount++;
    }
    // 2. Brand New TO-READ (Sends safely to Wishlist Col G)
    else if (status.includes("to-read")) {
      wishlistBatch.push({ title: rawTitle, author: fullAuthor });
      wishlistMap[searchTitle] = true; 
    } 
    // 3. Brand New READ / DNF / CURRENTLY READING (Sends to Tracker)
    else if (status.includes("read") || status.includes("currently-reading") || status.includes("did-not-finish") || status.includes("dnf")) {
      let mappedStatus = "";
      if (status.includes("did-not-finish") || status.includes("dnf")) mappedStatus = "Did Not Finish \u274C"; 
      else if (status.includes("currently-reading")) mappedStatus = "Currently Reading";
      else if (status.includes("read")) mappedStatus = "Finished";

      trackerBatch.push({
        title: rawTitle,
        series: series,
        author: fullAuthor,
        status: mappedStatus,
        dateAdded: dateAddedIdx !== -1 ? row[dateAddedIdx] : "",
        dateRead: dateReadIdx !== -1 ? row[dateReadIdx] : ""
      });
      trackerMap[searchTitle] = true; 
    }
  });

  // True append pasting (No gap overwriting!)
  if (trackerBatch.length > 0) {
    let tRow = getTrueLastRow(trackerSheet, TRACKER_TITLE_COL, TRACKER_START_ROW);
    trackerSheet.getRange(tRow, TRACKER_TITLE_COL, trackerBatch.length, 1).setValues(trackerBatch.map(b => [b.title]));
    trackerSheet.getRange(tRow, TRACKER_SERIES_COL, trackerBatch.length, 1).setValues(trackerBatch.map(b => [b.series]));
    trackerSheet.getRange(tRow, TRACKER_AUTHOR_COL, trackerBatch.length, 1).setValues(trackerBatch.map(b => [b.author]));
    trackerSheet.getRange(tRow, TRACKER_STATUS_COL, trackerBatch.length, 1).setValues(trackerBatch.map(b => [b.status]));
    trackerSheet.getRange(tRow, TRACKER_START_DATE, trackerBatch.length, 1).setValues(trackerBatch.map(b => [b.dateAdded]));
    trackerSheet.getRange(tRow, TRACKER_END_DATE, trackerBatch.length, 1).setValues(trackerBatch.map(b => [b.dateRead]));
  }
  
  if (wishlistBatch.length > 0) {
    let wRow = getTrueLastRow(wishlistSheet, WISH_TITLE_COL, WISH_START_ROW); 
    wishlistSheet.getRange(wRow, WISH_TITLE_COL, wishlistBatch.length, 1).setValues(wishlistBatch.map(b => [b.title]));
    wishlistSheet.getRange(wRow, WISH_AUTHOR_COL, wishlistBatch.length, 1).setValues(wishlistBatch.map(b => [b.author]));
  }

  ss.toast(`✅ Cleaned ${updateCount} existing books! Added ${trackerBatch.length} to Tracker and ${wishlistBatch.length} to Wishlist.`, "Smart Sync");
}

// Helper function that completely prevents the Gap Overwrite bug
function getTrueLastRow(sheet, colNum, startRow) {
  const values = sheet.getRange(1, colNum, sheet.getLastRow() || 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return Math.max(startRow, i + 2); 
    }
  }
  return startRow;
}