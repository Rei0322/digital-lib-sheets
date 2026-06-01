/**
 * Global List of Utility Sheets
 */
const UTILITY_SHEETS = [
  "READ ME", "eBook Page Counts for Book Tracker", "Settings", "Roasts",
    "Goodreads_Import", "Ratings Calculator", "Goodreads_Cleaned", "Data", "Data Drop",
    "Sync Data"
];

/**
 * Custom Menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 Guide and Sync')
      .addItem('Daily Tools & Ratings', 'showRatingsGuide')  
      .addItem('Setup & Import Workflow', 'showSidebar') 
      .addSeparator()
      .addItem('Run Master Sync', 'masterSync')
      .addItem('Show All Utility Sheets', 'showAllUtilitySheets')
      .addItem('Hide All Utility Sheets', 'hideAllUtilitySheets')
      .addItem('Sync Book Covers', 'updateBookCoversFlexible')
      .addToUi();
}

/**
 * Opens the Sidebar
 */
function showSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
        .setTitle('Guide & Tools')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Independent Function: Unhide all utility sheets
 */
function showAllUtilitySheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    UTILITY_SHEETS.forEach(name => {
        let sheet = ss.getSheetByName(name);
        if (sheet) sheet.showSheet();
    });
    ss.toast("All utility sheets are now visible.", "Workspace Expanded");
}

/**
 * Independent Function: Hide all utility sheets
 */
function hideAllUtilitySheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    UTILITY_SHEETS.forEach(name => {
        let sheet = ss.getSheetByName(name);
        if (sheet) sheet.hideSheet();
    });
    ss.toast("All utility sheets are hidden.", "Workspace Cleaned");
}

function showSheetByName(name) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(name);
    if (sheet) {
        sheet.showSheet(); // Ensures it's unhidden if it was hidden
        sheet.activate();
    }
}

/**
 * Master Sync: Runs the Smart Traffic Cop, processes other data, and cleans up!
 */
function masterSync() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    try {
        ss.toast("🔄 Running Master Sync...", "Status", 2);

        // 1. RUN THE MASTER TRAFFIC COP (Replaces the 3 old transfer scripts!)
        if (typeof masterGoodreadsSync === "function") masterGoodreadsSync();

        // 2. ROUTE TO SIDE SHEETS
        if (typeof syncEbookPages === "function") syncEbookPages();

        // 3. SAFE CLEANUP: Wipe the import sheet so it's ready for next time
        const importSheet = ss.getSheetByName("Goodreads_Import");
        if (importSheet && importSheet.getLastRow() > 1) {
            importSheet.getRange(2, 1, importSheet.getLastRow(), importSheet.getLastColumn()).clearContent();
        }

        // 4. WRAP UP
        if (typeof hideAllUtilitySheets === "function") hideAllUtilitySheets();

        ss.toast("✅ Master Sync Complete!", "Success", 3);
    } catch (e) {
        console.error(e);
        ss.toast("❌ Master Sync failed: " + e.message, "Error", -1);
    }
}

/**
 * FLEXIBLE BOOK COVER SYNC
 */
function updateBookCoversFlexible() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Reading list');
    const importSheet = ss.getSheetByName('Goodreads_Import');

    const folderId = 'token';

    // =====================================================================
    // ⚙️ COVER CONFIGURATION ZONE
    // =====================================================================
    const START_ROW = 4;
    const TITLE_COL_NUM = 3;  // C = 3
    const COVER_COL_NUM = 4;  // D = 4 (Change this if Cover is in a different column on Reading list!)
    // =====================================================================

    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW) return { count: 0 };

    // 1. Map Goodreads IDs (ISBNs)
    const idMap = {};
    if (importSheet) {
        const importData = importSheet.getDataRange().getValues();
        importData.forEach(row => {
            let bookId = String(row[0]).trim();
            let title = String(row[1]).trim();
            if (bookId && title) idMap[superFuzzy(title)] = bookId; 
        });
    }

    // 2. Connect to Drive Folder
    let driveFolder;
    try {
        driveFolder = DriveApp.getFolderById(folderId);
    } catch(e) {
        console.log("Could not access folder ID: " + folderId);
    }

    let matchCount = 0;
    
    // We fetch a 2-column wide block starting at Title (Col C to Col D)
    const dataRange = sheet.getRange(START_ROW, TITLE_COL_NUM, lastRow - START_ROW + 1, (COVER_COL_NUM - TITLE_COL_NUM) + 1);
    const dataValues = dataRange.getValues(); 

    for (let i = 0; i < dataValues.length; i++) {
        const title = String(dataValues[i][0]).trim();
        const currentCover = dataValues[i][COVER_COL_NUM - TITLE_COL_NUM]; 
        const rowIndex = START_ROW + i;

        // Skip if blank or already has a cover
        if (!title || (currentCover && currentCover.toString().trim() !== "")) continue;

        let imgUrl = null;
        let cleanTitle = superFuzzy(title);

        // STEP A: SEARCH DRIVE FIRST
        if (driveFolder) {
            try {
                const files = driveFolder.getFilesByName(cleanTitle + ".jpg");
                if (files.hasNext()) imgUrl = files.next().getDownloadUrl();
            } catch(e) {}
        }

        // STEP B: SEARCH WEB SECOND
        if (!imgUrl) {
            try {
                const bookId = idMap[cleanTitle];
                const query = bookId ? `isbn:${bookId}` : encodeURIComponent(title + " book cover");
                const response = UrlFetchApp.fetch(`https://www.googleapis.com/books/v1/volumes?q=${query}&maxResults=1`, {muteHttpExceptions: true});
                const res = JSON.parse(response.getContentText());

                if (res.items && res.items[0].volumeInfo.imageLinks) {
                    imgUrl = (res.items[0].volumeInfo.imageLinks.thumbnail || res.items[0].volumeInfo.imageLinks.smallThumbnail).replace('http:', 'https:');
                }
            } catch(e) {
                console.log("Skipping search for: " + title);
            }
        }

        // STEP C: APPLY IMAGE
        if (imgUrl) {
            try {
                const image = SpreadsheetApp.newCellImage().setSourceUrl(imgUrl).setAltTextDescription(title).build();
                sheet.getRange(rowIndex, COVER_COL_NUM).setValue(image);
                matchCount++;
                if (matchCount % 5 === 0) SpreadsheetApp.flush();
            } catch(e) {}
        }

        if (matchCount >= 15) break;
    }

    return { count: matchCount };
}


/** --- Helper Functions --- **/

function normalize(str) {
    if (!str) return "";
    return str.toString().toLowerCase().replace(/^(the |a |an )/i, "").replace(/[^a-z0-9]/g, "").trim();
}

/**
 * Global Normalizer (The "Super Fuzzy" Engine)
 */
function superFuzzy(str) {
    if (!str) return "";
    return str.toString()
        .toLowerCase()
        .split(/[:(#]/)[0]             
        .replace(/[^a-zA-Z0-9\s]/g, "") 
        .replace(/\s+/g, " ")           
        .split("vol")[0]                
        .trim();
}

function findFirstEmptyRow(sheet, col, startRow) {
    const vals = sheet.getRange(col + startRow + ":" + col).getValues();
    for (let i = 0; i < vals.length; i++) {
        if (vals[i][0] === "") return startRow + i;
    }
    return startRow + vals.length;
}

/**
 * Essential Helper: Prevents "#NAME?" errors
 * @customfunction
 */
function CELLIMAGE(url) {
    if (!url) return "";
    return SpreadsheetApp.newCellImage()
        .setSourceUrl(url)
        .setAltTextDescription('Book Cover')
        .build();
}

function resetCoverMemory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reading list'); 
  
  // =====================================================================
  // ⚙️ RESET CONFIGURATION
  // =====================================================================
  const START_ROW = 4;
  const COVER_COL_NUM = 4; // D = 4
  // =====================================================================
  
  const lastRow = sheet.getLastRow();
  
  if (lastRow >= START_ROW) {
    sheet.getRange(START_ROW, COVER_COL_NUM, lastRow - (START_ROW - 1), 1).clearContent();
    SpreadsheetApp.getUi().alert('Cover memory cleared! Now run Master Sync.');
  }
}

/**
 * 🧹 DATA CLEANER (Book tracker 🧸 EDITION)
 * Grabs raw CSV from Goodreads_Import, splits Title/Series, and pastes to Goodreads_Cleaned!
 */
function cleanGoodreadsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName("Goodreads_Import");
  const cleanSheet = ss.getSheetByName("Goodreads_Cleaned");

  if (!importSheet || importSheet.getLastRow() < 2) {
    return ss.toast("⚠️ No raw data found in Goodreads_Import!", "Data Cleaner", -1);
  }

  const rawData = importSheet.getDataRange().getValues();
  const headers = rawData[0];

  // Map the raw Goodreads columns
  const idIdx = headers.indexOf("Book Id");
  const titleIdx = headers.indexOf("Title");
  const authorIdx = headers.indexOf("Author");
  const coAuthorIdx = headers.indexOf("Additional Authors");
  const shelfIdx = headers.indexOf("Exclusive Shelf");
  const dateAddedIdx = headers.indexOf("Date Added");
  const dateReadIdx = headers.indexOf("Date Read");
  const pageIdx = headers.indexOf("Number of Pages");

  if (titleIdx === -1) return ss.toast("❌ No 'Title' column found in raw import!", "Error", -1);

  let cleanedData = [];

  for (let i = 1; i < rawData.length; i++) {
    let row = rawData[i];
    let rawTitle = row[titleIdx] ? row[titleIdx].toString().trim() : "";
    if (!rawTitle) continue;

    let bookId = idIdx > -1 ? row[idIdx] : "";
    let author = authorIdx > -1 ? row[authorIdx] : "";
    let coAuthor = coAuthorIdx > -1 ? row[coAuthorIdx] : "";
    let fullAuthor = coAuthor ? `${author} & ${coAuthor}` : author;
    let shelf = shelfIdx > -1 ? row[shelfIdx] : "";
    let dateAdded = dateAddedIdx > -1 ? row[dateAddedIdx] : "";
    let dateRead = dateReadIdx > -1 ? row[dateReadIdx] : "";
    let pages = pageIdx > -1 ? row[pageIdx] : "";

    let cleanTitle = rawTitle;
    let series = "";
    
    // Extracts Series from the Goodreads "Title (Series, #1)" format
    let match = rawTitle.match(/(.*?)\s*\((.*?(?:#\d+|Volume|Vol|Book).*?)\)/i);
    if (match) {
        cleanTitle = match[1].trim();
        series = match[2].trim();
    } else {
        let fallbackMatch = rawTitle.match(/(.*?)\s*\((.*?)\)$/);
        if (fallbackMatch && fallbackMatch[2].includes(",")) {
            cleanTitle = fallbackMatch[1].trim();
            series = fallbackMatch[2].trim();
        }
    }

    // Maps exactly to columns A through I on Goodreads_Cleaned
    cleanedData.push([i, bookId, cleanTitle, series, fullAuthor, shelf, dateAdded, dateRead, pages]);
  }

  // Set the exact headers so your Sync scripts never fail
  const cleanHeaders = [["#", "Book Id", "Clean Title", "Series", "Author", "Exclusive Shelf", "Date Added", "Date Read", "Number of Pages"]];
  cleanSheet.getRange(1, 1, 1, 9).setValues(cleanHeaders);

  // Clear old data and paste the newly cleaned data
  let cleanLastRow = cleanSheet.getLastRow();
  if (cleanLastRow > 1) {
      cleanSheet.getRange(2, 1, cleanLastRow - 1, 9).clearContent();
  }

  if (cleanedData.length > 0) {
      cleanSheet.getRange(2, 1, cleanedData.length, 9).setValues(cleanedData);
      ss.toast(`✨ Cleaned ${cleanedData.length} books and prepped them for routing!`, "Data Cleaner");
  }
}

/**
 * 🧮 TOGGLE CALCULATOR (Book tracker 🧸 EDITION)
 * Opens the Ratings Calculator sheet if hidden, hides it if open!
 */
function toggleCalculator() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName("Ratings Calculator");
  const mainSheet = ss.getSheetByName("Reading list");

  if (!calcSheet) {
    return ss.toast("⚠️ Could not find the 'Ratings Calculator' tab!", "Error", -1);
  }

  // If the calculator is hidden, show it and jump right to it!
  if (calcSheet.isSheetHidden()) {
    calcSheet.showSheet();
    ss.setActiveSheet(calcSheet);
    ss.toast("🧮 Calculator Open!", "Quick Tools", 3);
  } 
  // If the calculator is already visible, hide it and jump back to the Reading list!
  else {
    calcSheet.hideSheet();
    if (mainSheet) ss.setActiveSheet(mainSheet);
    ss.toast("🧮 Calculator Tucked Away!", "Quick Tools", 3);
  }
}

function showRatingsGuide() {
  var html = HtmlService.createHtmlOutputFromFile('RatingsGuide')
      .setTitle('Daily Tools & Ratings')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function resetAllGridCards() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dnfLog = ss.getSheetByName("DNF Log");
  
  // 1. Define your master source template block
  var sourceTemplate = dnfLog.getRange("AC45:BB54");
  
  // 2. Exact Grid Math settings from your layout mapping
  const maxCardsPerRow = 4;
  const maxRowsOfCards = 20; 
  const rowStride = 11;      // Jumps 11 rows down per grid row
  const colStride = 27;      // Jumps 27 columns right per grid column
  
  // Array to hold all target ranges for a faster batch execution
  var targetRanges = [];

  // 3. Loop through the entire layout grid
  for (let r = 0; r < maxRowsOfCards; r++) {
    // Calculates the starting row for the current card block (Row 12)
    let startRow = 12 + (r * rowStride); 

    for (let c = 0; c < maxCardsPerRow; c++) {
      // Calculates the starting column for the current card block (Col B / 2)
      let startCol = 2 + (c * colStride); 
      
      // Target range matching the exact 10x26 size of the template
      let cardRange = dnfLog.getRange(startRow, startCol, 10, 26);
      
      // Copy the blank template over this card block
      sourceTemplate.copyTo(cardRange);
    }
  }
  
  ss.toast("All DNF Log cards have been reset to default state! 🔄", "Reset Complete");
}