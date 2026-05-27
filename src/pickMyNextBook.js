/**
 * TBR RANDOMIZER & SMART SYNC ENGINE
 * VERSION: GOLD EDITION
 */

// --- PART 1: THE RANDOMIZER ENGINE ---
function pickMyNextBook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TBR Randomizer");
  
  // 1. Get Filters (Format=C6, Tag=D6, Author=E6)
  var filters = sheet.getRange("C6:E6").getValues()[0];
  var fFormat = String(filters[0]).toLowerCase().trim();
  var fTag    = String(filters[1]).toLowerCase().trim();
  var fAuth   = String(filters[2]).toLowerCase().trim();

  var lastRow = sheet.getLastRow();
  if (lastRow < 16) {
    SpreadsheetApp.getUi().alert("Your library is empty! Sync some books first.");
    return;
  }
  
  // Grab A16 to H (Cover to Tags) so we can see the covers!
  var data = sheet.getRange("A16:H" + lastRow).getValues(); 
  var choices = [];

  // 2. The Filter Loop
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var status = String(row[5]).toLowerCase().trim(); // Status is now index 5 (Col F)
    
    if (status === "unread") {
      var format = String(row[4]).toLowerCase().trim(); // Format is index 4 (Col E)
      var author = String(row[2]).toLowerCase().trim(); // Author is index 2 (Col C)
      var tags   = String(row[7]).toLowerCase().trim(); // Tags is index 7 (Col H)

      var matchF = (fFormat === "" || format === fFormat);
      var matchA = (fAuth === "" || author === fAuth);
      var matchT = (fTag === "" || tags.indexOf(fTag) !== -1);

      if (matchF && matchA && matchT) {
        // Save the row data AND the actual row number so we can grab the image later
        choices.push({
          title: row[1],
          author: row[2],
          series: row[3],
          hiddenWeight: row[4],
          actualRow: i + 16
        });
      }
    }
  }

  // 3. Output Setup
  var resCell = sheet.getRange("B12");
  var coverCell = sheet.getRange("A12"); // <--- THIS IS WHERE THE WINNING COVER GOES!
  
  if (choices.length === 0) {
    resCell.setValue("❌ No unread matches found!");
    coverCell.clearContent(); // Clear the cover box if nothing is found
  } else {
    var pick = choices[Math.floor(Math.random() * choices.length)];
    
    var t = String(pick.title).toUpperCase(); 
    var a = String(pick.author);               
    var s = (pick.series) ? " (" + String(pick.series).trim() + ")" : ""; 
    var f = " | " + String(pick.hiddenWeight).toUpperCase(); 

    var visiblePart = "📖 " + t + s + " by " + a;
    var finalStr = f + visiblePart + f; 

    resCell.setValue(finalStr);
    
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(finalStr)
      .setTextStyle(0, f.length, SpreadsheetApp.newTextStyle().setForegroundColor("white").build())
      .setTextStyle(f.length, f.length + visiblePart.length, SpreadsheetApp.newTextStyle().setForegroundColor("black").build())
      .setTextStyle(f.length + visiblePart.length, finalStr.length, SpreadsheetApp.newTextStyle().setForegroundColor("white").build())
      .build();
    
    resCell.setRichTextValue(richText);
    
    // --- 🪄 THE MAGIC COVER COPY ---
    // This perfectly copies the image from Column A of the winning book up to cell A12!
    sheet.getRange(pick.actualRow, 1).copyTo(coverCell);
  }
}

// --- PART 2: CLEAR FILTERS ---
function clearFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TBR Randomizer");
  
  // 1. Clears the dropdowns
  sheet.getRange("C6:E6").clearContent();
  
  // 2. Clears the result boxes
  sheet.getRange("B12").setValue("");
  sheet.getRange("A12").clearContent(); // Clears the cover too!
}

// --- PART 3: THE SMART SYNC SYSTEM ---
function syncDataDrop() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dropSheet = ss.getSheetByName("Data Drop");
  const trackerSheet = ss.getSheetByName("TBR Randomizer");
  const settingsSheet = ss.getSheetByName("Sync Settings");
  
  // 1. Get Mappings
  const colMap = {
    title:  Number(settingsSheet.getRange("B2").getValue()) - 1,
    author: Number(settingsSheet.getRange("B3").getValue()) - 1,
    series: Number(settingsSheet.getRange("B4").getValue()) - 1,
    format: Number(settingsSheet.getRange("B5").getValue()) - 1,
    tags:   Number(settingsSheet.getRange("B6").getValue()) - 1
  };

  // 2. Grab Data
  const startRow = 4;
  const lastRowWithContent = dropSheet.getLastRow();
  if (lastRowWithContent < startRow) {
    SpreadsheetApp.getUi().alert("No data found! Please paste starting at Row 4.");
    return;
  }
  
  const rawData = dropSheet.getRange(startRow, 1, lastRowWithContent - (startRow - 1), dropSheet.getLastColumn()).getValues();

  // 3. Process Data + Header Shield
  const formattedData = [];
  for (let i = 0; i < rawData.length; i++) {
    let row = rawData[i];
    let title = String(row[colMap.title] || "").trim();
    
    // Skip if empty or a header
    if (title !== "" && title.toLowerCase() !== "book title" && title.toLowerCase() !== "title") {
      formattedData.push([
        title, 
        String(row[colMap.author] || "Unknown Author").trim(),
        String(row[colMap.series] || "").trim(),
        String(row[colMap.format] || "Physical").trim(), // The "Trouble" column
        "Unread", 
        "",       
        String(row[colMap.tags] || "").trim()
      ]);
    }
  }

  if (formattedData.length === 0) return;

  // 4. Find Destination
  const trackerTitles = trackerSheet.getRange("B16:B").getValues();
  let relativeRow = 0;
  for (let j = 0; j < trackerTitles.length; j++) {
    if (trackerTitles[j][0] === "") {
      relativeRow = j;
      break;
    }
    if (j === trackerTitles.length - 1) relativeRow = trackerTitles.length;
  }
  let destRow = 16 + relativeRow;

  // 5. THE FIX: Clear validation for EVERY column we are pasting into
  // This prevents any "Violates Rules" errors from stopping the sync
  const destRange = trackerSheet.getRange(destRow, 2, formattedData.length, 7);
  
  // Capture existing rules from Row 16 before wiping
  const formatRule = trackerSheet.getRange("E16").getDataValidation();
  const statusRule = trackerSheet.getRange("F16").getDataValidation();
  const lengthRule = trackerSheet.getRange("G16").getDataValidation();

  // Wipe rules temporarily
  destRange.setDataValidation(null); 

  // Paste the data
  destRange.setValues(formattedData);
  
  // 6. RESTORE DROPDOWNS surgically
  if (formatRule != null) {
    trackerSheet.getRange(destRow, 5, formattedData.length, 1).setDataValidation(formatRule);
  }
  if (statusRule != null) {
    trackerSheet.getRange(destRow, 6, formattedData.length, 1).setDataValidation(statusRule);
  }
  if (lengthRule != null) {
    trackerSheet.getRange(destRow, 7, formattedData.length, 1).setDataValidation(lengthRule);
  }

  SpreadsheetApp.getUi().alert("🚀 Sync Success! " + formattedData.length + " books added.");
}

// --- PART 4: SYNC TO WISHLIST ---
function syncToWishlist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dropSheet = ss.getSheetByName("Data Drop");
  const wishlistSheet = ss.getSheetByName("Wishlist"); 
  const settingsSheet = ss.getSheetByName("Sync Settings");
  
  // 1. Get Mappings from Sync Settings
  const colMap = {
    title:  Number(settingsSheet.getRange("B2").getValue()) - 1,
    author: Number(settingsSheet.getRange("B3").getValue()) - 1,
    series: Number(settingsSheet.getRange("B4").getValue()) - 1,
    format: Number(settingsSheet.getRange("B5").getValue()) - 1,
    tags:   Number(settingsSheet.getRange("B6").getValue()) - 1
  };

  // 2. Grab Data
  const startRow = 4;
  const lastRowWithContent = dropSheet.getLastRow();
  
  if (lastRowWithContent < startRow) {
    SpreadsheetApp.getUi().alert("No data found! Please paste starting at Row 4.");
    return;
  }

  // Get all the raw data from the Data Drop tab
  const numRows = lastRowWithContent - startRow + 1;
  const rawData = dropSheet.getRange(startRow, 1, numRows, dropSheet.getLastColumn()).getValues();

  // 3. Map Data to the correct columns
  const mappedData = rawData.map(row => {
    return [
      row[colMap.title] || "",
      row[colMap.author] || "",
      row[colMap.series] || "",
      row[colMap.format] || "",
      row[colMap.tags] || ""
    ];
  });

  // 4. Append to Wishlist
  if (mappedData.length > 0) {
    // Find the next empty row in the Wishlist
    const wlNextRow = wishlistSheet.getLastRow() + 1;
    
    // Write the mapped data to the Wishlist
    wishlistSheet.getRange(wlNextRow, 1, mappedData.length, 5).setValues(mappedData);
    
    // Optional: Clear the Data Drop sheet after a successful sync
    // dropSheet.getRange(startRow, 1, numRows, dropSheet.getLastColumn()).clearContent();
    
    SpreadsheetApp.getUi().alert("Successfully synced to your Wishlist! ✨");
  }
}