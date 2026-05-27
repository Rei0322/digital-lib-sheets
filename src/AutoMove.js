/**
 * AUTO-MOVER MASTER SCRIPT (CRASH-PROOF + MULTI-SELECT MAGIC)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const ss = e.source;

  // =====================================================================
  // 1. MULTI-SELECT DROPDOWNS
  // =====================================================================
  if (sheetName === "DNF Log") {
    const rule = e.range.getDataValidation();
    if (rule != null && e.value && e.oldValue) {
      if (e.oldValue.indexOf(e.value) === -1) {
         e.range.setValue(e.oldValue + ", " + e.value);
      } else {
         e.range.setValue(e.oldValue);
      }
    }
    return; 
  }

  // =====================================================================
  // 2. WISHLIST TO READING LIST
  // =====================================================================
  if (sheetName === "Wishlist") {
    const STATUS_COL = 6;           
    const TITLE_COL = 3;            
    const AUTHOR_COL = 5;           
    const TRIGGER_WORD = "Started"; 
    
    if (col === STATUS_COL && e.value === TRIGGER_WORD && row >= 4) {
      const title = sheet.getRange(row, TITLE_COL).getValue();
      const author = sheet.getRange(row, AUTHOR_COL).getValue();
      if (!title) return; 
      
      const trackerSheet = ss.getSheetByName("Reading list");
      let tValues = trackerSheet.getRange("C4:C").getValues();
      let emptyRow = 4;
      while (emptyRow - 4 < tValues.length && tValues[emptyRow - 4][0] !== "") emptyRow++;
      
      trackerSheet.getRange(emptyRow, 3).setValue(title);   
      trackerSheet.getRange(emptyRow, 5).setValue(author);  
      
      sheet.getRange(row, TITLE_COL, 1, 5).setFontLine("line-through").setFontColor("#999999");
      ss.toast(`📖 Moved "${title}" to your Reading list!`, "Book Pipeline");
    }
  }
  
  // =====================================================================
  // 3. READING LIST AUTOMATIONS (Wrapped Grid DNF Move, Auto-Pages, Removal)
  // =====================================================================
  else if (sheetName === "Reading list") {
    if (col === 8 && row >= 4) { // Column 8 is Status (H)
      const status = String(e.range.getValue() || ""); 
      const title = sheet.getRange(row, 3).getValue();
      if (!title) return; 
      
      // --- Auto-Page Logic ---
      if (status === "Finished") {
        const totalPages = sheet.getRange(row, 13).getValue(); 
        if (totalPages !== "") sheet.getRange(row, 12).setValue(totalPages); 
      } else if (status === "Re-Read" || status === "Currently Reading") {
        sheet.getRange(row, 12).clearContent(); 
      }

      // --- DNF Log Syncing Logic (Wrapped Grid Layout) ---
      const isDNF = status.includes('Did Not Finish') || status.includes('❌');
      const dnfLog = ss.getSheetByName("DNF Log");
      
      // Perfect Grid Math based on your layout
      const maxCardsPerRow = 4;
      const maxRowsOfCards = 20; 
      const rowStride = 11;      // Jumps 11 rows down (e.g., Row 13 -> Row 24)
      const colStride = 27;      // Jumps 27 cols right (e.g., Col L(12) -> Col AM(39))

      // Grab all existing data to prevent script lag
      const allData = dnfLog.getDataRange().getValues();
      const totalRows = allData.length;
      const totalCols = allData[0].length;
      
      let foundRow = -1;
      let foundCol = -1;
      let emptyRow = -1;
      let emptyCol = -1;

      // Scan the grid: Left to right, 4 cards, then drop down 11 rows
      for (let r = 0; r < maxRowsOfCards; r++) {
        let absRow = 13 + (r * rowStride); // Starts at Row 13

        for (let c = 0; c < maxCardsPerRow; c++) {
          let absCol = 12 + (c * colStride); // Starts at Col L (12)
          
          let relativeRow = absRow - 1; // Array math (0-indexed)
          let relativeCol = absCol - 1; 
          
          let cardTitle = "";
          // Safely check if the cell exists in the loaded data
          if (relativeRow < totalRows && relativeCol < totalCols) {
            cardTitle = allData[relativeRow][relativeCol];
          }

          if (cardTitle === title) {
            foundRow = absRow;
            foundCol = absCol;
            break; 
          }
          if (cardTitle === "" && emptyRow === -1) {
            emptyRow = absRow;
            emptyCol = absCol;
          }
        }
        if (foundRow !== -1) break; // Stop outer loop if found
      }

      // --- ADD TO DNF LOG ---
      if (isDNF) {
        if (foundRow !== -1) {
          ss.toast("This book is already in your DNF Log! ❌", "Duplicate Prevented");
          return; 
        }
        if (emptyRow === -1) {
          SpreadsheetApp.getUi().alert("Your DNF Log is entirely full! Please add more cards.");
          return;
        }

        const coverValue = sheet.getRange(row, 4).getValue();  
        const coverFormula = sheet.getRange(row, 4).getFormula(); 
        const author = sheet.getRange(row, 5).getValue();      
        const genre = sheet.getRange(row, 6).getValue();       
        const series = sheet.getRange(row, 17).getValue();  
        const currentPage = sheet.getRange(row, 12).getValue(); 
        const totalPages = sheet.getRange(row, 13).getValue();   

        if (coverFormula && coverFormula.startsWith("=")) {
          dnfLog.getRange(emptyRow, emptyCol - 9).setFormula(coverFormula); 
        } else {
          dnfLog.getRange(emptyRow, emptyCol - 9).setValue(coverValue);
        }
        
        dnfLog.getRange(emptyRow, emptyCol).setValue(title); 
        dnfLog.getRange(emptyRow + 1, emptyCol).setValue(series); 
        dnfLog.getRange(emptyRow + 2, emptyCol).setValue(author); 
        dnfLog.getRange(emptyRow + 6, emptyCol).setValue(genre); 
        dnfLog.getRange(emptyRow + 7, emptyCol + 2).setValue(currentPage); 
        dnfLog.getRange(emptyRow + 7, emptyCol + 8).setValue(totalPages);  
        
        ss.toast(`Successfully moved "${title}" to your DNF Log! ❌`, "Success!");
      } 
      
      // --- SAFELY REMOVE FROM DNF LOG (Fix for merged cells) ---
      else {
        if (foundRow !== -1 && foundCol !== -1) {
          // Using .setValue("") instead of .clearContent() prevents the script from crashing on merged cells
          dnfLog.getRange(foundRow, foundCol - 9).setValue("");     // Cover Image
          dnfLog.getRange(foundRow, foundCol).setValue("");         // Title
          dnfLog.getRange(foundRow + 1, foundCol).setValue("");     // Series
          dnfLog.getRange(foundRow + 2, foundCol).setValue("");     // Author
          dnfLog.getRange(foundRow + 3, foundCol).setValue("");     // Main Offense
          dnfLog.getRange(foundRow + 4, foundCol).setValue("");     // Other Crimes
          dnfLog.getRange(foundRow + 6, foundCol).setValue("");     // Genre
          dnfLog.getRange(foundRow + 6, foundCol + 7).setValue(""); // Quit Date (Col S)
          dnfLog.getRange(foundRow + 7, foundCol + 2).setValue(""); // Pages Reached (Col N)
          dnfLog.getRange(foundRow + 7, foundCol + 8).setValue(""); // Total Pages (Col T)
          
          ss.toast(`Removed "${title}" from your DNF Log!`, "Tracker Updated");
        }
      }
    }
  }
}