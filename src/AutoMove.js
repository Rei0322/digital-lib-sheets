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
  // 1. MULTI-SELECT DROPDOWNS (For the DNF Log Offenses!)
  // =====================================================================
  if (sheetName === "DNF Log") {
    const rule = e.range.getDataValidation();
    
    // If you click a cell with a dropdown AND it has a previous value
    if (rule != null && e.value && e.oldValue) {
      
      // Combine the old selection with the new selection!
      if (e.oldValue.indexOf(e.value) === -1) {
         e.range.setValue(e.oldValue + ", " + e.value);
      } else {
         // If you accidentally click the same thing twice, just keep the list
         e.range.setValue(e.oldValue);
      }
    }
    return; // Stop here so it doesn't run the rest of the script!
  }

  // =====================================================================
  // 2. WISHLIST TO READING LIST
  // =====================================================================
  if (sheetName === "Wishlist") {
    const STATUS_COL = 3;           // C = 3 (Top-left of the merged checkbox)
    const TITLE_COL = 7;            // G = 7 (Top-left of the merged title)
    const AUTHOR_COL = 19;          // S = 19 (Top-left of the merged author)
    
    const COVER_COL = 14;            // N = 14 (Wishlist Cover)
    const GENRE_COL = 24;            // Y = 25 (Wishlist Genre)
    
    const isChecked = sheet.getRange(row, STATUS_COL).getValue();
    
    if (col === STATUS_COL && isChecked === true && row >= 4) {
      const title = sheet.getRange(row, TITLE_COL).getValue();
      const author = sheet.getRange(row, AUTHOR_COL).getValue();
      const genre = sheet.getRange(row, GENRE_COL).getValue();
      
      // We grab both the value and the formula, just in case your cover uses an =IMAGE() link!
      const coverFormula = sheet.getRange(row, COVER_COL).getFormula();
      const coverValue = sheet.getRange(row, COVER_COL).getValue();
      
      if (!title) return; 
      
      const trackerSheet = ss.getSheetByName("Reading list");
      
      const tValues = trackerSheet.getRange("C:C").getValues();
      let emptyRow = 4; 
      
      for (let i = 3; i < tValues.length; i++) { 
        if (tValues[i][0] === "") {
          emptyRow = i + 1;
          break;
        }
      }
      
      // Paste the data into the correct Reading List columns (C=3, D=4, E=5, F=6)
      trackerSheet.getRange(emptyRow, 3).setValue(title);   
      trackerSheet.getRange(emptyRow, 5).setValue(author);  
      trackerSheet.getRange(emptyRow, 6).setValue(genre);
      
      // Safely place the cover image
      if (coverFormula) {
        trackerSheet.getRange(emptyRow, 4).setFormula(coverFormula);
      } else {
        trackerSheet.getRange(emptyRow, 4).setValue(coverValue);
      }
      
      sheet.getRange(row, TITLE_COL).setFontLine("line-through").setFontColor("#999999");
      ss.toast(`📖 Moved "${title}" to your Reading list!`, "Book Pipeline");
    }
  }
  
  // =====================================================================
  // 3. READING LIST TO DNF LOG
  // =====================================================================
  else if (sheetName === "Reading list") {
    if (col === 8 && row >= 4) {
      const status = String(e.range.getValue()); 
      
      if (status.includes('Did Not Finish') || status.includes('❌')) {
        const dnfLog = ss.getSheetByName("DNF Log");
        const dnfMaxRows = dnfLog.getMaxRows(); 
        
        const title = sheet.getRange(row, 3).getValue();       
        const coverValue = sheet.getRange(row, 4).getValue();  
        const coverFormula = sheet.getRange(row, 4).getFormula(); 
        const author = sheet.getRange(row, 5).getValue();      
        const genre = sheet.getRange(row, 6).getValue();       
        const series = sheet.getRange(row, 17).getValue();     
        
        const cardSpacing = 11; 
        const maxCardsToCheck = 50; 
        
        for (let i = 0; i < maxCardsToCheck; i++) {
          let checkRow = 13 + (i * cardSpacing);
          if (checkRow > dnfMaxRows) break; 
          
          if (dnfLog.getRange(checkRow, 12).getValue() === title) {
            ss.toast("This book is already in your DNF Log! ❌", "Duplicate Prevented");
            return; 
          }
        }
        
        let currentCardIndex = 0;
        let targetRow = 13;
        
        while (currentCardIndex < maxCardsToCheck) {
          if (targetRow > dnfMaxRows) {
            SpreadsheetApp.getUi().alert("You are out of rows in your DNF Log! Please add more blank cards.");
            return;
          }
          if (dnfLog.getRange(targetRow, 12).getValue() === "") {
            break; 
          }
          currentCardIndex++;
          targetRow = 13 + (currentCardIndex * cardSpacing);
        }
        
        if (coverFormula && coverFormula.startsWith("=")) {
          dnfLog.getRange(targetRow, 3).setFormula(coverFormula); 
        } else {
          dnfLog.getRange(targetRow, 3).setValue(coverValue);
        }
        
        dnfLog.getRange(targetRow, 12).setValue(title); 
        dnfLog.getRange(targetRow + 1, 12).setValue(series); 
        dnfLog.getRange(targetRow + 2, 12).setValue(author); 
        dnfLog.getRange(targetRow + 6, 12).setValue(genre); 
        
        ss.toast(`Successfully moved "${title}" to your DNF Log! ❌`, "Success!");
      }
    }
  }
}