/**
 * Book tracker 🧸 RATINGS CALCULATOR - MASTER SYNC
 * Converts raw scores (CAWPILE/HEARTED/CRAFT & Spice) to emojis 
 * and pastes them as static values in the Reading list.
 */
function saveRatingsToTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName("Ratings Calculator");
  const trackerSheet = ss.getSheetByName("Reading list");

  // ======================================================================
  // ⚙️ CALC CONFIGURATION ZONE
  // ======================================================================
  const TITLE_COL_NUM = 3;         // C = 3 (Book titles in Reading list)
  const START_ROW = 4;             // Row where your book data starts

  // Destination Columns on 'Reading list'
  const STAR_COL = 18;             // R = 18 (Visual Stars)
  const SPICE_COL = 19;            // S = 19 (Visual Peppers)
  const RAW_STAR_COL = 25;         // Y = 25 (Raw Number for Stars)
  const RAW_SPICE_COL = 26;        // Z = 26 (Raw Number for Spice)

  // Location of your inputs on the NEW 'Ratings Calculator' Card
  const DROPDOWN_CELL = "G7";      
  
  // ⚠️ UPDATE THESE TWO CELLS to match wherever your new final average 
  // scores will live on your new card (e.g., at the bottom of the card)
  const FINAL_STAR_CELL = "F31";   
  const FINAL_SPICE_CELL = "K31";  
  // ======================================================================

  const bookTitle = calcSheet.getRange(DROPDOWN_CELL).getValue();
  const rawStarScore = calcSheet.getRange(FINAL_STAR_CELL).getValue();
  const rawSpiceScore = calcSheet.getRange(FINAL_SPICE_CELL).getValue();

  if (!bookTitle) {
    SpreadsheetApp.getUi().alert("⚠️ Please select a book from the dropdown first!");
    return;
  }

  // --- 1. Find the Book Row in the 'Reading list' ---
  const dataRange = trackerSheet.getRange(START_ROW, TITLE_COL_NUM, trackerSheet.getLastRow() - START_ROW + 1, 1);
  const titleData = dataRange.getValues();
  let targetRow = -1;

  for (let i = 0; i < titleData.length; i++) {
    if (titleData[i][0] === bookTitle) {
      targetRow = i + START_ROW;
      break;
    }
  }

  if (targetRow === -1) {
    SpreadsheetApp.getUi().alert("❌ Could not find '" + bookTitle + "' in your Reading list!");
    return;
  }

  // --- 2. Generate Visual Stars (Based on your exact scale!) ---
  let visualStars = "";
  if (rawStarScore !== "") {
    if (rawStarScore >= 9.0) { visualStars = "★★★★★"; }
    else if (rawStarScore >= 7.0) { visualStars = "★★★★☆"; }
    else if (rawStarScore >= 4.6) { visualStars = "★★★☆☆"; }
    else if (rawStarScore >= 2.3) { visualStars = "★★☆☆☆"; }
    else if (rawStarScore >= 1.1) { visualStars = "★☆☆☆☆"; }
    else { visualStars = "☆☆☆☆☆"; }
  }

  // --- 3. Generate Visual Peppers (1 to 5) ---
  let visualPeppers = "";
  if (rawSpiceScore !== "") {
    let roundedSpice = Math.round(rawSpiceScore);
    if (roundedSpice > 0 && roundedSpice <= 5) {
      visualPeppers = "🌶️".repeat(roundedSpice);
    } else if (roundedSpice > 5) {
      visualPeppers = "🌶️🌶️🌶️🌶️🌶️"; // Cap it at 5 just in case!
    }
  }

  // --- 4. Write the Values to the Target Row ---
  if (rawStarScore !== "") {
    trackerSheet.getRange(targetRow, STAR_COL).setValue(visualStars);
    trackerSheet.getRange(targetRow, RAW_STAR_COL).setValue(rawStarScore);
  }

  if (rawSpiceScore !== "") {
    trackerSheet.getRange(targetRow, SPICE_COL).setValue(visualPeppers);
    trackerSheet.getRange(targetRow, RAW_SPICE_COL).setValue(rawSpiceScore);
  }

  // --- 5. RESET THE CALCULATOR (Clear all inputs) ---
  // Clears the dropdown
  calcSheet.getRange(DROPDOWN_CELL).clearContent();
  
  // Clears Star Rating inputs (CAWPILE, CRAFT, HEARTED)
  calcSheet.getRange("E13:E19").clearContent();
  calcSheet.getRange("I13:I17").clearContent();
  calcSheet.getRange("M13:M19").clearContent();
  
  // Clears Spice Annex inputs (Climax & Flame)
  calcSheet.getRange("G23:G28").clearContent();
  calcSheet.getRange("M23:M27").clearContent();

  SpreadsheetApp.getUi().alert("✨ Success! Ratings for '" + bookTitle + "' have been synced.");
}