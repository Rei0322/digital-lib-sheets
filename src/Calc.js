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
  const TITLE_COL_NUM = 3;         // C = 3 (Book titles)
  const GENRE_COL_NUM = 6;         // F = 6 (Strict Dropdown Genres)
  const START_ROW = 4;             // Row where your book data starts

  // Destination Columns on 'Reading list'
  const STAR_COL = 18;             // R = 18 (Visual Stars)
  const SPICE_COL = 19;            // S = 19 (Visual Peppers)
  const RAW_STAR_COL = 25;         // Y = 25 (Raw Number for Stars)
  const RAW_SPICE_COL = 26;        // Z = 26 (Raw Number for Spice)

  // Location of inputs on the 'Ratings Calculator' Card
  const DROPDOWN_CELL = "G7";      
  const FINAL_SPICE_CELL = "K31";  

  // 🛑 ALLOWED GENRE KEYWORDS 
  // Mapped directly to your Settings!B7:B20 dropdown list
  const CAWPILE_GENRES = ["Fiction", "Mystery", "Thriller", "Horror", "Sci-Fi", "Fantasy", "Graphic Novel"];
  const HEARTED_GENRES = ["Romance", "Romantasy"];
  const CRAFT_GENRES   = ["Memoir", "Self-Help", "Business", "Biography", "Hobby", "Craft"];
  // ======================================================================

  const bookTitle = calcSheet.getRange(DROPDOWN_CELL).getValue();
  const rawSpiceScore = calcSheet.getRange(FINAL_SPICE_CELL).getValue();

  if (!bookTitle) {
    SpreadsheetApp.getUi().alert("⚠️ Please select a book from the dropdown first!");
    return;
  }

  // --- 1. FIND THE BOOK ROW & GRAB ITS GENRE ---
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

  // Fetch the genre for this specific book
  const bookGenre = trackerSheet.getRange(targetRow, GENRE_COL_NUM).getValue().toString();

  // --- 2. DETERMINE WHICH SYSTEM WAS USED ---
  const cawpileValues = calcSheet.getRange("E13:E19").getValues();
  const craftValues = calcSheet.getRange("I13:I17").getValues();
  const heartedValues = calcSheet.getRange("M13:M19").getValues();

  // Helper function to sum arrays
  const sumArray = (arr) => arr.reduce((sum, row) => sum + (Number(row[0]) || 0), 0);

  const cawpileSum = sumArray(cawpileValues);
  const craftSum = sumArray(craftValues);
  const heartedSum = sumArray(heartedValues);

  let totalScore = 0;
  let maxScore = 0;
  let activeSystem = "";

  if (cawpileSum > 0) {
    totalScore = cawpileSum;
    maxScore = 70; 
    activeSystem = "CAWPILE";
  } else if (heartedSum > 0) {
    totalScore = heartedSum;
    maxScore = 70; 
    activeSystem = "HEARTED";
  } else if (craftSum > 0) {
    totalScore = craftSum;
    maxScore = 50; 
    activeSystem = "CRAFT";
  }

  // --- 3. THE GENRE RESTRICTOR LOGIC ---
  if (activeSystem !== "") {
    // Helper to check if any allowed keyword is in the book's genre string
    const isValidGenre = (allowedList, genreString) => {
      return allowedList.some(keyword => genreString.toLowerCase().includes(keyword.toLowerCase()));
    };

    if (activeSystem === "HEARTED" && !isValidGenre(HEARTED_GENRES, bookGenre)) {
      SpreadsheetApp.getUi().alert(`⛔ Genre Mismatch!\nYou are using the HEARTED scale, but '${bookTitle}' is marked as '${bookGenre}'.\n\nPlease use CAWPILE or CRAFT instead.`);
      return; 
    }
    
    if (activeSystem === "CRAFT" && !isValidGenre(CRAFT_GENRES, bookGenre)) {
      SpreadsheetApp.getUi().alert(`⛔ Genre Mismatch!\nYou are using the CRAFT scale, but '${bookTitle}' is marked as '${bookGenre}'.\n\nPlease use CAWPILE or HEARTED instead.`);
      return;
    }

    if (activeSystem === "CAWPILE" && !isValidGenre(CAWPILE_GENRES, bookGenre)) {
      SpreadsheetApp.getUi().alert(`⛔ Genre Mismatch!\nYou are using the CAWPILE scale, but '${bookTitle}' is marked as '${bookGenre}'.\n\nPlease use HEARTED or CRAFT instead.`);
      return;
    }
  }

  // --- 4. CALCULATE FINAL STAR RATING ---
  let rawStarScore = "";
  let visualStars = "";

  if (maxScore > 0) {
    rawStarScore = (totalScore / maxScore) * 5;
    rawStarScore = Math.round(rawStarScore * 100) / 100; 

    if (rawStarScore >= 4.5) { visualStars = "★★★★★"; }
    else if (rawStarScore >= 3.5) { visualStars = "★★★★☆"; }
    else if (rawStarScore >= 2.5) { visualStars = "★★★☆☆"; }
    else if (rawStarScore >= 1.5) { visualStars = "★★☆☆☆"; }
    else if (rawStarScore > 0) { visualStars = "★☆☆☆☆"; }
    else { visualStars = "☆☆☆☆☆"; }
  }

  // --- 5. GENERATE VISUAL PEPPERS ---
  let visualPeppers = "";
  if (rawSpiceScore !== "") {
    let roundedSpice = Math.round(rawSpiceScore);
    if (roundedSpice > 0 && roundedSpice <= 5) {
      visualPeppers = "🌶️".repeat(roundedSpice);
    } else if (roundedSpice > 5) {
      visualPeppers = "🌶️🌶️🌶️🌶️🌶️"; 
    }
  }

  // --- 6. WRITE VALUES TO SPREADSHEET ---
  if (rawStarScore !== "") {
    trackerSheet.getRange(targetRow, STAR_COL).setValue(visualStars);
    trackerSheet.getRange(targetRow, RAW_STAR_COL).setValue(rawStarScore);
  }

  if (rawSpiceScore !== "") {
    trackerSheet.getRange(targetRow, SPICE_COL).setValue(visualPeppers);
    trackerSheet.getRange(targetRow, RAW_SPICE_COL).setValue(rawSpiceScore);
  }

  // --- 7. RESET THE CALCULATOR ---
  calcSheet.getRange(DROPDOWN_CELL).clearContent();
  calcSheet.getRange("E13:E19").clearContent();
  calcSheet.getRange("I13:I17").clearContent();
  calcSheet.getRange("M13:M19").clearContent();
  calcSheet.getRange("G23:G28").clearContent();
  calcSheet.getRange("M23:M27").clearContent();

  SpreadsheetApp.getUi().alert("✨ Success! Ratings for '" + bookTitle + "' have been synced.");
}