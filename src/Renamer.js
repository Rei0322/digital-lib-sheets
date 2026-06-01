/**
 * SMART RENAMER (Book tracker 🧸 EDITION)
 * Grabs Book IDs from the Import tab, cleans the titles on the fly, and renames the files!
 */
function smartRenameFromDrive() {
    const folderId = 'token';
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 🌟 THE FIX: Pointing back to the original sheet to grab those IDs!
    const importSheet = ss.getSheetByName('Goodreads_Import'); 
    
    if (!importSheet) return ss.toast("⚠️ Could not find the Goodreads_Import tab.", "Renamer", -1);
    
    const lastRow = importSheet.getLastRow();
    if (lastRow < 2) return;
    
    // Looking for the original Goodreads headers
    const headers = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues()[0];
    const idIdx = headers.indexOf("Book Id");
    const titleIdx = headers.indexOf("Title");
    
    if (idIdx === -1 || titleIdx === -1) {
        return ss.toast("⚠️ Missing 'Book Id' or 'Title' columns on the Import tab!", "Renamer", -1);
    }

    const importData = importSheet.getRange(2, 1, lastRow - 1, importSheet.getLastColumn()).getValues();
    const idMap = new Map();
    
    // Build the ID Dictionary
    importData.forEach(row => {
        if (row[idIdx] && row[titleIdx]) idMap.set(row[idIdx].toString(), row[titleIdx]);
    });

    let renameCount = 0;

    // Scan Drive and match IDs to Titles
    while (files.hasNext()) {
        let file = files.next();
        let fileName = file.getName().toLowerCase();
        let fileIdMatch = fileName.match(/\d+/); // Finds the Book ID in the image name

        if (fileIdMatch) {
            let fileId = fileIdMatch[0];
            if (idMap.has(fileId)) {
                let rawTitle = idMap.get(fileId);
                
                // 🌟 THE MAGIC: Strip the series info right here to make it perfectly "Clean"!
                let title = superFuzzy(rawTitle); 
                
                let newName = title + ".jpg";
                if (fileName !== newName) {
                    file.setName(newName);
                    renameCount++;
                }
            }
        }
    }
    
    if (renameCount > 0) {
        ss.toast(`✅ Successfully cleaned and renamed ${renameCount} files!`, "Renamer", 5);
    } else {
        ss.toast(`✅ Scan complete. No files needed renaming!`, "Renamer", 5);
    }
}

function superFuzzy(str) {
    if (!str) return "";
    return str.toString()
        .toLowerCase()
        .split(/[(\（]/)[0] // This automatically chops off the "(Series, #1)" part!
        .replace(/&amp;/g, " ")
        .replace(/&/g, " ")
        .replace(/\bamp\b/g, " ")
        .replace(/\band\b/g, " ")
        .replace(/[^a-z0-9]/g, " ") 
        .replace(/\s+/g, " ")       
        .trim();
}