/**
 * MULTI-LANGUAGE superFuzzy
 */
function superFuzzy(str) {
    if (!str) return "";
    let baseTitle = str.toString().toLowerCase().split(/[(\（]/)[0];
    baseTitle = baseTitle.replace(/&amp;/g, " ").replace(/&/g, " ").replace(/\bamp\b/g, " ").replace(/\band\b/g, " ");
    return baseTitle.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()！？!?]/g,"").replace(/\s+/g, " ").trim();
}

/**
 * ONE-CLICK SYNC (Book tracker 🧸 EDITION)
 * Scans both Reading list and Wishlist, then batch-pastes =IMAGE() formulas!
 */
function syncCoversFromDrive() {
    const startTime = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folderId = 'token';

    // ⚙️ Book tracker 🧸 CONFIGURATION
    const targets = [
        { sheet: ss.getSheetByName('Reading list'), startRow: 4, titleCol: 3, coverCol: 4 }, // C & D
        { sheet: ss.getSheetByName('Wishlist'), startRow: 13, titleCol: 7, coverCol: 14 }     // G & N
    ];

    let props = PropertiesService.getUserProperties();
    let brainToken = props.getProperty("brainBuildToken");

    // 1. BRAIN SETUP
    let brainSheet = ss.getSheetByName('Drive_Memory');
    if (!brainSheet) {
        brainSheet = ss.insertSheet('Drive_Memory');
        brainSheet.hideSheet();
    }

    // 2. BUILD THE BRAIN
    if (brainSheet.getLastRow() === 0 || brainToken) {
        const folder = DriveApp.getFolderById(folderId);
        let files;

        if (brainToken) {
            try {
                files = DriveApp.continueFileIterator(brainToken);
                ss.toast(`🧠 Resuming Brain scan... (Memorized ${brainSheet.getLastRow()} files)`, "Memory Build", 5);
            } catch(e) {
                files = folder.getFiles();
                ss.toast("🧠 Brain scan restarted...", "Memory Build", 5);
            }
        } else {
            files = folder.getFiles();
            ss.toast("🧠 Brain is empty! Scanning massive Drive folder...", "Memory Build", 5);
        }

        let newMemory = [];

        while (files.hasNext()) {
            if (Date.now() - startTime > 150000) {
                props.setProperty("brainBuildToken", files.getContinuationToken());
                if (newMemory.length > 0) {
                    brainSheet.getRange(brainSheet.getLastRow() + 1, 1, newMemory.length, 2).setValues(newMemory);
                }
                return ss.toast(`⏳ Memorized ${brainSheet.getLastRow()} files so far. Click Sync again!`, "Batch Paused", -1);
            }

            let file = files.next();
            if (file.isTrashed()) continue;
            let cleanName = superFuzzy(file.getName().replace(/\.jpg$/gi, ""));
            newMemory.push([cleanName, file.getId()]);
        }

        if (newMemory.length > 0) {
            brainSheet.getRange(brainSheet.getLastRow() + 1, 1, newMemory.length, 2).setValues(newMemory);
        }

        props.deleteProperty("brainBuildToken");
        ss.toast(`🧠 Drive fully memorized! (${brainSheet.getLastRow()} files). Starting batch sync...`, "Speed Boost", 5);
    }

    // 3. LOAD THE BRAIN
    let brainData = brainSheet.getDataRange().getValues();
    let brainMap = {};
    for (let b = 0; b < brainData.length; b++) {
        if (brainData[b][0]) brainMap[brainData[b][0].toString()] = brainData[b][1];
    }

    // 4. SYNC BOTH SHEETS
    let totalUpdated = 0;

    targets.forEach(target => {
        if (!target.sheet) return;
        const lastRow = target.sheet.getLastRow();
        if (lastRow < target.startRow) return;

        // Fetch just the columns we need
        const dataRange = target.sheet.getRange(target.startRow, target.titleCol, lastRow - target.startRow + 1, (target.coverCol - target.titleCol) + 1).getValues();
        let chunks = [];
        let currentChunk = null;

        for (let i = 0; i < dataRange.length; i++) {
            let rawTitle = dataRange[i][0];
            let existingCover = dataRange[i][target.coverCol - target.titleCol];

            if (!rawTitle || (existingCover !== "" && existingCover !== null)) continue;

            let cleanTitle = superFuzzy(rawTitle);
            let matchedFileId = brainMap[cleanTitle];

            if (matchedFileId) {
                let directUrl = "https://drive.google.com/uc?export=view&id=" + matchedFileId;
                let formula = `=IMAGE("${directUrl}")`;

                if (!currentChunk || currentChunk.endRow !== i + target.startRow - 1) {
                    if (currentChunk) chunks.push(currentChunk);
                    currentChunk = { startRow: i + target.startRow, endRow: i + target.startRow, formulas: [[formula]] };
                } else {
                    currentChunk.endRow = i + target.startRow;
                    currentChunk.formulas.push([formula]);
                }
                totalUpdated++;
            }
        }
        
        if (currentChunk) chunks.push(currentChunk);

        // Batch write to sheet
        for (let c of chunks) {
            target.sheet.getRange(c.startRow, target.coverCol, c.formulas.length, 1).setValues(c.formulas);
        }
    });

    if (totalUpdated > 0) {
        ss.toast(`✅ Successfully embedded ${totalUpdated} covers across your Tracker and Wishlist!`, "Success", 8);
    } else {
        ss.toast(`✅ Scan complete. Everything is perfectly synced!`, "Success", 5);
    }
}

/**
 * READWISE COVER UPDATER (NATIVE CELL IMAGES)
 * Scans both Tracker and Wishlist!
 */
function importFromReadwise() {
    const READWISE_TOKEN = "api key";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const targets = [
        { sheet: ss.getSheetByName("Reading list"), start: 4, title: 3, cover: 4 },
        { sheet: ss.getSheetByName("Wishlist"), start: 4, title: 7, cover: 14 }
    ];

    ss.toast("📚 Matching books and fetching covers from Readwise...", "Readwise Sync", 10);

    try {
        let trackerMap = {};

        // Build a map of books across BOTH sheets
        targets.forEach(t => {
            if (!t.sheet || t.sheet.getLastRow() < t.start) return;
            const data = t.sheet.getRange(t.start, t.title, t.sheet.getLastRow() - t.start + 1, 1).getValues();
            for (let i = 0; i < data.length; i++) {
                let cleanTitle = superFuzzy(data[i][0]);
                if (cleanTitle) trackerMap[cleanTitle] = { sheet: t.sheet, row: i + t.start, col: t.cover };
            }
        });

        let pageUrl = "https://readwise.io/api/v2/books/?category=books";
        let updatedCount = 0;

        while (pageUrl) {
            let response = UrlFetchApp.fetch(pageUrl, { headers: { "Authorization": "Token " + READWISE_TOKEN }, muteHttpExceptions: true });
            if (response.getResponseCode() !== 200) throw new Error("API Token failed!");

            let data = JSON.parse(response.getContentText());

            data.results.forEach(book => {
                let cleanTitle = superFuzzy(book.title);
                if (trackerMap[cleanTitle] && book.cover_image_url && !book.cover_image_url.includes("default")) {
                    let cellImage = SpreadsheetApp.newCellImage().setSourceUrl(book.cover_image_url).build();
                    let loc = trackerMap[cleanTitle];
                    
                    loc.sheet.getRange(loc.row, loc.col).setValue(cellImage);
                    updatedCount++;
                    delete trackerMap[cleanTitle];
                }
            });
            pageUrl = data.next;
        }

        ss.toast(`✅ Successfully injected ${updatedCount} covers from Readwise!`, "Success", 5);

    } catch (error) {
        ss.toast("❌ Readwise Sync failed: " + error.message, "Error", -1);
    }
}