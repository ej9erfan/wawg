function generateSpecialItemsCoordinator() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
    const outputSheetName = "SpecialItemsCoordinator";
    let outputSheet = ss.getSheetByName(outputSheetName);
    if (!outputSheet) outputSheet = ss.insertSheet(outputSheetName);

    // Completely clear previous content, formatting, filters
    outputSheet.clearContents();
    outputSheet.clearFormats();
    outputSheet.setFrozenRows(0);
    outputSheet.setFrozenColumns(0);
    if (outputSheet.getFilter()) outputSheet.getFilter().remove(); // remove filter

    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    const dataRows = data.slice(1);
    
    // --- KEYWORD DEFINITION FOR CONDITIONAL BOLDING ---
    const keywords = ["gluten", "dairy", "vegan", "vegetarian", "allergy"];
    // --------------------------------------------------

    // --- DIAPER SIZE DATA DEFINITION (START) ---
    const rawDiaperSizes = [
        // CODE | DESCRIPTION | RANGE
        ["N", "Newborn Diapers", "< 10 lbs"],
        ["1", "Size 1 Diapers", "8-14 lbs"],
        ["2", "Size 2 Diapers", "12-18 lbs"],
        ["3", "Size 3 Diapers", "16-28 lbs"],
        ["4", "Size 4 Diapers", "22-37 lbs"],
        ["5", "Size 5 Diapers", "27 + lbs"],
        ["6", "Size 6 Diapers", "35 + lbs"],
        ["7", "Size 7 Diapers", "41 + lbs"],
        ["2/3T", "Size 2T/3T PullUps", "16-34 lbs"],
        ["3/4T", "Size 3T/4T PullUps", "30-40 lbs"],
        ["4/5T", "Size 4T/5T PullUps", "37 + lbs"],
        ["5/6T", "Size 5T/6T PullUps", "41 + lbs"],
        ["6/7T", "Size 6T/7T PullUps", "64 + lbs"],
        ["MASMD", "Masc Adult Sm/Med Diapers", "28-40 inch waist"],
        ["MALXD", "Masc Adult Large/XL Diapers", "38-64 inch waist"],
        ["FASMD", "Fem Adult Sm/Med Diapers", "28-40 inch waist"],
        ["FALD", "Fem Adult Large Diapers", "38-50 inch waist"],
        ["FAXD", "Fem Adult XL Diapers", "48-64 inch waist"],
        ["FAXXD", "Fem Adult XXL Diapers", "64-80 inch waist"]
    ];
    
    // This array will hold the final list for the table rows (18 data + 1 total + 1 wipes = 20 rows)
    const allDiaperItems = [...rawDiaperSizes]; 
    
    // Add Total Diaper Packs row (Req 4)
    allDiaperItems.push(["Diapers Total", "Total Diaper Packs", ""]); 
    
    // Add Wipes row
    allDiaperItems.push(["Wipes", "Wipes Packs", ""]); 
    // --- DIAPER SIZE DATA DEFINITION (END) ---

    // Identify key indexes
    const routeIdx = headers.indexOf("routeDescription");
    const driverIdx = headers.indexOf("Driver");
    
    // Locate the "Conf for" column
    const confForHeader = headers.find(h => typeof h === 'string' && h.trim().startsWith("Conf for"));
    const confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;

    // Columns for coordinator view
    const includedCols = [
        "Driver", "Name", "# packs", "# eggs", "# milk", 
        "dogFood(qty)", "catFood(qty)", "diapers1(size)", 
        "diapers2(size)", "diapers3(size)", "wipes(qty)", 
        "toiletries(qty)", "fem hygiene(qty)", "Special Item Requests"
    ];
    const includedIndexes = includedCols.map(c => headers.indexOf(c)).filter(i => i !== -1);

    // Original data sheet indices needed for totaling/counting
    const packsOrigIdx = headers.indexOf("# packs");
    const eggsOrigIdx = headers.indexOf("# eggs");
    const milkOrigIdx = headers.indexOf("# milk");
    const wipesOrigIdx = headers.indexOf("wipes(qty)"); // Wipes index
    const diapers1Idx = headers.indexOf("diapers1(size)"); // Diaper indices
    const diapers2Idx = headers.indexOf("diapers2(size)");
    const diapers3Idx = headers.indexOf("diapers3(size)");
    const diaperSizeIndices = [diapers1Idx, diapers2Idx, diapers3Idx].filter(i => i !== -1);
    const dogFoodOrigIdx = headers.indexOf("dogFood(qty)");
    const catFoodOrigIdx = headers.indexOf("catFood(qty)");
    const toiletriesOrigIdx = headers.indexOf("toiletries(qty)");
    const femHygieneOrigIdx = headers.indexOf("fem hygiene(qty)");


    // Build route-driver mapping 
    const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
    const routeDriverMap = {};
    if (helperSheet) {
        const helperData = helperSheet.getDataRange().getValues().slice(1);
        helperData.forEach(row => { 
            const routeKey = String(row[0] || '').trim();
            const driverValue = String(row[1] || '').trim();
            if (routeKey) routeDriverMap[routeKey] = driverValue; 
        });
    }

    // --- Filter data to include only confirmed deliveries ('YES') for Item Summary ---
    let confirmedDataRows = dataRows;
    if (confForIdx !== -1) {
        confirmedDataRows = dataRows.filter(row => {
            const confValue = String(row[confForIdx] || '').toUpperCase().trim();
            return confValue === 'YES';
        });
    }

    // --- Filter data for Diaper/Wipes Count (YES or NO ANS) and Route Tables ---
    let dataRowsForDiaperCount = dataRows;
    if (confForIdx !== -1) {
        dataRowsForDiaperCount = dataRows.filter(r => {
            const confValue = String(r[confForIdx] || '').toUpperCase().trim();
            // Filter is YES or includes NO ANS
            return confValue === 'YES' || confValue.includes('NO ANS');
        });
    }
    // ------------------------------------------------------------------------------------------
    
    // --- CALCULATE DIAPER AND WIPES COUNTS ---
    const diaperCounts = allDiaperItems.reduce((acc, item) => { acc[item[0]] = 0; return acc; }, {});
    let totalDiaperPacks = 0; 

    dataRowsForDiaperCount.forEach(row => {
        // Count Diaper Codes
        diaperSizeIndices.forEach(idx => {
            const code = String(row[idx] || '').trim().toUpperCase();
            if (diaperCounts.hasOwnProperty(code) && code !== "Wipes" && code !== "DIAPER_TOTAL") {
                diaperCounts[code]++;
                totalDiaperPacks++; // Accumulate total diaper packs
            }
        });

        // Count Wipes (quantity)
        const wipesQty = parseFloat(row[wipesOrigIdx] || 0) || 0;
        diaperCounts["Wipes"] += wipesQty;
    });
    
    // Set the total diaper pack count
    diaperCounts["DIAPER_TOTAL"] = totalDiaperPacks;
    // ------------------------------------------

    // ====================================================================
    // === Item Summary Table (A1:E4) - Item in A, Total in B:E merged ===
    // ====================================================================
    const summaryStartRow = 1; 
    const summaryStartCol = 1; // Column A
    const summaryNumCols = 5; // A, B, C, D, E

    const totalPacks = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[packsOrigIdx] || 0) || 0), 0);
    const totalEggs = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[eggsOrigIdx] || 0) || 0), 0);
    const totalMilk = confirmedDataRows.reduce((s, r) => s + (parseFloat(String(r[milkOrigIdx]).replace('*', '') || 0)) || 0, 0);

    // 1. Set Header Row values (A1:E1)
    // A1: "Item", B1: "Total (Confirmed Only)", C1-E1: ""
    outputSheet.getRange(summaryStartRow, summaryStartCol, 1, summaryNumCols).setValues([
        ["Item", "Total (Confirmed Only)", "", "", ""]
    ]); 
    
    // 2. Apply Header Merge: B1:E1
    outputSheet.getRange("B1:E1").merge();

    // 3. Set Data Rows (A2:E4) - Item in A, Total in B
    outputSheet.getRange(summaryStartRow + 1, summaryStartCol, 3, summaryNumCols).setValues([ 
        ["# packs", totalPacks, "", "", ""],
        ["# eggs", totalEggs, "", "", ""],
        ["# milk", totalMilk, "", "", ""]
    ]);
    
    // 4. Apply Data Merges: B2:E2, B3:E3, B4:E4
    for (let r = summaryStartRow + 1; r <= summaryStartRow + 3; r++) {
        outputSheet.getRange(r, 2, 1, 4).merge(); // B:E merge
    }

    // 5. Apply Formatting 
    outputSheet.getRange(summaryStartRow, summaryStartCol, 4, summaryNumCols) 
        .setFontWeight("bold")
        .setFontSize("14")
        .setBackground("#D9EAD3")
        .setBorder(true, true, true, true, null, null);
    
    // Alignments (left for Item, right for Total)
    outputSheet.getRange("A1").setHorizontalAlignment("left"); 
    outputSheet.getRange("B1").setHorizontalAlignment("right"); 
    outputSheet.getRange("A2:A4").setHorizontalAlignment("left"); 
    outputSheet.getRange("B2:B4").setHorizontalAlignment("right"); 

    // ====================================================================
    // === Confirmation Status Table (H1:M5) - H:K merge, L:M merge ===
    // ====================================================================
    const confSummaryStartRow = 1; 
    const confSummaryStartCol = 8; // Column H
    const confSummaryNumCols = 6; // H, I, J, K, L, M

    let confForCounts = { "NO ANS": 0, "NO": 0, "YES": 0, "Total": 0 };
    if (confForIdx !== -1) {
        dataRows.forEach(row => {
            let val = String(row[confForIdx] || '').toUpperCase().trim();
            if (val.includes("NO ANS")) {
                confForCounts["NO ANS"]++;
            } else if (val === "NO") {
                confForCounts["NO"]++;
            } else if (val === "YES") {
                confForCounts["YES"]++;
            }
            if (val) {
                confForCounts["Total"]++;
            }
        });
    }

    // 1. Set Header Row values (H1:M1)
    // H1: "Confirmation Status", K1: "", L1: "Total Count", M1: ""
    outputSheet.getRange(confSummaryStartRow, confSummaryStartCol, 1, confSummaryNumCols).setValues([
        ["Confirmation Status", "", "", "", "Total Count", ""]
    ]);
    
    // 2. Apply Header Merges: H1:K1 and L1:M1
    outputSheet.getRange("H1:K1").merge();
    outputSheet.getRange("L1:M1").merge();
    
    // 3. Set Data Rows (H2:M5) - Status in H, Count in L
    outputSheet.getRange(confSummaryStartRow + 1, confSummaryStartCol, 4, confSummaryNumCols).setValues([
        ["No Ans", "", "", "", confForCounts["NO ANS"], ""],
        ["No", "", "", "", confForCounts["NO"], ""],
        ["Yes", "", "", "", confForCounts["YES"], ""],
        ["Total", "", "", "", confForCounts["Total"], ""]
    ]);

    // 4. Apply Data Merges: H2:K2, L2:M2, etc.
    for (let r = confSummaryStartRow + 1; r <= confSummaryStartRow + 4; r++) {
        outputSheet.getRange(r, 8, 1, 4).merge(); // H:K merge
        outputSheet.getRange(r, 12, 1, 2).merge(); // L:M merge
    }
    
    // 5. Apply Formatting
    outputSheet.getRange(confSummaryStartRow, confSummaryStartCol, 5, confSummaryNumCols)
        .setFontWeight("bold")
        .setFontSize("14")
        .setBackground("#D9D2E9") // Light Purple for distinction
        .setBorder(true, true, true, true, null, null);
        
    // Alignments (left for Status, right for Total)
    outputSheet.getRange("H1").setHorizontalAlignment("left"); 
    outputSheet.getRange("L1").setHorizontalAlignment("right");
    outputSheet.getRange("H2:H5").setHorizontalAlignment("left");
    outputSheet.getRange("L2:L5").setHorizontalAlignment("right");

    // ====================================================================
    // === DIAPER SIZE REFERENCE TABLE (A7:Oxx) - SWAPPED COLUMNS A and F === 
    // ====================================================================
    
    // 1 header + 20 data rows = 21 rows
    const DIAPER_NUM_ROWS = allDiaperItems.length + 1; 
    const diaperTableStartRow = 7; 
    const diaperTableStartCol = 1; // Column A
    const diaperTableNumCols = 15; // A through O 
    
    const diaperDataForSetValues = [
        // Header Row (15 elements: A | B-E | F-G | H-I | J-M | N-O)
        [
            "Size/Weight Ranges",  // A (Swapped from F)
            "Size Description",  // B
            "",      // C
            "",      // D
            "",      // E
            "Code",        // F (Swapped from A)
            "",      // G
            "Pre Ordered", // H
            "",      // I
            "Not Ordered Tally", // J
            "",    // K
            "",    // L
            "",    // M
            "Not Fulfilled Tally", // N
            ""] // O
    ]; 

    // Build the data array for setValues, inserting the calculated count
    allDiaperItems.forEach(row => {
        const code = row[0]; // Original Code
        const description = row[1]; // Original Description
        const range = row[2]; // Original Range

        let count = (code === "Diapers Total") ? diaperCounts["DIAPER_TOTAL"] : diaperCounts[code] || 0;

        // [A:Range, B:Desc, C, D, E, F:Code, G, H:Count, I, J:NotOrdered, K, L, M, N:Unfulfilled, O]
        diaperDataForSetValues.push([
            range, // A: Size/Weight Ranges (Swapped)
            description, // B: Size Description 
            "", "", "", // C, D, E: Blanks for B-E merge
            code, // F: Code (Swapped)
            "", // G: Blank for F-G merge
            count, // H: # of Pre Ordered (The calculated count)
            "", // I: Blank for H-I merge
            "", // J: # of Not Ordered (Blank for J-M merge)
            "", // K: Blank for J-M merge
            "", // L: Blank for J-M merge
            "", // M: Blank for J-M merge
            "", // N: Not Fulfilled Tally
            "" // O: Blank for the N-O merge
        ]);
    });


    // 1. Set the values for the whole range (A7:Oxx)
    const diaperTableRange = outputSheet.getRange(diaperTableStartRow, diaperTableStartCol, DIAPER_NUM_ROWS, diaperTableNumCols);
    diaperTableRange.setValues(diaperDataForSetValues);

    // 2. Apply Merging for the Header and Data Rows
    
    // Header Merging (Row 7)
    // A7: "Size/Weight Ranges" is NOT merged (It was F7:G7) -> Now A7:A7 (1 column)
    outputSheet.getRange(diaperTableStartRow, 2, 1, 4).merge(); // B7:E7 (Desc)
    // F7: "Code" is NOT merged (It was A7:A7) -> Now F7:G7 (2 columns - Adjusted for visual balance)
    outputSheet.getRange(diaperTableStartRow, 6, 1, 2).merge(); // F7:G7 (Code)
    outputSheet.getRange(diaperTableStartRow, 8, 1, 2).merge(); // H7:I7 (PreOrdered)
    outputSheet.getRange(diaperTableStartRow, 10, 1, 4).merge(); // J7:M7 (NotOrdered)
    outputSheet.getRange(diaperTableStartRow, 14, 1, 2).merge(); // N7:O7 (Not Fulfilled)

    // Data Rows Merging (Row 8 to End) - Merge PER ROW
    const firstDataRow = diaperTableStartRow + 1; 
    const numDataRows = DIAPER_NUM_ROWS - 1; 

    for (let r = firstDataRow; r < firstDataRow + numDataRows; r++) {
        // A: Size/Weight Ranges - No Merge (1 column)
        // Merge Size Description (B:E)
        outputSheet.getRange(r, 2, 1, 4).merge(); 
        // Merge Code (F:G) - Adjusted to 2 columns for visual balance
        outputSheet.getRange(r, 6, 1, 2).merge(); 
        // Merge PreOrdered (H:I)
        outputSheet.getRange(r, 8, 1, 2).merge(); 
        // Merge NotOrdered (J:M) 
        outputSheet.getRange(r, 10, 1, 4).merge();
        // Merge Not Fulfilled (N:O)
        outputSheet.getRange(r, 14, 1, 2).merge();
    }
    
    // 3. Apply Formatting
    diaperTableRange
        .setBorder(true, true, true, true, true, true)
        .setBackground("#F3F3F3") 
        .setFontSize(12);
    
    // Header Formatting (Row 7)
    outputSheet.getRange(diaperTableStartRow, diaperTableStartCol, 1, diaperTableNumCols)
        .setFontWeight("bold")
        .setBackground("#CCCCCC") 
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
        
    // Apply specific formatting to the "Total Diaper Packs" row
    const totalRowIndex = firstDataRow + rawDiaperSizes.length; 
    outputSheet.getRange(totalRowIndex, 1, 1, diaperTableNumCols)
        .setFontWeight("bold")
        .setBackground("#CFE2F3"); // Light Blue color
        
    // Set alignment for all data rows (A8 to Oxx)
    outputSheet.getRange(firstDataRow, diaperTableStartCol, numDataRows, 1).setHorizontalAlignment("center"); // Range (A) - Now the Range column
    outputSheet.getRange(firstDataRow, 2, numDataRows, 4).setHorizontalAlignment("left").setWrap(true); // Description (B:E merged)
    outputSheet.getRange(firstDataRow, 6, numDataRows, 2).setHorizontalAlignment("center"); // Code (F:G merged) - Now the Code column
    outputSheet.getRange(firstDataRow, 8, numDataRows, 2).setHorizontalAlignment("center"); // PreOrdered (H:I merged)
    outputSheet.getRange(firstDataRow, 10, numDataRows, 4).setHorizontalAlignment("center"); // NotOrdered (J:M merged)
    outputSheet.getRange(firstDataRow, 14, numDataRows, 2).setHorizontalAlignment("center"); // Not Fulfilled (N:O merged)
    
    // === ALTERNATING ROW SHADING (excluding Total and Wipes) ===
    const regularDiaperRows = rawDiaperSizes.length;
    for (let r = 0; r < regularDiaperRows; r++) {
        // Row index in the sheet is firstDataRow + r
        if (r % 2 !== 0) { // Apply shading to odd-indexed rows (r=1, 3, 5...)
            outputSheet.getRange(firstDataRow + r, 1, 1, diaperTableNumCols)
                .setBackground("#F0F0F0"); // Very light gray
        } else {
             outputSheet.getRange(firstDataRow + r, 1, 1, diaperTableNumCols)
                .setBackground("#FFFFFF"); // White/No color
        }
    }
    
    // Set the Wipes row to a neutral/light shade for distinction
    const wipesRowIndex = totalRowIndex + 1;
    outputSheet.getRange(wipesRowIndex, 1, 1, diaperTableNumCols)
        .setFontWeight("bold")
        .setBackground("#D9EAD3"); // Light Mint Green color

    // === HEAVIER BORDER LINES at logical breaks ===
    const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM; // Medium thickness border
    const borderColor = "#000000";

    // 1. After Size 7 Diapers (baby/toddler separation)
    const afterSize7RowIndex = firstDataRow + rawDiaperSizes.findIndex(d => d[0] === "7"); // Index 7 (8th row)
    outputSheet.getRange(afterSize7RowIndex, diaperTableStartCol, 1, diaperTableNumCols)
        .setBorder(null, null, true, null, null, null, borderColor, borderStyle);

    // 2. After 6T/7T PullUps (pull-ups separation)
    const after6T7TRowIndex = firstDataRow + rawDiaperSizes.findIndex(d => d[0] === "6/7T"); // Index 12 (13th row)
    outputSheet.getRange(after6T7TRowIndex, diaperTableStartCol, 1, diaperTableNumCols)
        .setBorder(null, null, true, null, null, null, borderColor, borderStyle);

    // 3. After MALXD (masculine adult separation)
    const afterMALXDRowIndex = firstDataRow + rawDiaperSizes.findIndex(d => d[0] === "MALXD"); // Index 14 (15th row)
    outputSheet.getRange(afterMALXDRowIndex, diaperTableStartCol, 1, diaperTableNumCols)
        .setBorder(null, null, true, null, null, null, borderColor, borderStyle);

    // ====================================================================

    // Row 30 timestamp 
    const LIGHT_GRAY = "#EFEFEF";
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    const timestampMessage = `This Packing List Tab was generated on ${timestamp}.`;
    const timestampRow = 30; // Adjusted row
    const timestampRange = outputSheet.getRange(timestampRow, 1, 1, 15); // Start at A30, span 15 columns (A30:O30)

    timestampRange.merge()
        .setValue(timestampMessage) // Set the actual formatted timestamp
        .setFontSize(16)
        .setFontWeight("bold")
        .setWrap(true)
        .setBackground(LIGHT_GRAY);

    let outputRow = 33; // Starts on Row 33

    // === Route sections ===
    const routes = [...new Set(dataRows.map(r => String(r[routeIdx]).trim()).filter(r => r))]; 
    const routeColors = [
        "#D9EAD3", // Light Mint Green
        "#FFF2CC", // Light Yellow / Cream
        "#FCE5CD", // Light Peach / Pale Orange
        "#D9D2E9", // Light Lavender / Pale Violet
        "#CFE2F3", // Pale Blue
        "#EEDD82", // New: Light Goldenrod
        "#E0F7FA", // Very Light Cyan / Aqua
        "#F8BBD0", // Light Pink / Rose
        "#E8F5E9", // Very Light Green / Mint
        "#FAEBD7", // New: Antique White (Off-White/Beige)
        "#FFCCBC", // Light Salmon / Light Coral
        "#C5CAE9", // Light Periwinkle
        "#9ACD32"  // New: Yellow-Green
    ];

    // Remove the driver column from the final output
    const includedIndexesNoDriver = includedIndexes.filter(i => i !== driverIdx);

    // Add 2 new checkbox columns
    const checkboxHeaders = ["Packed up & labeled?", "Picked up by driver?"];
    const finalHeaders = includedIndexesNoDriver.map(i => headers[i]).concat(checkboxHeaders);
    
    // Helper to get 1-based column index relative to finalHeaders
    const colIndex = name => finalHeaders.indexOf(name) + 1;

    routes.forEach((route, i) => {
        const driver = routeDriverMap[route] || "TBD";
        const routeColor = routeColors[i % routeColors.length];
        
        const checkStart = colIndex("Packed up & labeled?");
        
        // === Data filtering (Filter by Route AND Confirmed 'YES' OR 'NO ANS') ===
        const filteredDataRows = dataRowsForDiaperCount // Reuse the filtered data for consistency
            .filter(r => String(r[routeIdx]).trim() === route);

        const specialItemsColIndexInOutput = finalHeaders.indexOf("Special Item Requests");
        const rowsToBold = []; 

        const plainRows = filteredDataRows
            .map((r, rowIndex) => includedIndexesNoDriver.map((idx, col) => {
                let value = r[idx];
                
                // Truncate the Name column (col === 0 in the output)
                if (col === 0 && headers[idx] === "Name" && typeof value === 'string') {
                    const maxLength = 36;
                    value = value.length > maxLength ? value.substring(0, maxLength - 3).trim() + '...' : value;
                } 
                
                // Check for non-empty Special Item Requests and track for bolding
                if (col === specialItemsColIndexInOutput) {
                    const strValue = String(value || '').trim();
                    
                    const lowerStrValue = strValue.toLowerCase();
                    const shouldBold = keywords.some(keyword => 
                        lowerStrValue.includes(keyword)
                    );

                    if (shouldBold) {
                        rowsToBold.push(rowIndex); 
                    }
                    return strValue; 
                }

                return value;
            }).concat(["", ""])); // Add empty strings for checkbox columns
        

        if (plainRows.length === 0) {
            outputRow += 2; // Increased from 1 to 2 spacer rows
            return;
        }

        // === Totals Calculation ===
        const totals = {
            packs: 0, eggs: 0, milk: 0, dogFood: 0, catFood: 0, wipes: 0, toiletries: 0, femHygiene: 0
        };

        filteredDataRows.forEach(r => {
            totals.packs += parseFloat(r[packsOrigIdx] || 0) || 0;
            totals.eggs += parseFloat(r[eggsOrigIdx] || 0) || 0;
            totals.milk += parseFloat(String(r[milkOrigIdx]).replace('*', '') || 0) || 0;
            totals.dogFood += parseFloat(r[dogFoodOrigIdx] || 0) || 0;
            totals.catFood += parseFloat(r[catFoodOrigIdx] || 0) || 0;
            totals.wipes += parseFloat(r[wipesOrigIdx] || 0) || 0;
            totals.toiletries += parseFloat(r[toiletriesOrigIdx] || 0) || 0;
            totals.femHygiene += parseFloat(r[femHygieneOrigIdx] || 0) || 0;
        });

        // === Route header === (Row: outputRow)
        const startCol = 1; 
        const mergeCols = finalHeaders.length - (startCol - 1);
        const headerRange = outputSheet.getRange(outputRow, startCol, 1, mergeCols);

        outputSheet.getRange(outputRow, startCol) 
            .setValue(`Driver:   ${driver}   |   Route:   ${route}`)
            .setFontWeight("bold")
            .setBackground(routeColor)
            .setHorizontalAlignment("left")
            .setWrap(false)
            .setFontSize(14)
            .setFontWeight("bold");
        
        headerRange.mergeAcross(); 
        outputRow++;

        // === Column headers === (Row: outputRow)
        outputSheet.getRange(outputRow, 1, 1, finalHeaders.length)
            .setValues([finalHeaders])
            .setFontWeight("bold")
            .setBackground(routeColor)
            .setTextRotation(75)
            .setVerticalAlignment("middle");
        const dataHeaderRow = outputRow; // Store header row index for borders
        outputRow++;

        // === Data rows === (Start Row: outputRow)
        const range = outputSheet.getRange(outputRow, 1, plainRows.length, finalHeaders.length);
        
        range.setValues(plainRows);
        range.setFontSize(12);
        
        if (rowsToBold.length > 0) {
            rowsToBold.forEach(rowIndex => {
                outputSheet.getRange(outputRow + rowIndex, 1, 1, finalHeaders.length)
                    .setFontWeight("bold");
            });
        }
        
        for (let j = 0; j < plainRows.length; j++) {
            outputSheet.getRange(outputRow + j, 1, 1, finalHeaders.length)
                .setBackground(j % 2 === 0 ? "#FFFFFF" : "#F0F0F0");
        }

        const packedCol = includedIndexesNoDriver.length + 1;
        outputSheet.getRange(outputRow, packedCol, plainRows.length, 2).insertCheckboxes();

        outputRow += plainRows.length; 
        const dataEndRow = outputRow - 1; // Last data row index

        // === Totals Row === (Row: outputRow)
        const totalsRowData = new Array(finalHeaders.length).fill("");
        const packsCol = colIndex("# packs") - 1; 
        const eggsCol = colIndex("# eggs") - 1;
        const milkCol = colIndex("# milk") - 1;
        const dogFoodCol = colIndex("dogFood(qty)") - 1;
        const catFoodCol = colIndex("catFood(qty)") - 1;
        const wipesCol = colIndex("wipes(qty)") - 1;
        const toiletriesCol = colIndex("toiletries(qty)") - 1;
        const femHygieneCol = colIndex("fem hygiene(qty)") - 1;

        totalsRowData[0] = "TOTALS"; 
        totalsRowData[packsCol] = totals.packs;
        totalsRowData[eggsCol] = totals.eggs;
        totalsRowData[milkCol] = totals.milk;
        totalsRowData[dogFoodCol] = totals.dogFood;
        totalsRowData[catFoodCol] = totals.catFood;
        totalsRowData[wipesCol] = totals.wipes;
        totalsRowData[toiletriesCol] = totals.toiletries;
        totalsRowData[femHygieneCol] = totals.femHygiene;


        const totalsRange = outputSheet.getRange(outputRow, 1, 1, finalHeaders.length);
        totalsRange.setValues([totalsRowData]);
        totalsRange.setBackground(routeColor);
        totalsRange.setFontWeight("bold");
        totalsRange.setHorizontalAlignment("center");
        outputSheet.getRange(outputRow, 1).setHorizontalAlignment("left"); 

        if (checkStart > 0) {
            outputSheet.getRange(outputRow, checkStart, 1, 2).clearDataValidations();
        }
        const totalsRow = outputRow; // Store totals row index
        outputRow++; 

        // === Borders (Start at Header Row, End at Totals Row) ===
        // The total number of rows for the border is the header + data rows + totals row.
        const borderRangeStart = dataHeaderRow; // Start at the Column Headers row
        const totalRowsForBorder = plainRows.length + 2; // Header (1) + Data (N) + Totals (1)

        const borderStyleRoute = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
        
        // --- Logic to apply vertical borders from Header to Totals Row ---

        const milkColEnd = colIndex("# milk");
        if (milkColEnd > 0) {
             outputSheet.getRange(borderRangeStart, milkColEnd, totalRowsForBorder, 1)
                 .setBorder(null, null, null, true, null, null, "#000000", borderStyleRoute);
        }

        const packsStart = colIndex("# packs");
        const packsEnd = colIndex("catFood(qty)");
        if (packsStart > 0 && packsEnd > 0) {
            outputSheet.getRange(borderRangeStart, packsStart, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", borderStyleRoute);
            outputSheet.getRange(borderRangeStart, packsEnd, totalRowsForBorder, 1)
                .setBorder(null, null, null, true, null, null, "#000000", borderStyleRoute);
        }

        const diapersStart = colIndex("diapers1(size)");
        const diapersEnd = colIndex("wipes(qty)");
        if (diapersStart > 0 && diapersEnd > 0) {
            outputSheet.getRange(borderRangeStart, diapersStart, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", borderStyleRoute);
            outputSheet.getRange(borderRangeStart, diapersEnd, totalRowsForBorder, 1)
                .setBorder(null, null, null, true, null, null, "#000000", borderStyleRoute);
        }

        const hygieneStart = colIndex("toiletries(qty)");
        const hygieneEnd = colIndex("fem hygiene(qty)");
        if (hygieneStart > 0 && hygieneEnd > 0) {
            outputSheet.getRange(borderRangeStart, hygieneStart, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", borderStyleRoute);
            outputSheet.getRange(borderRangeStart, hygieneEnd, totalRowsForBorder, 1)
                .setBorder(null, null, null, true, null, null, "#000000", borderStyleRoute);
        }

        const specialIdx = colIndex("Special Item Requests");
        if (specialIdx > 0) {
            outputSheet.getRange(borderRangeStart, specialIdx, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.DOTTED);
        }

        const checkStartCol = colIndex("Packed up & labeled?");
        if (checkStartCol > 0) {
            outputSheet.getRange(borderRangeStart, checkStartCol, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.DOTTED);
        }
        
        // --- Add horizontal border below the Totals Row (outputRow - 1) ---
        outputSheet.getRange(totalsRow, startCol, 1, finalHeaders.length)
            .setBorder(null, null, true, null, null, null, "#000000", borderStyleRoute);


        // === Spacer Row 1 (Was Spacer Row 2) ===
        const unshadedSpacerRange1 = outputSheet.getRange(outputRow, 1, 1, finalHeaders.length);
        unshadedSpacerRange1.clearContent();
        unshadedSpacerRange1.clearFormat(); 
        if (checkStart > 0) {
            outputSheet.getRange(outputRow, checkStart, 1, 2).clearDataValidations();
        }
        outputRow++; 
        
        // === Spacer Row 2 (New) ===
        const unshadedSpacerRange2 = outputSheet.getRange(outputRow, 1, 1, finalHeaders.length);
        unshadedSpacerRange2.clearContent();
        unshadedSpacerRange2.clearFormat(); 
        if (checkStart > 0) {
            outputSheet.getRange(outputRow, checkStart, 1, 2).clearDataValidations();
        }
        outputRow++; 
    });

    // === Column sizing ===
    
    // Column A: Size/Weight Ranges (New position - was F)
    outputSheet.setColumnWidth(1, 175); 

    // Summary/Diaper Table Columns B-M
    outputSheet.setColumnWidth(2, 60); 
    outputSheet.setColumnWidth(3, 60); 
    outputSheet.setColumnWidth(4, 60); 
    outputSheet.setColumnWidth(5, 60); 
    outputSheet.setColumnWidth(6, 60); // Column F: Code (New position - was A)
    outputSheet.setColumnWidth(7, 60); 
    outputSheet.setColumnWidth(8, 60); 
    outputSheet.setColumnWidth(9, 60); 
    outputSheet.setColumnWidth(10, 60); 
    outputSheet.setColumnWidth(11, 60); 
    outputSheet.setColumnWidth(12, 60); 
    outputSheet.setColumnWidth(13, 60); 

    // Auto-size all data columns starting after M (Column N = 14) onwards
    if (finalHeaders.length > 13) {
        outputSheet.autoResizeColumns(14, finalHeaders.length - 13); 
    }
    
    const specialIdx = finalHeaders.indexOf("Special Item Requests");
    if (specialIdx !== -1) {
        const col = specialIdx + 1;
        outputSheet.setColumnWidth(col, 250);
        // Apply wrap starting from the route-specific data rows (Row 33)
        outputSheet.getRange(33, col, outputSheet.getLastRow() - 32, 1).setWrap(true);
    }

    const checkboxStartCol = finalHeaders.indexOf("Packed up & labeled?");
    if (checkboxStartCol !== -1) {
        const col1 = checkboxStartCol + 1;
        const col2 = checkboxStartCol + 2;
        outputSheet.setColumnWidth(col1, 140);
        outputSheet.setColumnWidth(col2, 140);
    }
    
    SpreadsheetApp.flush();

    return timestamp;
}

function makeTheLabels() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
    const outputSheetName = "DeliveryLabels";
    let outputSheet = ss.getSheetByName(outputSheetName);

    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");

    // --- NEW: Calculate the desired insertion index and log details ---
    let targetIndex;
    
    if (dataSheet) {
        // Log the details of the source sheet
        Logger.log(`Source Sheet ("${dataSheet.getName()}") Index: ${dataSheet.getIndex()}`);

        // FIX: Insert the new sheet immediately AFTER the source data sheet.
        // Sheets are 1-indexed. To place AFTER index N, we use N + 1.
        targetIndex = dataSheet.getIndex() + 1; 
        
        // Log the calculated index
        Logger.log(`Calculated Target Index (Insert AFTER data sheet): ${targetIndex}`);
        
        // If you wanted it at the very end, regardless of location, you would use:
        // targetIndex = ss.getNumSheets(); 
    } else {
        // Fallback: Insert at the very end of the sheet list if the data sheet is not found.
        targetIndex = ss.getNumSheets();
        Logger.log(`Source sheet "Deliveries-UPDATE HERE" not found. Inserting at the end (Index: ${targetIndex}).`);
    }
    // ----------------------------------------------------

    // Only insert the sheet if it doesn't already exist.
    if (!outputSheet) {
        // We now use the calculated targetIndex instead of the hardcoded 0.
        outputSheet = ss.insertSheet(outputSheetName, targetIndex); 
        Logger.log(`Output sheet created at index: ${outputSheet.getIndex()}`);
    } else {
        // If the sheet already exists, its position is not changed by this function
        Logger.log(`Output sheet ("${outputSheetName}") found at index: ${outputSheet.getIndex()}. Position not changed.`);
    }

    // Completely clear previous content, formatting, filters
    outputSheet.clear();
    outputSheet.setFrozenRows(0);
    outputSheet.setFrozenColumns(0);

    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    const dataRows = data.slice(1);

    // --- 1. Identify Key Indices ---
    const nameIdx = headers.indexOf("Name");
    const routeIdx = headers.indexOf("routeDescription");
    
    // Locate the "Conf for" column (case-insensitive find)
    const confForHeader = headers.find(h => typeof h === 'string' && h.trim().startsWith("Conf for"));
    const confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;

    // Indices for Bag Count Calculation
    const diapers1Idx = headers.indexOf("diapers1(size)");
    const diapers2Idx = headers.indexOf("diapers2(size)");
    const diapers3Idx = headers.indexOf("diapers3(size)");
    const wipesIdx = headers.indexOf("wipes(qty)");
    const dogFoodIdx = headers.indexOf("dogFood(qty)");
    const catFoodIdx = headers.indexOf("catFood(qty)");

    // Safety check for required columns
    if (nameIdx === -1 || routeIdx === -1 || confForIdx === -1) {
        Logger.log("Missing essential columns (Name, routeDescription, or Conf for *). Cannot generate labels.");
        SpreadsheetApp.getUi().alert("Missing essential columns (Name, routeDescription, or Conf for *). Cannot generate labels.");
        return;
    }

    // --- 2. Build Route-Driver Mapping ---
    const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
    const routeDriverMap = {};
    if (helperSheet) {
        const helperData = helperSheet.getDataRange().getValues().slice(1);
        helperData.forEach(row => { 
            const routeKey = String(row[0] || '').trim();
            const driverValue = String(row[1] || '').trim();
            if (routeKey) routeDriverMap[routeKey] = driverValue; 
        });
    }

    // --- 3. Filter Data and Process Each Delivery Row ---
    const deliveriesData = [];
    const labelFontSizeMap = {}; // Maps the delivery index to font sizes
    let deliveryIndexCounter = 0; // Counter for mapping purposes

    const filteredDataRows = dataRows
        // Filter by Conf for * = 'YES' or includes 'NO ANS'
        .filter(row => {
            const confValue = String(row[confForIdx] || '').toUpperCase().trim();
            return confValue === 'YES' || confValue.includes('NO ANS');
        });

    filteredDataRows.forEach(row => {
        const name = String(row[nameIdx] || 'UNKNOWN NAME');
        const route = String(row[routeIdx] || 'UNKNOWN ROUTE').trim();
        const driver = routeDriverMap[route] || "TBD";
        
        // Dynamic Font Size Emulation for Name
        const baseNameFontSize = 12;
        const reducedNameFontSize = 10;
        const maxNameLength = 20; 
        const nameFontSize = name.length > maxNameLength ? reducedNameFontSize : baseNameFontSize;

        // Dynamic Font Size Emulation for Driver (USING USER'S CUSTOM VALUES)
        const baseDriverFontSize = 13;
        const reducedDriverFontSize = 11;
        const maxDriverLength = 20;
        const driverFontSize = driver.length > maxDriverLength ? reducedDriverFontSize : baseDriverFontSize;


        // --- Bag Count Logic (Only counts bulky items) ---
        let bagCount = 0;

        // A. Diapers: 1 bag per non-empty column (max 3)
        if (diapers1Idx !== -1 && String(row[diapers1Idx]).trim() !== "") bagCount++;
        if (diapers2Idx !== -1 && String(row[diapers2Idx]).trim() !== "") bagCount++;
        if (diapers3Idx !== -1 && String(row[diapers3Idx]).trim() !== "") bagCount++;

        // B. Wipes: 1 bag per whole quantity
        const wipesQty = parseFloat(row[wipesIdx] || 0) || 0;
        bagCount += Math.ceil(wipesQty);
        
        // C. Pet Food: 1 bag per whole quantity
        const dogFoodQty = parseFloat(row[dogFoodIdx] || 0) || 0;
        const catFoodQty = parseFloat(row[catFoodIdx] || 0) || 0;
        bagCount += Math.ceil(dogFoodQty) + Math.ceil(catFoodQty);
        
        // --- IMPLEMENT HARD LIMIT OF 3 LABELS ---
        bagCount = Math.min(bagCount, 3);
        
        if (bagCount > 0) {
             const delivery = {
                Name: name,
                Driver: driver,
                TotalLabels: bagCount
            };
            
            deliveriesData.push(delivery);
            
            // Store font size metadata keyed by the delivery index
            labelFontSizeMap[deliveryIndexCounter] = { 
                name: nameFontSize,
                driver: driverFontSize
            };
            deliveryIndexCounter++;
        }
    });

    // --- 4. Format Output into Grouped Delivery Layout (2 Rows per Delivery) ---
    const finalOutput = [];
    const labelsPerRow = 3;
    const labelHeightRows = 2; // *** RESTORED: 2 rows per physical label ***
    const labelWidthCols = 2;  
    const totalCols = labelsPerRow * labelWidthCols; // 6 columns 
    
    // Store the delivery index corresponding to each new row pair for formatting later
    const outputRowMapping = []; 

    deliveriesData.forEach((delivery, index) => {
        const row1Data = []; // Name and Bag Numbers
        const row2Data = []; // Driver Names
        
        const numLabelsToGenerate = delivery.TotalLabels;

        for (let j = 0; j < labelsPerRow; j++) {
            if (j < numLabelsToGenerate) {
                // Generate content for this label slot
                const nameLine = `${delivery.Name}`;
                const detailLine = `${delivery.Driver}`; // Driver Name Only
                const numberLine = `BAG ${j + 1}/${delivery.TotalLabels}`;

                // Row 1: Name and Bag Number
                row1Data.push(nameLine, numberLine);
                // Row 2: Driver Name (Pushed twice for merging later)
                row2Data.push(detailLine, detailLine); 
            } else {
                // Leave the remaining slots blank
                row1Data.push("", ""); 
                row2Data.push("", ""); 
            }
        }
        
        finalOutput.push(row1Data, row2Data); // Push both rows for the delivery
        // Store the index of the delivery
        outputRowMapping.push(index);
    });
    
    if (finalOutput.length === 0) {
        SpreadsheetApp.getUi().alert("No confirmed deliveries with items found to generate labels.");
        return;
    }
    
    // --- 5. Write Data to Sheet ---
    
    outputSheet.getRange(1, 1, finalOutput.length, totalCols).setValues(finalOutput);

    // --- 6. Formatting and Sizing for Avery 6240 (30-up) ---
    
    // Set column widths for printing (Columns A/B, C/D, E/F form the 3 labels)
    //const width = 250;
    const width = 258; 
    outputSheet.setColumnWidth(1, width * 0.7); // Name/Route Line (70%)
    outputSheet.setColumnWidth(2, width * 0.3); // Label Number (30%)
    outputSheet.setColumnWidth(3, width * 0.7);
    outputSheet.setColumnWidth(4, width * 0.3);
    outputSheet.setColumnWidth(5, width * 0.7);
    outputSheet.setColumnWidth(6, width * 0.3);
    

    // Loop through the output data rows (2 rows per delivery)
    for (let r = 1; r <= finalOutput.length; r += labelHeightRows) { 
        
        // Calculate the index of the current delivery block in outputRowMapping
        const deliveryIndex = outputRowMapping[(r - 1) / labelHeightRows];
        const deliveryInfo = deliveriesData[deliveryIndex];
        const numLabels = deliveryInfo.TotalLabels;

        // *** RESTORED: Set both Row Heights to 36px (36px + 36px = 72px total label height) ***
        //outputSheet.setRowHeight(r, 36); // Row 1 (Name/Number)
        //outputSheet.setRowHeight(r + 1, 36); // Row 2 (Driver)
        outputSheet.setRowHeight(r, 51.5); // Row 1 (Name/Number)
        outputSheet.setRowHeight(r + 1, 51.5); // Row 2 (Driver)

        // General Formatting for the Label Block (2 rows x 6 columns)
        const labelBlock = outputSheet.getRange(r, 1, labelHeightRows, totalCols);

        // Explicitly remove all borders (Top, Bottom, Left, Right, Vertical, Horizontal)
        labelBlock.setBorder(false, false, false, false, false, false); 

        // Align all content to the middle vertically
        labelBlock.setFontWeight("bold").setVerticalAlignment("middle").setHorizontalAlignment("center");
        
        // --- Apply Dynamic Font Sizing for Names and Drivers ---

        // Get the font sizing metadata for this delivery
        const formatting = labelFontSizeMap[deliveryIndex];
        const nameFontSize = formatting ? formatting.name : 12; 
        const driverFontSize = formatting ? formatting.driver : 13; 

        // Apply formatting across all generated labels in this row pair
        for (let j = 0; j < numLabels; j++) {
            const colNameStart = j * 2 + 1; // 1, 3, 5
            const colBagStart = j * 2 + 2; // 2, 4, 6

            // Row 2: Driver Name Merging (Must happen first)
            outputSheet.getRange(r + 1, colNameStart, 1, 2).mergeAcross();
            
            // Row 1, Name Cell (Dynamic Font Size)
            outputSheet.getRange(r, colNameStart, 1, 1)
                .setFontSize(nameFontSize)
                .setWrap(false)
                .setFontColor("#000000") // Black color
                .setHorizontalAlignment("center");
            
            // Row 1, Bag Number Cell (Fixed Font Size)
            outputSheet.getRange(r, colBagStart, 1, 1)
                .setFontSize(10)
                .setFontWeight("normal")
                .setHorizontalAlignment("left");

            // Row 2, Driver Name Cell (Dynamic Font Size) - Applies to the merged range
            outputSheet.getRange(r + 1, colNameStart, 1, 2)
                .setFontSize(driverFontSize) // *** DYNAMIC FONT SIZE APPLIED HERE ***
                .setFontColor("#3C78D8") // Blue color
                .setHorizontalAlignment("center"); 
        }
    }

    // --- Safely delete excess rows and columns ---
    
    const maxRows = outputSheet.getMaxRows();
    const rowsToDelete = maxRows - finalOutput.length;
    if (rowsToDelete > 0) {
        outputSheet.deleteRows(finalOutput.length + 1, rowsToDelete);
    }
    
    const maxCols = outputSheet.getMaxColumns();
    const colsToDelete = maxCols - totalCols;
    if (colsToDelete > 0) {
        outputSheet.deleteColumns(totalCols + 1, colsToDelete);
    }
    
    SpreadsheetApp.flush();

    return timestamp;
}

function makeDriverRouteTabs() {
  Logger.log('Building driver route tabs...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
  const indexSheetName = "Driver_Deliveries";
  const helperFunction = "populateRouteTabs";
  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");

  //  CLEANUP PREVIOUS OUTPUT
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name.startsWith("Route_")) ss.deleteSheet(sh);
  });

  let indexSheet = ss.getSheetByName(indexSheetName);
  if (!indexSheet) indexSheet = ss.insertSheet(indexSheetName);
  indexSheet.clearContents();
  indexSheet.clearFormats();

  // LOAD DATA & INITIAL SETUP
  const dataRange = dataSheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("No data found in data sheet.");
    return;
  }

  const headers = data[0].map(h => String(h).trim());
  let dataRows = data.slice(1);

  // Column Index Lookups
  const routeColumnIndex = headers.indexOf("routeDescription");
  const nameColumnIndex = headers.indexOf("Name");
  const addressColumnIndex = headers.indexOf("Address");
  const phoneColumnIndex = headers.indexOf("Phone");
  const specReqsColumnIndex = headers.indexOf("Special Item Requests");
  const delivInstrColumnIndex = headers.indexOf("Special Delivery Instructions");
  const deliveryMapLinkIndex = headers.indexOf("Delivery Map Links");
  // New Indices for Diaper Consolidation
  const diapers1Idx = headers.indexOf("diapers1(size)");
  const diapers2Idx = headers.indexOf("diapers2(size)");
  const diapers3Idx = headers.indexOf("diapers3(size)");

  // Check for required columns
  if ([routeColumnIndex, nameColumnIndex, addressColumnIndex, phoneColumnIndex].some(idx => idx === -1)) {
    SpreadsheetApp.getUi().alert("Missing required columns. Ensure 'routeDescription', 'Name', 'Address', and 'Phone' exist.");
    return;
  }

  // Confirmation filtering (Yes / No Ans)
  const confForHeader = headers.find(h => h.startsWith("Conf for"));
  const confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;
  if (confForIdx !== -1) {
    const filteredData = [];
    dataRows.forEach((row) => {
      const val = String(row[confForIdx] || '').toUpperCase().trim();
      if (val === 'YES' || val.includes('NO ANS')) {
        filteredData.push(row);
      }
    });
    dataRows = filteredData;
  }

  // Helper table map: Route → Driver
  const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
  const routeDriverMap = {};
  if (helperSheet) {
    const helperData = helperSheet.getDataRange().getValues().slice(1);
    helperData.forEach(r => {
      const routeKey = String(r[0]).trim();
      if (routeKey) routeDriverMap[routeKey] = String(r[1]).trim();
    });
  }

  // --- START Item Columns to Keep (EMOJI HEADERS) ---
  const itemMap = [
    { label: "food", sourceHeader: "# packs", type: 'qty' },
    { label: "eggs", sourceHeader: "# eggs", type: 'qty' },
    { label: "milk", sourceHeader: "# milk", type: 'qty' },
    { label: "dog food", sourceHeader: "dogFood(qty)", type: 'qty' },
    { label: "cat food", sourceHeader: "catFood(qty)", type: 'qty' },
    { label: "diapers", sourceIndex: [diapers1Idx, diapers2Idx, diapers3Idx], type: 'diapers' },
    { label: "wipes", sourceHeader: "wipes(qty)", type: 'qty' },
    { label: "toiletries", sourceHeader: "toiletries(qty)", type: 'qty' },
    { label: "fem hygn", sourceHeader: "fem hygiene(qty)", type: 'qty' },
  ];

  const finalItemHeaders = itemMap.map(item => item.label);

  // --- END Item Columns to Keep ---

  // Group data by Route, now storing only plain rows
  const routesData = dataRows.reduce((acc, row) => {
    const route = String(row[routeColumnIndex]).trim();
    if (route) {
      if (!acc[route]) acc[route] = {
        plainRows: [],
        addresses: new Set(),
        firstRowIndex: row.join() // Store a key to find the row index later
      };
      acc[route].plainRows.push(row);
      const address = String(row[addressColumnIndex]).trim();
      if (address) acc[route].addresses.add(address);
    }
    return acc;
  }, {});

  // Colors - UNCHANGED
  const routeColors = [
    "#D9EAD3", "#FFF2CC", "#FCE5CD", "#D9D2E9", "#CFE2F3",
    "#F4CCCC", "#EAD1DC", "#E0F7FA", "#F8BBD0", "#B2EBF2",
    "#E8F5E9", "#FFF9C4", "#FFCCBC", "#C5CAE9", "#DCEDC8"
  ];

  // BUILD ROUTE SHEETS (WITH MULTIPLE TABLES)
  const summary = [["Route", "Driver", "Stops", "Open Sheet"]];
  const routes = Object.keys(routesData);

  routes.forEach((route, i) => {
    const routeSafe = route.replace(/[^a-zA-Z0-9]/g, "_").substring(0, 30);
    const routeColor = routeColors[i % routeColors.length];
    const driver = routeDriverMap[route] || "TBD";

    // Extract data for the current route
    const routeData = routesData[route];
    const routeRows = routeData.plainRows;
    const routeAddresses = Array.from(routeData.addresses);

    // --- START: MODIFIED LINK/LABEL EXTRACTION ---
    let routeMapLink = '';
    let routeMapLabel = '';
    const defaultLabel = `Open Map for ${route}`;
    
    if (deliveryMapLinkIndex !== -1 && routeRows.length > 0) {
      // Find the index of the first row in the source sheet (data array)
      const firstRowIndex = data.findIndex(row => row.join() === routeRows[0].join());
            
      if (firstRowIndex !== -1) {
        const sheetRowIndex = firstRowIndex + 1; // 1-based row index
        const sheetColIndex = deliveryMapLinkIndex + 1; // 1-based column index

        const range = dataSheet.getRange(sheetRowIndex, sheetColIndex);
        
        // 1. Get the LABEL first. This is the safest method and always works.
        // getDisplayValue() returns the visible text, whether it's rich text, a formula result, or plain text.
        routeMapLabel = String(range.getDisplayValue() || '').trim();
        
        // 2. Now, try to get the actual LINK URL.
        const richTextValue = range.getRichTextValue();
        
        // Case A: It's a true Rich Text hyperlink.
        if (richTextValue && richTextValue.getLinkUrl()) {
          routeMapLink = richTextValue.getLinkUrl();
        } else {
          // Case B: It's not Rich Text. Check if the raw value is a plain URL.
          const rawValue = String(range.getValue() || '').trim();
          if (rawValue.startsWith('http')) {
            routeMapLink = rawValue;
          }
          // (Note: This intentionally ignores HYPERLINK formulas as they are complex to parse)
        }
      }
    }
    
    // 3. Final label check: Use default if no link, or if label is just the raw link.
    if (!routeMapLink) {
      routeMapLabel = defaultLabel; // No link found, use default label
    } else if (routeMapLink === routeMapLabel || !routeMapLabel) {
      // If the label is just the raw link (Case B) or empty, use the cleaner default label.
      routeMapLabel = defaultLabel;
    }
    // --- END: MODIFIED LINK/LABEL EXTRACTION ---


    // Group stops by a unique identifier (Address + Name)
    const personStops = routeRows.reduce((acc, row) => {
      const address = String(row[addressColumnIndex]).trim();
      const name = String(row[nameColumnIndex]).trim();
      if (address) {
        const key = address + "|" + name;
        if (!acc[key]) acc[key] = [];
        acc[key].push(row);
      }
      return acc;
    }, {});

    // Create the Route Sheet
    const sheet = ss.insertSheet(`Route_${routeSafe}`);
    sheet.clearFormats();
    sheet.setTabColor(routeColor);

    // --- START: CORRECTED HELPER FUNCTION CALL ---
    // Call the helper function
    this[helperFunction](
      ss,
      sheet,
      route,
      driver,
      personStops,
      itemMap,
      finalItemHeaders,
      headers,
      routeColor,
      routeColors,
      {
        nameCol: nameColumnIndex,
        addressCol: addressColumnIndex,
        phoneCol: phoneColumnIndex,
        specCol: specReqsColumnIndex,
        delivCol: delivInstrColumnIndex
      },
      routeMapLink,
      routeMapLabel // <<< This argument was missing in your last version
    );
    // --- END: CORRECTED HELPER FUNCTION CALL ---


    // Add to index
    summary.push([
      route,
      driver,
      Object.keys(personStops).length,
      `=HYPERLINK("#gid=${sheet.getSheetId()}","Open")`
    ]);
  });

  // BUILD INDEX SHEET ("Driver_Deliveries")
  indexSheet.getRange(1, 1, summary.length, summary[0].length).setValues(summary);
  indexSheet.getRange("A1:D1")
    .setFontWeight("bold")
    .setBackground("#F0F0F0")
    .setHorizontalAlignment("center");

  indexSheet.autoResizeColumns(1, 4);
  SpreadsheetApp.flush();
  return timestamp;  
}

function populateRouteTabs(ss, sheet, route, driver, personStops, itemMap, finalItemHeaders, sourceHeaders, routeColor, allStopColors, colIndices, routeMapLink, routeMapLabel) {
  const finalTableWidth = 1 + finalItemHeaders.length;
  let row = 1;
  const itemStartCol = 2; // Column B

  const LIGHT_GRAY = "#EFEFEF";
  const OUTER_BORDER_STYLE = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  const SEPARATOR_BORDER_STYLE = SpreadsheetApp.BorderStyle.DOTTED;

  let stopIndex = 0;

  // Column setup
  sheet.setColumnWidth(1, 250);
  sheet.autoResizeColumns(2, finalItemHeaders.length);

  const mergeStartCol = 2;
  const mergeWidth = finalItemHeaders.length;
  const t0StartRow = row;

  // --- T0: Driver + Route Header Row ---
  // Define your merge regions explicitly
  const driverCols = 4; // A:D
  const routeStartCol = driverCols + 1; // E
  const routeCols = Math.max(1, finalTableWidth - driverCols); // E through J (or whatever fits)

  // Merge A:D for driver
  sheet.getRange(row, 1, 1, driverCols).merge();
  sheet.getRange(row, 1)
    .setValue(driver)
    .setBackground(LIGHT_GRAY)
    .setFontWeight("bold")
    .setFontSize(20)
    .setHorizontalAlignment("left");

  // Merge E:J (or end of table) for route
  sheet.getRange(row, routeStartCol, 1, routeCols).merge();
  sheet.getRange(row, routeStartCol)
    .setValue(route)
    .setBackground(LIGHT_GRAY)
    .setFontWeight("bold")
    .setFontSize(20)
    .setHorizontalAlignment("right");

  row++;

  // --- Map Link: (Row 2: Rich Text Hyperlink) ---
  sheet.getRange(row, 1)
    .setValue('Map Link:')
    .setFontSize(16)
    .setBackground(LIGHT_GRAY)
    .setHorizontalAlignment("left");

  const mapLinkRange = sheet.getRange(row, mergeStartCol, 1, mergeWidth);
  mapLinkRange.merge();

  // Use routeMapLabel for the display text
  if (routeMapLink) {
    // Build a RichTextValue manually using the new generated link
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(routeMapLabel) // Display text
      .setLinkUrl(routeMapLink)     // The actual URL
      .build();
    mapLinkRange.setRichTextValue(richText);
    //Logger.log(`Pre-populated link for ${route}: ${routeMapLink} with label: ${routeMapLabel}`);
  } else {
    // Fallback to message if no link was found in the data source
    mapLinkRange.setValue("Map Link Not Found in Data Sheet");
  }

  // Set to blank
  mapLinkRange.setFontWeight("bold").setFontSize(14);
  mapLinkRange.setBackground(LIGHT_GRAY);
  mapLinkRange.setHorizontalAlignment("right");
  row++;

  // --- Plain Text URL (Row 3: Short Link) ---
  sheet.getRange(row, 1)
    .setValue('Short/Typeable Link:')
    .setFontSize(16)
    .setBackground(LIGHT_GRAY)
    .setHorizontalAlignment("left");

  const blankMapLinkRange = sheet.getRange(row, mergeStartCol, 1, mergeWidth);
  blankMapLinkRange.merge();
  
  // Generate and set the short URL
  let shortUrlValue = "---";
  if (routeMapLink && routeMapLink.startsWith('http')) {
      const shortUrl = getShortUrl(routeMapLink);
      // Only set the short URL if it's different from the original link (meaning shortening succeeded)
      if (shortUrl && shortUrl !== routeMapLink) {
          shortUrlValue = shortUrl;
      } else {
          shortUrlValue = "Original link is too short or shortening failed.";
      }
  }

  blankMapLinkRange.setValue(shortUrlValue);
  blankMapLinkRange.setFontSize(16);
  blankMapLinkRange.setBackground(LIGHT_GRAY);
  blankMapLinkRange.setHorizontalAlignment("right");
  row++;

  // Rollup totals logic (unchanged)
  const rollupTotals = new Array(itemMap.length).fill(0);
  const allRows = Object.keys(personStops).flatMap(key => personStops[key]);

  const diapers1Idx = sourceHeaders.indexOf("diapers1(size)");
  const diapers2Idx = sourceHeaders.indexOf("diapers2(size)");
  const diapers3Idx = sourceHeaders.indexOf("diapers3(size)");

  itemMap.forEach((item, i) => {
    if (item.type === 'qty') {
      const sourceIdx = sourceHeaders.indexOf(item.sourceHeader);
      let totalQty = 0;
      allRows.forEach(r => {
        const val = r[sourceIdx];
        const numericVal = typeof val === 'number' ? val : (typeof val === 'string' ? parseFloat(val) : 0);
        if (!isNaN(numericVal)) totalQty += numericVal;
      });
      rollupTotals[i] = totalQty > 0 ? totalQty : '';
    } else if (item.type === 'diapers') {
      let totalDiaperRequests = 0;
      Object.keys(personStops).forEach(key => {
        const firstRow = personStops[key][0];
        if (diapers1Idx !== -1 && String(firstRow[diapers1Idx] || '').trim()) totalDiaperRequests++;
        if (diapers2Idx !== -1 && String(firstRow[diapers2Idx] || '').trim()) totalDiaperRequests++;
        if (diapers3Idx !== -1 && String(firstRow[diapers3Idx] || '').trim()) totalDiaperRequests++;
      });
      rollupTotals[i] = totalDiaperRequests > 0 ? totalDiaperRequests : '';
    }
  });

  // T0R4/R5 (unchanged)
  sheet.getRange(row, 1)
    .setValue("Route Items:")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(LIGHT_GRAY);
  sheet.getRange(row, itemStartCol, 1, finalItemHeaders.length)
    .setValues([finalItemHeaders])
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(LIGHT_GRAY)
    .setHorizontalAlignment("center");
  row++;
  sheet.getRange(row, 1)
    .setValue("Total Quantities:")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(LIGHT_GRAY);
  sheet.getRange(row, itemStartCol, 1, finalItemHeaders.length)
    .setValues([rollupTotals])
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(LIGHT_GRAY)
    .setHorizontalAlignment("center");
  row++;

  const notesRange = sheet.getRange(row, 1, 1, finalTableWidth);
  notesRange.merge()
    .setValue("In this table here, diapers shows the quantity of PACKS of Diapers and/or Pull-Ups for your route, but in the tables below, it will show the sizes each delivery gets instead. Everything else is a quantity in every table.")
    .setFontSize(10)
    .setWrap(true)
    .setBackground(LIGHT_GRAY);
  try { sheet.autoResizeRows(row, 1); } catch(e) {}
  row++;

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
  const timestampMessage = `This Driver/Route Tab was generated on ${timestamp}.`;

  const timestampRange = sheet.getRange(row, 1, 1, finalTableWidth);
  timestampRange.merge()
    .setValue(timestampMessage) // Set the actual formatted timestamp
    .setFontSize(10)
    .setWrap(true)
    .setBackground(LIGHT_GRAY);
  try { sheet.autoResizeRows(row, 1); } catch(e) {}
  //row++;
  sheet.getRange(t0StartRow, 1, row - t0StartRow + 1, finalTableWidth)
    .setBorder(true, true, true, true, null, null, "#000000", OUTER_BORDER_STYLE);

  row += 3;

  // --- INDIVIDUAL PERSON TABLES (rest of function is unchanged) ---
  Object.keys(personStops).forEach(key => {
    const personData = personStops[key];
    const firstRow = personData[0];
    const currentColor = allStopColors[stopIndex % allStopColors.length];
    const blockStartRow = row;

    const personName = String(firstRow[colIndices.nameCol] || '').trim();
    const address = String(firstRow[colIndices.addressCol] || '').trim();
    const personPhone = String(firstRow[colIndices.phoneCol] || '').trim();
    const specReqs = String(firstRow[colIndices.specCol] || '').trim();
    const delivInstr = String(firstRow[colIndices.delivCol] || '').trim();

    // Calculate item values
    const itemValues = [];
    itemMap.forEach(item => {
      if (item.type === 'qty') {
        const sourceIdx = sourceHeaders.indexOf(item.sourceHeader);
        let totalQty = 0;
        personData.forEach(r => {
          const val = r[sourceIdx];
          const numericVal = typeof val === 'number' ? val : (typeof val === 'string' ? parseFloat(val) : 0);
          if (!isNaN(numericVal)) totalQty += numericVal;
        });
        itemValues.push(totalQty > 0 ? totalQty : '');
      } else if (item.type === 'diapers') {
        const sizes = new Set();
        personData.forEach(r => {
          [diapers1Idx, diapers2Idx, diapers3Idx].forEach(idx => {
            if (idx !== -1) {
              const val = r[idx];
              if (val !== null && val !== undefined) {
                const size = String(val).trim();
                if (size) sizes.add(size);
              }
            }
          });
        });
        const diaperOrder = ['P','N','1','2','3','4','5','6','7','8','2/3T','3/4T','4/5T','5/6T','6/7T'];
        const sortedSizes = Array.from(sizes).sort((a,b)=>diaperOrder.indexOf(a)-diaperOrder.indexOf(b));
        itemValues.push(sortedSizes.join(' | '));
      }
    });

    // --- R1: Merge A+B for Driver/Phone, C–J for Address ---
  sheet.getRange(row, 1, 1, 4) // merge A–C so text can flow
    .merge()
    .setValue(`${personName} / ${personPhone || 'N/A'}`)
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("left")
    .setWrap(false)
    .setBackground(currentColor);
  sheet.getRange(row, 5, 1, finalTableWidth - 4) // starts one column later
    .merge()
    .setValue(address)
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("right")
    .setBackground(currentColor);
    row++;

    // --- R2: Special Item Requests ---
    if (specReqs) {
      sheet.getRange(row, 1)
        .setValue("Special Item Requests:")
        .setFontSize(12)
        .setBackground(currentColor)
        .setHorizontalAlignment("left");
      const range = sheet.getRange(row, 2, 1, finalTableWidth - 1)
        .merge()
        .setValue(specReqs)
        .setFontSize(16)
        .setBackground(currentColor)
        .setWrap(true);
      try { sheet.autoResizeRows(row, 1); } catch(e) {}
      row++;
    }

    // --- R3: Special Delivery Instructions ---
    if (delivInstr) {
      sheet.getRange(row, 1)
        .setValue("Special Instructions:")
        .setFontSize(12)
        .setBackground(currentColor)
        .setHorizontalAlignment("left");
      const range = sheet.getRange(row, 2, 1, finalTableWidth - 1)
        .merge()
        .setValue(delivInstr)
        .setFontSize(16)
        .setBackground(currentColor)
        .setWrap(true);
      try { sheet.autoResizeRows(row, 1); } catch(e) {}
      row++;
    }

    // --- R4: Item Headers ---
    sheet.getRange(row, 1)
      .setValue("Special Item Types:")
      .setFontSize(12)
      .setBackground(currentColor)
      .setHorizontalAlignment("left");
    sheet.getRange(row, itemStartCol, 1, finalItemHeaders.length)
      .setValues([finalItemHeaders])
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setBackground(currentColor);
    row++;

    // --- R5: Item Values ---
    sheet.getRange(row, 1)
      .setValue("Special Item Values:")
      .setFontSize(12)
      .setBackground(currentColor)
      .setHorizontalAlignment("left");
    sheet.getRange(row, itemStartCol, 1, itemValues.length)
      .setValues([itemValues])
      .setFontSize(14)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground(currentColor);
    row++;

    // Border & spacing
    sheet.getRange(blockStartRow, 1, row - blockStartRow, finalTableWidth)
      .setBorder(true, true, true, true, null, null, "#000000", OUTER_BORDER_STYLE);

    row += 2;
    stopIndex++;
  });
  return timestamp;
}

function writeRowsInChunks(sheet, startRow, dataRows, numCols, chunkSize = 50) {
  for (let i = 0; i < dataRows.length; i += chunkSize) {
    const chunk = dataRows.slice(i, i + chunkSize);
    sheet.getRange(startRow + i, 1, chunk.length, numCols).setValues(chunk);
  }
}

function deleteSheetsByPrefix() {
  var prefix = "Route_"; // <--- CHANGE THIS IF YOUR PREFIX CHANGES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetsToDelete = [];
  
  // Identify sheets to delete
  sheets.forEach(function(sheet) {
    if (sheet.getName().startsWith(prefix)) {
      sheetsToDelete.push(sheet);
    }
  });

  // Check if any sheets were found
  if (sheetsToDelete.length === 0) {
    SpreadsheetApp.getUi().alert('No sheets found with the prefix "' + prefix + '".');
    return;
  }
  
  // Confirm deletion with the user
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'Delete Confirmation',
    'Are you sure you want to delete ' + sheetsToDelete.length + ' sheets starting with "' + prefix + '"?',
    ui.ButtonSet.YES_NO
  );
  
  // Process deletion if confirmed
  if (result == ui.Button.YES) {
    sheetsToDelete.forEach(function(sheet) {
      ss.deleteSheet(sheet);
    });
    ui.alert(sheetsToDelete.length + ' sheets successfully deleted!');
  } else {
    ui.alert('Deletion cancelled.');
  }
}

function getShortUrl(longUrl) {
  if (!longUrl || !longUrl.startsWith('http')) {
    return ""; // Return empty if not a valid URL
  }
  
  // 💡 FIX: Using the correct modern API endpoint (create.php)
  // We use the shorturl= parameter with the simple format for a clean URL output.
  const API_URL = "https://is.gd/create.php?format=simple&url="; 
  
  try {
    const response = UrlFetchApp.fetch(API_URL + encodeURIComponent(longUrl));
    const shortUrl = response.getContentText().trim();
    
    // Check if the response is a valid URL (is.gd returns the URL if successful)
    if (shortUrl.startsWith('http')) {
      return shortUrl;
    } else {
      // If it returned an error (usually starts with "Error:"), log it and return the original URL
      Logger.log("is.gd shortener failed with error: " + shortUrl);
      return longUrl; 
    }
  } catch (e) {
    Logger.log("UrlFetchApp failed to reach shortener: " + e.toString());
    // Return the original long URL if the request fails
    return longUrl; 
  }
}

function formatKeywords(text) {
  if (typeof text !== 'string' || !text) {
    return text;
  }
  
  const keywords = ["gluten", "dairy", "vegan", "vegetarian", "allergy"];
  // Red color code
  const RED_COLOR = "#FF0000"; 
  
  // Create a default style for the entire string
  let builder = SpreadsheetApp.newRichTextValue().setText(text);
  let hasFormatting = false;
  
  // Convert the text to lowercase for case-insensitive searching
  const lowerText = text.toLowerCase();
  
  // Loop through keywords and apply formatting for all occurrences
  keywords.forEach(keyword => {
    let index = lowerText.indexOf(keyword);
    // Loop to find all instances of the keyword
    while (index !== -1) {
      hasFormatting = true;
      // Define the text style: bold and red
      const style = SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setForegroundColor(RED_COLOR)
          .build();
      
      // Apply the style to the substring
      // Note: The end index is exclusive, so it's index + keyword.length
      builder = builder.setTextStyle(index, index + keyword.length, style);
      
      // Look for the next occurrence starting from after the current match
      index = lowerText.indexOf(keyword, index + keyword.length);
    }
  });
  
  return hasFormatting ? builder.build() : text;
}

function makeDriverRouteTabsThenPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("Starting makeDriverRouteTabsThenPDF.");
    // 1. CAPTURE THE TIMESTAMP from the function that creates the tabs
    const timestamp = makeDriverRouteTabs(); 
    
    if (timestamp === "ERROR_NO_SHEET") return; 

    Logger.log("Timestamp captured. Flushing changes.");
    SpreadsheetApp.flush(); 
    
    // 2. SANITIZE & CONSTRUCT FILENAME
    const sanitizedTimestamp = timestamp.replace(/[\/:]/g, '-').replace(/\s/g, '_');
    const filename = `Driver_Route_Sheets_${sanitizedTimestamp}.pdf`;
    Logger.log(`PDF filename: ${filename}`);

    const allSheets = ss.getSheets();
    const hiddenSheets = [];

    // Hide all non-Route_ sheets
    allSheets.forEach(sh => {
      if (!sh.getName().startsWith('Route_') && sh.isSheetHidden() === false) {
          sh.hideSheet();
          hiddenSheets.push(sh); 
      }
    });
    Logger.log(`Hid ${hiddenSheets.length} non-route sheets.`);


    // Build PDF export URL using params object for clarity and consistency
    const urlBase = ss.getUrl().replace(/edit$/, '');
    
    const params = {
      'format': 'pdf',
      'portrait': 'true',
      'size': 'letter',
      'fzr': 'false',
      'fitw': 'true',
      'top_margin': '0.25',
      'bottom_margin': '0.25',
      'left_margin': '0.1',
      'right_margin': '0.1'
    };
    
    const options = Object.keys(params).map(key => `${key}=${params[key]}`).join('&');
    const pdfUrl = `${urlBase}export?${options}`;

    Logger.log("Starting PDF fetch. This is the part most likely to timeout.");
    // Fetch the PDF (all visible sheets)
    const token = ScriptApp.getOAuthToken();
    let blob;
    try {
      const response = UrlFetchApp.fetch(pdfUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      
      // Get the blob and set the sanitized filename
      blob = response.getBlob().setName(filename);
      Logger.log("PDF fetch successful.");
    } catch (e) {
      SpreadsheetApp.getUi().alert(`PDF Export Failed: ${e.toString()}`);
      Logger.log(`ERROR during UrlFetchApp.fetch: ${e.toString()}`);
      // Continue to try and restore sheets even on failure
    }
    
    Logger.log("Restoring previously hidden sheets.");
    // Restore previously hidden sheets
    hiddenSheets.forEach(sh => sh.showSheet());

    // Save PDF to Drive only if the blob was successfully created
    if (blob) {
      DriveApp.createFile(blob);
      Logger.log("PDF saved to Drive.");
      SpreadsheetApp.getUi().alert(`✅ Driver Route Sheets PDF created: "${filename}" - check your Google Drive!`);
    } else {
      // Alert that we couldn't save the PDF due to the earlier fetch error
      SpreadsheetApp.getUi().alert(`❌ PDF could not be saved due to an earlier error. Check logs for details.`);
    }
    Logger.log("makeDriverRouteTabsThenPDF finished.");
}

function makeSpecialItemsPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. CAPTURE THE TIMESTAMP from the first function
    const timestamp = generateSpecialItemsCoordinator(); 
    
    if (timestamp === "ERROR_NO_SHEET") return; // Exit if sheet was not found

    SpreadsheetApp.flush(); // Ensure all sheet changes are committed before PDF generation
    
    // 2. SANITIZE & CONSTRUCT FILENAME
    // Remove characters from the timestamp that are illegal or problematic in file paths (like ':', '/')
    const sanitizedTimestamp = timestamp.replace(/[\/:]/g, '-').replace(/\s/g, '_');
    const filename = `Special_Items_Packing_List_${sanitizedTimestamp}.pdf`;

    const sheet = ss.getSheetByName("SpecialItemsCoordinator");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Error: 'SpecialItemsCoordinator' sheet not found for PDF export.");
      return;
    }

    // --- 3. PDF EXPORT LOGIC ---
    // Get the sheet ID for PDF generation
    const sheetId = sheet.getSheetId();
    
    // Base URL for the export
    const url = ss.getUrl().replace(/edit$/, ''); 
    
    // Standard PDF export parameters (adjust these if needed for your specific layout)
    const params = {
      'format': 'pdf',              // export as pdf
      'size': 'letter',                 // A4 paper size is common
      'portrait': 'false',           // portrait orientation
      'fitw': 'true',               // fit to width
      'sheetnames': 'false',        // suppress sheet names
      'printtitle': 'false',        // suppress table title
      'pagenumbers': 'false',       // suppress page numbers
      'gridlines': 'false',         // suppress gridlines
      'fzr': 'false',               // do not repeat row headers
      'gid': sheetId,                // the specific sheet to export
      'top_margin': '0.25',          // Added custom margins
      'bottom_margin': '0.25',
      'left_margin': '0.25',
      'right_margin': '0.25'
    };

    const options = Object.keys(params).map(key => `${key}=${params[key]}`).join('&');
    const pdfUrl = `${url}export?${options}`;

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(pdfUrl, {
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    });
    
    // Get the blob and set the sanitized filename
    const blob = response.getBlob().setName(filename); 
    
    // Save to Google Drive
    DriveApp.createFile(blob);
    SpreadsheetApp.getUi().alert(`✅ Special Items PDF created: "${filename}" - check your Google Drive!`);
}

function makeLabelsPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. CAPTURE THE TIMESTAMP from the first function
    const timestamp = makeTheLabels(); 
    
    if (timestamp === "ERROR_NO_SHEET") return; // Exit if sheet was not found

    SpreadsheetApp.flush(); // Ensure all sheet changes are committed before PDF generation
    
    // 2. SANITIZE & CONSTRUCT FILENAME
    // Remove characters from the timestamp that are illegal or problematic in file paths (like ':', '/')
    const sanitizedTimestamp = timestamp.replace(/[\/:]/g, '-').replace(/\s/g, '_');
    const filename = `Printable_Labels_${sanitizedTimestamp}.pdf`;

    const sheet = ss.getSheetByName("DeliveryLabels");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Error: 'DeliveryLabels' sheet not found for PDF export.");
      return;
    }

    // --- 3. PDF EXPORT LOGIC ---
    // Get the sheet ID for PDF generation
    const sheetId = sheet.getSheetId();
    
    // Base URL for the export
    const url = ss.getUrl().replace(/edit$/, ''); 
    
    // Standard PDF export parameters (adjust these if needed for your specific layout)
    const params = {
      'format': 'pdf',              // export as pdf
      'size': 'letter',                 // A4 paper size is common
      'portrait': 'true',           // portrait orientation
      'fitw': 'true',               // fit to width
      'sheetnames': 'false',        // suppress sheet names
      'printtitle': 'false',        // suppress table title
      'pagenumbers': 'false',       // suppress page numbers
      'gridlines': 'false',         // suppress gridlines
      'fzr': 'false',               // do not repeat row headers
      'gid': sheetId,                // the specific sheet to export
      'top_margin': '0.30',          // Added custom margins
      'bottom_margin': '0.0',
      'left_margin': '0.19',
      'right_margin': '0.19'
    };

    const options = Object.keys(params).map(key => `${key}=${params[key]}`).join('&');
    const pdfUrl = `${url}export?${options}`;

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(pdfUrl, {
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    });
    
    // Get the blob and set the sanitized filename
    const blob = response.getBlob().setName(filename); 
    
    // Save to Google Drive
    DriveApp.createFile(blob);
    SpreadsheetApp.getUi().alert(`✅ Printable Lables PDF created: "${filename}" - check your Google Drive! Download it from Google Drive, Open it in Adobe Reader, when you go to print it, select "Fit" instead of "Actual Size"`);
}

function makeAllPDFs() {
  makeDriverRouteTabsThenPDF();
  makeSpecialItemsPDF();
  makeLabelsPDF();
  SpreadsheetApp.getUi().alert('🎉 All 3 PDFs generated successfully!');
}

function filterForMapCreation() {
  const sheetName = "Deliveries-UPDATE HERE";
  const columnIndex = 9; // Column I is the 9th column
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
    return;
  }
  
  // Get the full data range of the sheet
  const range = sheet.getDataRange();
  
  // If a filter already exists, remove it before applying a new one
  let filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
  
  // Create a new filter on the range
  filter = range.createFilter();
  
  // Create criteria to hide "blanks" and any variation of "no"
  const criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(["", "no", "No", "NO"]) // Hides blanks and case-variations of "no"
    .build();
    
  // Apply the criteria to Column I
  filter.setColumnFilterCriteria(columnIndex, criteria);
  
  SpreadsheetApp.getUi().alert('Filter applied to Column I: Hiding "blanks" and "no".');
}

function clearMapFilter() {
  const sheetName = "Deliveries-UPDATE HERE";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet '${sheetName}' not found.`);
    return;
  }
  
  const filter = sheet.getFilter();
  
  if (filter) {
    filter.remove();
    SpreadsheetApp.getUi().alert('Filter cleared from Data-UPDATE HERE sheet.');
  } else {
    SpreadsheetApp.getUi().alert('No filter was found to clear.');
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📦 Distribution Day Tasks")
    .addItem("1. Filter Data for Map (Col I)", "filterForMapCreation") // <-- NEW
    .addItem("2. Clear Map Data Filter", "clearMapFilter")           // <-- NEW  
    .addSeparator() 
    .addItem("3. Make/Update Driver Route Tabs", "makeDriverRouteTabs")
    .addItem("4. Make/Update Special Items Tab", "generateSpecialItemsCoordinator")
    .addItem("5. Make/Update Printable Special Items Labels Tab", "makeTheLabels")
    .addSeparator()
    .addItem("6. Make Driver Route Tabs then the PDF", "makeDriverRouteTabsThenPDF")
    .addItem("7. Make Special Items Tab then PDF", "makeSpecialItemsPDF")
    .addItem("8. Make Printable Special Items Labels Tab then PDF", "makeLabelsPDF") 
    .addSeparator()
    .addItem("9. Make All PDFs (Click OK on all 4 pop up messages)", "makeAllPDFs")
    .addToUi();
  ui.createMenu("📦 Sheet Tools")
    .addItem('Delete Route Sheets', 'deleteSheetsByPrefix')
    .addToUi();
}