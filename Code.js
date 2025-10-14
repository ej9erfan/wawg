// This function called by manual execution in Google Sheets
function generateDeliveryTables() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Data1");
    const outputSheet = ss.getSheetByName("Deliveries");
    outputSheet.clearContents();
    outputSheet.clearFormats();

    const data = dataSheet.getDataRange().getValues();
    if (data.length < 2) {
        SpreadsheetApp.getUi().alert("No data found in 'Data1' sheet.");
        return;
    }

    const headers = data[0];
    let dataRows = data.slice(1);

    // Identify Key Columns
    const routeColumnIndex = headers.indexOf("routeDescription");
    const driverColumnIndex = headers.indexOf("driver");
    const addressColumnIndex = headers.indexOf("Address");
    
    // Locate the "Conf for" column (which starts with "Conf for ")
    const confForHeader = headers.find(h => typeof h === 'string' && h.trim().startsWith("Conf for"));
    const confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;

    // === Filter dataRows to ONLY include "Yes" confirmations ===
    const confirmedDataRows = (confForIdx !== -1)
        ? dataRows.filter(r => {
            const confValue = String(r[confForIdx] || '').toUpperCase().trim();
            return confValue === 'YES';
        })
        : dataRows; 
    
    dataRows = confirmedDataRows; // Use the filtered list for all subsequent operations
    // =====================================================================

    // Get the route/driver map data from DeliveriesHelpTable and normalize keys
    const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
    const helperData = helperSheet.getDataRange().getValues().slice(1);
    const routeDriverMap = {};
    helperData.forEach(row => {
        const routeKey = String(row[0]).trim();
        if (routeKey) routeDriverMap[routeKey] = String(row[1]).trim();
    });

    // Get all unique routes from the FILTERED data
    const routes = (routeColumnIndex !== -1)
        ? [...new Set(dataRows
            .map(r => String(r[routeColumnIndex]).trim())
            .filter(r => r)
        )]
        : [];

    // Filter out unwanted columns
    const columnsToAlwaysExclude = ["routeDescription", "driver", "Source", "methodWipesEntry", "Updated", "lastName"];

    const excludedColumns = headers.filter(header => {
        if (typeof header === 'string' && header.trim() === "") return true;
        if (typeof header === 'string' && header.trim().startsWith("Conf for")) return true;
        if (columnsToAlwaysExclude.includes(header)) return true;
        return false;
    });

    const includedIndexes = headers.map((h, i) => excludedColumns.includes(h) ? -1 : i).filter(i => i !== -1);
    const filteredHeaders = includedIndexes.map(i => headers[i]);

    // *** IMPORTANT CHANGE: Remove the temporary "Route Map Link" header. It will be added dynamically. ***
    // Note: The original code didn't actually have this line here, but the previous logic relied on it.
    // Since we are changing the placement, we ensure it's not present in the filteredHeaders array.
    
    // Define specific columns to total
    const columnsToTotalHeaders = [
        "# packs", "# eggs", "# milk",
        "dogFood(qty)", "catFood(qty)",
        "wipes(qty)", "toiletries(qty)", "fem hygiene(qty)"
    ];

    const numericColumns = columnsToTotalHeaders
        .map(header => filteredHeaders.indexOf(header))
        .filter(idx => idx !== -1);

    const diaperColumns = ["diapers1(size)", "diapers2(size)", "diapers3(size)"]
        .map(c => filteredHeaders.indexOf(c))
        .filter(idx => idx !== -1);

    const supportedDiaperSizes = ["N","1","2","3","4","5","6"];
    const supportedPullupSizes = ["2/3T","3/4T","4/5T","5/6T"];

    const routeColors = [
        "#D9EAD3", "#FFF2CC", "#FCE5CD", "#D9D2E9", "#CFE2F3", "#F4CCCC", 
        "#EAD1DC", "#E0F7FA", "#F8BBD0", "#B2EBF2", "#E8F5E9", "#FFF9C4", 
        "#FFCCBC", "#C5CAE9", "#DCEDC8"
    ];

    const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;

    // --- Column Index Lookups for Borders (1-based index) ---
    // These indices are based on filteredHeaders (not the final sheet columns)
    const packsStart = filteredHeaders.indexOf("# packs") + 1;
    const packsEnd = filteredHeaders.indexOf("# milk") + 1;

    const petFoodStart = filteredHeaders.indexOf("dogFood(qty)") + 1;
    const petFoodEnd = filteredHeaders.indexOf("catFood(qty)") + 1;

    const diapersStart = filteredHeaders.indexOf("diapers1(size)") + 1;
    const diapersEnd = filteredHeaders.indexOf("wipes(qty)") + 1;

    const hygieneStart = filteredHeaders.indexOf("toiletries(qty)") + 1;
    const hygieneEnd = filteredHeaders.indexOf("fem hygiene(qty)") + 1;
    
    // -----------------------------------------------------------

    // === START OUTPUT LOGIC ===
    let outputRow = 1;
    const mapLinkTargetCol = 15; // Target is Column O (15th column)
    
    // Determine the max width needed for all rows (data columns or the map link column, whichever is wider)
    const finalTableWidth = Math.max(filteredHeaders.length, mapLinkTargetCol);

    // --- 1. Global Rollup Link (Row 1, Column E) ---
    const allConfirmedAddresses = dataRows
        .map(r => String(r[addressColumnIndex]).trim())
        .filter(a => a);

    let allRoutesMapLink = "";
    if (allConfirmedAddresses.length > 0) {
        const encodedAddresses = allConfirmedAddresses.map(addr => encodeURIComponent(addr)).join('/');
        allRoutesMapLink = `https://www.google.com/maps/dir/${encodedAddresses}`;
    }

    // Set header in A1:D1
    outputSheet.getRange(outputRow, 1, 1, 4)
        .merge()
        .setValue("Global Delivery Overview:")
        .setFontWeight("bold")
        .setBackground("#F0F0F0")
        .setVerticalAlignment("middle");
        
    const globalLinkCell = outputSheet.getRange(outputRow, 5); // Column E, Row 1

    if (allRoutesMapLink) {
        globalLinkCell.setFormula(`=HYPERLINK("${allRoutesMapLink}", "View ALL ${routes.length} Routes Map")`);
        globalLinkCell.setFontWeight("bold").setBackground("#B6D7A8").setHorizontalAlignment("center");
    } else {
        globalLinkCell.setValue("No Confirmed Stops Across All Routes").setFontWeight("bold").setBackground("#F4CCCC").setHorizontalAlignment("center");
    }

    outputRow++; // Now outputRow is 2, ready for the first route table.
    // --- End Global Rollup Link ---


    routes.forEach((route, routeIdx) => {
        const driver = routeDriverMap[route] || "TBD";
        // 1. Filter the already filtered dataRows by route.
        const rawRouteRows = dataRows.filter(r => String(r[routeColumnIndex]).trim() === route);
        
        // Collect confirmed addresses for map link generation
        const addresses = rawRouteRows
            .map(r => String(r[addressColumnIndex]).trim())
            .filter(a => a);

        let mapLink = "";
        if (addresses.length > 0) {
            // Generates a multi-stop directions link
            const encodedAddresses = addresses.map(addr => encodeURIComponent(addr)).join('/');
            mapLink = `https://www.google.com/maps/dir/${encodedAddresses}`;
        }
        
        // 2. Apply column filtering to get the final rows for output.
        const routeRows = rawRouteRows.map(r => includedIndexes.map(i => r[i]));

        const routeColor = routeColors[routeIdx % routeColors.length];
        
        // Remember the row before writing:
        const headerRowStart = outputRow;

        // Route header (Value in Col A, Shade full width) 
        const routeHeaderRange = outputSheet.getRange(outputRow, 1, 1, finalTableWidth);
        outputSheet.getRange(outputRow, 1).setValue(`Driver: ${driver} | Route: ${route}`);
        routeHeaderRange.setBackground(routeColor);
        outputRow++;

        // Table header (Rotation handled here)
        const tableHeaderRange = outputSheet.getRange(outputRow, 1, 1, filteredHeaders.length);
        tableHeaderRange.setValues([filteredHeaders]);
        tableHeaderRange.setFontWeight("bold")
                        .setBackground(routeColor)
                        .setVerticalAlignment("middle")
                        .setTextRotation(75);
        outputRow++;

        // Client rows with alternating colors
        const numDataRows = routeRows.length;
        // Padded rows now include padding up to the final table width
        const paddedRows = routeRows.map(r => { 
            let row = [...r]; 
            while (row.length < finalTableWidth) row.push(""); 
            return row; 
        });
        
        paddedRows.forEach((r, i) => {
            const rowRange = outputSheet.getRange(outputRow + i, 1, 1, finalTableWidth);
            rowRange.setValues([r]);
            rowRange.setBackground(i % 2 === 0 ? "#FFFFFF" : "#F0F0F0"); 
        });
        outputRow += numDataRows;

        // Apply borders to the table (Header + Client Rows + Totals Row)
        const tableRowsStart = headerRowStart + 1; // Row *after* route header (The table header)
        const tableRows = numDataRows + 2; // 1 header row + client rows + 1 totals row
        const borderRangeStart = tableRowsStart;
        
        // Packs/Food Border
        if (packsStart !== 0 && packsEnd !== 0) {
            outputSheet.getRange(borderRangeStart, packsStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle);
            outputSheet.getRange(borderRangeStart, packsEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }

        // Pet Food Border
        if (petFoodStart !== 0 && petFoodEnd !== 0) {
            outputSheet.getRange(borderRangeStart, petFoodStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle);
            outputSheet.getRange(borderRangeStart, petFoodEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }

        // Diapers/Wipes Border
        if (diapersStart !== 0 && diapersEnd !== 0) {
            outputSheet.getRange(borderRangeStart, diapersStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle);
            outputSheet.getRange(borderRangeStart, diapersEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }

        // Hygiene/Toiletries Border
        if (hygieneStart !== 0 && hygieneEnd !== 0) {
            outputSheet.getRange(borderRangeStart, hygieneStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle);
            outputSheet.getRange(borderRangeStart, hygieneEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }
        // End Borders


        // Totals row
        const totalsRow = Array(finalTableWidth).fill(""); // Totals row is initialized to finalTableWidth
        totalsRow[0] = "Totals";
        
        // Totals for numeric columns
        // This is safe because filteredIdx (index in filteredHeaders) must be less than finalTableWidth
        numericColumns.forEach(filteredIdx => {
            totalsRow[filteredIdx] = routeRows.reduce((sum, r) => sum + (parseFloat(r[filteredIdx]) || 0), 0);
        });
        
        // Shade the Totals row (up to finalTableWidth)
        const totalsRange = outputSheet.getRange(outputRow, 1, 1, finalTableWidth);
        totalsRange
            .setValues([totalsRow])
            .setBackground(routeColor)
            .setFontWeight("bold");

        // Set the Map Link cell in Column O (15)
        const mapLinkCell = outputSheet.getRange(outputRow, mapLinkTargetCol); 
        if (mapLink) {
            mapLinkCell.setFormula(`=HYPERLINK("${mapLink}", "View Route Map")`);
            mapLinkCell.setHorizontalAlignment("center");
        } else {
            mapLinkCell.setValue("No Confirmed Stops");
            mapLinkCell.setHorizontalAlignment("center");
        }

        outputRow++;

        // Diapers & Pull-Ups summary row
        const diaperSummaryMap = {};
        const pullupSummaryMap = {};
        routeRows.forEach(r => diaperColumns.forEach(idx => {
            let val = r[idx];
            if (val) {
                val = val.toString().trim();
                if (val.toUpperCase() === "N" || /^[1-6]$/.test(val)) diaperSummaryMap[val] = (diaperSummaryMap[val] || 0) + 1;
                else if (/^\d\/\dT$/.test(val) || /^\d\/\d\d?T$/.test(val)) pullupSummaryMap[val] = (pullupSummaryMap[val] || 0) + 1;
            }
        }));
        const diaperSummaryText = supportedDiaperSizes.filter(s => diaperSummaryMap[s]).map(s => `${s} (qty${diaperSummaryMap[s]})`).join(", ");
        const pullupSummaryText = supportedPullupSizes.filter(s => pullupSummaryMap[s]).map(s => `${s} (qty${pullupSummaryMap[s]})`).join(", ");
        const summaryRow = Array(finalTableWidth).fill(""); // Summary row is initialized to finalTableWidth
        summaryRow[0] = `Diapers: ${diaperSummaryText || "None"} | Pull-Ups: ${pullupSummaryText || "None"}`;

        // Shade the Summary row (up to finalTableWidth)
        outputSheet.getRange(outputRow, 1, 1, finalTableWidth)
            .setValues([summaryRow])
            .setBackground(routeColor);
        outputRow++;

        // Spacer row
        outputRow++;
    });

    // Autofit column widths
    const specialWideCols = ["Special Item Requests", "Special Delivery Instructions", "Address"];
    
    // Get 0-based indices for special wide columns
    const specialWideIndexes = specialWideCols
        .map(name => filteredHeaders.indexOf(name))
        .filter(i => i !== -1);
        
    // Handle special wide + wrapped columns
    specialWideIndexes.forEach(idx => {
        outputSheet.setColumnWidth(idx + 1, 36 * 7); // ~36 chars wide
        // Data starts from row 2 (header row)
        outputSheet.getRange(2, idx + 1, outputSheet.getLastRow() - 1, 1).setWrap(true);
    });
    
    // Set fixed width for the new map link column O (15)
    outputSheet.setColumnWidth(mapLinkTargetCol, 150);
    
    const diaperSummaryIndex = 0; // Column A (1)

    // Autosize all other columns up to finalTableWidth
    for (let c = 1; c <= finalTableWidth; c++) {
        const isSpecialWideCol = specialWideIndexes.includes(c - 1);
        const isMapLinkCol = c === mapLinkTargetCol;
        const isSummaryCol = c === diaperSummaryIndex + 1; 
        
        if (!isSpecialWideCol && !isMapLinkCol && !isSummaryCol) {
            outputSheet.autoResizeColumn(c);
        }
    }

    SpreadsheetApp.flush();
}



function generateAndEmailDriverPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  // 1. Identify the sheets to include in the PDF (e.g., sheets containing "Route")
  const sheetsToExport = allSheets.filter(sheet => {
    const name = sheet.getName();
    // Exclude the source data and helper sheets, include all generated route sheets
    return name.includes("Route") && 
           !name.includes("Data") && 
           !name.includes("HelpTable");
  });
  
  if (sheetsToExport.length === 0) {
    SpreadsheetApp.getUi().alert("No driver sheets found to export.");
    return;
  }

  // 2. Generate the PDF Blob
  const ssId = ss.getId();
  
  // Construct the URL parameters for the sheets to include
  const sheetIds = sheetsToExport.map(s => `&gid=${s.getSheetId()}`).join('');
  
  // URL for exporting multiple GIDs (sheets) as a single PDF.
  // We specify landscape (portrait=false) and fit to width (fitw=true).
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?exportFormat=pdf&format=pdf` +
              `&size=A4&portrait=false&fitw=true` + 
              `&fzr=true` + // Repeat frozen rows (if you freeze the header row)
              `${sheetIds}`;

  const token = ScriptApp.getOAuthToken();
  
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`
    },
    muteHttpExceptions: true
  });
  
  const pdfBlob = response.getBlob().setName(`DeliveryTables_${new Date().toLocaleDateString()}.pdf`);
  
  // 3. Email the PDF
  const userEmail = Session.getActiveUser().getEmail(); 
  MailApp.sendEmail({
    to: userEmail,
    subject: "Automated Delivery Tables PDF",
    body: "Attached is the consolidated PDF for all delivery routes. Each driver's table is guaranteed to start on a new page.",
    attachments: [pdfBlob]
  });

  SpreadsheetApp.getUi().alert(`Consolidated Delivery Tables PDF emailed to ${userEmail}.`);
}

function generateSpecialItemsCoordinator() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Data1");
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

    // Identify key indexes
    const routeIdx = headers.indexOf("routeDescription");
    const driverIdx = headers.indexOf("driver");
    
    // Locate the "Conf for" column (which starts with "Conf for ")
    const confForHeader = headers.find(h => typeof h === 'string' && h.trim().startsWith("Conf for"));
    const confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;

    // Columns for coordinator view
    const includedCols = [
        "driver",
        "firstName",
        "# packs",
        "# eggs",
        "# milk",
        "dogFood(qty)",
        "catFood(qty)",
        "diapers1(size)",
        "diapers2(size)",
        "diapers3(size)",
        "wipes(qty)",
        "toiletries(qty)",
        "fem hygiene(qty)",
        "Special Item Requests"
    ];
    const includedIndexes = includedCols.map(c => headers.indexOf(c)).filter(i => i !== -1);

    // Original data sheet indices needed for totaling
    const packsOrigIdx = headers.indexOf("# packs");
    const eggsOrigIdx = headers.indexOf("# eggs");
    const milkOrigIdx = headers.indexOf("# milk");
    const dogFoodOrigIdx = headers.indexOf("dogFood(qty)");
    const catFoodOrigIdx = headers.indexOf("catFood(qty)");
    const wipesOrigIdx = headers.indexOf("wipes(qty)");
    const toiletriesOrigIdx = headers.indexOf("toiletries(qty)");
    const femHygieneOrigIdx = headers.indexOf("fem hygiene(qty)");

    // Build route-driver mapping from helper table
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

    // --- NEW: Filter data to include only confirmed deliveries ('YES') for accurate ITEM TOTALS ---
    let confirmedDataRows = dataRows;
    if (confForIdx !== -1) {
        confirmedDataRows = dataRows.filter(row => {
            const confValue = String(row[confForIdx] || '').toUpperCase().trim();
            return confValue === 'YES';
        });
    }
    // ------------------------------------------------------------------------------------------

    // === Top Summary Table (C1:D4 for items) ===
    // Use confirmedDataRows for the total items needed for confirmed deliveries
    const totalPacks = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[packsOrigIdx]) || 0), 0);
    const totalEggs = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[eggsOrigIdx]) || 0), 0);
    const totalMilk = confirmedDataRows.reduce((s, r) => s + (parseFloat(String(r[milkOrigIdx]).replace('*', '') || 0)) || 0, 0);

    const summaryStartCol = 3; // Column C

    outputSheet.getRange(1, summaryStartCol, 1, 2).setValues([["Item", "Total (Confirmed Only)"]]); // Changed header for clarity
    outputSheet.getRange(2, summaryStartCol, 3, 2).setValues([
        ["# packs", totalPacks],
        ["# eggs", totalEggs],
        ["# milk", totalMilk]
    ]);
    outputSheet.getRange(1, summaryStartCol, 4, 2)
        .setFontWeight("bold")
        .setBackground("#D9EAD3")
        .setBorder(true, true, true, true, null, null);

    // === "Conf for" Summary Table (F1:G4) ===
    // This still uses the *unfiltered* dataRows to count the status of all rows.
    let confForCounts = { "NO ANS": 0, "NO": 0, "YES": 0, "Total": 0 };
    if (confForIdx !== -1) {
        dataRows.forEach(row => {
            // Trim and uppercase the value for reliable counting
            let val = String(row[confForIdx] || '').toUpperCase().trim();
            if (val.includes("NO ANS")) {
                confForCounts["NO ANS"]++;
            } else if (val === "NO") {
                confForCounts["NO"]++;
            } else if (val === "YES") {
                confForCounts["YES"]++;
            }
            // Only count if there's any value in the cell (not empty string)
            if (val) {
                 confForCounts["Total"]++;
            }
        });
    }

    const confSummaryStartCol = 6; // Column F (F1:G4)

    outputSheet.getRange(1, confSummaryStartCol, 1, 2).setValues([["Confirmation Status", "Count"]]);
    outputSheet.getRange(2, confSummaryStartCol, 4, 2).setValues([
        ["No Ans", confForCounts["NO ANS"]],
        ["No", confForCounts["NO"]],
        ["Yes", confForCounts["YES"]],
        ["Total", confForCounts["Total"]]
    ]);
    outputSheet.getRange(1, confSummaryStartCol, 4, 2)
        .setFontWeight("bold")
        .setBackground("#D9D2E9") // Light Purple for distinction
        .setBorder(true, true, true, true, null, null);

    let outputRow = 6; // leave gap before detailed tables

    // === Route sections ===
    const routes = [...new Set(dataRows.map(r => String(r[routeIdx]).trim()).filter(r => r))]; // ensure routes are trimmed for consistency
    const routeColors = [
        "#D9EAD3", // Light Green (Original)
        "#FFF2CC", // Light Yellow (Original)
        "#FCE5CD", // Light Orange (Original)
        "#D9D2E9", // Light Purple (Original)
        "#CFE2F3", // Light Blue (Original)
        "#F4CCCC", // Light Red/Pink (Original)
        "#EAD1DC", // Dusty Rose
        "#E0F7FA", // Light Cyan/Aqua
        "#F8BBD0", // Pale Pink
        "#B2EBF2", // Sky Blue
        "#E8F5E9", // Pale Mint
        "#FFF9C4", // Creamy Yellow
        "#FFCCBC", // Pale Peach
        "#C5CAE9", // Pale Lavender
        "#DCEDC8" // Soft Lime
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
        
        // Declare checkStart once at the beginning of the loop scope to avoid the redeclaration error
        const checkStart = colIndex("Packed up & labeled?");
        
        // === Data filtering (Filter by Route AND Confirmed 'YES') ===
        const filteredDataRows = dataRows
            .filter(r => String(r[routeIdx]).trim() === route) // Filter by route
            .filter(r => { 
                if (confForIdx === -1) return true; // If no conf column, include all
                const confValue = String(r[confForIdx] || '').toUpperCase().trim();
                // Only include rows where Conf for is 'YES'.
                return confValue === 'YES'; 
            });

        // Map filtered data rows to the final output columns
        const rows = filteredDataRows
            .map(r => includedIndexesNoDriver.map(i => r[i]).concat(["", ""]));

        if (rows.length === 0) {
            // If no confirmed rows, skip generating table for this route, but still add a separator space.
            outputRow++;
            return;
        }

        // === Totals Calculation (NEW) ===
        const totals = {
            packs: 0, eggs: 0, milk: 0, dogFood: 0, catFood: 0, wipes: 0, toiletries: 0, femHygiene: 0
        };

        filteredDataRows.forEach(r => {
            // Use String().replace() to handle potential '*' in milk or other cells if present.
            totals.packs += parseFloat(r[packsOrigIdx] || 0) || 0;
            totals.eggs += parseFloat(r[eggsOrigIdx] || 0) || 0;
            totals.milk += parseFloat(String(r[milkOrigIdx]).replace('*', '') || 0) || 0;
            totals.dogFood += parseFloat(r[dogFoodOrigIdx] || 0) || 0;
            totals.catFood += parseFloat(r[catFoodOrigIdx] || 0) || 0;
            totals.wipes += parseFloat(r[wipesOrigIdx] || 0) || 0;
            totals.toiletries += parseFloat(r[toiletriesOrigIdx] || 0) || 0;
            totals.femHygiene += parseFloat(r[femHygieneOrigIdx] || 0) || 0;
        });

        // === Route header (start at col A but without width expansion) ===
        const startCol = 1; 
        const mergeCols = finalHeaders.length - (startCol - 1);
        const headerRange = outputSheet.getRange(outputRow, startCol, 1, mergeCols);

        // Temporarily turn off auto-resize & lock width
        const colAWidth = outputSheet.getColumnWidth(1);
        outputSheet.getRange(outputRow, startCol)
            .setValue(`Driver: ${driver} | Route: ${route}`)
            .setFontWeight("bold")
            .setBackground(routeColor)
            .setHorizontalAlignment("left")
            .setWrap(false)
            .setFontSize(12)
            .setFontWeight("bold");

        headerRange.mergeAcross(); // allow text overflow across merged cells
        outputSheet.setColumnWidth(1, colAWidth); // restore width

        outputRow++;

        // === Column headers ===
        outputSheet.getRange(outputRow, 1, 1, finalHeaders.length)
            .setValues([finalHeaders])
            .setFontWeight("bold")
            .setBackground(routeColor)
            .setTextRotation(75)
            .setVerticalAlignment("middle");
        outputRow++;

        // === Data rows ===
        const range = outputSheet.getRange(outputRow, 1, rows.length, finalHeaders.length);
        range.setValues(rows);

        // Alternate row colors
        for (let j = 0; j < rows.length; j++) {
            outputSheet.getRange(outputRow + j, 1, 1, finalHeaders.length)
                .setBackground(j % 2 === 0 ? "#FFFFFF" : "#F0F0F0");
        }

        // Add checkboxes only to the data rows
        const packedCol = includedIndexesNoDriver.length + 1;
        outputSheet.getRange(outputRow, packedCol, rows.length, 2).insertCheckboxes();

        outputRow += rows.length; // outputRow is now the start of the Totals Row

        // === Totals Row (Shaded and Filled with Data) ===
        const totalsRowData = new Array(finalHeaders.length).fill("");
        const packsCol = colIndex("# packs") - 1; // 0-based index for array placement
        const eggsCol = colIndex("# eggs") - 1;
        const milkCol = colIndex("# milk") - 1;
        const dogFoodCol = colIndex("dogFood(qty)") - 1;
        const catFoodCol = colIndex("catFood(qty)") - 1;
        const wipesCol = colIndex("wipes(qty)") - 1;
        const toiletriesCol = colIndex("toiletries(qty)") - 1;
        const femHygieneCol = colIndex("fem hygiene(qty)") - 1;

        totalsRowData[0] = "TOTALS"; // firstName column
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
        outputSheet.getRange(outputRow, 1).setHorizontalAlignment("left"); // Keep "TOTALS" left-aligned

        // Ensure residual checkboxes are removed from the totals row
        // const checkStart = colIndex("Packed up & labeled?"); // Removed redeclaration
        if (checkStart > 0) {
            outputSheet.getRange(outputRow, checkStart, 1, 2).clearDataValidations();
        }

        outputRow++; // outputRow is now the blank spacer row (Spacer Row 2)

        // === Spacer Row 2 (Unshaded) ===
        const unshadedSpacerRange = outputSheet.getRange(outputRow, 1, 1, finalHeaders.length);
        unshadedSpacerRange.clearContent();
        unshadedSpacerRange.clearFormat(); 

        // Ensure residual checkboxes are removed
        if (checkStart > 0) {
            outputSheet.getRange(outputRow, checkStart, 1, 2).clearDataValidations();
        }
        
        outputRow++; // Advance by 1 (unshaded spacer)


        // === Borders (Apply borders to the Header + Data Rows + Totals Row) ===
        const headerRow = outputRow - rows.length - 2; // Start row for borders (Header)
        const borderRangeStart = headerRow;
        const totalRowsForBorder = rows.length + 2; // Header (1) + Data (N) + Totals (1)

        const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
        
        // NEW BORDER: Between # milk and dogFood(qty)
        const milkColEnd = colIndex("# milk");
        if (milkColEnd > 0) {
             outputSheet.getRange(borderRangeStart, milkColEnd, totalRowsForBorder, 1)
                 .setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }

        // ORIGINAL BORDERS (updated to use totalRowsForBorder):
        const packsStart = colIndex("# packs");
        const packsEnd = colIndex("catFood(qty)");
        if (packsStart > 0 && packsEnd > 0) {
            outputSheet.getRange(borderRangeStart, packsStart, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", borderStyle);
            outputSheet.getRange(borderRangeStart, packsEnd, totalRowsForBorder, 1)
                .setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }

        const diapersStart = colIndex("diapers1(size)");
        const diapersEnd = colIndex("wipes(qty)");
        if (diapersStart > 0 && diapersEnd > 0) {
            outputSheet.getRange(borderRangeStart, diapersStart, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", borderStyle);
            outputSheet.getRange(borderRangeStart, diapersEnd, totalRowsForBorder, 1)
                .setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }

        const hygieneStart = colIndex("toiletries(qty)");
        const hygieneEnd = colIndex("fem hygiene(qty)");
        if (hygieneStart > 0 && hygieneEnd > 0) {
            outputSheet.getRange(borderRangeStart, hygieneStart, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", borderStyle);
            outputSheet.getRange(borderRangeStart, hygieneEnd, totalRowsForBorder, 1)
                .setBorder(null, null, null, true, null, null, "#000000", borderStyle);
        }

        const specialIdx = colIndex("Special Item Requests");
        if (specialIdx > 0) {
            outputSheet.getRange(borderRangeStart, specialIdx, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.DOTTED);
        }

        // const checkStart = colIndex("Packed up & labeled?"); // Removed redeclaration
        if (checkStart > 0) {
            outputSheet.getRange(borderRangeStart, checkStart, totalRowsForBorder, 1)
                .setBorder(null, true, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.DOTTED);
        }
    });

    // === Column sizing ===
    // Auto resize all columns first
    outputSheet.autoResizeColumns(1, finalHeaders.length);

    const specialIdx = finalHeaders.indexOf("Special Item Requests");
    if (specialIdx !== -1) {
        const col = specialIdx + 1;
        outputSheet.setColumnWidth(col, 250);
        // Only apply wrap to the data rows and below (skip top summary tables)
        outputSheet.getRange(6, col, outputSheet.getLastRow() - 5, 1).setWrap(true);
    }

    const checkboxStartCol = finalHeaders.indexOf("Packed up & labeled?");
    if (checkboxStartCol !== -1) {
        const col1 = checkboxStartCol + 1;
        const col2 = checkboxStartCol + 2;
        outputSheet.setColumnWidth(col1, 140);
        outputSheet.setColumnWidth(col2, 140);
    }


    SpreadsheetApp.flush();
}



// This function called by generateDeliveryTables
function writeRowsInChunks(sheet, startRow, dataRows, numCols, chunkSize = 50) {
  for (let i = 0; i < dataRows.length; i += chunkSize) {
    const chunk = dataRows.slice(i, i + chunkSize);
    sheet.getRange(startRow + i, 1, chunk.length, numCols).setValues(chunk);
  }
}