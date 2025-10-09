function myFunction() {
  var range = SpreadsheetApp.getActive().getActiveSheet().getActiveRange();
  var values = range.getValues();

  const urlBeginning = "https://maps.googleapis.com/maps/api/staticmap?size=640x640";
  let urlMiddle = ""
  const urlEnd = "&key=AIzaSyDUmU95JYQAEuQw0vixU3SUKm7qZqaXahA";

  values.forEach(function (row) {
    var label = row[0];
    var address = row[1].split(" ").join("+");
    const stringRow = "&markers=label:" + label + "%7C" + address + "%22";
    urlMiddle += stringRow;
  });

  const finalUrl = urlBeginning + urlMiddle + urlEnd;
  showURL("Neighbors", finalUrl);
}

function showURL(name, url){
  var html = '<html><body><img style="width: 100%; height: auto;" src="'+url+'" /><a href="'+ url +'" target="_blank">Open in new tab</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html).setWidth(640).setHeight(670);
  SpreadsheetApp.getUi().showModelessDialog(ui,name);
}

function writeRowsInChunks(sheet, startRow, dataRows, numCols, chunkSize = 50) {
  for (let i = 0; i < dataRows.length; i += chunkSize) {
    const chunk = dataRows.slice(i, i + chunkSize);
    sheet.getRange(startRow + i, 1, chunk.length, numCols).setValues(chunk);
  }
}

function generateDeliveryTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data1");
  const outputSheet = ss.getSheetByName("Deliveries");
  outputSheet.clearContents();
  outputSheet.clearFormats();

  const data = dataSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("No data found in 'Data' sheet.");
    return;
  }

  const headers = data[0];
  let dataRows = data.slice(1);

  // Identify Key Columns and Extract Route/Driver Data ---
  const routeColumnIndex = headers.indexOf("routeDescription");
  const driverColumnIndex = headers.indexOf("driver");

  // Get the route/driver map data BEFORE filtering the dataRows
  const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
  const helperData = helperSheet.getDataRange().getValues().slice(1);
  const routeDriverMap = {};
  helperData.forEach(row => { if (row[0]) routeDriverMap[row[0]] = row[1]; });

  // Get all unique routes from the original data (must use original column index)
  // Ensure the index is valid before trying to map
  const routes = (routeColumnIndex !== -1)
    ? [...new Set(dataRows.map(r => r[routeColumnIndex]).filter(r => r))]
    : [];
  let outputRow = 1;

  // Filter out unwanted columns ---
  const columnsToAlwaysExclude = ["routeDescription", "driver", "N, S, E, Central", "Source", "methodWipesEntry"];

  const excludedColumns = headers.filter(header => {
    // Exclude headers that are empty or contain only whitespace (the unnamed columns)
    if (typeof header === 'string' && header.trim() === "") {
        return true;
    }

    // Exclude any column that starts with "Conf for"
     if (typeof header === 'string' && header.trim().startsWith("Conf for")) {
      return true;
    }
    
    // Check against the always exclude list
    if (columnsToAlwaysExclude.includes(header)) {
      return true;
    }
    return false;
  });


  const includedIndexes = headers.map((h, i) => excludedColumns.includes(h) ? -1 : i).filter(i => i !== -1);
  const filteredHeaders = includedIndexes.map(i => headers[i]);

  const numericColumns = filteredHeaders
    .map((h, idx) => dataRows.some(r => !isNaN(parseFloat(r[idx])) && r[idx] !== "") ? idx : -1)
    .filter(idx => idx !== -1);

  const diaperColumns = ["diapers1(size)", "diapers2(size)", "diapers3(size)"]
    .map(c => filteredHeaders.indexOf(c))
    .filter(idx => idx !== -1);

  const supportedDiaperSizes = ["N","1","2","3","4","5","6"];
  const supportedPullupSizes = ["2/3T","3/4T","4/5T","5/6T"];

  const routeColors = ["#D9EAD3","#FFF2CC","#FCE5CD","#D9D2E9","#CFE2F3","#F4CCCC"]; // rotate if more routes

  // --- Column Index Lookups for Borders (1-based index) ---
  // The indices are based on the filteredHeaders array.
  const packsStart = filteredHeaders.indexOf("# packs") + 1;
  const packsEnd = filteredHeaders.indexOf("# milk") + 1;

  const petFoodStart = filteredHeaders.indexOf("dogFood(qty)") + 1;
  const petFoodEnd = filteredHeaders.indexOf("catFood(qty)") + 1;

  const diapersStart = filteredHeaders.indexOf("diapers1(size)") + 1;
  const diapersEnd = filteredHeaders.indexOf("wipes(qty)") + 1;

  const hygieneStart = filteredHeaders.indexOf("toiletries(qty)") + 1;
  const hygieneEnd = filteredHeaders.indexOf("fem hygiene(qty)") + 1;

  const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM; // Used for "2pt" border


  routes.forEach((route, routeIdx) => {
    const driver = routeDriverMap[route] || "TBD";
    // 1. Filter the UNFILTERED dataRows by route, using the correct index.
    const rawRouteRows = dataRows.filter(r => r[routeColumnIndex] === route);
    
    // 2. NOW apply the column filtering to get the final rows for output.
    const routeRows = rawRouteRows.map(r => includedIndexes.map(i => r[i]));

    const routeColor = routeColors[routeIdx % routeColors.length];
    
    // Remember the row before writing:
    const headerRowStart = outputRow;

// --- Route header (FIXED: Value in Col A, Shade full width) ---
// 1. Define the full range for shading
const routeHeaderRange = outputSheet.getRange(outputRow, 1, 1, filteredHeaders.length);

// 2. Define the single cell for the text
const routeHeaderCell = outputSheet.getRange(outputRow, 1);

// Set value ONLY in the first cell (allows text to spill over)
routeHeaderCell.setValue(`Driver: ${driver} | Route: ${route}`); 

// Apply background across the full range
routeHeaderRange.setBackground(routeColor); 

// ⭐️ FINAL, ROBUST PAGE BREAK: This function is often more reliable
// in environments where setPageBreak() fails.
// if (outputRow > 1) { 
//     outputSheet.insertPageBreak(outputRow);
// }

outputRow++; // Correctly moves to the table header row (only one increment here!)


    // --- Table header (Rotation handled here) ---
    const tableHeaderRange = outputSheet.getRange(outputRow, 1, 1, filteredHeaders.length);
    tableHeaderRange.setValues([filteredHeaders]);
    tableHeaderRange.setFontWeight("bold")
                    .setBackground(routeColor)
                    .setVerticalAlignment("middle")
                    .setTextRotation(75); // <-- Rotation already correctly implemented
    outputRow++;

    // --- Client rows with alternating colors ---
    const numDataRows = routeRows.length; // Use this to calculate range length
    const paddedRows = routeRows.map(r => { let row = [...r]; while (row.length < filteredHeaders.length) row.push(""); return row; });
    paddedRows.forEach((r, i) => {
      const rowRange = outputSheet.getRange(outputRow + i, 1, 1, filteredHeaders.length);
      rowRange.setValues([r]);
      rowRange.setBackground(i % 2 === 0 ? "#FFFFFF" : "#F0F0F0"); // white / light gray
    });
    outputRow += numDataRows;

    // --- Apply THICK BORDERS to the table (Header + Client Rows) ---
    const tableRowsStart = headerRowStart + 1; // Row *after* route header
    const tableRows = numDataRows + 1; // 1 header row + client rows

    const borderRangeStart = tableRowsStart;
    const borderRangeEnd = tableRowsStart + tableRows - 1;

    // Packs/Food Border
    outputSheet.getRange(borderRangeStart, packsStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle); // Left of packsStart
    outputSheet.getRange(borderRangeStart, packsEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle); // Right of packsEnd

    // Pet Food Border
    outputSheet.getRange(borderRangeStart, petFoodStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle); // Left of petFoodStart
    outputSheet.getRange(borderRangeStart, petFoodEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle); // Right of petFoodEnd

    // Diapers/Wipes Border
    outputSheet.getRange(borderRangeStart, diapersStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle); // Left of diapersStart
    outputSheet.getRange(borderRangeStart, diapersEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle); // Right of diapersEnd

    // Hygiene/Toiletries Border
    outputSheet.getRange(borderRangeStart, hygieneStart, tableRows, 1).setBorder(null, true, null, null, null, null, "#000000", borderStyle); // Left of hygieneStart
    outputSheet.getRange(borderRangeStart, hygieneEnd, tableRows, 1).setBorder(null, null, null, true, null, null, "#000000", borderStyle); // Right of hygieneEnd
    // --- End Borders ---


    // --- Totals row ---
    const totalsRow = Array(filteredHeaders.length).fill("");
    totalsRow[0] = "Totals";
    numericColumns.forEach(idx => { totalsRow[idx] = routeRows.reduce((sum, r) => sum + (parseFloat(r[idx]) || 0), 0); });
    outputSheet.getRange(outputRow, 1, 1, filteredHeaders.length).setValues([totalsRow]);
    outputRow++;

    // --- Diapers & Pull-Ups summary row ---
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
    const summaryRow = Array(filteredHeaders.length).fill("");
    summaryRow[0] = `Diapers: ${diaperSummaryText || "None"} | Pull-Ups: ${pullupSummaryText || "None"}`;
    outputSheet.getRange(outputRow, 1, 1, filteredHeaders.length).setValues([summaryRow]);
    outputRow++;

    // --- Spacer row ---
    outputRow++;
  });

  // --- Autofit column widths (Existing logic maintained) ---
  const specialWideCols = ["Special Item Requests", "Special Delivery Instructions"];
  const diaperSummaryCols = ["Diapers:"]; // we'll match by header start

  // Find column indexes
  const specialWideIndexes = specialWideCols.map(name => filteredHeaders.indexOf(name)).filter(i => i !== -1);
  const diaperSummaryIndex = 0; // summary always written into col A

  // Handle special wide + wrapped columns
  specialWideIndexes.forEach(idx => {
    outputSheet.setColumnWidth(idx + 1, 36 * 7); // ~36 chars wide
    outputSheet.getRange(2, idx + 1, outputSheet.getLastRow() - 1, 1).setWrap(true);
  });

  // Autosize all other columns except diaper summary col
  for (let c = 1; c <= filteredHeaders.length; c++) {
    if (!specialWideIndexes.includes(c - 1) && c !== diaperSummaryIndex + 1) {
      outputSheet.autoResizeColumn(c);
    }
  }

  SpreadsheetApp.flush();
}