function getDataAndHeaders(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet '${sheetName}' not found`);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error(`No data found in '${sheetName}'`);
  return [data[0], data.slice(1)];
}

function getConfirmedRows(headers, rows) {
  const confIdx = headers.findIndex(h => typeof h === 'string' && h.trim().startsWith("Conf for"));
  if (confIdx === -1) return rows;
  return rows.filter(r => String(r[confIdx] || '').toUpperCase().trim() === 'YES');
}

function getRouteDriverMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
  if (!helperSheet) return {};
  const helperData = helperSheet.getDataRange().getValues().slice(1);
  const map = {};
  helperData.forEach(([route, driver]) => {
    const key = String(route || '').trim();
    if (key) map[key] = String(driver || '').trim();
  });
  return map;
}

function getUniqueRoutes(rows, routeColIdx) {
  return [...new Set(
    rows.map(r => String(r[routeColIdx]).trim()).filter(Boolean)
  )];
}

function makeMapLink(addresses) {
  if (!addresses || addresses.length === 0) return '';
  const encoded = addresses.map(a => encodeURIComponent(a)).join('/');
  return `https://www.google.com/maps/dir/${encoded}`;
}

function getRouteColors() {
  return [
    "#D9EAD3", "#FFF2CC", "#FCE5CD", "#D9D2E9", "#CFE2F3", "#F4CCCC", 
    "#EAD1DC", "#E0F7FA", "#F8BBD0", "#B2EBF2", "#E8F5E9", "#FFF9C4", 
    "#FFCCBC", "#C5CAE9", "#DCEDC8"
  ];
}

function writeRowsInChunks(sheet, startRow, dataRows, numCols, chunkSize = 50) {
  for (let i = 0; i < dataRows.length; i += chunkSize) {
    const chunk = dataRows.slice(i, i + chunkSize);
    sheet.getRange(startRow + i, 1, chunk.length, numCols).setValues(chunk);
  }
}

function generateDeliveryTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Data1";
  const outputName = "Deliveries";

  const outputSheet = ss.getSheetByName(outputName) || ss.insertSheet(outputName);
  outputSheet.clearContents().clearFormats();

  const [headers, allRows] = getDataAndHeaders(sheetName);
  const confirmedRows = getConfirmedRows(headers, allRows);
  const routeDriverMap = getRouteDriverMap();
  const routeIdx = headers.indexOf("routeDescription");
  const addressIdx = headers.indexOf("Address");

  if (routeIdx === -1) {
    SpreadsheetApp.getUi().alert("No 'routeDescription' column found.");
    return;
  }

  const routes = getUniqueRoutes(confirmedRows, routeIdx);
  const colors = getRouteColors();

  if (routes.length === 0) {
    SpreadsheetApp.getUi().alert("No confirmed deliveries found.");
    return;
  }

  let currentRow = 1;

  routes.forEach((route, i) => {
    const driver = routeDriverMap[route] || "Unassigned";
    const routeColor = colors[i % colors.length];

    // Header
    outputSheet.getRange(currentRow, 1, 1, 3).merge();
    outputSheet.getRange(currentRow, 1).setValue(`${driver} – ${route}`);
    outputSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground(routeColor);
    currentRow++;

    // Rows for this route
    const routeRows = confirmedRows.filter(r => String(r[routeIdx]).trim() === route);
    if (routeRows.length > 0) {
      const routeData = routeRows.map(r => {
        const name = r[headers.indexOf("Name")] || "";
        const addr = r[addressIdx] || "";
        const phone = r[headers.indexOf("Phone")] || "";
        return [name, addr, phone];
      });

      writeRowsInChunks(outputSheet, currentRow, routeData, 3);
      currentRow += routeData.length;

      // Map link row
      const addresses = routeRows.map(r => r[addressIdx]).filter(Boolean);
      const mapLink = makeMapLink(addresses);
      outputSheet.getRange(currentRow, 1, 1, 3).merge();
      outputSheet.getRange(currentRow, 1)
        .setValue(`Map: ${mapLink}`)
        .setFontSize(10)
        .setFontStyle("italic")
        .setFontColor("#1155CC")
        .setBackground("#F3F3F3");
      currentRow += 2;
    }
  });

  outputSheet.autoResizeColumns(1, 3);
  outputSheet.setFrozenRows(0);
}

function generateSpecialItemsCoordinator() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Data1";
  const outputName = "Special Items Coordinator";

  const outputSheet = ss.getSheetByName(outputName) || ss.insertSheet(outputName);
  outputSheet.clearContents().clearFormats();

  const [headers, allRows] = getDataAndHeaders(sheetName);
  const confirmedRows = getConfirmedRows(headers, allRows);
  const routeDriverMap = getRouteDriverMap();
  const routeIdx = headers.indexOf("routeDescription");
  const specialItemIdx = headers.indexOf("Special Item Request");

  if (routeIdx === -1 || specialItemIdx === -1) {
    SpreadsheetApp.getUi().alert("Missing required columns for routes or special items.");
    return;
  }

  const routes = getUniqueRoutes(confirmedRows, routeIdx);
  const colors = getRouteColors();

  if (routes.length === 0) {
    SpreadsheetApp.getUi().alert("No confirmed routes found.");
    return;
  }

  // Title
  outputSheet.getRange("C1").setValue("Special Item Requests")
    .setFontSize(16).setFontWeight("bold").setFontColor("#333333");

  let currentRow = 3;

  routes.forEach((route, i) => {
    const driver = routeDriverMap[route] || "Unassigned";
    const routeColor = colors[i % colors.length];

    // Route header
    outputSheet.getRange(currentRow, 2, 1, 3).merge();
    outputSheet.getRange(currentRow, 2)
      .setValue(`${driver} – ${route}`)
      .setFontWeight("bold")
      .setBackground(routeColor);
    currentRow++;

    // Rows for this route
    const routeRows = confirmedRows.filter(r => String(r[routeIdx]).trim() === route);
    const specialRows = routeRows.filter(r => String(r[specialItemIdx]).trim() !== "");

    if (specialRows.length > 0) {
      const displayRows = specialRows.map(r => {
        const name = r[headers.indexOf("Name")] || "";
        const item = r[specialItemIdx] || "";
        const address = r[headers.indexOf("Address")] || "";
        return [name, item, address, ""];
      });

      // Header row for this section
      outputSheet.getRange(currentRow, 2, 1, 4).setValues([["Name", "Item", "Address", "✓"]]);
      outputSheet.getRange(currentRow, 2, 1, 4)
        .setFontWeight("bold").setBackground("#E0E0E0");
      currentRow++;

      // Data rows
      writeRowsInChunks(outputSheet, currentRow, displayRows, 4);
      currentRow += displayRows.length + 1;
    }
  });

  outputSheet.autoResizeColumns(2, 4);
  outputSheet.setColumnWidths(2, 1, 150); // narrow route header column
  outputSheet.setColumnWidths(3, 1, 200); // wider for item
  outputSheet.setColumnWidths(4, 1, 250); // wider for address
  outputSheet.setColumnWidths(5, 1, 40);  // checkbox col
}
