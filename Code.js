/**
 * Automatically rebuilds the DeliveriesHelpTable from the source data.
 * This ensures Driver names are always current before reports run.
 */
function refreshDeliveriesHelpTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
  const helperSheetName = "DeliveriesHelpTable";
  
  // 1. Get or Create the helper sheet
  let helperSheet = ss.getSheetByName(helperSheetName);
  if (!helperSheet) {
    helperSheet = ss.insertSheet(helperSheetName);
  }
  helperSheet.clear(); // Clear old data

  // 2. Load Source Data
  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const routeIdx = headers.indexOf("routeDescription");
  const driverIdx = headers.indexOf("Driver");

  if (routeIdx === -1 || driverIdx === -1) {
    console.error("Could not find 'routeDescription' or 'Driver' columns in source.");
    return;
  }

  // 3. Extract Unique Route -> Driver mapping
  const mapping = {};
  data.slice(1).forEach(row => {
    const route = String(row[routeIdx] || "").trim();
    const driver = String(row[driverIdx] || "").trim();
    
    // Only add if route isn't empty; update driver if one is found
    if (route && (!mapping[route] || mapping[route] === "TBD" || mapping[route] === "")) {
      mapping[route] = driver || "TBD";
    }
  });

  // 4. Convert to array and sort by route (alphabetical)
  const outputRows = [["routeDescription", "Driver"]];
  Object.keys(mapping).sort().forEach(route => {
    outputRows.push([route, mapping[route]]);
  });

  // 5. Write back to DeliveriesHelpTable
  helperSheet.getRange(1, 1, outputRows.length, 2).setValues(outputRows);
  helperSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#F0F0F0");
  helperSheet.autoResizeColumns(1, 2);
  
  console.log("DeliveriesHelpTable refreshed successfully.");
}

// New drivers tab functions

/***************
 Batch 1: Controller + Data Loading / Grouping Helpers
 (Drop-in replacement for the top of makeDriverRouteTabs)
***************/

function makeDriverRouteTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  refreshDeliveriesHelpTable();
  const DATA_SHEET_NAME = "Deliveries-UPDATE HERE";
  const INDEX_SHEET_NAME = "Driver_Deliveries";
  const ROUTE_SHEET_PREFIX = "Route_";
  // const helperWriter = "buildRouteSheet"; // the per-route writer we'll call (Batch 2 will replace with new functions)
  const helperWriter = "populateRouteTabs";

  Logger.log("makeDriverRouteTabs: start");

  // 1) cleanup old route sheets
  cleanupPreviousRouteSheets(ss, ROUTE_SHEET_PREFIX);

  // 2) load raw data + headers
  const { dataSheet, headers, dataRows } = loadDriverRouteData(ss, DATA_SHEET_NAME);
  if (!dataSheet || headers.length === 0) {
    SpreadsheetApp.getUi().alert(`Data sheet "${DATA_SHEET_NAME}" not found or empty.`);
    return;
  }
  if (dataRows.length === 0) {
    SpreadsheetApp.getUi().alert("No data rows found in the data sheet.");
    return;
  }

  // 3) apply confirmation filter (Conf for *)
  const confForHeader = headers.find(h => typeof h === 'string' && h.trim().startsWith("Conf for"));
  const confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;
  const filteredRows = filterDriverRouteRows(dataRows, confForIdx);

  // 4) helper mapping (route -> driver)
  const routeDriverMap = buildRouteDriverMap(ss, "DeliveriesHelpTable");

  // 5) group rows by route (object keyed by routeDescription)
  const grouped = groupRowsByRoute(filteredRows, headers, "routeDescription");

  // 6) gather route-level map links (pre-extraction for each route)
  const routeMapInfo = {};
  Object.keys(grouped).forEach(routeKey => {
    routeMapInfo[routeKey] = extractMapLinkForRoute({
      ss,
      dataSheet,
      headers,
      dataRows,          // full data array (including header-less rows)
      routeValue: routeKey,
      deliveryMapColumnHeader: "Delivery Map Links"
    });
  });

  // 7) orchestrate sheet creation using existing writer (populateRouteTabs)
  // We'll call populateRouteTabs per route. Batch 2 will replace this with new builders.
  const summaryRows = [["Route", "Driver", "Stops", "Open Sheet", "Map"]];
  const routeNames = Object.keys(grouped).sort();

  routeNames.forEach((route, idx) => {
    const safeName = route.replace(/[^a-zA-Z0-9]/g, "_").substring(0, 30);
    const sheetName = ROUTE_SHEET_PREFIX + safeName;
    const routeColor = pickRouteColor(idx); // small utility below

    const driverName = routeDriverMap[route] || "TBD";
    const personStops = grouped[route].personStops;

    // create new sheet
    const sheet = ss.insertSheet(sheetName);
    sheet.clearFormats();
    try { sheet.setTabColor(routeColor); } catch(e) { /* ignore colors if invalid */ }

    // map info for this route
    const mapInfo = routeMapInfo[route] || { link: "", label: "" };

    // Call the route writer (current code uses populateRouteTabs)
    // Signature we expect (matches your old function):
    // populateRouteTabs(ss, sheet, route, driver, personStops, itemMap, finalItemHeaders, headers, routeColor, allStopColors, colIndices, routeMapLink, routeMapLabel)
    // We'll re-create the minimal param set here from the old code so existing populateRouteTabs keeps working.
    try {
      // Build a small column index object used by the existing populateRouteTabs
      const colIndices = {
        nameCol: headers.indexOf("Name"),
        addressCol: headers.indexOf("Address"),
        phoneCol: headers.indexOf("Phone"),
        specCol: headers.indexOf("Special Item Requests"),
        delivCol: headers.indexOf("Special Delivery Instructions")
      };

      // Build a simple itemMap + finalItemHeaders consistent with current implementation
      const diapers1Idx = headers.indexOf("diapers1(size)");
      const diapers2Idx = headers.indexOf("diapers2(size)");
      const diapers3Idx = headers.indexOf("diapers3(size)");
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
      const finalItemHeaders = itemMap.map(i => i.label);

      // call the writer
      buildRouteSheet({
        ss,
        sheet,
        route,
        driverName,
        personStops,
        itemMap,
        finalItemHeaders,
        headers,
        routeColor,
        stopColors: pickStopColors(),
        colIndices,
        mapInfo
      });

      // Build index entry (with hyperlink to the sheet and auto map URL)
      const stopAddresses = Object.keys(personStops).map(key => key.split("|")[0]);
      const autoMapUrl = buildAutoMapUrl(stopAddresses);
      const mapFormula = autoMapUrl
        ? `=HYPERLINK("${autoMapUrl}","View Map")`
        : "No addresses";

      summaryRows.push([
        route,
        driverName,
        Object.keys(personStops).length,
        `=HYPERLINK("#gid=${sheet.getSheetId()}","Open")`,
        mapFormula
      ]);
    } catch (err) {
      Logger.log(`Error writing route ${route}: ${err}`);
      // remove the sheet if something failed to avoid half-created outputs
      try { ss.deleteSheet(sheet); } catch (e) {}
    }
  });

  // 8) Write index sheet
  writeRouteIndexSheet(ss, INDEX_SHEET_NAME, summaryRows);

  const finishedTimestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
  Logger.log(`makeDriverRouteTabs: finished at ${finishedTimestamp}`);
  return finishedTimestamp;
}

/* -------------------------
   Helper: cleanupPreviousRouteSheets
   - deletes sheets whose name starts with prefix
------------------------- */
function cleanupPreviousRouteSheets(ss, prefix) {
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    try {
      if (sh.getName().startsWith(prefix)) ss.deleteSheet(sh);
    } catch (e) {
      // ignore delete errors for protected or active sheets
      Logger.log(`cleanupPreviousRouteSheets: ${e}`);
    }
  });
}

/* -------------------------
   Helper: loadDriverRouteData
   - returns { dataSheet, headers, dataRows }
------------------------- */
function loadDriverRouteData(ss, sheetName) {
  const dataSheet = ss.getSheetByName(sheetName);
  if (!dataSheet) return { dataSheet: null, headers: [], dataRows: [] };

  const dataRange = dataSheet.getDataRange();
  const data = dataRange.getValues() || [];
  if (data.length === 0) return { dataSheet, headers: [], dataRows: [] };

  const headers = data[0].map(h => String(h).trim());
  const dataRows = data.slice(1);
  return { dataSheet, headers, dataRows };
}

/* -------------------------
   Helper: filterDriverRouteRows
   - keeps only rows with Conf for == YES or contains NO ANS
   - if confForIdx === -1, returns original rows
------------------------- */
function filterDriverRouteRows(dataRows, confForIdx) {
  if (confForIdx === -1) return dataRows.slice(); // no filtering column present

  const filtered = dataRows.filter(row => {
    const val = String(row[confForIdx] || "").toUpperCase().trim();
    return val === "YES" || val.includes("NO ANS");
  });
  return filtered;
}

/* -------------------------
   Helper: buildRouteDriverMap
   - reads DeliveriesHelpTable (route -> driver)
------------------------- */
function buildRouteDriverMap(ss, helperSheetName) {
  const map = {};
  const helper = ss.getSheetByName(helperSheetName);
  if (!helper) return map;
  const rows = helper.getDataRange().getValues().slice(1);
  rows.forEach(r => {
    const key = String(r[0] || "").trim();
    const drv = String(r[1] || "").trim();
    if (key) map[key] = drv;
  });
  return map;
}

/* -------------------------
   Helper: groupRowsByRoute
   - groups by headerName (routeDescription)
   - returns { route: { plainRows: [], personStops: { address|name: [rows] } } }
------------------------- */
function groupRowsByRoute(dataRows, headers, headerName) {
  const map = {};
  const routeIdx = headers.indexOf(headerName);
  const addressIdx = headers.indexOf("Address");
  const nameIdx = headers.indexOf("Name");

  dataRows.forEach(row => {
    const route = String(routeIdx !== -1 ? row[routeIdx] : "").trim();
    if (!route) return;
    if (!map[route]) map[route] = { plainRows: [], personStops: {}, firstRowKey: row.join() };
    map[route].plainRows.push(row);

    const address = String(addressIdx !== -1 ? row[addressIdx] : "").trim();
    const name = String(nameIdx !== -1 ? row[nameIdx] : "").trim();
    const personKey = address ? (address + "|" + name) : (name || Math.random().toString(36).slice(2,8));
    if (!map[route].personStops[personKey]) map[route].personStops[personKey] = [];
    map[route].personStops[personKey].push(row);
  });

  return map;
}

/* -------------------------
   Helper: extractMapLinkForRoute
   - Finds the first Delivery Map Links value for a route and attempts to extract:
     { link: urlOrEmpty, label: displayLabelOrDefault }
------------------------- */
function extractMapLinkForRoute(opts) {
  const { ss, dataSheet, headers, dataRows, routeValue, deliveryMapColumnHeader } = opts || {};
  if (!dataSheet || !headers || !dataRows) return { link: "", label: "" };

  const deliveryMapLinkIndex = headers.indexOf(deliveryMapColumnHeader);
  if (deliveryMapLinkIndex === -1) return { link: "", label: "" };

  // Find a row in the full data table that belongs to this route
  const routeIdx = headers.indexOf("routeDescription");
  if (routeIdx === -1) return { link: "", label: "" };

  // Locate first matching row in the original dataRows array
  const matchRow = dataRows.find(r => String(r[routeIdx]).trim() === routeValue);
  if (!matchRow) return { link: "", label: "" };

  // Find its 1-based position in the sheet (we need to search the sheet rows to get Range object)
  const stringified = matchRow.join();
  const allValues = dataSheet.getDataRange().getValues();
  const foundIndex = allValues.findIndex(r => r.join() === stringified);
  if (foundIndex === -1) return { link: "", label: "" };

  const rowNumber = foundIndex + 1; // 1-based
  const colNumber = deliveryMapLinkIndex + 1;
  const range = dataSheet.getRange(rowNumber, colNumber);

  // Attempt to get display label & link
  let label = String(range.getDisplayValue() || '').trim();
  let link = "";

  const rt = range.getRichTextValue && range.getRichTextValue();
  if (rt && rt.getLinkUrl && rt.getLinkUrl()) {
    link = rt.getLinkUrl();
  } else {
    const raw = String(range.getValue() || "").trim();
    if (raw && raw.toLowerCase().startsWith("http")) link = raw;
  }

  // Fallback label logic
  if (!link) label = `Open Map for ${routeValue}`;
  else if (!label || label === link) label = `Open Map for ${routeValue}`;

  return { link, label };
}

/* -------------------------
   Helper: shortenUrlForTyping
------------------------- */
function shortenUrlForTyping(url) {
  if (!url) return "";
  try {
    return UrlFetchApp.fetch(
      "https://is.gd/create.php?format=simple&url=" + encodeURIComponent(url)
    ).getContentText();
  } catch (e) {
    return url; // fallback
  }
}

/* -------------------------
   Helper: buildAutoMapUrl
   - Builds a Google Maps URL that shows all addresses for a route as pins.
   - Uses the /maps/search/ format which accepts a pipe-separated list of
     addresses and renders them as individual markers — no API key needed,
     works in any browser, shareable as-is.
   - Google caps the address list at around 10 locations; excess are silently
     dropped by Maps, so we enforce the cap here to keep URLs predictable.
------------------------- */
function buildAutoMapUrl(addresses) {
  const MAX_PINS = 10;
  const clean = addresses
    .map(a => String(a || "").trim())
    .filter(a => a.length > 0)
    .slice(0, MAX_PINS);

  if (clean.length === 0) return "";

  // /maps/search/ with a pipe-separated query shows multiple pins on one map
  const query = clean.map(a => encodeURIComponent(a)).join("%7C"); // %7C = pipe
  return `https://www.google.com/maps/search/${query}`;
}

/* -------------------------
   Helper: writeRouteIndexSheet
   - writes or creates an index sheet summary
------------------------- */
function writeRouteIndexSheet(ss, indexSheetName, summaryRows) {
  let indexSheet = ss.getSheetByName(indexSheetName);
  if (!indexSheet) indexSheet = ss.insertSheet(indexSheetName);
  indexSheet.clearContents();
  indexSheet.clearFormats();

  if (!summaryRows || summaryRows.length === 0) return;
  indexSheet.getRange(1,1,summaryRows.length, summaryRows[0].length).setValues(summaryRows);
  indexSheet.getRange(1,1,1,summaryRows[0].length).setFontWeight("bold").setBackground("#F0F0F0").setHorizontalAlignment("center");
  indexSheet.autoResizeColumns(1, summaryRows[0].length);
  // Make the Map column a bit wider so "View Map" links are easy to click
  const mapCol = summaryRows[0].length;
  indexSheet.setColumnWidth(mapCol, 120);
}

/* -------------------------
   Small utilities
------------------------- */
function pickRouteColor(index) {
  const routeColors = [
    "#D9EAD3", "#FFF2CC", "#FCE5CD", "#D9D2E9", "#CFE2F3",
    "#F4CCCC", "#EAD1DC", "#E0F7FA", "#F8BBD0", "#B2EBF2",
    "#E8F5E9", "#FFF9C4", "#FFCCBC", "#C5CAE9", "#DCEDC8"
  ];
  return routeColors[index % routeColors.length];
}

function pickStopColors() {
  return [
    "#D9EAD3", "#FFF2CC", "#FCE5CD", "#D9D2E9", "#CFE2F3",
    "#F4CCCC", "#EAD1DC", "#E0F7FA", "#F8BBD0", "#B2EBF2",
    "#E8F5E9", "#FFF9C4", "#FFCCBC", "#C5CAE9", "#DCEDC8"
  ];
}

/**********************************************************************
 Batch 2 — ROUTE SHEET BUILDERS
 These functions replace the route-level responsibilities of
 populateRouteTabs (headers, map links, item summaries, timestamps)
***********************************************************************/

/**
 * Main route sheet builder (Batch 2)
 * Batch 3 will add: write person blocks inside this controller.
 */
function buildRouteSheet(opts) {
  const {
    ss,
    sheet,
    route,
    driverName,
    personStops,
    itemMap,
    finalItemHeaders,
    headers,
    routeColor,
    colIndices,
    mapInfo
  } = opts;

  sheet.clearFormats();

  const rollup = computeRouteItemRollup(personStops, itemMap, headers);

  // 1. Write the Header (Colors A1:J7 Gray)
  writeFullRouteHeader({
    sheet,
    route,
    driverName,
    mapInfo,
    rollup,
    itemMap
  });

  // --- ADD THIS LINE BELOW ---
  // This clears any "bleeding" gray background from row 8 down to the end of the sheet
  sheet.getRange(8, 1, sheet.getMaxRows() - 7, 10).setBackground("white");
  // ---------------------------

  const stopColors = pickStopColors();
  let row = 9; // row 8 intentionally blank and now forced to white

  Object.values(personStops).forEach((rowsForPerson, idx) => {
    const currentColor = stopColors[idx % stopColors.length];
    const tableStartRow = row;

    row = writeStopTable({
      sheet,
      rowsForPerson,
      headers,
      itemMap,
      routeColor: currentColor, 
      startRow: row
    });

    applyStopTableBorders({
      sheet,
      startRow: tableStartRow,
      endRow: row - 2, 
      stopColor: currentColor
    });
  });

  return row;
}

/**********************************************************************
 ROUTE SHEET Header
***********************************************************************/
function writeFullRouteHeader(opts) {
  const { sheet, route, driverName, mapInfo, rollup, itemMap } = opts;

  // Column widths
  sheet.setColumnWidth(1, 250);
  for (let c = 2; c <= 10; c++) sheet.setColumnWidth(c, 100);

  // HR1 — Driver / Route
  sheet.getRange("A1:D1").merge()
    .setValue(driverName)
    .setFontWeight("bold")
    .setFontSize(24)
    .setHorizontalAlignment("left");

  sheet.getRange("E1:J1").merge()
    .setValue(route)
    .setFontWeight("bold")
    .setFontSize(24)
    .setHorizontalAlignment("right");

  // HR2 — Map Link
  sheet.getRange("A2").setValue("Map Link:")
    .setFontSize(16);

  const mapCell = sheet.getRange("B2:J2").merge()
    .setFontSize(16)
    .setHorizontalAlignment("right");

  if (mapInfo?.link) {
    mapCell.setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText(mapInfo.label || mapInfo.link)
        .setLinkUrl(mapInfo.link)
        .build()
    );
  } else {
    mapCell.setValue("No map available");
  }

  // HR3 — Short Link
  sheet.getRange("A3").setValue("Short/Typeable Link:")
    .setFontSize(16);

  const shortUrl = mapInfo?.link ? shortenUrlForTyping(mapInfo.link) : "";
  sheet.getRange("B3:J3").merge()
    .setValue(shortUrl)
    .setFontSize(16)
    .setHorizontalAlignment("right");

  // --- HR4: Route Items (Headers) ---
  sheet.getRange("A4").setValue("Route Items")
    .setFontWeight("bold")
    .setFontSize(12);

  // We define itemNames ONCE using the labels from your itemMap
  const itemNames = itemMap.map(i => i.label);

  sheet.getRange(4, 2, 1, itemNames.length)
    .setValues([itemNames])
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center");

  // --- HR5: Totals Row (Quantities) ---
  sheet.getRange("A5").setValue("Total Quantities")
    .setFontWeight("bold")
    .setFontSize(14);

  // Map the rollup values to the headers defined above
  const totalValues = itemNames.map(name => rollup[name] || "");

  // Write the values starting at Column B (2)
  const totalsRange = sheet.getRange(5, 2, 1, totalValues.length);
  totalsRange.setValues([totalValues])
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center");

  // Specific check: If 'diapers' is in your map, it will now land 
  // exactly in its assigned column in Row 5.

  // HR6 — Static explanation (wrapped, auto-height)
  const hr6 = sheet.getRange("A6:J6").merge();

  hr6.setValue(
    "In this table here, diapers shows the quantity of PACKS of Diapers and/or Pull-Ups for your route, but in the tables below, it will show the sizes each delivery gets instead. Everything else is a quantity."
  );

  hr6
    .setWrap(true)
    .setFontStyle("italic");


  // HR7 — Timestamp (merged, centered, bold)
  const stamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd HH:mm:ss"
  );

  sheet.getRange("A7:J7")
    .merge()
    .setValue(`Generated: ${stamp}`)
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");


  // Background shading HR1–HR7
  sheet.getRange("A1:J7").setBackground("#f3f3f3");

    // Medium border around entire header (HR1–HR7)
  sheet.getRange("A1:J7").setBorder(
    true,  // top
    true,  // left
    true,  // bottom
    true,  // right
    false, // vertical inner
    false, // horizontal inner
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

}

/**********************************************************************
 ITEM ROLLUP LOGIC
***********************************************************************/
function computeRouteItemRollup(personStops, itemMap, headers) {
  const rollup = {};
  const allRows = Object.values(personStops).flat();
  
  const diaperIndices = [
    headers.indexOf("diapers1(size)"),
    headers.indexOf("diapers2(size)"),
    headers.indexOf("diapers3(size)")
  ].filter(idx => idx !== -1);

  itemMap.forEach(item => {
    const label = item.label;
    if (!label) return;

    if (item.type === "qty" && item.sourceHeader) {
      const idx = headers.indexOf(item.sourceHeader);
      const total = allRows.reduce((sum, row) => {
        const v = Number(row[idx] || 0);
        return sum + (isFinite(v) ? v : 0);
      }, 0);
      rollup[label] = total > 0 ? total : "";
    }

    if (item.type === "diapers") {
      let totalPacks = 0;
      allRows.forEach(row => {
        diaperIndices.forEach(idx => {
          // If the cell contains a size (like "1" or "4"), it counts as 1 pack
          if (String(row[idx] || '').trim() !== "") {
            totalPacks++;
          }
        });
      });
      rollup[label] = totalPacks > 0 ? totalPacks : "";
    }
  });

  return rollup;
}

/**********************************************************************
 UTILITIES
***********************************************************************/
function applyRouteColumnWidths(sheet) {
  sheet.setColumnWidth(1, 180); // labels / people names
  sheet.setColumnWidth(2, 300); // text / links
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 90);
}

/*
function shortenUrl(url) {
  if (!url || url.length <= 60) return url;
  return url.substring(0, 55) + "...";
}
*/

/**BATCH 3 Alt:
 * *****/

function writeStopTable(opts) {
  const {
    sheet,
    rowsForPerson,
    headers,
    itemMap,
    routeColor,
    startRow
  } = opts;

  let row = startRow;

  row = writeSTR1(sheet, rowsForPerson, headers, row);
  //row = writeSTR2(sheet, rowsForPerson, headers, row);
  row = renderSpecialItemRequestsRow({ sheet, rowsForPerson, headers, row });
  //row = writeSTR3(sheet, rowsForPerson, headers, row);
  row = renderSpecialDeliveryInstructionsRow({ sheet, rowsForPerson, headers, row });
  row = writeSTR4(sheet, rowsForPerson, headers, itemMap, row);
  row = writeSTR5(sheet, rowsForPerson, headers, itemMap, row);

  const endRow = row - 1; // last written row

  // Border around entire stop table (A:J)
  sheet.getRange(startRow, 1, endRow - startRow + 1, 10).setBorder(
    true, true, true, true,   // outer borders
    false, false,             // no internal lines
    "black",
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  return row + 1; // single blank row after table
}


function writeSTR1(sheet, rows, headers, row) {
  const nameIdx = headers.indexOf("Name");
  const phoneIdx = headers.indexOf("Phone");
  const addrIdx = headers.indexOf("Address");

  const r = rows[0];
  const leftText = `${r[nameIdx] || ""}  /  ${r[phoneIdx] || ""}`;
  const rightText = r[addrIdx] || "";

  sheet.getRange(`A${row}:D${row}`).merge()
    .setValue(leftText)
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("left");

  sheet.getRange(`E${row}:J${row}`).merge()
    .setValue(rightText)
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("right");

  return row + 1;
}

function writeSTR2(sheet, rows, headers, row) {
  const idx = headers.indexOf("Special Item Requests");
  if (idx === -1) return row;

  const text = rows.map(r => r[idx]).filter(Boolean).join("\n");
  if (!text) return row;

  sheet.getRange(`A${row}`).setValue("Special Item Requests:")
    .setFontSize(12);

  sheet.getRange(`B${row}:J${row}`).merge()
    .setValue(text)
    .setFontSize(16)
    .setWrap(true);

  return row + 1;
}

function writeSTR3(sheet, rows, headers, row) {
  const idx = headers.indexOf("Special Delivery Instructions");
  if (idx === -1) return row;

  const text = rows.map(r => r[idx]).filter(Boolean).join("\n");
  if (!text) return row;

  sheet.getRange(`A${row}`).setValue("Special Instructions:")
    .setFontSize(12);

  sheet.getRange(`B${row}:J${row}`).merge()
    .setValue(text)
    .setFontSize(16)
    .setWrap(true);

  return row + 1;
}

function writeSTR4(sheet, rows, headers, itemMap, row) {
  sheet.getRange(`A${row}`).setValue("Special Item Types:")
    .setFontSize(12);

  const labels = itemMap.map(i => i.label);
  sheet.getRange(row, 2, 1, labels.length)
    .setValues([labels])
    .setFontSize(10)
    .setHorizontalAlignment("center");

  return row + 1;
}

function writeSTR5(sheet, rows, headers, itemMap, row) {
  sheet.getRange(`A${row}`).setValue("Special Item Values:")
    .setFontSize(12);

  // Diaper sorting order for the consolidated list
  const diaperOrder = ['P','N','1','2','3','4','5','6','7','8','2/3T','3/4T','4/5T','5/6T','6/7T'];

  const values = itemMap.map(item => {
    // 1. Handle standard quantity items (Food, Eggs, etc.)
    if (item.type === "qty") {
      const idx = headers.indexOf(item.sourceHeader);
      const total = rows.reduce((s, r) => s + Number(r[idx] || 0), 0);
      return total > 0 ? total : "";
    }

    // 2. Handle Diapers (Consolidating sizes into a string like "1 | 3")
    if (item.type === "diapers") {
      const sizes = new Set();
      
      // Look through every row for this person and every diaper column index
      rows.forEach(r => {
        item.sourceIndex.forEach(i => {
          if (i !== -1) {
            const val = String(r[i] || '').trim();
            if (val) sizes.add(val);
          }
        });
      });

      if (sizes.size === 0) return "";

      // Sort based on the diaperOrder array and join with a pipe
      return Array.from(sizes)
        .sort((a, b) => diaperOrder.indexOf(a) - diaperOrder.indexOf(b))
        .join(' | ');
    }
    
    return "";
  });

  sheet.getRange(row, 2, 1, values.length)
    .setValues([values])
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center");

  return row + 1;
}

/**********************************************************************
 SECTION 2 — Special Item Requests (Modular)
 ***********************************************************************/
function renderSpecialItemRequestsRow(opts) {
  const { sheet, rowsForPerson, headers, row } = opts;
  const idx = headers.indexOf("Special Item Requests");
  if (idx === -1) return row;

  const all = rowsForPerson.map(r => r[idx]).filter(v => v);
  if (all.length === 0) return row;

  const text = all.join("\n");
  const keywords = ["gluten", "dairy", "vegan", "vegetarian", "allergy", "nut", "tomato", "tomatos", "tomatoes", "soy", "lactose", "plant", "extra", "no", "non", "not"];

  // 1. Prepare the cell structure
  sheet.getRange(row, 1).setValue("Special Item Requests:").setFontSize(12);
  const target = sheet.getRange(row, 2, 1, 9);
  target.merge().setWrap(true);

  // 2. Generate the RichText using your function
  // Passing text and keywords as two separate arguments to match your definition
  const rich = createPartiallyBoldedText(text, keywords);

  // 3. Apply the RichText to the top-left cell of the merged range
  //sheet.getRange(row, 2).setRichTextValue(rich);
  // 3. Apply the RichText to the top-left cell of the merged range
  const cell = sheet.getRange(row, 2);
  cell.setRichTextValue(rich);
  cell.setFontSize(16); // Force the container to 16
  
  return row + 1;
}

/**********************************************************************
 SECTION 3 — Special Delivery Instructions (Modular)
 ***********************************************************************/
function renderSpecialDeliveryInstructionsRow(opts) {
  const { sheet, rowsForPerson, headers, row } = opts;
  const idx = headers.indexOf("Special Delivery Instructions");
  if (idx === -1) return row;

  const all = rowsForPerson.map(r => r[idx]).filter(v => v);
  if (all.length === 0) return row;

  const text = all.join("\n");
  // Optional: Add delivery-specific keywords if you want bolding here too
  const deliveryKeywords = ["gate", "code", "door", "ring", "back", "side", "call"];

  sheet.getRange(row, 1).setValue("Special Instructions:").setFontSize(12);
  const target = sheet.getRange(row, 2, 1, 9);
  target.merge().setWrap(true);

  const rich = createPartiallyBoldedText(text, deliveryKeywords);
  const cell = sheet.getRange(row, 2);
  cell.setRichTextValue(rich);
  cell.setFontSize(16); // Force the container to 16

  return row + 1;
}
/**********************************************************************
 STOP TABLE FORMATTING — Borders, shading
***********************************************************************/
function applyStopTableBorders(opts) {
  const { sheet, startRow, endRow, stopColor } = opts;
  const totalCols = 10; // A:J

  const rng = sheet.getRange(startRow, 1, endRow - startRow + 1, totalCols);

  // 1. Fill the entire stop with its assigned color
  rng.setBackground(stopColor);

  // 2. Apply a solid medium outer border
  rng.setBorder(
    true, true, true, true,   // Top, Left, Bottom, Right
    false, false,             // No internal lines
    "black", 
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // 3. Optional: Add a subtle dotted line between the Name row and the Items row
  // This helps scanability if the tables are long.
  sheet.getRange(startRow, 1, 1, totalCols).setBorder(
    null, null, true, null, 
    false, false, 
    "black", 
    SpreadsheetApp.BorderStyle.DOTTED
  );
}


/**********************************************************************
 ITEM EXTRACTION HELPERS
***********************************************************************/
function extractItemValuesForPerson(rows, itemMap, headers) {
  const out = {};

  itemMap.forEach(item => {
    const label = item.label;

    if (item.type === "qty" && item.sourceHeader) {
      const idx = headers.indexOf(item.sourceHeader);
      out[label] = rows.reduce((sum, r) => sum + Number(r[idx] || 0), 0);
    }

    if (item.type === "diapers") {
      out["diapers"] = 0;
      const indices = item.sourceIndex || [];
      indices.forEach(idx => {
        out["diapers"] += rows.reduce((sum, r) => sum + Number(r[idx] || 0), 0);
      });
    }
  });

  return out;
}



// End New drivers tab functions



// New packing list tab functions

function makePackingLists() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    refreshDeliveriesHelpTable();
    const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
    const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
    const outputSheet = setupPackingListTab(ss, "PackingLists");

    // 1. Load and Prepare Data
    const { 
        dataRows, 
        headers, 
        confForIdx, 
        includedIndexes, 
        diaperSizeIndices, 
        packsOrigIdx, 
        eggsOrigIdx, 
        milkOrigIdx, 
        wipesOrigIdx, 
        dogFoodOrigIdx, 
        catFoodOrigIdx, 
        toiletriesOrigIdx, 
        femHygieneOrigIdx,
        routeIdx,
        driverIdx
    } = loadAndIndexPackingData(dataSheet);

    const routeDriverMap = buildPackingListRouteDriverMap(helperSheet);
    
    // 2. Filter Data for Calculations and Route Tables
    const { confirmedDataRows, dataRowsForDiaperCount } = filterPackingDataForReports(dataRows, confForIdx);
    
    // 3. Perform Calculations
    const allDiaperItems = defineDiaperItems();
    const diaperCounts = calculatePackingDiaperAndWipesCounts(dataRowsForDiaperCount, allDiaperItems, diaperSizeIndices, wipesOrigIdx);
    const confForCounts = calculatePackingConfirmationCounts(dataRows, confForIdx);
    
    // 4. Generate Tables
    let outputRow = 1;
    
    // A. Item Summary (A1)
    outputRow = generateItemSummaryTable(outputSheet, confirmedDataRows, packsOrigIdx, eggsOrigIdx, milkOrigIdx, outputRow);
    
    // B. Confirmation Status (H1) - This table starts at the same row as Item Summary
    // It does NOT update outputRow as it's side-by-side.
    generateConfirmationStatusTable(outputSheet, confForCounts, 1); // Start Row 1 is still correct here

    // C. Diaper/Wipes Reference (Starts after Item Summary)
    // We add a 2-row buffer below the first two tables, which both end at Row 4/5.
    // The Diaper table should start at Row 7 (5 + 2).
    outputRow = generateDiaperReferenceTable(outputSheet, allDiaperItems, diaperCounts, outputRow + 2); 
    
    // D. Timestamp and Route Sections
    // Add a 2-row buffer before the timestamp
    outputRow = generateTimestamp(outputSheet, outputRow + 2); 

    generatePackingListRouteSections(
        outputSheet, dataRowsForDiaperCount, headers, routeDriverMap, diaperCounts,
        {
            routeIdx, driverIdx, confForIdx, includedIndexes,
            packsOrigIdx, eggsOrigIdx, milkOrigIdx, dogFoodOrigIdx, catFoodOrigIdx,
            wipesOrigIdx, toiletriesOrigIdx, femHygieneOrigIdx
        },
        // This clears all existing "baked-in" green formatting before writing new data
        outputRow
    );

    // 5. Final Formatting
    applyPackingListFinalColumnSizing(outputSheet, includedIndexes, driverIdx, headers);
    
    SpreadsheetApp.flush();
    return outputSheet.getName();
}

function setupPackingListTab(ss, outputSheetName) {
    let outputSheet = ss.getSheetByName(outputSheetName);
    if (!outputSheet) outputSheet = ss.insertSheet(outputSheetName);

    // Get the maximum size of the sheet to ensure full clearing
    const maxRows = outputSheet.getMaxRows();
    const maxCols = outputSheet.getMaxColumns();

    // Get the range covering the entire sheet
    const fullRange = outputSheet.getRange(1, 1, maxRows, maxCols);

    // Completely clear previous content, formatting, filters, and validations
    // Note: If you use outputSheet.clear(), you don't need all the clears below,
    // but using the Range object is safer if the sheet might be large.
    
    // Using Range methods is more specific and often faster than relying on Sheet.clear()
    // fullRange.clearContents();
    // fullRange.clearFormats();
    // fullRange.clearDataValidations();
    fullRange.clear();

    // Clear filters (if any)
    if (outputSheet.getFilter()) {
        outputSheet.getFilter().remove();
    }
    
    // Clear frozen rows/columns (already in your original code, kept for completeness)
    outputSheet.setFrozenRows(0);
    outputSheet.setFrozenColumns(0);
    
    return outputSheet;
}

function defineDiaperItems() {
    const rawDiaperSizes = [
        ["N", "Newborn Diapers", "< 10 lbs"], ["1", "Size 1 Diapers", "8-14 lbs"],
        ["2", "Size 2 Diapers", "12-18 lbs"], ["3", "Size 3 Diapers", "16-28 lbs"],
        ["4", "Size 4 Diapers", "22-37 lbs"], ["5", "Size 5 Diapers", "27 + lbs"],
        ["6", "Size 6 Diapers", "35 + lbs"], ["7", "Size 7 Diapers", "41 + lbs"],
        ["2/3T", "Size 2T/3T PullUps", "16-34 lbs"], ["3/4T", "Size 3T/4T PullUps", "30-40 lbs"],
        ["4/5T", "Size 4T/5T PullUps", "37 + lbs"], ["5/6T", "Size 5T/6T PullUps", "41 + lbs"],
        ["6/7T", "Size 6T/7T PullUps", "64 + lbs"], ["AdSml", "Adult Small Diapers", ""],
        ["AdMed", "Adult Medium Diapers", ""], ["AdLrg", "Adult Large Diapers", ""],
        ["AdXL", "Adult Xtra Large Diapers", ""], ["AdXXL", "Adult XXL Diapers", ""]
    ];
    
    const allDiaperItems = [...rawDiaperSizes]; 
    allDiaperItems.push(["Diapers Total", "Total Diaper Packs", ""]); 
    allDiaperItems.push(["Wipes", "Wipes Packs", ""]); 
    
    return allDiaperItems;
}

function loadAndIndexPackingData(dataSheet) {
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim()); // Ensure headers are strings and trimmed
    const dataRows = data.slice(1);
    
    const includedCols = [
        "Driver", "Name", "# packs", "# eggs", "# milk", 
        "dogFood(qty)", "catFood(qty)", "diapers1(size)", 
        "diapers2(size)", "diapers3(size)", "wipes(qty)", 
        "toiletries(qty)", "fem hygiene(qty)", "Special Item Requests"
    ];

    const getIndex = (h) => headers.indexOf(h);
    
    const confForHeader = headers.find(h => h.startsWith("Conf for"));
    const confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;
    
    const routeIdx = getIndex("routeDescription");
    const driverIdx = getIndex("Driver");
    const diapers1Idx = getIndex("diapers1(size)");
    const diapers2Idx = getIndex("diapers2(size)");
    const diapers3Idx = getIndex("diapers3(size)");

    return {
        dataRows,
        headers,
        confForIdx,
        routeIdx,
        driverIdx,
        includedIndexes: includedCols.map(getIndex).filter(i => i !== -1),
        diaperSizeIndices: [diapers1Idx, diapers2Idx, diapers3Idx].filter(i => i !== -1),
        packsOrigIdx: getIndex("# packs"),
        eggsOrigIdx: getIndex("# eggs"),
        milkOrigIdx: getIndex("# milk"),
        wipesOrigIdx: getIndex("wipes(qty)"),
        dogFoodOrigIdx: getIndex("dogFood(qty)"),
        catFoodOrigIdx: getIndex("catFood(qty)"),
        toiletriesOrigIdx: getIndex("toiletries(qty)"),
        femHygieneOrigIdx: getIndex("fem hygiene(qty)")
    };
}

function buildPackingListRouteDriverMap(helperSheet) {
    const routeDriverMap = {};
    if (helperSheet) {
        const helperData = helperSheet.getDataRange().getValues().slice(1);
        helperData.forEach(row => { 
            const routeKey = String(row[0] || '').trim();
            const driverValue = String(row[1] || '').trim();
            if (routeKey) routeDriverMap[routeKey] = driverValue; 
        });
    }
    return routeDriverMap;
}

function filterPackingDataForReports(dataRows, confForIdx) {
    let confirmedDataRows = dataRows;
    let dataRowsForDiaperCount = dataRows;

    if (confForIdx !== -1) {
        // Confirmed only (for Item Summary Totals)
        confirmedDataRows = dataRows.filter(row => {
            const confValue = String(row[confForIdx] || '').toUpperCase().trim();
            return confValue === 'YES';
        });

        // Diaper/Wipes Count and Route Tables (YES or NO ANS)
        dataRowsForDiaperCount = dataRows.filter(r => {
            const confValue = String(r[confForIdx] || '').toUpperCase().trim();
            return confValue === 'YES' || confValue.includes('NO ANS');
        });
    }
    return { confirmedDataRows, dataRowsForDiaperCount };
}

function calculatePackingDiaperAndWipesCounts(dataRowsForDiaperCount, allDiaperItems, diaperSizeIndices, wipesOrigIdx) {
    // Create a temporary map where all keys are UPPERCASE for matching
    const diaperCounts = allDiaperItems.reduce((acc, item) => { 
      acc[item[0].toUpperCase()] = 0; 
      return acc; 
    }, {});
    
    let totalDiaperPacks = 0; 

    dataRowsForDiaperCount.forEach(row => {
        // Count Diaper Codes
        diaperSizeIndices.forEach(idx => {
            const rawCode = String(row[idx] || '').trim().toUpperCase();
            
            // Match against the uppercase keys
            if (diaperCounts.hasOwnProperty(rawCode) && rawCode !== "WIPES" && rawCode !== "DIAPERS TOTAL") {
                diaperCounts[rawCode]++;
                totalDiaperPacks++;
            }
        });

        // Count Wipes (quantity) - Note the uppercase key here
        const wipesQty = parseFloat(row[wipesOrigIdx] || 0) || 0;
        diaperCounts["WIPES"] += wipesQty;
    });
    
    // Map the uppercase counts back to the original casing for the table display
    const finalCounts = {};
    allDiaperItems.forEach(item => {
      const originalCode = item[0];
      finalCounts[originalCode] = diaperCounts[originalCode.toUpperCase()] || 0;
    });

    finalCounts["DIAPER_TOTAL"] = totalDiaperPacks;
    return finalCounts;
}

function calculatePackingConfirmationCounts(dataRows, confForIdx) {
    const counts = { "NO ANS": 0, "NO": 0, "YES": 0, "Total": 0 };
    if (confForIdx !== -1) {
        dataRows.forEach(row => {
            let val = String(row[confForIdx] || '').toUpperCase().trim();
            if (val.includes("NO ANS")) {
                counts["NO ANS"]++;
            } else if (val === "NO") {
                counts["NO"]++;
            } else if (val === "YES") {
                counts["YES"]++;
            }
            if (val) {
                counts["Total"]++;
            }
        });
    }
    return counts;
}

function generateItemSummaryTable(outputSheet, confirmedDataRows, packsIdx, eggsIdx, milkIdx, startRow) {
    const summaryStartCol = 1; 
    const summaryNumCols = 5; 
    
    const totalPacks = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[packsIdx] || 0) || 0), 0);
    const totalEggs = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[eggsIdx] || 0) || 0), 0);
    const totalMilk = confirmedDataRows.reduce((s, r) => s + (parseFloat(String(r[milkIdx]).replace('*', '') || 0)) || 0, 0);

    // Set Header and Data
    outputSheet.getRange(startRow, summaryStartCol, 1, summaryNumCols).setValues([["Standard Items", "Total (Confirmed Only)", "", "", ""]]); 
    outputSheet.getRange(startRow + 1, summaryStartCol, 3, summaryNumCols).setValues([ 
        ["# packs", totalPacks, "", "", ""],
        ["# eggs", totalEggs, "", "", ""],
        ["# milk", totalMilk, "", "", ""]
    ]);
    
    // Apply Merges
    outputSheet.getRange("B1:E1").merge();
    for (let r = startRow + 1; r <= startRow + 3; r++) {
        outputSheet.getRange(r, 2, 1, 4).merge(); // B:E merge
    }

    // Apply Formatting
    outputSheet.getRange(startRow, summaryStartCol, 4, summaryNumCols) 
        .setFontWeight("bold")
        .setFontSize("14")
        .setBackground("#D9EAD3")
        .setBorder(true, true, true, true, null, null);
    
    // Alignments
    outputSheet.getRange("A1").setHorizontalAlignment("left"); 
    outputSheet.getRange("B1").setHorizontalAlignment("right"); 
    outputSheet.getRange("A2:A4").setHorizontalAlignment("left"); 
    outputSheet.getRange("B2:B4").setHorizontalAlignment("right");
    return startRow + 4; // Return the next available row (Row 5)
}

function generateConfirmationStatusTable(outputSheet, confForCounts, startRow) {
    const confSummaryStartCol = 8; // Column H
    const confSummaryNumCols = 6; // H, I, J, K, L, M

    // Set Header and Data
    outputSheet.getRange(startRow, confSummaryStartCol, 1, confSummaryNumCols).setValues([
        ["Confirmation Status", "", "", "", "Total Count", ""]
    ]);
    outputSheet.getRange(startRow + 1, confSummaryStartCol, 4, confSummaryNumCols).setValues([
        ["No Ans", "", "", "", confForCounts["NO ANS"], ""],
        ["No", "", "", "", confForCounts["NO"], ""],
        ["Yes", "", "", "", confForCounts["YES"], ""],
        ["Total", "", "", "", confForCounts["Total"], ""]
    ]);

    // Apply Merges
    outputSheet.getRange("H1:K1").merge();
    outputSheet.getRange("L1:M1").merge();
    for (let r = startRow + 1; r <= startRow + 4; r++) {
        outputSheet.getRange(r, 8, 1, 4).merge(); // H:K merge
        outputSheet.getRange(r, 12, 1, 2).merge(); // L:M merge
    }
    
    // Apply Formatting
    outputSheet.getRange(startRow, confSummaryStartCol, 5, confSummaryNumCols)
        .setFontWeight("bold")
        .setFontSize("14")
        .setBackground("#D9D2E9") 
        .setBorder(true, true, true, true, null, null);

    // Alignments
    outputSheet.getRange("H1").setHorizontalAlignment("left"); 
    outputSheet.getRange("L1").setHorizontalAlignment("right");
    outputSheet.getRange("H2:H5").setHorizontalAlignment("left");
    outputSheet.getRange("L2:L5").setHorizontalAlignment("right");    
    return startRow + 5; // Return the next available row (Row 6)
}

function generateDiaperReferenceTable(outputSheet, allDiaperItems, diaperCounts, startRow) {
    const diaperTableStartCol = 1; 
    const diaperTableNumCols = 15;
    const DIAPER_NUM_ROWS = allDiaperItems.length + 1; // Header + 21 items = 22 rows
    
    // Define the list of *only* raw sizes (excluding the last two: Total and Wipes)
    // This is used for calculating the correct row indices for shading/borders.
    const rawDiaperSizes = allDiaperItems.slice(0, allDiaperItems.length - 2); 

    // --- 1. Set Values ---
    
    const diaperDataForSetValues = [
        ["Size/Weight Ranges", "Size Description", "", "", "", "Code", "", "Pre Ordered", "", "Not Ordered Tally", "", "", "", "Not Fulfilled Tally", ""], // Header
        ...allDiaperItems.map(row => {
            // Count logic depends on the specific row
            const code = row[0];
            const count = (code === "Diapers Total") 
                          ? diaperCounts["DIAPER_TOTAL"] || 0
                          : diaperCounts[code] || 0;
                          
            return [row[2], row[1], "", "", "", code, "", count, "", "", "", "", "", "", ""];
        })
    ];

    const diaperTableRange = outputSheet.getRange(startRow, diaperTableStartCol, DIAPER_NUM_ROWS, diaperTableNumCols);
    diaperTableRange.setValues(diaperDataForSetValues);

    // --- 2. Apply Merging ---
    
    // Header Merging (startRow)
    outputSheet.getRange(startRow, 2, 1, 4).merge();  // B:E (Desc)
    outputSheet.getRange(startRow, 6, 1, 2).merge();  // F:G (Code)
    outputSheet.getRange(startRow, 8, 1, 2).merge();  // H:I (PreOrdered)
    outputSheet.getRange(startRow, 10, 1, 4).merge(); // J:M (NotOrdered)
    outputSheet.getRange(startRow, 14, 1, 2).merge(); // N:O (Not Fulfilled)

    // Data Rows Merging (startRow + 1 to End)
    const firstDataRow = startRow + 1; 
    const numDataRows = DIAPER_NUM_ROWS - 1; // 21 rows of data

    for (let r = firstDataRow; r < firstDataRow + numDataRows; r++) {
        // Merge Size Description (B:E)
        outputSheet.getRange(r, 2, 1, 4).merge(); 
        // Merge Code (F:G)
        outputSheet.getRange(r, 6, 1, 2).merge(); 
        // Merge PreOrdered (H:I)
        outputSheet.getRange(r, 8, 1, 2).merge(); 
        // Merge NotOrdered (J:M) 
        outputSheet.getRange(r, 10, 1, 4).merge();
        // Merge Not Fulfilled (N:O)
        outputSheet.getRange(r, 14, 1, 2).merge();
    }
    
    // --- 3. Apply Formatting ---
    
    // Apply basic borders and background to the whole table
    diaperTableRange
        .setBorder(true, true, true, true, true, true)
        .setBackground("#F3F3F3") 
        .setFontSize(12);

    // Header Formatting (Row startRow)
    outputSheet.getRange(startRow, diaperTableStartCol, 1, diaperTableNumCols)
        .setFontWeight("bold")
        .setBackground("#CCCCCC") // Darker Gray
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
        
    // Set alignment for all data rows (startRow + 1 to End)
    // Range (A)
    outputSheet.getRange(firstDataRow, diaperTableStartCol, numDataRows, 1).setHorizontalAlignment("center"); 
    // Description (B:E merged)
    outputSheet.getRange(firstDataRow, 2, numDataRows, 4).setHorizontalAlignment("left").setWrap(true); 
    // Code, PreOrdered, NotOrdered, Not Fulfilled (F to O)
    outputSheet.getRange(firstDataRow, 6, numDataRows, 10).setHorizontalAlignment("center"); 
    
    // === ALTERNATING ROW SHADING (excluding Total and Wipes) ===
    const regularDiaperRows = rawDiaperSizes.length; // 19 rows of actual sizes
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
    
    // Apply specific formatting to the "Total Diaper Packs" row
    const totalRowIndex = firstDataRow + rawDiaperSizes.length; // 19th row of data (index 19 in data array)
    outputSheet.getRange(totalRowIndex, 1, 1, diaperTableNumCols)
        .setFontWeight("bold")
        .setBackground("#CFE2F3"); // Light Blue color
        
    // Set the Wipes row to a neutral/light shade for distinction
    const wipesRowIndex = totalRowIndex + 1; // 20th row of data
    outputSheet.getRange(wipesRowIndex, 1, 1, diaperTableNumCols)
        .setFontWeight("bold")
        .setBackground("#D9EAD3"); // Light Mint Green color

    // === HEAVIER BORDER LINES at logical breaks ===
    const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM; 
    const borderColor = "#000000";

    // Helper to find the 0-based index of a code in the rawDiaperSizes array
    const findIndexInRawSizes = (code) => rawDiaperSizes.findIndex(d => d[0] === code);

    // 1. After Size 7 Diapers (baby/toddler separation)
    // Index 7 in rawDiaperSizes (8th row of data)
    const afterSize7RowIndex = firstDataRow + findIndexInRawSizes("7"); 
    outputSheet.getRange(afterSize7RowIndex, diaperTableStartCol, 1, diaperTableNumCols)
        .setBorder(null, null, true, null, null, null, borderColor, borderStyle);

    // 2. After 6T/7T PullUps (pull-ups separation)
    // Index 12 in rawDiaperSizes (13th row of data)
    const after6T7TRowIndex = firstDataRow + findIndexInRawSizes("6/7T"); 
    outputSheet.getRange(after6T7TRowIndex, diaperTableStartCol, 1, diaperTableNumCols)
        .setBorder(null, null, true, null, null, null, borderColor, borderStyle);

    // 3. After AdXXL (adult separation)
    // Index 14 in rawDiaperSizes (15th row of data)
    const afterAdSmlRowIndex = firstDataRow + findIndexInRawSizes("AdXXL");
    outputSheet.getRange(afterAdSmlRowIndex, diaperTableStartCol, 1, diaperTableNumCols)
        .setBorder(null, null, true, null, null, null, borderColor, borderStyle);
    
    return startRow + DIAPER_NUM_ROWS; // Return next available row after the table
}

function generateTimestamp(outputSheet, timestampRow) {
    const timestamp = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    const timestampMessage = `This Packing List Tab was generated on ${timestamp}.`;
    const timestampRange = outputSheet.getRange(timestampRow, 1, 1, 15); 

    timestampRange.merge()
        .setValue(timestampMessage)
        .setFontSize(16)
        .setFontWeight("bold")
        .setBackground("#EFEFEF");
    
    return timestampRow + 3; // Space buffer for routes to start
}

function generatePackingListRouteSections(outputSheet, dataRowsForDiaperCount, headers, routeDriverMap, diaperCounts, indices, outputRow) {
    const keywords = ["gluten", "dairy", "vegan", "vegetarian", "allergy", "nut", "tomato", "tomatos", "tomatoes", "soy", "lactose", "plant", "extra", "no", "non", "not"];
    
    // 1. Get unique routes and SORT them alphabetically
    const routes = [...new Set(dataRowsForDiaperCount.map(r => String(r[indices.routeIdx]).trim()).filter(r => r))].sort();

    // --- CONFIGURATION OPTIONS ---
    const USE_GREY_SHADE = true;
    const NEUTRAL_GREY = "#E0E0E0";
    const USE_DARKER_SHADE = true;

    const alternatingColorPairs = [
        ["#D9EAD3", "#F0FDF0"], ["#FFF2CC", "#FFFFFA"], ["#FCE5CD", "#FFF9F4"], 
        ["#D9D2E9", "#F0F0FC"], ["#CFE2F3", "#F0F8FF"], ["#EEDD82", "#FFFFE5"], 
        ["#E0F7FA", "#F0FFFF"], ["#F8BBD0", "#FFF0F5"], ["#E8F5E9", "#F0FFF0"], 
        ["#FAEBD7", "#FFFBF5"], ["#FFCCBC", "#FFF5F0"], ["#C5CAE9", "#E5E8F5"], 
        ["#9ACD32", "#F5FFF0"]  
    ];

    const routeColors = alternatingColorPairs.map(pair => pair[0]);

    const headerMap = {
        "Name": "Name", "# packs": "#\npacks", "# eggs": "#\neggs", "# milk": "#\nmilk", 
        "dogFood(qty)": "#\ndog\nfood", "catFood(qty)": "#\ncat\nfood",
        "diapers1(size)": "size of\ndiapers", "diapers2(size)": "size of\ndiapers",
        "diapers3(size)": "size of\ndiapers", "wipes(qty)": "#\nwipes",
        "toiletries(qty)": "#\nToilet-\nries", "fem hygiene(qty)": "#\nFem\nHygiene",
        "Special Item Requests": "Special Item\nRequests",
        "Packed up & labeled?": "Packed\nup and\nlabeled?", "Picked up by driver?": "Picked\nup by\ndriver?" 
    };

    const includedIndexesNoDriver = indices.includedIndexes.filter(i => i !== indices.driverIdx);
    const checkboxHeaders = ["Packed up & labeled?", "Picked up by driver?"];
    const finalHeaders = includedIndexesNoDriver.map(i => headerMap[headers[i]] || headers[i]).concat(checkboxHeaders.map(h => headerMap[h] || h));
    const numCols = finalHeaders.length; 
    const lastColIndex = numCols;
    const colIndex = name => finalHeaders.indexOf(name) + 1; 

    const DOTTED = SpreadsheetApp.BorderStyle.DOTTED;
    const MEDIUM_SOLID = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    const BLACK = "#000000";
    const WHITE = "#FFFFFF";

    // 2. MAIN LOOP - This now iterates through the SORTED routes list
    routes.forEach((route, i) => {
        const driver = routeDriverMap[route] || "TBD";
        const routeColor = routeColors[i % routeColors.length];
        const colorPair = alternatingColorPairs[i % alternatingColorPairs.length];
        
        let shadedColor = USE_GREY_SHADE ? NEUTRAL_GREY : (USE_DARKER_SHADE ? colorPair[0] : colorPair[1]);
        
        // Filter rows specifically for THIS route in the loop
        const filteredDataRows = dataRowsForDiaperCount.filter(r => String(r[indices.routeIdx]).trim() === route);
        
        if (filteredDataRows.length === 0) return;

        const specialItemsColIndexInOutput = finalHeaders.indexOf("Special Item\nRequests"); 
        const rowsForSheet = [];
        const richTextObjects = []; 

        filteredDataRows.forEach((r, rowIndex) => {
            const rowData = includedIndexesNoDriver.map((idx, col) => {
                let value = r[idx];
                if (col === 0 && headers[idx] === "Name" && typeof value === 'string') {
                    const maxLength = 36;
                    value = value.length > maxLength ? value.substring(0, maxLength - 3).trim() + '...' : value;
                } 
                if (col === specialItemsColIndexInOutput) {
                    const strValue = String(value || '').trim();
                    if (strValue.length > 0) {
                        richTextObjects.push({
                            rowIndex: rowIndex,
                            colIndex: colIndex("Special Item\nRequests"),
                            richText: createPartiallyBoldedText(strValue, keywords)
                        });
                    }
                    return strValue; 
                }
                return value;
            });
            rowsForSheet.push(rowData.concat(["", ""]));
        });

        // DRAW HEADER
        const headerRange = outputSheet.getRange(outputRow, 1, 1, numCols);
        headerRange.mergeAcross().setValue(`Driver:   ${driver}   |   Route:   ${route}`)
            .setFontSize(14).setFontWeight("bold").setBackground(routeColor).setHorizontalAlignment("left");
        outputRow++;

        // DRAW COLUMN HEADERS
        const dataHeaderRow = outputRow;
        const columnHeaderRange = outputSheet.getRange(dataHeaderRow, 1, 1, numCols);
        columnHeaderRange.setValues([finalHeaders]).setFontWeight("bold").setBackground(routeColor)
            .setVerticalAlignment("middle").setWrap(true).setHorizontalAlignment("center")
            .setBorder(true, true, true, true, false, true, BLACK, DOTTED);
        outputSheet.getRange(dataHeaderRow, lastColIndex, 1, 1).setBorder(true, true, true, true, false, false, BLACK, DOTTED);
        outputRow++;

        // DRAW DATA ROWS
        const dataStartRow = outputRow;
        const numDataRows = rowsForSheet.length;
        const dataRange = outputSheet.getRange(dataStartRow, 1, numDataRows, numCols);
        dataRange.setValues(rowsForSheet).setFontSize(12);

        for (let j = 0; j < numDataRows; j++) {
            const rowRange = outputSheet.getRange(dataStartRow + j, 1, 1, numCols);
            rowRange.setBackground(j % 2 === 0 ? WHITE : shadedColor);
            outputSheet.getRange(dataStartRow + j, 1, 1, numCols - 1).setBorder(true, true, true, true, true, false, BLACK, DOTTED);
            outputSheet.getRange(dataStartRow + j, lastColIndex, 1, 1).setBorder(true, true, true, true, true, false, BLACK, DOTTED); 
        }

        richTextObjects.forEach(item => {
            outputSheet.getRange(dataStartRow + item.rowIndex, item.colIndex).setRichTextValue(item.richText);
        });

        outputSheet.getRange(dataStartRow, includedIndexesNoDriver.length + 1, numDataRows, 2).insertCheckboxes();
        outputRow += numDataRows; 

        // DRAW TOTALS ROW
        const totalsRange = outputSheet.getRange(outputRow, 1, 1, numCols);
        totalsRange.setBackground(routeColor).setFontWeight("bold").setBorder(true, true, true, true, false, true);
        outputSheet.getRange(outputRow, lastColIndex).setBorder(true, true, true, true, false, false);
        outputRow++; 
        
        // APPLY THICK BORDERS
        const borderAfterColumns = [
            colIndex("Name"), colIndex("#\nmilk"), colIndex("#\ncat\nfood"),
            colIndex("#\nwipes"), colIndex("#\nFem\nHygiene"),
            colIndex("Special Item\nRequests"), colIndex("Picked\nup by\ndriver?")
        ];
        
        borderAfterColumns.forEach(col => {
            if (col > 0 && col <= numCols) {
                outputSheet.getRange(dataHeaderRow, col, numDataRows + 2, 1).setBorder(null, null, null, true, null, null, BLACK, MEDIUM_SOLID);
            }
        });
        
        outputRow += 1; // Spacer
    });

    return outputRow;
}

function applyPackingListFinalColumnSizing(outputSheet, includedIndexes, driverIdx, headers) {
    
    // Find final headers (excluding driver, adding checkboxes)
    const includedIndexesNoDriver = includedIndexes.filter(i => i !== driverIdx);
    const checkboxHeaders = ["Packed up & labeled?", "Picked up by driver?"];
    const finalHeaders = includedIndexesNoDriver.map(i => headers[i]).concat(checkboxHeaders);
    
    // Set widths for Summary/Diaper Table Columns A-M
    outputSheet.setColumnWidth(1, 175); 
    for(let i = 2; i <= 13; i++) {
        outputSheet.setColumnWidth(i, 60); 
    }

    // Auto-size all data columns starting after M (Column 14) onwards
    if (finalHeaders.length > 13) {
        outputSheet.autoResizeColumns(14, finalHeaders.length - 13); 
    }
    
    // Specific width for 'Special Item Requests' and text wrap
    const specialIdx = finalHeaders.indexOf("Special Item Requests");
    if (specialIdx !== -1) {
        const col = specialIdx + 1;
        outputSheet.setColumnWidth(col, 410);
        // Apply wrap starting from the route-specific data rows (Row 33)
        outputSheet.getRange(33, col, outputSheet.getLastRow() - 32, 1).setWrap(true);
    }

    // Specific width for Checkboxes
    const checkboxStartCol = finalHeaders.indexOf("Packed up & labeled?");
    if (checkboxStartCol !== -1) {
        outputSheet.setColumnWidth(checkboxStartCol + 1, 60);
        outputSheet.setColumnWidth(checkboxStartCol + 2, 60);
    }
}

function simpleRegexpEscape(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function createPartiallyBoldedText(text, keywords) {
  if (!text || typeof text !== 'string') {
    return SpreadsheetApp.newRichTextValue().setText(text || '').build();
  }

  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

  const modifiers = ["no", "not", "non", "extra"];
  const ranges = [];

  // Normalize keywords
  const normalizedKeywords = (keywords || []).map(k => String(k).toLowerCase());

  // 1) Find keyword matches with STRICT WORD BOUNDARIES
  normalizedKeywords.forEach(keyword => {
    if (!keyword) return;
    // The \\b ensures "ring" matches "ring the bell" but NOT "spring"
    const re = new RegExp('\\b' + simpleRegexpEscape(keyword) + '\\b', 'gi');
    let m;
    while ((m = re.exec(text)) !== null) {
      ranges.push([m.index, m.index + m[0].length]);
    }
  });

  // 2) Find modifier occurrences with STRICT WORD BOUNDARIES
  modifiers.forEach(mod => {
    const re = new RegExp('\\b' + simpleRegexpEscape(mod) + '\\b', 'gi');
    let m;
    while ((m = re.exec(text)) !== null) {
      const modStart = m.index;
      const modEnd = modStart + m[0].length;
      ranges.push([modStart, modEnd]);

      // Bold the next word after the modifier
      const after = text.slice(modEnd);
      const nextMatch = after.match(/^\s*([^\s,.;:!?()]+)/);
      if (nextMatch) {
        const nextWord = nextMatch[1];
        const leading = nextMatch[0];
        const wordStart = modEnd + (leading.length - nextWord.length);
        const wordEnd = wordStart + nextWord.length;
        ranges.push([wordStart, wordEnd]);
      }
    }
  });

  if (ranges.length === 0) {
    return builder.build();
  }

  // 3) Merge overlapping/adjacent ranges
  ranges.sort((a, b) => a[0] - b[0]);
  const merged = [];
  for (let i = 0; i < ranges.length; i++) {
    const [s, e] = ranges[i];
    if (merged.length === 0 || s > merged[merged.length - 1][1]) {
      merged.push([s, e]);
    } else {
      merged[merged.length - 1][1] = Math.max(merged[merged.length - 1][1], e);
    }
  }

  // 4) Apply the bold style
  merged.forEach(([s, e]) => {
    if (s >= 0 && e > s && e <= text.length) {
      builder.setTextStyle(s, e, boldStyle);
    }
  });

  return builder.build();
}

// end new packing list work

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
  
  const keywords = ["gluten", "dairy", "vegan", "vegetarian", "allergy", "nut", "tomato", "extra", "no", "soy", "non", "not"];
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

/**
 * Converts a full routeDescription string into a short label for printing on bags.
 *
 * Design goals:
 *   - Every label fits easily on one line at readable font size
 *   - Each route is instantly distinguishable — no two codes look alike at a glance
 *   - Dual-area routes (Encanto/Lemon Grove, etc.) get ONE city initial + number,
 *     so the slash and second city name disappear entirely from the sticker
 *   - Single-area routes keep the city name or a natural short form
 *   - The number suffix (1, 2, 3…) is always preserved so routes within a city
 *     are still sorted correctly
 *
 * To add a new route: add one line to ROUTE_MAP below with the exact
 * routeDescription value as the key and the desired label as the value.
 */
function abbreviateRoute(route) {
  if (!route || route === "UNKNOWN ROUTE") return "No Route";

  // --- Direct lookup table ---
  // Key   = exact routeDescription value from the spreadsheet
  // Value = what prints on the sticker
  //
  // Convention for dual-area routes: pick the most geographically distinct
  // city initial (or 2-letter combo) + route number.  No slashes.
  //
  // Convention for single-area routes: keep the city name if short enough,
  // or shorten to an obvious abbreviation.
  // ─── ROUTE CODE FORMAT ──────────────────────────────────────────────────
  // Standard:    6 alpha + optional integer             e.g. NCOAST or TERALT1
  // Directional: 5 alpha + 1 direction (N/S/E/W) + optional integer  e.g. CHVSTW1
  // Numbers only added when 2+ routes share the same base name/abbreviation.
  // Codes should sound or look like the place — instant mental link.
  //
  // ADDING A NEW ROUTE: one line below. Key = exact routeDescription value
  // from the spreadsheet. Un-abbreviated label in production = missing entry.
  // ─────────────────────────────────────────────────────────────────────────
  const ROUTE_MAP = {
    // DTOWN — Downtown  (multiple routes, numbered)
    "Downtown1":                      "DTOWN1",
    "Downtown2":                      "DTOWN2",
    "Downtown3":                      "DTOWN3",

    // NCOAST — North Coast  (multiple routes, numbered)
    "North Coast1":                   "NCOAST1",
    "North Coast2":                   "NCOAST2",
    "North Coast3":                   "NCOAST3",

    // NORMHT — Normal Heights  (single route, no number)
    "Normal Heights":                 "NORMHT",
    "Normal Heights1":                "NORMHT1",
    "Normal Heights2":                "NORMHT2",

    // AZALEA — Azalea  (single route, no number — it's just the word)
    "Azalea":                         "AZALEA",
    "Azalea1":                        "AZALEA1",
    "Azalea2":                        "AZALEA2",

    // TERALT — Teralta  (multiple routes, numbered)
    "Teralta1":                       "TERALT1",
    "Teralta2":                       "TERALT2",
    "Teralta3":                       "TERALT3",

    // ELCAJO — El Cajon  (EL CAJOn)  (multiple routes, numbered)
    "El Cajon 1":                     "ELCAJO1",
    "El Cajon 2":                     "ELCAJO2",
    "El Cajon 3":                     "ELCAJO3",

    // CHVST + E/W — Chula Vista  (directional: 5 alpha + direction + number)
    "Chula Vista E1":                 "CHVSTE1",
    "Chula Vista E2":                 "CHVSTE2",
    "Chula Vista W 1":                "CHVSTW1",
    "Chula Vista W 2":                "CHVSTW2",
    "Chula Vista W 3":                "CHVSTW3",

    // ENCNTO — Encanto / Lemon Grove  (multiple routes, numbered)
    "Encanto/Lemon Grove1":           "ENCNTO1",
    "Encanto/Lemon Grove 2":          "ENCNTO2",
    "Encanto/Lemon Grove 3":          "ENCNTO3",

    // MTNVIW — Mountain View / National City  (multiple routes, numbered)
    "Mountain View/National City1":   "MTNVCY1",
    "Mountain View/National City2":   "MTNVCY2",
    "Mountain View/National City3":   "MTNVCY3",

    // GRNTVI — Grantville / College East  (single route, no number)
    "Grantville/College East":        "GRNTVI",
    "Grantville/College East1":       "GRNTVI1",
    "Grantville/College East2":       "GRNTVI2",
    "Grantville/College East3":       "GRNTVI3",

    // LMRAY — Lake Murray / La Mesa  (LaKe MuRrAY)
    "Lake Murray/La Mesa1":           "LMRAY1",
    "Lake Murray/La Mesa2":           "LMRAY2",
    "Lake Murray/La Mesa3":           "LMRAY3",

    // LOGANH — Logan Heights  (multiple routes, numbered)
    "Logan Heights1":                 "LOGANH1",
    "Logan Heights2":                 "LOGANH2",
    "Logan Heights3":                 "LOGANH3",

    // WLKDLV — Walking Delivery  (multiple routes, numbered)
    "Walking Delivery 1":             "WLKDLV1",
    "Walking Delivery 2":             "WLKDLV2",
    "Walking Delivery 3":             "WLKDLV3",

    // VOLNTR — Volunteers  (single route — no number unless multiples exist)
    "Volunteers":                     "VOLNTR",
    "Volunteers1":                    "VOLNTR1",
    "Volunteers2":                    "VOLNTR2",
    "Volunteers3":                    "VOLNTR3",
  };

  // Exact match — then format the code into spaced tokens before returning.
  // Stored format:  LOGANH2  or  CHVSTW3  (no spaces)
  // Printed format: LOGANH 2 or  CHVST W 3  (spaces between alpha and number,
  //                                           and before directional if present)
  const code = ROUTE_MAP[route];
  if (code) return formatRouteCode(code);

  // Fallback: return the route as-is so nothing silently breaks when
  // a new route name is added to the spreadsheet before it's added here.
  // When you see an un-abbreviated label, just add it to ROUTE_MAP above.
  return route.replace(/\s+/g, ' ').trim();
}

/**
 * Splits a compact route code into spaced display tokens.
 *
 * Rules:
 *   All-alpha code (no number):          "AZALEA"   → "AZALEA"
 *   Alpha + number:                      "LOGANH2"  → "LOGANH 2"
 *   5-alpha + directional + number:      "CHVSTW3"  → "CHVST W 3"
 *   5-alpha + directional (no number):   "CHVSTW"   → "CHVST W"
 *
 * Detection logic:
 *   A directional is present when the code has exactly one letter immediately
 *   before the trailing digit(s), AND that letter is N, S, E, or W,
 *   AND the alpha prefix before it is 5 characters long.
 */
function formatRouteCode(code) {
  // Split into alpha prefix and optional trailing number
  const match = code.match(/^([A-Z]+?)(\d+)?$/);
  if (!match) return code; // shouldn't happen, but safe fallback

  const alpha = match[1];        // e.g. "CHVSTW" or "LOGANH" or "AZALEA"
  const num   = match[2] || "";  // e.g. "3" or ""

  // Check for directional: last letter of alpha is N/S/E/W AND prefix is 5 chars
  const directionals = new Set(["N","S","E","W"]);
  const lastLetter = alpha.slice(-1);
  const prefix     = alpha.slice(0, -1);

  if (prefix.length === 5 && directionals.has(lastLetter)) {
    // Directional format: "CHVST W 3" or "CHVST W"
    return num ? `${prefix} ${lastLetter} ${num}` : `${prefix} ${lastLetter}`;
  }

  // Standard format: "LOGANH 2" or "AZALEA"
  return num ? `${alpha} ${num}` : alpha;
}

/**
 * DIAGNOSTIC TOOL: 
 * Run this to see exactly what the script is "seeing" for Row 2 
 * before any grid logic or merging happens.
 */
function debugRow2Contents(config = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
  const debugSheetName = "DEBUG_Row2_Contents";
  
  // 1. Setup or Clear the Debug Tab
  let debugSheet = ss.getSheetByName(debugSheetName);
  if (!debugSheet) {
    debugSheet = ss.insertSheet(debugSheetName);
  }
  debugSheet.clear().appendRow(["Recipient Name", "Selected Row 2 Content", "Mode Used"]);

  if (!dataSheet) return;

  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const dataRows = data.slice(1);

  // 2. Map Indices and Helpers
  const nameIdx = headers.indexOf("Name");
  const routeIdx = headers.indexOf("routeDescription");
  const confForIdx = headers.findIndex(h => h.toString().startsWith("Conf for"));
  
  const routeDriverMap = {};
  const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
  if (helperSheet) {
    helperSheet.getDataRange().getValues().slice(1).forEach(row => {
      if (row[0]) routeDriverMap[String(row[0]).trim()] = String(row[1] || 'TBD').trim();
    });
  }

  const debugLog = [];
  const modeLabel = config.includeDriverNames ? "Driver" : 
                    config.includeRouteNames ? "Route" : 
                    config.includeAbbreviatedRouteNames ? "Abbrev" : "None";

  // 3. Process the data exactly like the label maker does
  dataRows.filter(row => {
    const confValue = String(row[confForIdx] || '').toUpperCase();
    return confValue === 'YES' || confValue.includes('NO ANS');
  }).forEach(row => {
    const name = String(row[nameIdx] || 'UNKNOWN');
    const fullRoute = String(row[routeIdx] || 'UNKNOWN ROUTE').trim();
    const driverName = routeDriverMap[fullRoute] || "TBD";
    const abbrevRoute = abbreviateRoute(fullRoute);

    // THE DECISION POINT (Same logic as the label generator)
    let selection = "";
    if (config.includeDriverNames) selection = driverName;
    else if (config.includeRouteNames) selection = fullRoute;
    else if (config.includeAbbreviatedRouteNames) selection = abbrevRoute;

    debugLog.push([name, selection, modeLabel]);
  });

  // 4. Output to Column A (and B for reference)
  if (debugLog.length > 0) {
    debugSheet.getRange(2, 1, debugLog.length, 3).setValues(debugLog);
    debugSheet.autoResizeColumns(1, 3);
    SpreadsheetApp.getUi().alert(`Debug Complete. Check the "${debugSheetName}" tab.`);
  } else {
    SpreadsheetApp.getUi().alert("No confirmed deliveries found during debug.");
  }
}

/**
 * ULTRA-FAST DIAGNOSTIC
 * This bypasses all formatting to avoid the 6-minute timeout.
 * It outputs only the Row 2 content for every delivery to a new tab.
 */
function fastDebugRow2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
  const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
  const debugSheetName = "DEBUG_RESULTS";
  
  // 1. Setup Sheet
  let debugSheet = ss.getSheetByName(debugSheetName) || ss.insertSheet(debugSheetName);
  debugSheet.clear();

  // 2. Load ALL Data into Memory at once (Crucial for speed)
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = headers.indexOf("Name");
  const routeIdx = headers.indexOf("routeDescription");
  const confForIdx = headers.findIndex(h => h.toString().startsWith("Conf for"));

  const routeDriverMap = {};
  if (helperSheet) {
    const helperData = helperSheet.getDataRange().getValues();
    for (let i = 1; i < helperData.length; i++) {
      if (helperData[i][0]) routeDriverMap[String(helperData[i][0]).trim()] = helperData[i][1];
    }
  }

  // 3. Process Logic
  const outputValues = [["RECIPIENT", "ROW 2 CONTENT (ABBREVIATED)"]];
  
  data.slice(1).forEach(row => {
    const confValue = String(row[confForIdx] || '').toUpperCase();
    if (confValue === 'YES' || confValue.includes('NO ANS')) {
      const name = row[nameIdx];
      const fullRoute = String(row[routeIdx] || '').trim();
      
      // We are testing the abbreviation specifically here
      const result = abbreviateRoute(fullRoute); 
      
      outputValues.push([name, result]);
    }
  });

  // 4. One single write operation (Fast!)
  if (outputValues.length > 0) {
    debugSheet.getRange(1, 1, outputValues.length, 2).setValues(outputValues);
    SpreadsheetApp.getUi().alert("Done! Check the DEBUG_RESULTS tab.");
  }
}

// --- Menu Wrappers for Debugging ---
function debugDrivers() { debugRow2Contents({ includeDriverNames: true }); }
function debugRoutes() { debugRow2Contents({ includeRouteNames: true }); }
function debugAbbrev() { debugRow2Contents({ includeAbbreviatedRouteNames: true }); }
function debugNull() { debugRow2Contents }

function generateDeliveryLabels(config = {}) {
    const defaults = {
        includeDriverNames: false,
        includeRouteNames: false,
        includeAbbreviatedRouteNames: false  // All false — caller must opt in explicitly
    };
    const finalConfig = { ...defaults, ...config };
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
    const outputSheetName = "DeliveryLabels";

    if (!dataSheet) {
        SpreadsheetApp.getUi().alert("Source sheet 'Deliveries-UPDATE HERE' not found.");
        return;
    }
    
    const outputSheet = setupOutputSheet(ss, dataSheet, outputSheetName);
    const { headers, dataRows } = getSourceData(dataSheet);
    const indices = getColumnIndices(headers);

    if (!indices.isValid) {
        SpreadsheetApp.getUi().alert("Missing essential columns. Cannot generate labels.");
        return;
    }

    const routeDriverMap = getRouteDriverMap(ss);

    // Generate the labels - pass the config here if your label generator 
    // needs to know about routes early on
    const { allLabels, labelFontSizeMap } = generateAllIndividualLabels(dataRows, indices, routeDriverMap);

    if (allLabels.length === 0) {
        SpreadsheetApp.getUi().alert("No confirmed deliveries found.");
        return;
    }

    // --- STEP 4: This is where the magic happens ---
    // Pass the finalConfig so the formatter knows which string to build
    const { finalOutput, outputLabelMapping, stampSlots } = formatContinuousLabelData(
        allLabels, 
        finalConfig
    );
    
    writeAndFormatContinuousLabels(outputSheet, finalOutput, outputLabelMapping, stampSlots, allLabels, labelFontSizeMap, finalConfig);
    
    SpreadsheetApp.flush();
    return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
}

// --- Wrapper functions for easy use from the menu ---

function generateWithDrivers() {
    // Adding 'return' ensures the timestamp reaches makeLabelsPDF
    return generateDeliveryLabels({ includeDriverNames: true });
}

function generateWithoutDrivers() {
    // Adding 'return' ensures the timestamp reaches makeLabelsPDF
    return generateDeliveryLabels({ includeDriverNames: false });
}

function generateWithRoutes() {
    // Adding 'return' ensures the timestamp reaches makeLabelsPDF
    return generateDeliveryLabels({ includeRouteNames: true });
}

function generateWithAbbreviatedRoutes() {
    // Adding 'return' ensures the timestamp reaches makeLabelsPDF
    return generateDeliveryLabels({ includeAbbreviatedRouteNames: true });
}

// --- Constants for Label Dimensions (Avery 6240) ---
const LABELS_PER_ROW = 3;
const LABEL_HEIGHT_ROWS = 2; // 2 rows of Google Sheets = 1 physical label
const LABEL_WIDTH_COLS = 2;
const TOTAL_COLS = LABELS_PER_ROW * LABEL_WIDTH_COLS; // 6 columns total
const LABEL_ROWS_PER_PAGE = 10;                        // Avery 6240: 10 label rows per sheet
const LABELS_PER_PAGE = LABELS_PER_ROW * LABEL_ROWS_PER_PAGE; // 30 slots per sheet
const DELIVERY_SLOTS_PER_PAGE = LABELS_PER_PAGE - 1;          // 29 delivery slots; slot 30 = stamp

// (setupOutputSheet, getSourceData, getColumnIndices, getRouteDriverMap, calculateBagCount remain the same)

/**
 * Handles sheet creation/retrieval and clearing.
 */
function setupOutputSheet(ss, dataSheet, outputSheetName) {
    let outputSheet = ss.getSheetByName(outputSheetName);
    
    // Calculate the target index to place the new sheet after the data sheet
    const targetIndex = dataSheet.getIndex() + 1; 

    if (!outputSheet) {
        outputSheet = ss.insertSheet(outputSheetName, targetIndex); 
    }
    
    // Clear and reset the sheet
    outputSheet.clear();
    outputSheet.setFrozenRows(0);
    outputSheet.setFrozenColumns(0);
    
    return outputSheet;
}

/**
 * Gets all values from the data sheet, excluding headers.
 */
function getSourceData(dataSheet) {
    const data = dataSheet.getDataRange().getValues();
    return {
        headers: data[0],
        dataRows: data.slice(1)
    };
}

/**
 * Finds and returns all necessary column indices.
 */
function getColumnIndices(headers) {
    const indices = {};
    indices.nameIdx = headers.indexOf("Name");
    indices.routeIdx = headers.indexOf("routeDescription");
    
    // Locate the "Conf for" column (case-insensitive find)
    const confForHeader = headers.find(h => typeof h === 'string' && h.trim().startsWith("Conf for"));
    indices.confForIdx = confForHeader ? headers.indexOf(confForHeader) : -1;

    // Indices for Bag Count Calculation
    indices.diapers1Idx = headers.indexOf("diapers1(size)");
    indices.diapers2Idx = headers.indexOf("diapers2(size)");
    indices.diapers3Idx = headers.indexOf("diapers3(size)");
    indices.wipesIdx = headers.indexOf("wipes(qty)");
    indices.dogFoodIdx = headers.indexOf("dogFood(qty)");
    indices.catFoodIdx = headers.indexOf("catFood(qty)");
    
    indices.isValid = indices.nameIdx !== -1 && indices.routeIdx !== -1 && indices.confForIdx !== -1;
    
    return indices;
}

/**
 * Builds a map of route to driver from the helper sheet.
 */
function getRouteDriverMap(ss) {
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
    return routeDriverMap;
}

/**
 * Calculates the number of bags needed for a single delivery row.
 */
function calculateBagCount(row, indices) {
    let bagCount = 0;
    const MAX_LABELS = 3;
    
    // A. Diapers: 1 bag per non-empty column (max 3)
    if (indices.diapers1Idx !== -1 && String(row[indices.diapers1Idx]).trim() !== "") bagCount++;
    if (indices.diapers2Idx !== -1 && String(row[indices.diapers2Idx]).trim() !== "") bagCount++;
    if (indices.diapers3Idx !== -1 && String(row[indices.diapers3Idx]).trim() !== "") bagCount++;

    // B. Wipes: 1 bag per whole quantity (Math.ceil handles decimal quantities)
    const wipesQty = parseFloat(row[indices.wipesIdx] || 0) || 0;
    bagCount += Math.ceil(wipesQty);
    
    // C. Pet Food: 1 bag per whole quantity
    const dogFoodQty = parseFloat(row[indices.dogFoodIdx] || 0) || 0;
    const catFoodQty = parseFloat(row[indices.catFoodIdx] || 0) || 0;
    bagCount += Math.ceil(dogFoodQty) + Math.ceil(catFoodQty);
    
    // Enforce hard limit
    return Math.min(bagCount, MAX_LABELS);
}



/**
 * NEW: Processes deliveries and generates a flat array of required individual labels.
 */
function generateAllIndividualLabels(dataRows, indices, routeDriverMap) {
    const allLabels = [];
    // Stores formatting keyed by the sequential index of the delivery the label belongs to.
    const labelFontSizeMap = {}; 
    let deliveryIndexCounter = 0;

    const filteredDataRows = dataRows
        .filter(row => {
            const confValue = String(row[indices.confForIdx] || '').toUpperCase().trim();
            return confValue === 'YES' || confValue.includes('NO ANS');
        });

    filteredDataRows.forEach(row => {
        const name = String(row[indices.nameIdx] || 'UNKNOWN NAME');
        const route = String(row[indices.routeIdx] || 'UNKNOWN ROUTE').trim();
        const driver = routeDriverMap[route] || "TBD";
        const totalLabels = calculateBagCount(row, indices);
        
        if (totalLabels > 0) {
            // Determine Font Sizes
            const baseNameFontSize = 12;
            const reducedNameFontSize = 10;
            const maxNameLength = 20; 
            const nameFontSize = name.length > maxNameLength ? reducedNameFontSize : baseNameFontSize;
            
            const baseDriverFontSize = 13;
            const reducedDriverFontSize = 11;
            const maxDriverLength = 20;
            const driverFontSize = driver.length > maxDriverLength ? reducedDriverFontSize : baseDriverFontSize;

            const abbrevRoute = abbreviateRoute(route);

            // Generate an entry for *each* required label
            for (let i = 1; i <= totalLabels; i++) {
                allLabels.push({
                    Name: name,
                    Driver: driver,
                    Route: route,
                    AbbrevRoute: abbrevRoute,
                    BagNumber: i,
                    TotalLabels: totalLabels,
                    DeliveryIndex: deliveryIndexCounter // Reference for formatting map
                });
            }

            // Store font size metadata keyed by the delivery index
            labelFontSizeMap[deliveryIndexCounter] = { 
                name: nameFontSize,
                driver: driverFontSize
            };
            deliveryIndexCounter++;
        }
    });

    return { allLabels, labelFontSizeMap };
}

/**
 * NEW: Formats the flat list of labels into the final continuous 2D array.
 * config must have: includeDriverNames, includeRouteNames, includeAbbreviatedRouteNames
 */
/**
 * Builds the 2D array written to the sheet.
 *
 * Page layout (Avery 6240, 5 rows × 3 cols = 15 slots per page):
 *   Slots 1–14  delivery labels
 *   Slot  15    stamp label — always bottom-right corner, every page
 *
 * If a delivery label would land in slot 15 it is bumped to the next page,
 * keeping slot 15 clear for the stamp on every page without exception.
 *
 * outputLabelMapping: array parallel to finalOutput row-pairs.
 *   value ≥ 0  → DeliveryIndex of first label in that row (for font sizing)
 *   value = -1 → row is filler or stamp (no delivery formatting needed)
 *   value = -2 → row contains the stamp label in its rightmost slot
 */
function formatContinuousLabelData(allLabels, config) {
    const finalOutput      = [];
    const outputLabelMapping = [];
    const stampSlots       = new Set(); // sheet row-pair indices that carry a stamp

    const timestamp = Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy/MM/dd HH:mm:ss"
    );

    let labelIndex = 0;   // position in allLabels
    let pageNumber = 1;
    let slotOnPage = 0;   // 0-based slot within current page (0–14)

    // Helper: push one label-row (3 slots = 6 columns) to finalOutput
    function pushLabelRow(slots) {
        // slots: array of 3 objects { row1a, row1b, row2 }
        const row1 = [], row2 = [];
        slots.forEach(s => { row1.push(s.row1a, s.row1b); row2.push(s.row2, ""); });
        finalOutput.push(row1, row2);
        outputLabelMapping.push(slots[0].deliveryIndex ?? -1);
    }

    function emptySlot()  { return { row1a: "", row1b: "", row2: "", deliveryIndex: -1 }; }
    function stampSlot(pg, ts) {
        return {
            row1a: `Page ${pg}`,
            row1b: "",
            row2:  ts,
            deliveryIndex: -2   // sentinel: stamp
        };
    }
    function deliverySlot(label, config) {
        let detailLine = "";
        if (config.includeDriverNames)           detailLine = label.Driver;
        else if (config.includeRouteNames)       detailLine = label.Route;
        else if (config.includeAbbreviatedRouteNames) detailLine = label.AbbrevRoute;
        return {
            row1a: label.Name,
            row1b: `BAG ${label.BagNumber}/${label.TotalLabels}`,
            row2:  detailLine,
            deliveryIndex: label.DeliveryIndex
        };
    }

    // Build pages until all delivery labels are placed
    while (labelIndex < allLabels.length) {
        const pageSlots = []; // will hold up to 15 slot-objects for this page

        for (let s = 0; s < LABELS_PER_PAGE; s++) {
            if (s === DELIVERY_SLOTS_PER_PAGE) {
                // Slot 15 (0-based: 14) is always the stamp
                pageSlots.push(stampSlot(pageNumber, timestamp));
            } else if (labelIndex < allLabels.length) {
                pageSlots.push(deliverySlot(allLabels[labelIndex++], config));
            } else {
                pageSlots.push(emptySlot());
            }
        }

        // Write the 5 label-rows for this page (3 slots each)
        const stampRowPairIndex = (finalOutput.length / 2) + (LABEL_ROWS_PER_PAGE - 1);
        for (let row = 0; row < LABEL_ROWS_PER_PAGE; row++) {
            const s = pageSlots.slice(row * LABELS_PER_ROW, row * LABELS_PER_ROW + LABELS_PER_ROW);
            pushLabelRow(s);
            if (row === LABEL_ROWS_PER_PAGE - 1) {
                // Mark the last row-pair of this page as containing the stamp
                stampSlots.add(finalOutput.length / 2 - 1); // row-pair index just pushed
            }
        }

        pageNumber++;
    }

    // If allLabels was empty we still need at least one page of stamps
    if (finalOutput.length === 0) {
        const pageSlots = [];
        for (let s = 0; s < LABELS_PER_PAGE; s++) {
            pageSlots.push(s === DELIVERY_SLOTS_PER_PAGE ? stampSlot(1, timestamp) : emptySlot());
        }
        for (let row = 0; row < LABEL_ROWS_PER_PAGE; row++) {
            const s = pageSlots.slice(row * LABELS_PER_ROW, row * LABELS_PER_ROW + LABELS_PER_ROW);
            pushLabelRow(s);
        }
        stampSlots.add(LABEL_ROWS_PER_PAGE - 1);
    }

    return { finalOutput, outputLabelMapping, stampSlots };
}

/**
 * Writes data to the sheet and applies formatting.
 * stampSlots: Set of row-pair indices (0-based) whose rightmost slot is a stamp label.
 */
function writeAndFormatContinuousLabels(outputSheet, finalOutput, outputLabelMapping, stampSlots, allLabels, labelFontSizeMap, config) {
    
    // 1. Write Data to Sheet
    outputSheet.getRange(1, 1, finalOutput.length, TOTAL_COLS).setValues(finalOutput);

    // 2. Set column widths
    const width = 258; 
    outputSheet.setColumnWidth(1, width * 0.7); 
    outputSheet.setColumnWidth(2, width * 0.3);
    outputSheet.setColumnWidth(3, width * 0.7);
    outputSheet.setColumnWidth(4, width * 0.3);
    outputSheet.setColumnWidth(5, width * 0.7);
    outputSheet.setColumnWidth(6, width * 0.3);
    
    let labelCounter = 0; // tracks position in allLabels
    let rowPairIndex = 0; // tracks which row-pair we are formatting

    // 3. Loop through output data rows (2 sheet rows per physical label row)
    for (let r = 1; r <= finalOutput.length; r += LABEL_HEIGHT_ROWS) {

        // Set row heights
        outputSheet.setRowHeight(r,     51.5); 
        outputSheet.setRowHeight(r + 1, 51.5); 

        // General block formatting
        const labelBlock = outputSheet.getRange(r, 1, LABEL_HEIGHT_ROWS, TOTAL_COLS);
        labelBlock.setBorder(false, false, false, false, false, false); 
        labelBlock.setFontWeight("bold").setVerticalAlignment("middle").setHorizontalAlignment("center");

        const isStampRow = stampSlots.has(rowPairIndex);

        // Format each of the 3 label slots in this row pair
        for (let j = 0; j < LABELS_PER_ROW; j++) {
            const colNameStart = j * 2 + 1; // 1, 3, 5
            const colBagStart  = j * 2 + 2; // 2, 4, 6
            const isStampSlot  = isStampRow && (j === LABELS_PER_ROW - 1); // rightmost slot of stamp row

            if (isStampSlot) {
                // ── STAMP LABEL ───────────────────────────────────────────────
                // Spans the full 2-col slot; row 1 = "Page N", row 2 = timestamp
                const stampRow1 = outputSheet.getRange(r,     colNameStart, 1, 2);
                const stampRow2 = outputSheet.getRange(r + 1, colNameStart, 1, 2);

                if (!stampRow1.isPartOfMerge()) stampRow1.mergeAcross();
                if (!stampRow2.isPartOfMerge()) stampRow2.mergeAcross();

                stampRow1
                    .setFontSize(14)
                    .setFontWeight("bold")
                    .setFontColor("#000000")
                    .setHorizontalAlignment("center")
                    .setVerticalAlignment("middle");

                stampRow2
                    .setFontSize(9)
                    .setFontWeight("normal")
                    .setFontColor("#000000")
                    .setHorizontalAlignment("center")
                    .setVerticalAlignment("middle")
                    .setWrap(false);

            } else if (labelCounter < allLabels.length) {
                // ── DELIVERY LABEL ────────────────────────────────────────────
                const currentLabel = allLabels[labelCounter];
                const formatting   = labelFontSizeMap[currentLabel.DeliveryIndex];
                const nameFontSize = formatting ? formatting.name : 12;

                let row2Content = "";
                if (config.includeDriverNames)               row2Content = currentLabel.Driver;
                else if (config.includeRouteNames)           row2Content = currentLabel.Route;
                else if (config.includeAbbreviatedRouteNames) row2Content = currentLabel.AbbrevRoute;

                const detailFontSize = row2Content.length > 20 ? 11 : 13;

                const row2Range = outputSheet.getRange(r + 1, colNameStart, 1, 2);
                if (!row2Range.isPartOfMerge()) row2Range.mergeAcross();

                outputSheet.getRange(r, colNameStart, 1, 1)
                    .setFontSize(nameFontSize)
                    .setWrap(false)
                    .setFontColor("#000000")
                    .setHorizontalAlignment("center");

                outputSheet.getRange(r, colBagStart, 1, 1)
                    .setFontSize(10)
                    .setFontWeight("normal")
                    .setHorizontalAlignment("left");

                row2Range
                    .setFontSize(detailFontSize)
                    .setFontColor(row2Content !== "" ? "#3C78D8" : "#000000")
                    .setHorizontalAlignment("center");

                labelCounter++;
            } else {
                // ── EMPTY FILLER ──────────────────────────────────────────────
                // Nothing to format; labelCounter intentionally not advanced
                // (filler slots don't consume allLabels entries)
            }
        }

        rowPairIndex++;
    }
    
    // 4. Safely delete excess rows and columns
    const maxRows = outputSheet.getMaxRows();
    const rowsToDelete = maxRows - finalOutput.length;
    if (rowsToDelete > 0) {
        outputSheet.deleteRows(finalOutput.length + 1, rowsToDelete);
    }
    
    const maxCols = outputSheet.getMaxColumns();
    const colsToDelete = maxCols - TOTAL_COLS;
    if (colsToDelete > 0) {
        outputSheet.deleteColumns(TOTAL_COLS + 1, colsToDelete);
    }
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

function makePackingListThenPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. CAPTURE THE TIMESTAMP from the first function
    //const timestamp = generateSpecialItemsCoordinator();
    const timestamp = makePackingLists(); 
    
    if (timestamp === "ERROR_NO_SHEET") return; // Exit if sheet was not found

    SpreadsheetApp.flush(); // Ensure all sheet changes are committed before PDF generation
    
    // 2. SANITIZE & CONSTRUCT FILENAME
    // Remove characters from the timestamp that are illegal or problematic in file paths (like ':', '/')
    const sanitizedTimestamp = timestamp.replace(/[\/:]/g, '-').replace(/\s/g, '_');
    const filename = `Packing_Lists_${sanitizedTimestamp}.pdf`;

    const sheet = ss.getSheetByName("PackingLists");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Error: 'PackingLists' sheet not found for PDF export.");
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

/**
 * Asks which label mode to use, generates the sheet, then exports to PDF.
 * Called directly from the menu OR from makeAllPDFs (which passes a mode).
 * @param {string} [forcedMode] - optional: "drivers"|"routes"|"abbreviated"|"none"
 *                                Pass this from makeAllPDFs to skip the prompt.
 */
function makeLabelsPDF(forcedMode) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // 1. Determine mode — prompt the user unless a mode was passed in
    let mode = forcedMode;
    if (!mode) {
      const response = ui.prompt(
        '📦 Label Mode',
        'Which label style?\n\n' +
        '  1 — Driver names\n' +
        '  2 — Full route names\n' +
        '  3 — Abbreviated route codes\n' +
        '  4 — No extra info (name + bag count only)\n\n' +
        'Type a number (1–4) and click OK:',
        ui.ButtonSet.OK_CANCEL
      );
      if (response.getSelectedButton() !== ui.Button.OK) return;
      const choice = response.getResponseText().trim();
      const modeMap = { '1':'drivers', '2':'routes', '3':'abbreviated', '4':'none' };
      mode = modeMap[choice];
      if (!mode) {
        ui.alert('Invalid choice. Please enter 1, 2, 3, or 4.');
        return;
      }
    }

    // 2. Generate the label sheet in the chosen mode
    const modeConfig = {
      drivers:      { includeDriverNames: true },
      routes:       { includeRouteNames: true },
      abbreviated:  { includeAbbreviatedRouteNames: true },
      none:         {}
    };
    const timestamp = generateDeliveryLabels(modeConfig[mode]);

    if (!timestamp || timestamp === "ERROR_NO_SHEET") {
      console.log("PDF generation cancelled: No timestamp returned.");
      return; 
    }

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

/**
 * Generates the DeliveryLabels sheet only (no PDF). Prompts for mode.
 * Wired to Distribution Day Tasks menu item 5.
 */
function makeTheLabels() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '📦 Label Mode',
    'Which label style?\n\n' +
    '  1 — Driver names\n' +
    '  2 — Full route names\n' +
    '  3 — Abbreviated route codes\n' +
    '  4 — No extra info (name + bag count only)\n\n' +
    'Type a number (1–4) and click OK:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const choice = response.getResponseText().trim();
  const modeMap = {
    '1': { includeDriverNames: true },
    '2': { includeRouteNames: true },
    '3': { includeAbbreviatedRouteNames: true },
    '4': {}
  };
  const config = modeMap[choice];
  if (!config) {
    ui.alert('Invalid choice. Please enter 1, 2, 3, or 4.');
    return;
  }
  generateDeliveryLabels(config);
}

function makeAllPDFs() {
  makeDriverRouteTabsThenPDF();
  makePackingListThenPDF();
  // Use abbreviated route codes as the default for the bulk "make everything" run.
  // Change the argument to 'drivers', 'routes', or 'none' if preferred.
  makeLabelsPDF('abbreviated');
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
    .addItem("1. Filter Data for Map (Col I)", "filterForMapCreation")
    .addItem("2. Clear Map Data Filter", "clearMapFilter")
    .addSeparator()
    .addItem("3. Make/Update Driver Route Tabs", "makeDriverRouteTabs")
    .addItem("4. Make/Update Packing Lists Tab", "makePackingLists")
    .addItem("5. Make/Update Delivery Labels Tab", "makeTheLabels")
    .addSeparator()
    .addItem("6. Make Driver Route Tabs then PDF", "makeDriverRouteTabsThenPDF")
    .addItem("7. Make Packing Lists Tab then PDF", "makePackingListThenPDF")
    .addItem("8. Make Delivery Labels Tab then PDF", "makeLabelsPDF")
    .addSeparator()
    .addItem("9. Make All PDFs", "makeAllPDFs")
    .addToUi();

  // Label Generator: each item generates the sheet only (no PDF prompt).
  // Use item 8 above or the PDF option below to also export.
  ui.createMenu("📦 Label Generator")
    .addItem("Labels — Driver Names",          "generateWithDrivers")
    .addItem("Labels — Full Route Names",       "generateWithRoutes")
    .addItem("Labels — Abbreviated Route Codes","generateWithAbbreviatedRoutes")
    .addItem("Labels — No Route Info",          "generateWithoutDrivers")
    .addSeparator()
    .addItem("Labels — Choose Mode + Export PDF","makeLabelsPDF")
    .addToUi();

  ui.createMenu("📦 Sheet Tools")
    .addItem("Delete Route Sheets", "deleteSheetsByPrefix")
    .addToUi();
}