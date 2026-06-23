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

/*******************************************************************************
 DRIVER ROUTE TAB GENERATION — BATCHED CONTINUATION SYSTEM
 
 With 27 routes at ~13s per insertSheet, generation takes ~378s — over the
 6-minute Apps Script limit. This system splits work into batches of 8 routes,
 each batch auto-triggering the next via ScriptApp.newTrigger + PropertiesService.

 Entry points:
   makeDriverRouteTabs()     — menu item 3, always starts fresh
   continueDriverRouteTabs() — called by auto-trigger, never by user directly
   cancelDriverRouteTabs()   — emergency stop, clears state + triggers
*******************************************************************************/

// How many routes to process per execution. 8 × ~14s = ~112s, well within 360s.
const ROUTE_BATCH_SIZE = 8;

// MAP LINK SOURCE — flip this to switch behavior for route sheet header map links:
//   true  = auto-generate waypoint map from stop addresses (no manual maintenance needed)
//   false = pull from "Delivery Map Links" column in source sheet (legacy, manual)
// The index sheet (Driver_Deliveries) always auto-generates its map links regardless.
const AUTO_GENERATE_MAP_LINKS = true;

// PropertiesService keys
const PROP_NEXT_IDX     = "WAWG_ROUTE_NEXT_IDX";
const PROP_SUMMARY_ROWS = "WAWG_ROUTE_SUMMARY_ROWS";
const PROP_ROUTE_NAMES  = "WAWG_ROUTE_NAMES";
const PROP_IN_PROGRESS  = "WAWG_ROUTE_IN_PROGRESS";
const PROP_TRIGGER_ID   = "WAWG_ROUTE_TRIGGER_ID";

/**
 * Menu item 3 — always starts fresh.
 * Cleans up previous sheets + any stale state, then kicks off batch 1.
 */
function makeDriverRouteTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = Date.now();
  const lap = (label) => Logger.log(`⏱ +${((Date.now()-t0)/1000).toFixed(1)}s — ${label}`);

  // Cancel any in-progress batch run before starting fresh
  _routeBatch_cancelTriggers();
  _routeBatch_clearState();

  refreshDeliveriesHelpTable();
  const DATA_SHEET_NAME    = "Deliveries-UPDATE HERE";
  const ROUTE_SHEET_PREFIX = "Route_";

  lap("start — refreshDeliveriesHelpTable done");

  // 1) cleanup old route sheets
  cleanupPreviousRouteSheets(ss, ROUTE_SHEET_PREFIX);
  lap("1) cleanup done");

  // 2–7) load, filter, group, pre-compute — same as before
  const prepared = _routeBatch_prepareData(ss, DATA_SHEET_NAME);
  if (!prepared) return; // alerts already shown inside
  const { routeNames } = prepared;
  lap(`prep done — ${routeNames.length} routes`);

  // Save state for continuation batches
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_ROUTE_NAMES,  JSON.stringify(routeNames));
  props.setProperty(PROP_SUMMARY_ROWS, JSON.stringify([["Route", "Driver", "Stops", "Open Sheet", "Map"]]));
  props.setProperty(PROP_NEXT_IDX,     "0");
  props.setProperty(PROP_IN_PROGRESS,  "true");

  // Run batch 1 immediately
  const nextIdx = _routeBatch_runBatch(ss, prepared, 0);
  lap(`batch 1 done — processed routes 0–${nextIdx - 1}`);

  if (nextIdx >= routeNames.length) {
    // All routes fit in one batch — finish now
    _routeBatch_finish(ss);
    lap("all routes done in one batch");
  } else {
    // Schedule continuation
    props.setProperty(PROP_NEXT_IDX, String(nextIdx));
    _routeBatch_scheduleTrigger();
    const batchNum  = Math.ceil(nextIdx / ROUTE_BATCH_SIZE);
    const totalBatches = Math.ceil(routeNames.length / ROUTE_BATCH_SIZE);
    SpreadsheetApp.getUi().alert(
      `✅ Batch ${batchNum} of ${totalBatches} complete (routes 1–${nextIdx} of ${routeNames.length}).

` +
      `The remaining ${routeNames.length - nextIdx} routes will generate automatically in ~90 seconds.

` +
      `You'll see a "✅ All route tabs complete!" alert when finished. Do not run item 3 again until then.`
    );
  }
}

/**
 * Called automatically by the trigger — never call from menu.
 * Picks up from where the last batch left off.
 */
function continueDriverRouteTabs() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const t0    = Date.now();
  const lap   = (label) => Logger.log(`⏱ +${((Date.now()-t0)/1000).toFixed(1)}s — ${label}`);

  // Safety check — bail if no job is running
  if (props.getProperty(PROP_IN_PROGRESS) !== "true") {
    Logger.log("continueDriverRouteTabs: no job in progress, exiting.");
    _routeBatch_cancelTriggers();
    return;
  }

  const routeNames = JSON.parse(props.getProperty(PROP_ROUTE_NAMES) || "[]");
  const nextIdx    = parseInt(props.getProperty(PROP_NEXT_IDX) || "0", 10);
  Logger.log(`continueDriverRouteTabs: picked up at index ${nextIdx} of ${routeNames.length}`);

  if (nextIdx >= routeNames.length) {
    // Nothing left — finish
    _routeBatch_finish(ss);
    return;
  }

  // Cancel the trigger that fired us (we'll create a new one if needed)
  _routeBatch_cancelTriggers();

  lap(`continuation batch starting at route index ${nextIdx}`);

  const prepared = _routeBatch_prepareData(ss, "Deliveries-UPDATE HERE");
  if (!prepared) {
    _routeBatch_clearState();
    return;
  }

  const newNextIdx = _routeBatch_runBatch(ss, prepared, nextIdx);
  lap(`batch done — processed up to index ${newNextIdx - 1}`);

  if (newNextIdx >= routeNames.length) {
    _routeBatch_finish(ss);
    const totalBatches = Math.ceil(routeNames.length / ROUTE_BATCH_SIZE);
    Logger.log(`All ${routeNames.length} routes complete after ${totalBatches} batches.`);
  } else {
    props.setProperty(PROP_NEXT_IDX, String(newNextIdx));
    _routeBatch_scheduleTrigger();
    const batchNum     = Math.ceil(newNextIdx / ROUTE_BATCH_SIZE);
    const totalBatches = Math.ceil(routeNames.length / ROUTE_BATCH_SIZE);
    Logger.log(`Batch ${batchNum}/${totalBatches} done — scheduled next batch at index ${newNextIdx}`);
  }
}

/**
 * Emergency stop — clears all state and pending triggers.
 * Wire to a menu item if needed for troubleshooting.
 */
function cancelDriverRouteTabs() {
  _routeBatch_cancelTriggers();
  _routeBatch_clearState();
  _routeBatch_uiAlert("Route tab generation cancelled. Run item 3 to start fresh.");
}

// ── Private helpers ───────────────────────────────────────────────────────────

/**
 * Loads and pre-computes everything needed to build route sheets.
 * Returns null if data is missing (alerts shown internally).
 */
function _routeBatch_prepareData(ss, dataSheetName) {
  const { dataSheet, headers, dataRows } = loadDriverRouteData(ss, dataSheetName);
  if (!dataSheet || headers.length === 0) {
    _routeBatch_uiAlert(`Data sheet "${dataSheetName}" not found or empty.`);
    return null;
  }
  if (dataRows.length === 0) {
    _routeBatch_uiAlert("No data rows found in the data sheet.");
    return null;
  }

  const confForHeader = headers.find(h => typeof h === "string" && h.trim().startsWith("Conf for"));
  const confForIdx    = confForHeader ? headers.indexOf(confForHeader) : -1;
  const filteredRows  = filterDriverRouteRows(dataRows, confForIdx);
  const routeDriverMap = buildRouteDriverMap(ss, "DeliveriesHelpTable");
  const grouped        = groupRowsByRoute(filteredRows, headers, "routeDescription");
  const routeNames     = Object.keys(grouped).sort();

  // Build route map info — either auto-generated from stop addresses or
  // pulled from the manual "Delivery Map Links" column, controlled by the toggle.
  const routeMapInfo = {};

  if (AUTO_GENERATE_MAP_LINKS) {
    // Auto path: build waypoint URLs from addresses. Shortening happens per-route
    // inside buildRouteMapInfo via shortenUrlForTyping (~1s each, fine in batches).
    // We defer actual building to _routeBatch_runBatch so the short URL call
    // is spread across executions rather than all happening in prepareData.
    // Mark as deferred — runBatch will call buildRouteMapInfo per route.
    routeNames.forEach(routeKey => {
      routeMapInfo[routeKey] = null; // sentinel — resolved in runBatch
    });
  } else {
    // Manual path: read from source sheet column (legacy behavior)
    const allSheetValues = dataSheet.getDataRange().getValues();
    const mapLinkColIdx  = headers.indexOf("Delivery Map Links");
    const richTextColumn = mapLinkColIdx !== -1
      ? dataSheet.getRange(1, mapLinkColIdx + 1, dataSheet.getLastRow(), 1).getRichTextValues()
      : null;
    routeNames.forEach(routeKey => {
      routeMapInfo[routeKey] = extractMapLinkForRoute({
        headers, dataRows, allSheetValues, richTextColumn,
        routeValue: routeKey,
        deliveryMapColumnHeader: "Delivery Map Links"
      });
    });
  }

  const diapers1Idx = headers.indexOf("diapers1(size)");
  const diapers2Idx = headers.indexOf("diapers2(size)");
  const diapers3Idx = headers.indexOf("diapers3(size)");
  const itemMap = [
    { label: "food",       sourceHeader: "# packs",         type: "qty" },
    { label: "eggs",       sourceHeader: "# eggs",          type: "qty" },
    { label: "milk",       sourceHeader: "# milk",          type: "qty" },
    { label: "dog food",   sourceHeader: "dogFood(qty)",     type: "qty" },
    { label: "cat food",   sourceHeader: "catFood(qty)",     type: "qty" },
    { label: "diapers",    sourceIndex: [diapers1Idx, diapers2Idx, diapers3Idx], type: "diapers" },
    { label: "wipes",      sourceHeader: "wipes(qty)",       type: "qty" },
    { label: "toiletries", sourceHeader: "toiletries(qty)",  type: "qty" },
    { label: "fem hygn",   sourceHeader: "fem hygiene(qty)", type: "qty" },
  ];

  return { dataSheet, headers, dataRows, routeNames, routeDriverMap, grouped, routeMapInfo, itemMap };
}

/**
 * Processes ROUTE_BATCH_SIZE routes starting at startIdx.
 * Appends to PROP_SUMMARY_ROWS in PropertiesService.
 * Returns the next unprocessed index.
 */
function _routeBatch_runBatch(ss, prepared, startIdx) {
  const { headers, routeNames, routeDriverMap, grouped, routeMapInfo, itemMap } = prepared;
  const props       = PropertiesService.getScriptProperties();
  const summaryRows = JSON.parse(props.getProperty(PROP_SUMMARY_ROWS) || "[[]]");
  const endIdx      = Math.min(startIdx + ROUTE_BATCH_SIZE, routeNames.length);
  const t0          = Date.now();

  for (let idx = startIdx; idx < endIdx; idx++) {
    const route      = routeNames[idx];
    const tRoute     = Date.now();
    const stopCount  = Object.keys(grouped[route].personStops).length;
    Logger.log(`  → route ${idx+1}/${routeNames.length}: "${route}" (${stopCount} stops)`);

    const safeName    = route.replace(/[^a-zA-Z0-9]/g, "_").substring(0, 30);
    const sheetName   = "Route_" + safeName;
    const routeColor  = pickRouteColor(idx);
    const driverName  = routeDriverMap[route] || "TBD";
    const personStops = grouped[route].personStops;
    // Resolve map info — auto-generate if toggle is on and sentinel is null
    let mapInfo = routeMapInfo[route];
    if (AUTO_GENERATE_MAP_LINKS || mapInfo === null) {
      mapInfo = buildRouteMapInfo(personStops, route, headers);
    }
    mapInfo = mapInfo || { link: "", label: "" };

    const sheet = ss.insertSheet(sheetName);
    sheet.clearFormats();
    try { sheet.setTabColor(routeColor); } catch(e) {}
    Logger.log(`    insertSheet+clearFormats: +${((Date.now()-tRoute)/1000).toFixed(1)}s`);

    try {
      buildRouteSheet({
        ss, sheet, route, driverName, personStops,
        itemMap, finalItemHeaders: itemMap.map(i => i.label), headers,
        routeColor, stopColors: pickStopColors(),
        colIndices: {
          nameCol:    headers.indexOf("Name"),
          addressCol: headers.indexOf("Address"),
          phoneCol:   headers.indexOf("Phone"),
          specCol:    headers.indexOf("Special Item Requests"),
          delivCol:   headers.indexOf("Special Delivery Instructions")
        },
        mapInfo
      });

      // For the index sheet, always auto-generate the waypoint URL (no shortening needed there)
      const stopAddresses = Object.values(personStops).map(rows => {
        const r   = rows[0];
        const addrIdx = headers.indexOf("Address");
        const raw = addrIdx !== -1 ? String(r[addrIdx] || "").trim() : "";
        // Strip route name from address
        return raw.split(",").map(p => p.trim()).filter(p => p !== route).join(", ");
      }).filter(a => a.length > 0);
      const autoMapUrl = buildAutoMapUrl(stopAddresses, "");
      summaryRows.push([
        route, driverName, stopCount,
        `=HYPERLINK("#gid=${sheet.getSheetId()}","Open")`,
        autoMapUrl ? `=HYPERLINK("${autoMapUrl}","View Map")` : "No addresses"
      ]);
      Logger.log(`    route done: +${((Date.now()-tRoute)/1000).toFixed(1)}s  cumulative: +${((Date.now()-t0)/1000).toFixed(1)}s`);
    } catch (err) {
      Logger.log(`    ❌ Error on route "${route}": ${err}`);
      try { ss.deleteSheet(sheet); } catch(e) {}
    }
  }

  // Persist updated summary rows
  props.setProperty(PROP_SUMMARY_ROWS, JSON.stringify(summaryRows));
  return endIdx;
}

/**
 * Called when all batches are done.
 * Writes the index sheet, cleans up state + triggers, alerts the user.
 */
function _routeBatch_finish(ss) {
  const props       = PropertiesService.getScriptProperties();
  const summaryRows = JSON.parse(props.getProperty(PROP_SUMMARY_ROWS) || "[[]]");

  writeRouteIndexSheet(ss, "Driver_Deliveries", summaryRows);
  _routeBatch_clearState();
  _routeBatch_cancelTriggers();

  // Write a visible completion stamp to the index sheet so it's obvious
  // even when this runs from a trigger and can't show a UI alert.
  try {
    const indexSheet = ss.getSheetByName("Driver_Deliveries");
    if (indexSheet) {
      const ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
      indexSheet.getRange(1, 7).setValue(`✅ Complete: ${ts}`).setFontWeight("bold").setBackground("#B6D7A8");
    }
  } catch(e) { Logger.log(`Status write error: ${e}`); }

  Logger.log("✅ All route tabs complete.");
  _routeBatch_uiAlert("✅ All route tabs complete! Check your Route_ sheets and Driver_Deliveries index.");
}

/**
 * Safe alert wrapper — shows UI alert when called from menu context,
 * falls back to Logger when called from a trigger (no UI available).
 */
function _routeBatch_uiAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    Logger.log(`[alert] ${msg}`);
  }
}

/** Schedule a one-time trigger to call continueDriverRouteTabs in 2 minutes. */
function _routeBatch_scheduleTrigger() {
  const trigger = ScriptApp.newTrigger("continueDriverRouteTabs")
    .timeBased()
    .after(30 * 1000) // 30 seconds
    .create();
  PropertiesService.getScriptProperties().setProperty(PROP_TRIGGER_ID, trigger.getUniqueId());
  Logger.log(`Trigger scheduled: ${trigger.getUniqueId()} (fires in ~30s)`);
}

/** Cancel all WAWG route triggers (by function name, in case ID tracking drifts). */
function _routeBatch_cancelTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "continueDriverRouteTabs") {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
  PropertiesService.getScriptProperties().deleteProperty(PROP_TRIGGER_ID);
}

/** Clear all batch state from PropertiesService. */
function _routeBatch_clearState() {
  const props = PropertiesService.getScriptProperties();
  [PROP_NEXT_IDX, PROP_SUMMARY_ROWS, PROP_ROUTE_NAMES, PROP_IN_PROGRESS, PROP_TRIGGER_ID]
    .forEach(k => props.deleteProperty(k));
}


/* -------------------------
   Helper: cleanupPreviousRouteSheets
   Deletes all sheets whose name starts with the given prefix.
------------------------- */
function cleanupPreviousRouteSheets(ss, prefix) {
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    try {
      if (sh.getName().startsWith(prefix)) ss.deleteSheet(sh);
    } catch (e) {
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
  const {
    headers, dataRows, routeValue, deliveryMapColumnHeader,
    allSheetValues,   // pre-loaded getValues() — required
    richTextColumn    // pre-loaded getRichTextValues() for the map link column — required
  } = opts || {};
  if (!headers || !dataRows || !allSheetValues || !richTextColumn) return { link: "", label: "" };

  const deliveryMapLinkIndex = headers.indexOf(deliveryMapColumnHeader);
  if (deliveryMapLinkIndex === -1) return { link: "", label: "" };

  const routeIdx = headers.indexOf("routeDescription");
  if (routeIdx === -1) return { link: "", label: "" };

  const matchRow = dataRows.find(r => String(r[routeIdx]).trim() === routeValue);
  if (!matchRow) return { link: "", label: "" };

  const stringified = matchRow.join();
  const foundIndex  = allSheetValues.findIndex(r => r.join() === stringified);
  if (foundIndex === -1) return { link: "", label: "" };

  // richTextColumn is a 2D array [[RichTextValue], [RichTextValue], ...]
  const rt    = richTextColumn[foundIndex] && richTextColumn[foundIndex][0];
  const raw   = String(allSheetValues[foundIndex][deliveryMapLinkIndex] || "").trim();
  const label_raw = raw;

  let link  = "";
  let label = "";

  if (rt && typeof rt.getLinkUrl === "function" && rt.getLinkUrl()) {
    link  = rt.getLinkUrl();
    const rtText = typeof rt.getText === "function" ? rt.getText() : "";
    label = rtText.trim() || raw;
  } else if (raw.toLowerCase().startsWith("http")) {
    link  = raw;
    label = raw;
  }

  if (!link)                  label = `Open Map for ${routeValue}`;
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
/**
 * Builds a Google Maps waypoint URL from a list of addresses.
 * Uses /maps/search/ with pipe-separated addresses — shows all stops as
 * individual pins on one map with no imposed driving order.
 * Strips route name segments from addresses before encoding.
 * No hard cap — Maps handles 20+ addresses in search mode fine.
 */
function buildAutoMapUrl(addresses, routeName) {
  const clean = addresses
    .map(a => {
      // Strip route name from address if present (e.g. "123 Main, Teralta1, San Diego, CA")
      const raw  = String(a || "").trim();
      const name = String(routeName || "").trim();
      if (!name) return raw;
      return raw.split(",").map(p => p.trim()).filter(p => p !== name).join(", ");
    })
    .filter(a => a.length > 0);

  if (clean.length === 0) return "";

  const query = clean.map(a => encodeURIComponent(a)).join("%7C"); // %7C = pipe
  return `https://www.google.com/maps/search/${query}`;
}

/**
 * AUTO_GENERATE_MAP_LINKS path:
 * Builds a waypoint map URL from a route's stop addresses, shortens it,
 * and returns { link, label } in the same shape as extractMapLinkForRoute.
 */
function buildRouteMapInfo(personStops, route, headers) {
  const addrIdx   = headers.indexOf("Address");
  const routeIdx  = headers.indexOf("routeDescription");

  const addresses = Object.values(personStops).map(rows => {
    const r   = rows[0];
    const raw = addrIdx !== -1 ? String(r[addrIdx] || "").trim() : "";
    // Strip route name segment from address
    const routeName = routeIdx !== -1 ? String(r[routeIdx] || "").trim() : route;
    return raw.split(",").map(p => p.trim()).filter(p => p !== routeName).join(", ");
  }).filter(a => a.length > 0);

  if (addresses.length === 0) return { link: "", label: "" };

  const url      = buildAutoMapUrl(addresses, ""); // already stripped above
  const short    = url ? shortenUrlForTyping(url) : "";
  const link     = short || url;
  const label    = `Map: ${route}`;

  return { link, label, fullUrl: url, shortUrl: short };
}

/* -------------------------
   Helper: writeRouteIndexSheet
   Writes or refreshes the Driver_Deliveries index tab.
------------------------- */
function writeRouteIndexSheet(ss, indexSheetName, summaryRows) {
  let indexSheet = ss.getSheetByName(indexSheetName);
  if (!indexSheet) indexSheet = ss.insertSheet(indexSheetName);
  indexSheet.clearContents();
  indexSheet.clearFormats();
  if (!summaryRows || summaryRows.length === 0) return;
  indexSheet.getRange(1, 1, summaryRows.length, summaryRows[0].length).setValues(summaryRows);
  indexSheet.getRange(1, 1, 1, summaryRows[0].length)
    .setFontWeight("bold").setBackground("#F0F0F0").setHorizontalAlignment("center");
  indexSheet.autoResizeColumns(1, summaryRows[0].length);
  indexSheet.setColumnWidth(summaryRows[0].length, 120);
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
/**
 * Assigns a sort bucket to a driver route stop (array of rows for one person/address)
 * using the same priority order as the packing list sort:
 * 1. No special items  2. Dog only  3. Dog+Cat  4. Cat only
 * 5. Diapers/Wipes  6. Hygiene  7. Special requests (always last)
 *
 * itemMap is used to resolve column indices by label rather than assuming positions.
 */
function routeSheetStopBucket(rowsForPerson, headers, itemMap) {
    const getIdx = label => {
        const item = itemMap.find(i => i.label === label);
        if (!item) return -1;
        if (item.type === 'qty') return headers.indexOf(item.sourceHeader);
        return -1;
    };
    const getDiaperIdxs = () => {
        const item = itemMap.find(i => i.type === 'diapers');
        return item ? item.sourceIndex.filter(i => i !== -1) : [];
    };

    const dogIdx     = getIdx("dog food");
    const catIdx     = getIdx("cat food");
    const wipesIdx   = getIdx("wipes");
    const toiIdx     = getIdx("toiletries");
    const femIdx     = getIdx("fem hygn");
    const diaperIdxs = getDiaperIdxs();
    const specIdx    = headers.indexOf("Special Item Requests");

    // Union values across all rows for this stop
    const qty = idx => idx !== -1
        ? rowsForPerson.reduce((s, r) => s + (parseFloat(r[idx] || 0) || 0), 0)
        : 0;
    const str = idx => idx !== -1
        ? rowsForPerson.some(r => String(r[idx] || "").trim() !== "")
        : false;

    const hasDog        = qty(dogIdx) > 0;
    const hasCat        = qty(catIdx) > 0;
    const hasWipes      = qty(wipesIdx) > 0;
    const hasToi        = qty(toiIdx) > 0;
    const hasFem        = qty(femIdx) > 0;
    const hasDiapers    = diaperIdxs.some(idx =>
        rowsForPerson.some(r => String(r[idx] || "").trim() !== "")
    );
    const hasSpecialReq = str(specIdx);

    if (hasSpecialReq)                              return 7;
    if (!hasDog && !hasCat && !hasDiapers &&
        !hasWipes && !hasToi && !hasFem)            return 1;
    if (hasDog && hasCat)                           return 3;
    if (hasDog && !hasCat)                          return 2;
    if (hasCat && !hasDog)                          return 4;
    if (hasDiapers || hasWipes)                     return 5;
    if (hasToi || hasFem)                           return 6;
    return 8;
}

function buildRouteSheet(opts) {
  const {
    ss, sheet, route, driverName, personStops,
    itemMap, finalItemHeaders, headers, routeColor, colIndices, mapInfo
  } = opts;

  const tb = Date.now();
  const lapB = (label) => Logger.log(`    ⏱ buildRouteSheet +${((Date.now()-tb)/1000).toFixed(2)}s — ${label}`);

  sheet.clearFormats();
  lapB("clearFormats");

  const rollup = computeRouteItemRollup(personStops, itemMap, headers);
  lapB("computeRouteItemRollup");

  // 1. Write the route header block (rows 1-7)
  writeFullRouteHeader({ sheet, route, driverName, mapInfo, rollup, itemMap });
  lapB("writeFullRouteHeader");

  const stopColorList = pickStopColors();
  const stopEntries   = Object.values(personStops);
  const COLS          = 10;
  const BODY_START    = 9;
  const totalBodyRows = Math.max(stopEntries.length * 10 + 5, 50);

  // 2. White base in one call
  sheet.getRange(BODY_START - 1, 1, totalBodyRows + 1, COLS).setBackground("#FFFFFF");
  lapB(`white base (${totalBodyRows} rows)`);

  // 3. Sort stops by special-items priority (same order as packing list per Gavi's request)
  const sortedStopEntries = stopEntries.slice().sort((a, b) =>
    routeSheetStopBucket(a, headers, itemMap) - routeSheetStopBucket(b, headers, itemMap)
  );

  // 4. Write stops — collect border targets
  const borderTargets = [];
  let row = BODY_START;

  sortedStopEntries.forEach((rowsForPerson, idx) => {
    const currentColor  = stopColorList[idx % stopColorList.length];
    const tableStartRow = row;
    row = writeStopTable({ sheet, rowsForPerson, headers, itemMap, routeColor: currentColor, startRow: row });
    borderTargets.push({ startRow: tableStartRow, endRow: row - 2, color: currentColor });
  });
  lapB(`writeStopTable x${sortedStopEntries.length} done`);

  // 4. Batch apply backgrounds + borders
  borderTargets.forEach(({ startRow, endRow, color }) => {
    const numRows = endRow - startRow + 1;
    if (numRows < 1) return;
    sheet.getRange(startRow, 1, numRows, COLS).setBackground(color);
    sheet.getRange(startRow, 1, numRows, COLS).setBorder(
      true, true, true, true, false, false,
      "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
    sheet.getRange(startRow, 1, 1, COLS).setBorder(
      null, null, true, null, false, false,
      "black", SpreadsheetApp.BorderStyle.DOTTED
    );
  });
  lapB(`borders+backgrounds x${borderTargets.length} done`);

  return row;
}

/**********************************************************************
 ROUTE SHEET Header
***********************************************************************/
function writeFullRouteHeader(opts) {
  const { sheet, route, driverName, mapInfo, rollup, itemMap } = opts;

  // Column widths — batch as fast as Apps Script allows
  sheet.setColumnWidth(1, 250);
  for (let c = 2; c <= 10; c++) sheet.setColumnWidth(c, 100);

  const itemNames   = itemMap.map(i => i.label);
  const totalValues = itemNames.map(name => rollup[name] || "");
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  const COLS = 10;

  // ── Bulk value write: all 7 header rows at once ───────────────────────────
  // Col A values (rows 1-7), then col B+ where applicable
  // Use pre-shortened URL if auto-generate provided one, otherwise shorten now.
  // When AUTO_GENERATE_MAP_LINKS is true, buildRouteMapInfo already shortened it.
  const shortUrl = mapInfo?.shortUrl || (mapInfo?.link ? shortenUrlForTyping(mapInfo.link) : "");

  const colAValues = [
    [driverName], ["Map Link:"], ["Short Link:"],
    ["Route Items"], ["Total Quantities"],
    ["In this table here, diapers shows the quantity of PACKS of Diapers and/or Pull-Ups for your route, but in the tables below, it will show the sizes each delivery gets instead. Everything else is a quantity."],
    [`Generated: ${stamp}`]
  ];
  sheet.getRange(1, 1, 7, 1).setValues(colAValues);

  // ── Merges ────────────────────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, 4).merge();  // A1:D1 driver
  sheet.getRange(1, 5, 1, 6).merge();  // E1:J1 route
  sheet.getRange(2, 2, 1, 9).merge();  // B2:J2 map link
  sheet.getRange(3, 2, 1, 9).merge();  // B3:J3 map url
  sheet.getRange(6, 1, 1, COLS).merge(); // A6:J6 explanation
  sheet.getRange(7, 1, 1, COLS).merge(); // A7:J7 timestamp

  // ── Route name in E1:J1 ───────────────────────────────────────────────────
  sheet.getRange(1, 5).setValue(route)
    .setHorizontalAlignment("right");

  // ── Map link (row 2) ──────────────────────────────────────────────────────
  const mapCell = sheet.getRange(2, 2);
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

  // ── Short/typeable URL (row 3) ───────────────────────────────────────────
  sheet.getRange(3, 2).setValue(shortUrl || mapInfo?.link || "")
    .setFontSize(14).setHorizontalAlignment("left");

  // ── Item headers row 4, totals row 5 (B onwards) ─────────────────────────
  sheet.getRange(4, 2, 1, itemNames.length).setValues([itemNames]);
  sheet.getRange(5, 2, 1, totalValues.length).setValues([totalValues]);

  // ── Bulk formatting (2 range calls cover all 7 rows) ─────────────────────
  const headerBlock = sheet.getRange(1, 1, 7, COLS);
  headerBlock
    .setBackground("#f3f3f3")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Row-specific font sizes and weights via grid arrays (1 call each)
  const fontSizeGrid = [
    [24,24,24,24,24,24,24,24,24,24],  // row 1 driver
    [16,16,16,16,16,16,16,16,16,16],  // row 2 map link
    [11,11,11,11,11,11,11,11,11,11],  // row 3 url
    [12,12,12,12,12,12,12,12,12,12],  // row 4 item headers
    [14,14,14,14,14,14,14,14,14,14],  // row 5 totals
    [11,11,11,11,11,11,11,11,11,11],  // row 6 explanation
    [16,16,16,16,16,16,16,16,16,16],  // row 7 timestamp
  ];
  const fontWeightGrid = [
    ["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],
    ["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],
    ["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],
    ["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],
    ["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],
    ["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],
    ["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],
  ];
  headerBlock.setFontSizes(fontSizeGrid);
  headerBlock.setFontWeights(fontWeightGrid);

  // Individual alignment corrections
  sheet.getRange(1, 1).setHorizontalAlignment("left");   // driver left
  sheet.getRange(1, 5).setHorizontalAlignment("right");  // route right
  sheet.getRange(2, 2).setHorizontalAlignment("right");  // map link right
  sheet.getRange(3, 2).setHorizontalAlignment("right");  // url right
  sheet.getRange(4, 2, 1, itemNames.length).setHorizontalAlignment("center");
  sheet.getRange(5, 2, 1, totalValues.length).setHorizontalAlignment("center");
  sheet.getRange(6, 1).setWrap(true).setFontStyle("italic").setHorizontalAlignment("left");
  sheet.getRange(7, 1).setHorizontalAlignment("center").setVerticalAlignment("middle");
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

  // Borders and background are applied by the caller (buildRouteSheet) in a batched pass.
  return row + 1; // single blank row after table
}


function writeSTR1(sheet, rows, headers, row) {
  const nameIdx   = headers.indexOf("Name");
  const phoneIdx  = headers.indexOf("Phone");
  const addrIdx   = headers.indexOf("Address");
  const routeIdx  = headers.indexOf("routeDescription");

  const r         = rows[0];
  const namePhone = `${r[nameIdx] || ""}  /  ${r[phoneIdx] || ""}`;

  // Addresses are stored as "Street, RouteName, City, State".
  // The route name is already on the tab header — strip it from the address
  // so drivers see a clean street address without a redundant or misleading label.
  const rawAddress  = String(r[addrIdx] || "").trim();
  const routeName   = routeIdx !== -1 ? String(r[routeIdx] || "").trim() : "";
  const addrParts   = rawAddress.split(",").map(p => p.trim());
  const cleanParts  = addrParts.filter(p => p !== routeName);
  const address     = cleanParts.join(", ");

  // Row 1: Name / Phone — full width A:J
  sheet.getRange(row, 1, 1, 10).merge()
    .setValue(namePhone)
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("left");

  // Row 2: Address — full width A:J, always its own line so no truncation
  sheet.getRange(row + 1, 1, 1, 10).merge()
    .setValue(address)
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("left");

  return row + 2;
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

    // 2. Handle Diapers — collect ALL size entries including duplicates.
    // A person with two bags of the same size (e.g. 4/5T, 4/5T) should show "4/5T | 4/5T".
    if (item.type === "diapers") {
      const sizes = [];

      rows.forEach(r => {
        item.sourceIndex.forEach(i => {
          if (i !== -1) {
            const val = String(r[i] || '').trim();
            if (val) sizes.push(val);
          }
        });
      });

      if (sizes.length === 0) return "";

      return sizes
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
  // Background and borders are now applied in batch by buildRouteSheet's borderTargets pass.
  // This function is retained for any legacy callers (populateRouteTabs) but is a no-op
  // when called from buildRouteSheet.
  const { sheet, startRow, endRow, stopColor } = opts;
  if (!sheet || !startRow || !endRow || !stopColor) return;
  const totalCols = 10;
  const numRows = endRow - startRow + 1;
  if (numRows < 1) return;
  sheet.getRange(startRow, 1, numRows, totalCols).setBackground(stopColor);
  sheet.getRange(startRow, 1, numRows, totalCols).setBorder(
    true, true, true, true, false, false,
    "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  sheet.getRange(startRow, 1, 1, totalCols).setBorder(
    null, null, true, null, false, false,
    "black", SpreadsheetApp.BorderStyle.DOTTED
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

/**
 * Generates the Runners Dashboard tab: confirmation status, standard item totals,
 * special items pull list (cross-route and per-route), diaper reference, and
 * special requests flag list. This is the coordinator/runner lead view.
 */
function makeNaviDash() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    refreshDeliveriesHelpTable();
    const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
    const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
    const dashSheet = setupPackingListTab(ss, "NaviDash");

    const {
        dataRows, headers, confForIdx, includedIndexes, diaperSizeIndices,
        packsOrigIdx, eggsOrigIdx, milkOrigIdx, wipesOrigIdx,
        dogFoodOrigIdx, catFoodOrigIdx, toiletriesOrigIdx, femHygieneOrigIdx,
        routeIdx, driverIdx
    } = loadAndIndexPackingData(dataSheet);

    const routeDriverMap = buildPackingListRouteDriverMap(helperSheet);
    const { confirmedDataRows, dataRowsForDiaperCount } = filterPackingDataForReports(dataRows, confForIdx);
    const allDiaperItems = defineDiaperItems();
    const diaperCounts = calculatePackingDiaperAndWipesCounts(dataRowsForDiaperCount, allDiaperItems, diaperSizeIndices, wipesOrigIdx);
    const confForCounts = calculatePackingConfirmationCounts(dataRows, confForIdx);

    const specialReqIdxNaviDash = headers ? headers.indexOf("Special Item Requests") : -1;
    const indices = {
        routeIdx, driverIdx, confForIdx, includedIndexes,
        packsOrigIdx, eggsOrigIdx, milkOrigIdx, dogFoodOrigIdx, catFoodOrigIdx,
        wipesOrigIdx, toiletriesOrigIdx, femHygieneOrigIdx, diaperSizeIndices,
        specialReqIdx: specialReqIdxNaviDash
    };

    // -- Column sizing: portrait layout, A:J only (10 columns) --
    dashSheet.setColumnWidth(1, 280); // Col A: item names, route names, flag list labels
    dashSheet.setColumnWidth(2, 120); // Col B: category / driver name
    dashSheet.setColumnWidth(3, 130); // Col C: recipient names / total qty / dog food
    dashSheet.setColumnWidth(4, 55);  // Col D: same as C — ☐ pulled / cat food
    dashSheet.setColumnWidth(5, 70);  // Col E: wipes / tracking col 1a
    dashSheet.setColumnWidth(6, 70);  // Col F: toiletries / tracking col 1b
    dashSheet.setColumnWidth(7, 70);  // Col G: fem hygiene / tracking col 2a
    dashSheet.setColumnWidth(8, 70);  // Col H: diaper packs / tracking col 2b
    dashSheet.setColumnWidth(9, 70);  // Col I: special requests flag / tracking col 3a
    dashSheet.setColumnWidth(10, 55); // Col J: ☐ checkbox / tracking col 3b

    let outputRow = 1;

    // === SECTION 1: Title Banner ===
    outputRow = writeDashboardTitle(dashSheet, outputRow, dataRowsForDiaperCount.length, confirmedDataRows.length);

    // === SECTION 2: Standard Items then Confirmation Status (stacked, A:J) ===
    // Items table first (4 rows), conf status below it (5 rows), both full width.
    const summaryEndRow = generateItemSummaryTable(dashSheet, confirmedDataRows, packsOrigIdx, eggsOrigIdx, milkOrigIdx, outputRow);
    const confEndRow    = generateConfirmationStatusTable(dashSheet, confForCounts, summaryEndRow);
    outputRow = confEndRow + 1;

    // === SECTION 3: Special Items Cross-Route Pull List ===
    outputRow = generateSpecialItemsTotalsTable(dashSheet, dataRowsForDiaperCount, allDiaperItems, diaperCounts, indices, outputRow);
    outputRow += 1;

    // === SECTION 4: Per-Route Special Items Breakdown ===
    outputRow = generatePerRouteSpecialItemsBreakdown(dashSheet, dataRowsForDiaperCount, allDiaperItems, diaperCounts, indices, routeDriverMap, outputRow);
    outputRow += 1;

    // SECTION 5 (Diaper Reference Table) removed — covered by Section 3 pull list.

    // === SECTION 5: Special Requests Flag List ===
    outputRow = generateSpecialRequestsFlagList(dashSheet, dataRowsForDiaperCount, headers, indices, routeDriverMap, outputRow);

    // === SECTION 6: Timestamp ===
    outputRow += 1;
    generateTimestamp(dashSheet, outputRow);

    SpreadsheetApp.flush();
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    return timestamp;
}

/**
 * Writes a prominent title banner at the top of the dashboard.
 */
function writeDashboardTitle(sheet, startRow, totalRows, confirmedRows) {
    const COLS = 10;
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    sheet.getRange(startRow, 1, 1, COLS).merge()
        .setValue("🏃 Navigator Dashboard — Distribution Day")
        .setFontSize(20).setFontWeight("bold")
        .setBackground("#1C4587").setFontColor("#FFFFFF")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(startRow, 40);

    sheet.getRange(startRow + 1, 1, 1, COLS).merge()
        .setValue(`Generated: ${ts}    |    Total deliveries (YES + NO ANS): ${totalRows}    |    Confirmed (YES only): ${confirmedRows}`)
        .setFontSize(11).setFontWeight("normal")
        .setBackground("#C9DAF8").setFontColor("#000000")
        .setHorizontalAlignment("center");

    return startRow + 2;
}

/**
 * GAVI'S REQUEST — Cross-route totals for all special items.
 * Dog food, cat food, toiletries, fem hygiene, wipes, and every diaper size
 * as individual rows. One table the runner lead can read top to bottom and
 * pull everything before packing stations open.
 */
function generateSpecialItemsTotalsTable(sheet, dataRows, allDiaperItems, diaperCounts, indices, startRow) {
    const BG_HEADER  = "#274E13";  // dark green
    const FG_HEADER  = "#FFFFFF";
    const BG_SECTION = "#D9EAD3";  // light green section headers
    const BG_ALT     = "#F3F3F3";
    const BG_TOTAL   = "#B6D7A8";
    const BG_PET     = "#FCE5CD";  // orange tint for pet food
    const BG_HYGIENE = "#D9D2E9";  // purple tint for hygiene
    const BG_DIAPER  = "#CFE2F3";  // blue tint for diapers
    // Layout: A=Item, B=Category, C=Total Qty, D=Pulled☐, E:F=On Hand, G:H=Given Out, I:J=Not on Hand
    const NUM_COLS   = 10;
    const TRACK_BG   = "#F8F8F8"; // light grey for the three manual tracking col pairs

    // Section banner
    sheet.getRange(startRow, 1, 1, NUM_COLS).merge()
        .setValue("SPECIAL ITEMS — TOTAL PULL LIST  (all routes combined, YES + NO ANS)")
        .setFontSize(14).setFontWeight("bold")
        .setBackground(BG_HEADER).setFontColor(FG_HEADER)
        .setHorizontalAlignment("left");
    startRow++;

    // Column headers — A:D fixed, E:F / G:H / I:J merged tracking pairs
    sheet.getRange(startRow, 1, 1, NUM_COLS)
        .setValues([["Item", "Category", "Total", "Pulled?", "On Hand", "", "Given Out", "", "Not on Hand", ""]])
        .setFontWeight("bold").setFontSize(11)
        .setBackground("#3D5F3C").setFontColor("#FFFFFF")
        .setHorizontalAlignment("center");
    // Merge the three tracking header pairs
    sheet.getRange(startRow, 5, 1, 2).merge(); // E:F — On Hand
    sheet.getRange(startRow, 7, 1, 2).merge(); // G:H — Given Out
    sheet.getRange(startRow, 9, 1, 2).merge(); // I:J — Not on Hand
    startRow++;

    const rows = [];
    const rowMeta = [];

    // --- Pet Food ---
    const dogTotal  = dataRows.reduce((s, r) => s + (parseFloat(r[indices.dogFoodOrigIdx] || 0) || 0), 0);
    const catTotal  = dataRows.reduce((s, r) => s + (parseFloat(r[indices.catFoodOrigIdx] || 0) || 0), 0);
    const petTotal  = dogTotal + catTotal;
    rows.push(["Dog Food",       "Pet Food",  dogTotal,  "☐", "", "", "", "", "", ""]); rowMeta.push(BG_PET);
    rows.push(["Cat Food",       "Pet Food",  catTotal,  "☐", "", "", "", "", "", ""]); rowMeta.push(BG_PET);
    rows.push(["Pet Food Total", "Pet Food",  petTotal,  "☐", "", "", "", "", "", ""]); rowMeta.push("#F6B26B");

    // --- Hygiene ---
    const wipesTotal   = dataRows.reduce((s, r) => s + (parseFloat(r[indices.wipesOrigIdx]      || 0) || 0), 0);
    const toiTotal     = dataRows.reduce((s, r) => s + (parseFloat(r[indices.toiletriesOrigIdx] || 0) || 0), 0);
    const femTotal     = dataRows.reduce((s, r) => s + (parseFloat(r[indices.femHygieneOrigIdx] || 0) || 0), 0);
    const hygieneTotal = wipesTotal + toiTotal + femTotal;
    rows.push(["Wipes",          "Hygiene", wipesTotal,   "☐", "", "", "", "", "", ""]); rowMeta.push(BG_HYGIENE);
    rows.push(["Toiletries",     "Hygiene", toiTotal,     "☐", "", "", "", "", "", ""]); rowMeta.push(BG_HYGIENE);
    rows.push(["Fem Hygiene",    "Hygiene", femTotal,     "☐", "", "", "", "", "", ""]); rowMeta.push(BG_HYGIENE);
    rows.push(["Hygiene Total",  "Hygiene", hygieneTotal, "☐", "", "", "", "", "", ""]); rowMeta.push("#B4A7D6");

    // --- Diapers by size ---
    const rawDiaperSizes = allDiaperItems.filter(d => d[0] !== "Diapers Total" && d[0] !== "Wipes");
    rawDiaperSizes.forEach((item, i) => {
        const qty = diaperCounts[item[0]] || 0;
        rows.push([`Diapers — ${item[0]} (${item[1]})`, "Diapers", qty, "☐", "", "", "", "", "", ""]);
        rowMeta.push(i % 2 === 0 ? "#FFFFFF" : BG_ALT);
    });
    rows.push(["Diapers Total (all sizes)", "Diapers", diaperCounts["DIAPER_TOTAL"] || 0, "☐", "", "", "", "", "", ""]);
    rowMeta.push(BG_DIAPER);

    // Write data
    sheet.getRange(startRow, 1, rows.length, NUM_COLS).setValues(rows);

    // Per-row: backgrounds, bold totals, merge tracking pairs E:F / G:H / I:J
    rows.forEach((row, i) => {
        const rng = sheet.getRange(startRow + i, 1, 1, NUM_COLS);
        rng.setBackground(rowMeta[i]);
        if (row[0].includes("Total")) rng.setFontWeight("bold");
        // Merge tracking column pairs on each data row
        sheet.getRange(startRow + i, 5, 1, 2).merge().setBackground(TRACK_BG); // E:F
        sheet.getRange(startRow + i, 7, 1, 2).merge().setBackground(TRACK_BG); // G:H
        sheet.getRange(startRow + i, 9, 1, 2).merge().setBackground(TRACK_BG); // I:J
    });

    // Borders and alignment
    sheet.getRange(startRow - 1, 1, rows.length + 1, NUM_COLS)
        .setBorder(true, true, true, true, true, true)
        .setFontSize(12).setVerticalAlignment("middle");
    sheet.getRange(startRow, 3, rows.length, 1).setHorizontalAlignment("center"); // Total Qty
    sheet.getRange(startRow, 4, rows.length, 1).setHorizontalAlignment("center").setFontSize(16); // ☐

    // Col D same width as C (both narrow), tracking cols wider for writing space
    sheet.setColumnWidth(4, 55);  // D = ☐ pulled
    sheet.setColumnWidth(5, 70);  // E (merges to F)
    sheet.setColumnWidth(6, 70);  // F
    sheet.setColumnWidth(7, 70);  // G (merges to H)
    sheet.setColumnWidth(8, 70);  // H
    sheet.setColumnWidth(9, 70);  // I (merges to J)
    sheet.setColumnWidth(10, 70); // J

    return startRow + rows.length;
}

/**
 * Per-route breakdown of special items so the runner lead can stage
 * items by route pile before drivers arrive.
 */
function generatePerRouteSpecialItemsBreakdown(sheet, dataRows, allDiaperItems, diaperCounts, indices, routeDriverMap, startRow) {
    const BG_HEADER  = "#0B5394";
    const FG_HEADER  = "#FFFFFF";
    const BG_COL_HDR = "#1C84C6";
    const NUM_COLS   = 10; // route, driver, dog, cat, wipes, toiletries, fem, diapers(packs), wipes(packs), special flag

    // Section banner
    sheet.getRange(startRow, 1, 1, 10).merge()
        .setValue("SPECIAL ITEMS — PER ROUTE BREAKDOWN  (stage by driver pile)")
        .setFontSize(14).setFontWeight("bold")
        .setBackground(BG_HEADER).setFontColor(FG_HEADER)
        .setHorizontalAlignment("left");
    startRow++;

    const colHeaders = [
        "Route", "Driver",
        "Dog\nFood", "Cat\nFood",
        "Wipes", "Toilet-\nries", "Fem\nHygiene",
        "Diaper\nPacks", "Special\nRequests?", "Staged? ✓"
    ];
    sheet.getRange(startRow, 1, 1, NUM_COLS)
        .setValues([colHeaders])
        .setFontWeight("bold").setFontSize(11)
        .setBackground(BG_COL_HDR).setFontColor(FG_HEADER)
        .setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");
    sheet.setRowHeight(startRow, 45);
    startRow++;

    const routes = [...new Set(dataRows.map(r => String(r[indices.routeIdx] || "").trim()).filter(r => r))].sort();
    const routeColors = [
        "#D9EAD3","#FFF2CC","#FCE5CD","#D9D2E9","#CFE2F3",
        "#F4CCCC","#EAD1DC","#E0F7FA","#F8BBD0","#B2EBF2"
    ];

    const tableDataRows = [];
    const rowColors = [];

    routes.forEach((route, i) => {
        const routeRows = dataRows.filter(r => String(r[indices.routeIdx] || "").trim() === route);
        const driver = routeDriverMap[route] || "TBD";
        const color = routeColors[i % routeColors.length];

        const dog   = routeRows.reduce((s, r) => s + (parseFloat(r[indices.dogFoodOrigIdx]    || 0) || 0), 0);
        const cat   = routeRows.reduce((s, r) => s + (parseFloat(r[indices.catFoodOrigIdx]    || 0) || 0), 0);
        const wipes = routeRows.reduce((s, r) => s + (parseFloat(r[indices.wipesOrigIdx]      || 0) || 0), 0);
        const toi   = routeRows.reduce((s, r) => s + (parseFloat(r[indices.toiletriesOrigIdx] || 0) || 0), 0);
        const fem   = routeRows.reduce((s, r) => s + (parseFloat(r[indices.femHygieneOrigIdx] || 0) || 0), 0);

        // Diaper packs: count non-empty cells across the 3 diaper columns
        let diaperPacks = 0;
        routeRows.forEach(r => {
            indices.diaperSizeIndices.forEach(idx => {
                if (String(r[idx] || "").trim() !== "") diaperPacks++;
            });
        });

        const specIdx = indices.specialReqIdx !== undefined
            ? indices.specialReqIdx
            : indices.femHygieneOrigIdx + 1; // fallback
        const hasSpecialReqs = routeRows.some(r => String(r[specIdx] || "").trim() !== "");

        tableDataRows.push([route, driver, dog || "", cat || "", wipes || "", toi || "", fem || "", diaperPacks || "", hasSpecialReqs ? "YES" : "", "☐"]);
        rowColors.push(color);
    });

    // Totals row
    const totDog   = tableDataRows.reduce((s, r) => s + (Number(r[2]) || 0), 0);
    const totCat   = tableDataRows.reduce((s, r) => s + (Number(r[3]) || 0), 0);
    const totWipes = tableDataRows.reduce((s, r) => s + (Number(r[4]) || 0), 0);
    const totToi   = tableDataRows.reduce((s, r) => s + (Number(r[5]) || 0), 0);
    const totFem   = tableDataRows.reduce((s, r) => s + (Number(r[6]) || 0), 0);
    const totDiap  = tableDataRows.reduce((s, r) => s + (Number(r[7]) || 0), 0);
    tableDataRows.push(["TOTALS", "", totDog, totCat, totWipes, totToi, totFem, totDiap, "", ""]);
    rowColors.push("#B6D7A8"); // no ☐ on totals row

    sheet.getRange(startRow, 1, tableDataRows.length, NUM_COLS).setValues(tableDataRows);

    tableDataRows.forEach((row, i) => {
        const rng = sheet.getRange(startRow + i, 1, 1, NUM_COLS);
        rng.setBackground(rowColors[i]);
        if (row[0] === "TOTALS") rng.setFontWeight("bold");
    });

    sheet.getRange(startRow - 1, 1, tableDataRows.length + 1, NUM_COLS)
        .setBorder(true, true, true, true, true, true)
        .setFontSize(12).setVerticalAlignment("middle").setHorizontalAlignment("center");
    sheet.getRange(startRow, 1, tableDataRows.length, 2).setHorizontalAlignment("left"); // route/driver left
    sheet.getRange(startRow, NUM_COLS, tableDataRows.length - 1, 1).setFontSize(16); // ☐ col
    sheet.setColumnWidth(NUM_COLS, 70);

    // Footer header row — lighter repeat of column headers for orientation after a page break
    const footerRow = startRow + tableDataRows.length;
    const footerColHeaders = [
        "Route", "Driver",
        "Dog Food", "Cat Food",
        "Wipes", "Toiletries", "Fem Hygiene",
        "Diaper Packs", "Special Requests?", "Staged? ✓"
    ];
    sheet.getRange(footerRow, 1, 1, NUM_COLS)
        .setValues([footerColHeaders])
        .setFontWeight("normal").setFontStyle("italic").setFontSize(10)
        .setBackground("#D0E4C8").setFontColor("#444444")
        .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    sheet.setRowHeight(footerRow, 35);
    sheet.getRange(footerRow, 1, 1, NUM_COLS)
        .setBorder(true, true, true, true, true, true);

    return footerRow + 1;
}

/**
 * Cross-route flag list of all deliveries with non-empty Special Item Requests.
 * Gives the runner lead one view of every special need without flipping through route tables.
 */
function generateSpecialRequestsFlagList(sheet, dataRows, headers, indices, routeDriverMap, startRow) {
    const BG_HEADER = "#7F6000";
    const FG_HEADER = "#FFFFFF";
    const NUM_COLS  = 10; // A=Route B=Driver C=Name D:I=Request(merged) J=☐
    const specialReqIdx = headers.indexOf("Special Item Requests");

    // Section banner
    sheet.getRange(startRow, 1, 1, NUM_COLS).merge()
        .setValue("SPECIAL ITEM REQUESTS — FLAG LIST  (all routes, pre-stage these separately)")
        .setFontSize(14).setFontWeight("bold")
        .setBackground(BG_HEADER).setFontColor(FG_HEADER)
        .setHorizontalAlignment("left");
    startRow++;

    if (specialReqIdx === -1) {
        sheet.getRange(startRow, 1, 1, NUM_COLS).merge()
            .setValue("(No 'Special Item Requests' column found in source data)")
            .setFontStyle("italic").setBackground("#FFF3CD");
        return startRow + 2;
    }

    // Header row — D:I merged for "Special Item Request", J = "Handled? ✓"
    const headerRow = ["Route", "Driver", "Recipient", "Special Item Request", "", "", "", "", "", "Handled? ✓"];
    sheet.getRange(startRow, 1, 1, NUM_COLS)
        .setValues([headerRow])
        .setFontWeight("bold").setFontSize(11)
        .setBackground("#BF9000").setFontColor("#FFFFFF")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(startRow, 4, 1, 6).merge(); // D:I header merge
    startRow++;

    const nameIdx = headers.indexOf("Name");
    const routeColors = [
        "#FFF2CC","#FCE5CD","#D9EAD3","#D9D2E9","#CFE2F3",
        "#F4CCCC","#EAD1DC","#E0F7FA","#F8BBD0","#B2EBF2"
    ];
    const routeColorMap = {};
    const routes = [...new Set(dataRows.map(r => String(r[indices.routeIdx] || "").trim()).filter(r => r))].sort();
    routes.forEach((route, i) => { routeColorMap[route] = routeColors[i % routeColors.length]; });

    const flagRows = [];
    const flagColors = [];

    routes.forEach(route => {
        dataRows
            .filter(r => String(r[indices.routeIdx] || "").trim() === route)
            .forEach(r => {
                const special = String(r[specialReqIdx] || "").trim();
                if (!special) return;
                const name   = nameIdx !== -1 ? String(r[nameIdx] || "").trim() : "";
                const driver = routeDriverMap[route] || "TBD";
                // D:I holds the request text (D is written, E:I are empty for merge)
                flagRows.push([route, driver, name, special, "", "", "", "", "", "☐"]);
                flagColors.push(routeColorMap[route]);
            });
    });

    if (flagRows.length === 0) {
        sheet.getRange(startRow, 1, 1, NUM_COLS).merge()
            .setValue("(No special item requests found)")
            .setFontStyle("italic").setBackground("#F3F3F3");
        return startRow + 2;
    }

    sheet.getRange(startRow, 1, flagRows.length, NUM_COLS).setValues(flagRows);

    // Merge D:I on every data row so request text has full width to render on one line
    for (let i = 0; i < flagRows.length; i++) {
        sheet.getRange(startRow + i, 4, 1, 6).merge();
        sheet.getRange(startRow + i, 1, 1, NUM_COLS).setBackground(flagColors[i]);
    }

    // Column widths: A=Route, B=Driver, C=Name match, D (merged D:I) wide, J=☐ narrow
    sheet.setColumnWidth(4, 70);  // D same as C — merge handles the visual width
    sheet.setColumnWidth(10, 55); // J = ☐ checkbox col

    // Formatting
    sheet.getRange(startRow, 1, flagRows.length, 3).setHorizontalAlignment("left");
    sheet.getRange(startRow, 4, flagRows.length, 6).setWrap(true).setHorizontalAlignment("left");
    sheet.getRange(startRow, 10, flagRows.length, 1).setHorizontalAlignment("center").setFontSize(16);

    sheet.getRange(startRow - 1, 1, flagRows.length + 1, NUM_COLS)
        .setBorder(true, true, true, true, true, true)
        .setFontSize(12).setVerticalAlignment("middle");

    // Footer header row — lighter repeat of column headers for orientation after a page break
    const footerRow = startRow + flagRows.length;
    const footerHeaders = ["Route", "Driver", "Recipient", "Special Item Request", "", "", "", "", "", "Handled? ✓"];
    sheet.getRange(footerRow, 1, 1, NUM_COLS)
        .setValues([footerHeaders])
        .setFontWeight("normal").setFontStyle("italic").setFontSize(10)
        .setBackground("#C9A84C").setFontColor("#444444")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(footerRow, 4, 1, 6).merge(); // D:I merged same as header
    sheet.getRange(footerRow, 4, 1, 1).setHorizontalAlignment("left");
    sheet.getRange(footerRow, 1, 1, NUM_COLS)
        .setBorder(true, true, true, true, true, true);

    return footerRow + 1;
}

/**
 * Generates ONLY the per-route packing tables (bag station view).
 * Summary/dashboard tables have moved to NaviDash.
 */
function makePackingLists() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    refreshDeliveriesHelpTable();
    const dataSheet = ss.getSheetByName("Deliveries-UPDATE HERE");
    const helperSheet = ss.getSheetByName("DeliveriesHelpTable");
    const outputSheet = setupPackingListTab(ss, "PackingLists");

    // 1. Load and Prepare Data
    const {
        dataRows, headers, confForIdx, includedIndexes, diaperSizeIndices,
        packsOrigIdx, eggsOrigIdx, milkOrigIdx, wipesOrigIdx,
        dogFoodOrigIdx, catFoodOrigIdx, toiletriesOrigIdx, femHygieneOrigIdx,
        routeIdx, driverIdx
    } = loadAndIndexPackingData(dataSheet);

    const routeDriverMap = buildPackingListRouteDriverMap(helperSheet);

    // 2. Filter Data
    const { confirmedDataRows, dataRowsForDiaperCount } = filterPackingDataForReports(dataRows, confForIdx);

    // 3. Calculations (still needed by route sections)
    const allDiaperItems = defineDiaperItems();
    const diaperCounts = calculatePackingDiaperAndWipesCounts(dataRowsForDiaperCount, allDiaperItems, diaperSizeIndices, wipesOrigIdx);

    // 4. Timestamp at top of packing list tab
    let outputRow = 1;
    outputRow = generateTimestamp(outputSheet, outputRow);

    // 5. Route packing tables
    generatePackingListRouteSections(
        outputSheet, dataRowsForDiaperCount, headers, routeDriverMap, diaperCounts,
        {
            routeIdx, driverIdx, confForIdx, includedIndexes,
            packsOrigIdx, eggsOrigIdx, milkOrigIdx, dogFoodOrigIdx, catFoodOrigIdx,
            wipesOrigIdx, toiletriesOrigIdx, femHygieneOrigIdx,
            diaperSizeIndices,
            specialReqIdx: headers.indexOf("Special Item Requests")
        },
        outputRow
    );

    // 6. Final Formatting
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
        femHygieneOrigIdx: getIndex("fem hygiene(qty)"),
        specialReqIdx:     getIndex("Special Item Requests")
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
    const summaryNumCols  = 10;

    const totalPacks = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[packsIdx] || 0) || 0), 0);
    const totalEggs  = confirmedDataRows.reduce((s, r) => s + (parseFloat(r[eggsIdx] || 0) || 0), 0);
    const totalMilk  = confirmedDataRows.reduce((s, r) => s + (parseFloat(String(r[milkIdx]).replace('*', '') || 0)) || 0, 0);

    // Set Header and Data (10 cols wide)
    outputSheet.getRange(startRow, summaryStartCol, 1, summaryNumCols).setValues([["Standard Items", "Total (Confirmed Only)", "", "", "", "", "", "", "", ""]]);
    outputSheet.getRange(startRow + 1, summaryStartCol, 3, summaryNumCols).setValues([
        ["# packs", totalPacks, "", "", "", "", "", "", "", ""],
        ["# eggs",  totalEggs,  "", "", "", "", "", "", "", ""],
        ["# milk",  totalMilk,  "", "", "", "", "", "", "", ""]
    ]);
    
    // Apply Merges — value col (B) spans B:J
    outputSheet.getRange(startRow, 2, 1, 9).merge();
    for (let r = startRow + 1; r <= startRow + 3; r++) {
        outputSheet.getRange(r, 2, 1, 9).merge();
    }

    // Apply Formatting
    outputSheet.getRange(startRow, summaryStartCol, 4, summaryNumCols)
        .setFontWeight("bold")
        .setFontSize("13")
        .setBackground("#D9EAD3")
        .setBorder(true, true, true, true, null, null);

    // Alignments
    outputSheet.getRange(startRow, 1, 4, 1).setHorizontalAlignment("left");
    outputSheet.getRange(startRow, 2, 4, 1).setHorizontalAlignment("left").setFontSize("13");
    return startRow + 4;
}

function generateConfirmationStatusTable(outputSheet, confForCounts, startRow) {
    // Portrait layout: conf status stacks below item summary.
    // Uses cols A:J (1:10) — label col A, value col B:J merged.
    const COLS = 10;

    // Set Header and Data
    outputSheet.getRange(startRow, 1, 1, COLS).setValues([
        ["Confirmation Status", "Total Count", "", "", "", "", "", "", "", ""]
    ]);
    outputSheet.getRange(startRow + 1, 1, 4, COLS).setValues([
        ["No Ans", confForCounts["NO ANS"], "", "", "", "", "", "", "", ""],
        ["No",     confForCounts["NO"],     "", "", "", "", "", "", "", ""],
        ["Yes",    confForCounts["YES"],    "", "", "", "", "", "", "", ""],
        ["Total",  confForCounts["Total"],  "", "", "", "", "", "", "", ""]
    ]);

    // Merges: header spans A:H, count spans I:J; data rows label A, count B:J
    outputSheet.getRange(startRow, 1, 1, 8).merge();  // label header A:H
    outputSheet.getRange(startRow, 9, 1, 2).merge();  // count header I:J
    for (let r = startRow + 1; r <= startRow + 4; r++) {
        outputSheet.getRange(r, 2, 1, 9).merge();     // count value B:J
    }

    // Apply Formatting
    outputSheet.getRange(startRow, 1, 5, COLS)
        .setFontWeight("bold")
        .setFontSize("13")
        .setBackground("#D9D2E9")
        .setBorder(true, true, true, true, null, null);

    // Alignments
    outputSheet.getRange(startRow, 1, 1, 1).setHorizontalAlignment("left");
    outputSheet.getRange(startRow, 9, 1, 1).setHorizontalAlignment("right");
    outputSheet.getRange(startRow + 1, 1, 4, 1).setHorizontalAlignment("left");
    outputSheet.getRange(startRow + 1, 2, 4, 1).setHorizontalAlignment("left");
    return startRow + 5;
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
    const timestampRange = outputSheet.getRange(timestampRow, 1, 1, 10); 

    timestampRange.merge()
        .setValue(timestampMessage)
        .setFontSize(16)
        .setFontWeight("bold")
        .setBackground("#EFEFEF");
    
    return timestampRow + 3; // Space buffer for routes to start
}

/**
 * Assigns a sort bucket to a packing list row so entries within each route
 * are ordered from simplest to most complex, supporting human data parsing.
 * Per Gavi: no special items first → dog → dog+cat → cat → diapers/wipes →
 * hygiene → special requests last.
 *
 * Rows with special requests are always pushed to the end regardless of
 * what other items they have, since those need extra attention at packing.
 */
function packingListSortBucket(row, indices) {
    const qty = idx => parseFloat(row[idx] || 0) || 0;
    const str = idx => String(row[idx] || "").trim();

    const hasDog      = qty(indices.dogFoodOrigIdx) > 0;
    const hasCat      = qty(indices.catFoodOrigIdx) > 0;
    const hasWipes    = qty(indices.wipesOrigIdx) > 0;
    const hasToi      = qty(indices.toiletriesOrigIdx) > 0;
    const hasFem      = qty(indices.femHygieneOrigIdx) > 0;
    const hasDiapers    = indices.diaperSizeIndices.some(idx => str(idx) !== "");
    const specIdx       = indices.specialReqIdx !== undefined
        ? indices.specialReqIdx
        : indices.femHygieneOrigIdx + 1; // fallback
    const hasSpecialReq = String(row[specIdx] || "").trim() !== "";

    // Special requests always sort last — bucket 7
    if (hasSpecialReq) return 7;

    const hasHygiene      = hasWipes || hasToi || hasFem;
    const hasDiaperWipes  = hasDiapers || hasWipes;

    // No special items at all → bucket 1
    if (!hasDog && !hasCat && !hasDiapers && !hasWipes && !hasToi && !hasFem) return 1;

    // Pet food buckets
    if (hasDog && hasCat)  return 3; // dog + cat
    if (hasDog && !hasCat) return 2; // dog only
    if (hasCat && !hasDog) return 4; // cat only

    // Diapers/wipes (no pet food)
    if (hasDiapers || hasWipes) return 5;

    // Hygiene only (toiletries/fem hygiene, no pet food, no diapers/wipes)
    if (hasToi || hasFem) return 6;

    return 8; // fallback — shouldn't reach here
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
        
        // Filter rows for this route, then sort by special-items priority per Gavi's request:
        // 1. No special items  2. Dog only  3. Dog+Cat  4. Cat only
        // 5. Diapers/Wipes  6. Hygiene  7. Has special requests (always last)
        const filteredDataRows = dataRowsForDiaperCount
            .filter(r => String(r[indices.routeIdx]).trim() === route)
            .sort((a, b) => packingListSortBucket(a, indices) - packingListSortBucket(b, indices));

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
        // Use an array, not a Set — duplicates are intentional.
        // Two bags of 4/5T should display as "4/5T | 4/5T", not "4/5T".
        const sizes = [];
        personData.forEach(r => {
          [diapers1Idx, diapers2Idx, diapers3Idx].forEach(idx => {
            if (idx !== -1) {
              const val = r[idx];
              if (val !== null && val !== undefined) {
                const size = String(val).trim();
                if (size) sizes.push(size);
              }
            }
          });
        });
        const diaperOrder = ['P','N','1','2','3','4','5','6','7','8','2/3T','3/4T','4/5T','5/6T','6/7T'];
        const sortedSizes = sizes.sort((a,b)=>diaperOrder.indexOf(a)-diaperOrder.indexOf(b));
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
  var prefix = "Route_"; // targets legacy multi-sheet Route_ tabs (now replaced by single Driver_Routes sheet)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetsToDelete = [];
  
  // Identify legacy Route_ sheets
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

    // Group filtered rows by route, then sort routes alphabetically —
    // matching the sort order used by makeDriverRouteTabs and generatePackingListRouteSections.
    const groupedByRoute = {};
    filteredDataRows.forEach(row => {
        const route = String(row[indices.routeIdx] || 'UNKNOWN ROUTE').trim();
        if (!groupedByRoute[route]) groupedByRoute[route] = [];
        groupedByRoute[route].push(row);
    });
    const sortedRoutes = Object.keys(groupedByRoute).sort();

    sortedRoutes.forEach(route => {
        groupedByRoute[route].forEach(row => {
            const name = String(row[indices.nameIdx] || 'UNKNOWN NAME');
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
    });

    // Pad route boundaries: insert empty filler slots so every new route
    // starts at the beginning of a fresh label row (column 1 of 3).
    // This lets drivers/volunteers cut the sheet cleanly by route.
    const paddedLabels = [];
    let slotCount = 0; // tracks position within current label row (0, 1, or 2)
    for (let i = 0; i < allLabels.length; i++) {
        const isNewRoute = i > 0 && allLabels[i].Route !== allLabels[i - 1].Route;
        if (isNewRoute && slotCount !== 0) {
            // Fill remaining slots in this row with blanks, then reset
            const padNeeded = LABELS_PER_ROW - slotCount;
            for (let p = 0; p < padNeeded; p++) {
                paddedLabels.push({ _pad: true, DeliveryIndex: -1 });
                slotCount = (slotCount + 1) % LABELS_PER_ROW;
            }
        }
        paddedLabels.push(allLabels[i]);
        slotCount = (slotCount + 1) % LABELS_PER_ROW;
    }

    return { allLabels: paddedLabels, labelFontSizeMap };
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
                const lbl = allLabels[labelIndex++];
                pageSlots.push(lbl._pad ? emptySlot() : deliverySlot(lbl, config));
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
 *
 * PERFORMANCE: All formatting is collected into batch arrays first, then written
 * in as few setValues/setFontSizes/setFontColors calls as possible.
 * Individual getRange calls inside a loop are the #1 cause of timeouts in Apps Script.
 */
function writeAndFormatContinuousLabels(outputSheet, finalOutput, outputLabelMapping, stampSlots, allLabels, labelFontSizeMap, config) {
    const totalRows = finalOutput.length;
    const totalCols = TOTAL_COLS;

    // 1. Write all data in one call
    outputSheet.getRange(1, 1, totalRows, totalCols).setValues(finalOutput);

    // 2. Set column widths
    const width = 258;
    outputSheet.setColumnWidth(1, width * 0.7);
    outputSheet.setColumnWidth(2, width * 0.3);
    outputSheet.setColumnWidth(3, width * 0.7);
    outputSheet.setColumnWidth(4, width * 0.3);
    outputSheet.setColumnWidth(5, width * 0.7);
    outputSheet.setColumnWidth(6, width * 0.3);

    // 3. Set all row heights in two batch calls (odd rows and even rows)
    //    setRowHeights(startRow, numRows, height) is much faster than one-by-one
    for (let r = 1; r <= totalRows; r++) {
        outputSheet.setRowHeight(r, 51.5);
    }

    // 4. Apply base formatting to the entire sheet in one pass
    const fullRange = outputSheet.getRange(1, 1, totalRows, totalCols);
    fullRange
        .setFontWeight("bold")
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setFontColor("#000000")
        .setBorder(false, false, false, false, false, false);

    // 5. Build batch arrays for the per-slot formatting that varies by slot
    //    fontSizes[r][c] and fontColors[r][c] are 1-indexed to match sheet rows/cols
    //    We build them as 2D arrays spanning the full sheet then write once per property.
    const fontSizeGrid  = Array.from({length: totalRows}, () => new Array(totalCols).fill(12));
    const fontColorGrid = Array.from({length: totalRows}, () => new Array(totalCols).fill("#000000"));
    const fontWeightGrid = Array.from({length: totalRows}, () => new Array(totalCols).fill("bold"));
    const hAlignGrid    = Array.from({length: totalRows}, () => new Array(totalCols).fill("center"));

    let labelCounter = 0;
    let rowPairIndex = 0;
    const mergeTargets = []; // collect merge operations — do them all after the grid writes

    for (let r = 1; r <= totalRows; r += LABEL_HEIGHT_ROWS) {
        const ri  = r - 1;       // 0-based row index for grid arrays
        const ri2 = r;            // 0-based index of the second row in this pair

        const isStampRow = stampSlots.has(rowPairIndex);

        for (let j = 0; j < LABELS_PER_ROW; j++) {
            const colNameStart = j * 2 + 1; // sheet col, 1-based
            const colBagStart  = j * 2 + 2;
            const ci           = colNameStart - 1; // 0-based col index
            const ciBag        = colBagStart  - 1;

            const isStampSlot = isStampRow && (j === LABELS_PER_ROW - 1);

            if (isStampSlot) {
                // Row 1 of stamp: "Page N" — larger, centered across 2 cols
                fontSizeGrid[ri][ci]    = 14;
                fontSizeGrid[ri][ciBag] = 14;
                // Row 2 of stamp: timestamp — smaller
                fontSizeGrid[ri2][ci]    = 9;
                fontSizeGrid[ri2][ciBag] = 9;
                fontWeightGrid[ri2][ci]    = "normal";
                fontWeightGrid[ri2][ciBag] = "normal";
                mergeTargets.push([r,     colNameStart, 1, 2]);
                mergeTargets.push([r + 1, colNameStart, 1, 2]);

            } else if (labelCounter < allLabels.length) {
                const currentLabel = allLabels[labelCounter];
                // _pad sentinels are empty slots — skip formatting, just advance counter
                if (currentLabel._pad) {
                    labelCounter++;
                    continue;
                }
                const formatting   = labelFontSizeMap[currentLabel.DeliveryIndex];
                const nameFontSize = formatting ? formatting.name : 12;

                let row2Content = "";
                if (config.includeDriverNames)                row2Content = currentLabel.Driver;
                else if (config.includeRouteNames)            row2Content = currentLabel.Route;
                else if (config.includeAbbreviatedRouteNames) row2Content = currentLabel.AbbrevRoute;

                const detailFontSize = row2Content.length > 20 ? 11 : 13;
                const detailColor    = row2Content !== "" ? "#3C78D8" : "#000000";

                // Row 1, name cell
                fontSizeGrid[ri][ci]  = nameFontSize;
                // Row 1, bag number cell — normal weight, left-aligned, smaller
                fontSizeGrid[ri][ciBag]   = 10;
                fontWeightGrid[ri][ciBag] = "normal";
                hAlignGrid[ri][ciBag]     = "left";
                // Row 2, detail line (spans 2 cols after merge)
                fontSizeGrid[ri2][ci]    = detailFontSize;
                fontSizeGrid[ri2][ciBag] = detailFontSize;
                fontColorGrid[ri2][ci]   = detailColor;
                fontColorGrid[ri2][ciBag]= detailColor;

                mergeTargets.push([r + 1, colNameStart, 1, 2]);
                labelCounter++;
            }
            // empty filler slots: base formatting from step 4 is fine, nothing extra needed
        }
        rowPairIndex++;
    }

    // 6. Write the batch grids — 4 calls instead of hundreds
    outputSheet.getRange(1, 1, totalRows, totalCols).setFontSizes(fontSizeGrid);
    outputSheet.getRange(1, 1, totalRows, totalCols).setFontColors(fontColorGrid);
    outputSheet.getRange(1, 1, totalRows, totalCols).setFontWeights(fontWeightGrid);
    outputSheet.getRange(1, 1, totalRows, totalCols).setHorizontalAlignments(hAlignGrid);

    // 7. Merges — must happen after values are written
    mergeTargets.forEach(([row, col, numRows, numCols]) => {
        const rng = outputSheet.getRange(row, col, numRows, numCols);
        if (!rng.isPartOfMerge()) rng.mergeAcross();
    });

    // 8. Safely delete excess rows and columns
    const maxRows = outputSheet.getMaxRows();
    if (maxRows > totalRows) outputSheet.deleteRows(totalRows + 1, maxRows - totalRows);
    const maxCols = outputSheet.getMaxColumns();
    if (maxCols > totalCols) outputSheet.deleteColumns(totalCols + 1, maxCols - totalCols);
}

/**
 * Exports the already-generated Route_ tabs to PDF.
 * Does NOT regenerate the tabs — run item 3 first, then run this.
 * Keeping generation and export separate avoids hitting the 6-minute
 * execution limit (generation alone takes ~360s for 27 routes).
 */
function makeDriverRouteTabsThenPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("makeDriverRouteTabsThenPDF: export only (no regeneration).");

    // Verify at least one Route_ sheet exists
    const routeSheets = ss.getSheets().filter(sh => sh.getName().startsWith('Route_'));
    if (routeSheets.length === 0) {
        SpreadsheetApp.getUi().alert('No Route_ sheets found. Run item 3 (Make/Update Driver Route Tabs) first.');
        return;
    }
    Logger.log(`Found ${routeSheets.length} Route_ sheets to export.`);

    SpreadsheetApp.flush();

    const sanitizedTimestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss")
        .replace(/[\/:]/g, '-').replace(/\s/g, '_');
    const filename = `Driver_Route_Sheets_${sanitizedTimestamp}.pdf`;
    Logger.log(`PDF filename: ${filename}`);

    // Hide all non-Route_ sheets temporarily
    const allSheets    = ss.getSheets();
    const hiddenSheets = [];
    allSheets.forEach(sh => {
        if (!sh.getName().startsWith('Route_') && !sh.isSheetHidden()) {
            sh.hideSheet();
            hiddenSheets.push(sh);
        }
    });
    Logger.log(`Hid ${hiddenSheets.length} non-route sheets.`);

    const urlBase = ss.getUrl().replace(/edit$/, '');
    const params  = {
        'format': 'pdf', 'portrait': 'true', 'size': 'letter',
        'fzr': 'false', 'fitw': 'true',
        'top_margin': '0.25', 'bottom_margin': '0.25',
        'left_margin': '0.1', 'right_margin': '0.1'
    };
    const pdfUrl = `${urlBase}export?${Object.keys(params).map(k => `${k}=${params[k]}`).join('&')}`;

    Logger.log("Starting PDF fetch.");
    const token = ScriptApp.getOAuthToken();
    let blob;
    try {
        const response = UrlFetchApp.fetch(pdfUrl, {
            headers: { Authorization: 'Bearer ' + token },
            muteHttpExceptions: true
        });
        blob = response.getBlob().setName(filename);
        Logger.log("PDF fetch successful.");
    } catch (e) {
        SpreadsheetApp.getUi().alert(`PDF Export Failed: ${e.toString()}`);
        Logger.log(`ERROR during PDF fetch: ${e.toString()}`);
    }

    hiddenSheets.forEach(sh => sh.showSheet());
    Logger.log("Sheets restored.");

    if (blob) {
        DriveApp.createFile(blob);
        SpreadsheetApp.getUi().alert(`✅ Driver Route Sheets PDF created: "${filename}" — check your Google Drive!`);
    } else {
        SpreadsheetApp.getUi().alert(`❌ PDF could not be saved. Check logs for details.`);
    }
    Logger.log("makeDriverRouteTabsThenPDF finished.");
}

function makePackingListThenPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Generate both tabs
    makeNaviDash();
    const timestamp = makePackingLists();

    if (timestamp === "ERROR_NO_SHEET") return;

    SpreadsheetApp.flush();

    const sanitizedTimestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss")
        .replace(/[\/:]/g, '-').replace(/\s/g, '_');
    const filename = `Packing_Lists_${sanitizedTimestamp}.pdf`;

    const sheet = ss.getSheetByName("PackingLists");
    if (!sheet) {
        SpreadsheetApp.getUi().alert("Error: 'PackingLists' sheet not found for PDF export.");
        return;
    }

    const sheetId = sheet.getSheetId();
    const url = ss.getUrl().replace(/edit$/, '');
    const params = {
        'format': 'pdf', 'size': 'letter', 'portrait': 'false',
        'fitw': 'true', 'sheetnames': 'false', 'printtitle': 'false',
        'pagenumbers': 'false', 'gridlines': 'false', 'fzr': 'false',
        'gid': sheetId,
        'top_margin': '0.25', 'bottom_margin': '0.25',
        'left_margin': '0.25', 'right_margin': '0.25'
    };
    const pdfUrl = `${url}export?${Object.keys(params).map(k => `${k}=${params[k]}`).join('&')}`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(pdfUrl, { headers: { 'Authorization': `Bearer ${token}` }, muteHttpExceptions: true });
    DriveApp.createFile(response.getBlob().setName(filename));
    SpreadsheetApp.getUi().alert(`✅ Packing Lists PDF created: "${filename}" — check your Google Drive!`);
}

/**
 * Generates the Runners Dashboard tab then exports it as a PDF.
 * Wired to menu item 9.
 */
function makeNaviDashThenPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    makeNaviDash();
    SpreadsheetApp.flush();

    const sanitizedTimestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss")
        .replace(/[\/:]/g, '-').replace(/\s/g, '_');
    const filename = `Navigator_Dashboard_${sanitizedTimestamp}.pdf`;

    const sheet = ss.getSheetByName("NaviDash");
    if (!sheet) {
        SpreadsheetApp.getUi().alert("Error: 'NaviDash' sheet not found for PDF export.");
        return;
    }

    const sheetId = sheet.getSheetId();
    const url = ss.getUrl().replace(/edit$/, '');
    const params = {
        'format': 'pdf', 'size': 'letter', 'portrait': 'true',
        'fitw': 'true', 'sheetnames': 'false', 'printtitle': 'false',
        'pagenumbers': 'false', 'gridlines': 'false', 'fzr': 'false',
        'gid': sheetId,
        'top_margin': '0.25', 'bottom_margin': '0.25',
        'left_margin': '0.25', 'right_margin': '0.25'
    };
    const pdfUrl = `${url}export?${Object.keys(params).map(k => `${k}=${params[k]}`).join('&')}`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(pdfUrl, { headers: { 'Authorization': `Bearer ${token}` }, muteHttpExceptions: true });
    DriveApp.createFile(response.getBlob().setName(filename));
    SpreadsheetApp.getUi().alert(`✅ Navigator Dashboard PDF created: "${filename}" — check your Google Drive!`);
}

/**
 * Asks which label mode to use, generates the sheet, then exports to PDF.
 * Called directly from the menu, or pass a forcedMode to skip the prompt.
 * @param {string} [forcedMode] - optional: "drivers"|"routes"|"abbreviated"|"none"
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
    .addItem("1. Filter Data for Map (Col I)",          "filterForMapCreation")
    .addItem("2. Clear Map Data Filter",                "clearMapFilter")
    .addSeparator()
    .addItem("3. Make/Update Driver Route Tabs",        "makeDriverRouteTabs")
    .addItem("4. Make/Update NaviDash",                 "makeNaviDash")
    .addItem("5. Make/Update Packing Lists Tab",        "makePackingLists")
    .addItem("6. Make/Update Delivery Labels Tab",      "makeTheLabels")
    .addSeparator()
    .addItem("7. Export Route Tabs → PDF  (run 3 first)",      "makeDriverRouteTabsThenPDF")
    .addItem("8. Make NaviDash then PDF",               "makeNaviDashThenPDF")
    .addItem("9. Make Delivery Labels Tab then PDF",    "makeLabelsPDF")
    .addToUi();

  ui.createMenu("📦 Label Generator")
    .addItem("Labels — Driver Names",                   "generateWithDrivers")
    .addItem("Labels — Full Route Names",               "generateWithRoutes")
    .addItem("Labels — Abbreviated Route Codes",        "generateWithAbbreviatedRoutes")
    .addItem("Labels — No Route Info",                  "generateWithoutDrivers")
    .addSeparator()
    .addItem("Labels — Choose Mode + Export PDF",       "makeLabelsPDF")
    .addToUi();

  ui.createMenu("📦 Sheet Tools")
    .addItem("Delete Route Sheets",                     "deleteSheetsByPrefix")
    .addSeparator()
    .addItem("⛔ Cancel Route Tab Generation",          "cancelDriverRouteTabs")
    .addToUi();
}