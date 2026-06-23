# Contributing to WAWG Distribution Day Scripts

## How the Code Gets Deployed

This is a Google Apps Script project. There is no build step and no npm. The entire codebase lives in one file — `Code.js` — which gets copy-pasted directly into the Apps Script editor inside the WAWG Google Sheet.

**Deploy process:**
1. Make and test your changes locally in `Code.js`
2. Open the WAWG Google Sheet
3. Click **Extensions → Apps Script**
4. Select all existing code and replace it with the contents of `Code.js`
5. Replace `appsscript.json` contents with the repo version (preserves the required OAuth scopes)
6. Click **Save** (floppy disk icon or Ctrl+S)
7. Reload the Google Sheet and test from the menus

---

## Project Structure

One file: `Code.js`. Organized into logical sections:

| Section | Key functions |
|---|---|
| Top of file | Constants (`ROUTE_BATCH_SIZE`, `AUTO_GENERATE_MAP_LINKS`, `PROP_*`), `refreshDeliveriesHelpTable` |
| Batched route generation | `makeDriverRouteTabs`, `continueDriverRouteTabs`, `cancelDriverRouteTabs`, `_routeBatch_*` helpers |
| Route sheet helpers | `buildRouteSheet`, `writeFullRouteHeader`, `writeStopTable`, `writeSTR1`–`writeSTR5`, `routeSheetStopBucket` |
| Map utilities | `buildAutoMapUrl`, `buildRouteMapInfo`, `extractMapLinkForRoute`, `shortenUrlForTyping` |
| Navigator Dashboard | `makeNaviDash`, `writeDashboardTitle`, `generateSpecialItemsTotalsTable`, `generatePerRouteSpecialItemsBreakdown`, `generateSpecialRequestsFlagList`, `generateItemSummaryTable`, `generateConfirmationStatusTable` |
| Packing lists | `makePackingLists`, `generatePackingListRouteSections`, `packingListSortBucket`, `loadAndIndexPackingData` |
| Label generation | `generateDeliveryLabels`, `generateAllIndividualLabels`, `formatContinuousLabelData`, `writeAndFormatContinuousLabels`, `abbreviateRoute`, `formatRouteCode` |
| PDF exports | `makeDriverRouteTabsThenPDF`, `makeNaviDashThenPDF`, `makePackingListThenPDF`, `makeLabelsPDF` |
| Sheet tools | `makeTheLabels`, `filterForMapCreation`, `clearMapFilter`, `deleteSheetsByPrefix` |
| Menu | `onOpen` |

---

## Key Constraints

**Apps Script execution limit is 6 minutes per function call.** Driver route sheet generation (~14s per route × 27 routes = ~380s) exceeds this limit. The `makeDriverRouteTabs` function uses a `PropertiesService` + `ScriptApp.newTrigger` continuation system to split work into batches of 8 routes with a 30-second gap between batches. Total wall time is ~8 minutes across 4 executions.

**Route tab generation and PDF export are separate steps.** `makeDriverRouteTabsThenPDF` (menu item 7) does NOT regenerate the tabs — it exports whatever `Route_` tabs already exist. Run item 3 first, wait for completion, then run item 7. This split is intentional and must be preserved.

**`continueDriverRouteTabs` must never appear in the menu.** It is the trigger target function and is called automatically. Adding it to the menu risks users firing it manually and corrupting batch state.

**Minimize individual `getRange` calls inside loops.** Sheets API round-trip overhead dominates execution time. Collect formatting into 2D arrays and write in one batch call. Never call `setFontSize`, `setFontColor`, etc. one cell at a time inside a loop.

**`SpreadsheetApp.flush()` is called before PDF export.** Ensures all sheet writes are committed before the export URL is fetched. Don't remove it.

**The dashboard uses exactly 10 columns (A:J), portrait orientation.** All table banners, merges, and column widths in the NaviDash functions are calibrated for this layout. If you extend a table beyond column J it will break the print layout.

---

## Batch Continuation System

The `_routeBatch_*` private helpers manage state across executions:

| Function | Purpose |
|---|---|
| `_routeBatch_prepareData` | Loads and pre-computes all data needed for route sheets |
| `_routeBatch_runBatch` | Processes N routes, appends to summary rows in PropertiesService |
| `_routeBatch_finish` | Writes index sheet, clears state, writes completion stamp to Driver_Deliveries col G |
| `_routeBatch_scheduleTrigger` | Creates a one-shot 30-second trigger calling `continueDriverRouteTabs` |
| `_routeBatch_cancelTriggers` | Deletes all `continueDriverRouteTabs` triggers by function name |
| `_routeBatch_clearState` | Deletes all `WAWG_ROUTE_*` PropertiesService keys |
| `_routeBatch_uiAlert` | Safe alert wrapper — falls back to Logger when called from trigger context |

PropertiesService keys used:

| Key | Contains |
|---|---|
| `WAWG_ROUTE_NEXT_IDX` | Integer index of next unprocessed route |
| `WAWG_ROUTE_NAMES` | JSON array of all route names in order |
| `WAWG_ROUTE_SUMMARY_ROWS` | JSON array of index sheet rows accumulated across batches |
| `WAWG_ROUTE_IN_PROGRESS` | `"true"` while a batch job is running |
| `WAWG_ROUTE_TRIGGER_ID` | Unique ID of the most recently scheduled trigger |

---

## Stop Sort Order

Both driver route sheets and packing lists sort stops/rows within each route by special-items complexity. The sort bucket function (`routeSheetStopBucket` / `packingListSortBucket`) assigns:

| Bucket | Condition |
|---|---|
| 1 | No special items |
| 2 | Dog food only |
| 3 | Dog + cat food |
| 4 | Cat food only |
| 5 | Diapers and/or wipes |
| 6 | Hygiene only (toiletries / fem hygiene) |
| 7 | Has special requests — always last |

Special requests always sort last regardless of what other items the delivery has.

---

## Map Link Toggle

```javascript
const AUTO_GENERATE_MAP_LINKS = false; // top of Code.js
```

- `false` (default) — reads from the `Delivery Map Links` column in the source sheet
- `true` — auto-generates a Google Maps waypoint URL (`/maps/search/`) from stop addresses, strips the route name segment from each address, and shortens via is.gd

The toggle only affects the route sheet header map link. The `Driver_Deliveries` index sheet always auto-generates its map URLs regardless of this setting.

---

## Adding a New Route

All route abbreviations live in `ROUTE_MAP` inside `abbreviateRoute()`. To add a new route:

1. Find `const ROUTE_MAP = {` in `Code.js`
2. Add one line: `"Exact Route Name As In Sheet": "NEWCOD",`
3. Rules:
   - Standard: 6 alpha + optional integer (`NCOAST1`)
   - Directional: 5 alpha + direction letter + optional integer (`CHVSTW1`)
   - Omit the number if only one route exists with that base name
   - Avoid ending a 6-letter code in N, S, E, or W — `formatRouteCode()` will misread it as a directional

If a route has no entry in `ROUTE_MAP`, the full route name prints as a fallback. Nothing breaks silently.

---

## Changing the Default Label Mode

Menu item 9 (and the Label Generator menu) prompts the user. No hardcoded default. Choices:

- `1` — Driver Names
- `2` — Full Route Names
- `3` — Abbreviated Route Codes
- `4` — No Route Info

---

## Testing Checklist

Before opening a PR, test manually in the live Google Sheet:

- [ ] Item 3 completes all routes across batches without error
- [ ] `Driver_Deliveries` col G shows `✅ Complete: <timestamp>` after batch finishes
- [ ] Item 7 exports PDF from existing Route_ tabs (does not regenerate)
- [ ] Item 4 generates NaviDash — all six sections present, portrait A:J layout
- [ ] Item 8 exports NaviDash PDF — portrait, filename `Navigator_Dashboard_*`
- [ ] Item 5 generates PackingLists — routes sorted alphabetically, stops sorted by bucket
- [ ] All four label modes generate without error
- [ ] Abbreviated codes display with spaces: `CHVST W 1` not `CHVSTW1`
- [ ] Stamp label appears in slot 30 of every Avery sheet; slot 30 is never a delivery label
- [ ] New routes start on a fresh label row (not mid-row)
- [ ] ⛔ Cancel Route Tab Generation clears all PropertiesService state and deletes pending triggers

---

## What Not to Do

- **Do not add a combined "Make All PDFs" runner.** Each deliverable runs close to the 6-minute limit on its own. They cannot be chained.
- **Do not call formatting methods one cell at a time inside a label loop.** Use batch 2D array writes.
- **Do not change the Avery 6240 layout constants** (`LABEL_ROWS_PER_PAGE = 10`, `LABELS_PER_ROW = 3`) without re-testing print alignment on physical label stock.
- **Do not extend NaviDash tables beyond column J.** The PDF export is calibrated for portrait A:J. Going wider breaks the print layout silently.
- **Do not put `continueDriverRouteTabs` in a menu.** It is a trigger target only.

---

## Tech Debt

- `Packing_Lists_PackingLists.pdf` — the packing list PDF filename does not include a timestamp like other PDF exports. Low priority since the packing list is not printed from PDF anyway.
- Packing list page breaks cannot be set programmatically in Apps Script. The operator must configure them manually in Sheets before printing. Until Google exposes a page break API (or until the packing list is restructured as a separate per-route tab system similar to the driver route sheets), this manual step cannot be eliminated.
- `populateRouteTabs` — legacy route sheet writer, retained as a fallback but superseded by `buildRouteSheet`. Can be removed once `buildRouteSheet` has been validated across several distribution cycles.
- `applyRouteColumnWidths` — small utility function that appears to be unused. Safe to remove in a cleanup pass.

---

## License & Sharing

MIT licensed — deliberately. Contributors and adopting orgs are not required to share their changes. You decide what you share and with whom.

🐙 **[github.com/ej9erfan](https://github.com/ej9erfan)**
💬 **[Open a Discussion](https://github.com/ej9erfan/WAWG-Distribution-Scripts/discussions)**
