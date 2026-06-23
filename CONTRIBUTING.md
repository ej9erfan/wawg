# Contributing to WAWG Distribution Day Scripts

## How the Code Gets Deployed

This is a Google Apps Script project. There is no build step and no npm. The entire codebase lives in one file — `Code.js` — which gets copy-pasted directly into the Apps Script editor inside the WAWG Google Sheet.

**Deploy process:**
1. Make and test your changes locally in `Code.js`
2. Open the WAWG Google Sheet
3. Click **Extensions → Apps Script**
4. Select all existing code and replace it with the contents of `Code.js`
5. Click **Save** (floppy disk icon or Ctrl+S)
6. Reload the Google Sheet and test from the menus

---

## Project Structure

There is one file: `Code.js`. It is organized into logical sections:

| Section | What's in it |
|---|---|
| Top of file | `refreshDeliveriesHelpTable`, helper utilities |
| `makeDriverRouteTabs` | Builds per-route sheets for drivers |
| `makePackingLists` | Builds the PackingLists tab |
| Label generation | `generateDeliveryLabels`, `formatContinuousLabelData`, `writeAndFormatContinuousLabels`, `abbreviateRoute`, `formatRouteCode` |
| PDF exports | `makeDriverRouteTabsThenPDF`, `makePackingListThenPDF`, `makeLabelsPDF` |
| Sheet tools | `makeTheLabels`, `filterForMapCreation`, `clearMapFilter`, `deleteSheetsByPrefix` |
| Menu | `onOpen` — defines all three custom menus |

---

## Key Constraints

**Apps Script execution limit is 6 minutes per function call.** Each of the three main deliverable functions (`makeDriverRouteTabs`, `makePackingLists`, `generateDeliveryLabels`) runs close to this limit on a full dataset. Do not attempt to chain them into a single function call. The `makeAllPDFs` combined runner was removed for this reason.

**Minimize individual `getRange` calls inside loops.** The Sheets API call overhead is the primary driver of execution time. Collect formatting into 2D arrays and write them in one batch call (see `writeAndFormatContinuousLabels` for the pattern). Never call `setFontSize`, `setFontColor`, etc. one cell at a time inside a loop over labels or rows.

**`SpreadsheetApp.flush()` is called before PDF export.** This ensures all sheet writes are committed before the export URL is fetched. Don't remove it.

---

## Adding a New Route

All route abbreviations live in `abbreviateRoute()` inside a `ROUTE_MAP` lookup table. To add a new route:

1. Find `const ROUTE_MAP = {` in `Code.js`
2. Add one line: `"Exact Route Name As In Sheet": "NEWCOD",`
3. Follow the format rules in the comment header above the table:
   - Standard: 6 alpha + optional integer (`NCOAST1`)
   - Directional: 5 alpha + direction letter + optional integer (`CHVSTW1`)
   - Omit the number if only one route exists with that base name
   - Avoid ending a 6-letter code in N, S, E, or W — `formatRouteCode()` will misread it as a directional

If a route exists in the sheet but has no entry in `ROUTE_MAP`, the full route name prints as a fallback. Nothing breaks silently — you'll see the long name on the label.

---

## Changing the Default Label Mode

When **Make Delivery Labels Tab then PDF** (menu item 8) is run from the menu it prompts the user. The prompt accepts:

- `1` — Driver Names
- `2` — Full Route Names
- `3` — Abbreviated Route Codes
- `4` — No Route Info

There is no hardcoded default for the interactive prompt — the user must choose.

---

## Testing

Before opening a PR, test the following manually in the live Google Sheet:

- [ ] All four label modes generate without error
- [ ] Row 2 is blank in No Route Info mode
- [ ] Blue text appears on row 2 in all non-blank modes
- [ ] Abbreviated codes display with spaces: `CHVST W 1` not `CHVSTW1`
- [ ] Stamp label appears in slot 30 of every Avery sheet; slot 30 is never a delivery label
- [ ] `makeTheLabels` (menu item 5) prompts and runs without error
- [ ] `makeLabelsPDF` (menu item 8) prompts for mode and exports to Drive
- [ ] `makeDriverRouteTabsThenPDF` (menu item 6) completes within ~5 minutes
- [ ] `makePackingListThenPDF` (menu item 7) completes within ~5 minutes
- [ ] Print alignment on Avery 6240 stock is not affected by label changes

---

## What Not to Do

- **Do not add a `makeAllPDFs` combined runner.** It was removed intentionally. See the PR history for context.
- **Do not use `PropertiesService` for trigger-based continuation chains.** This approach was attempted and removed — it is too slow and too fragile for the data sizes involved.
- **Do not call formatting methods one cell at a time inside a label loop.** Use batch 2D array writes.
- **Do not change the Avery 6240 layout constants** (`LABEL_ROWS_PER_PAGE = 10`, `LABELS_PER_ROW = 3`) without re-testing print alignment on physical label stock.
---

## License & Sharing

This project is MIT licensed. That was a deliberate choice — contributors and adopting orgs are not required to share their changes publicly or with anyone. You decide what you share and with whom. That boundary is yours.

If you do want to share improvements, compare notes, or just say what your org did with this — reach out:

🐙 **[github.com/ej9erfan](https://github.com/ej9erfan)**
💬 **[Open a Discussion](https://github.com/ej9erfan/WAWG-Distribution-Scripts/discussions)**

### Tech Debt

- `Packing_Lists_PackingLists.pdf` - This filename result in Google Drive does not get a date/time stamp like the other "make X then PDF" functions. Not really important. This PDF does not get used, because the page breaks are not auto aligned and printing must occur in Google Sheets to align them correctly. Please leave this here until it is somehow possible to make a print job of all of this informatoion in one tab and not split tables across pages. Alternatively, rewrite the Packing lists to act just like the Driver tabs/sheets or roll them all into an all-in-one solution. Which is not desireable becaue the Packing Lists include some administrative information tabls as well. Perhaps the asnwer is to split the driver and packing data from the admin tables....the make a all-in-one route-centric solution for both drivers and preparers.
