# WAWG Distribution Day Scripts — README

This document explains how the Google Sheets automation works, what it produces, how to run it, and how to print everything correctly on Distribution Day.

---

## What Gets Generated

Running the automation produces **four deliverables**:

| # | Deliverable | Sheet Tab | PDF File | Printed On |
|---|---|---|---|---|
| 1 | Driver Route Sheets | `Route_<name>` (one per route) | `Driver_Route_Sheets_<timestamp>.pdf` | Regular paper |
| 2 | Navigator Dashboard | `NaviDash` | `Navigator_Dashboard_<timestamp>.pdf` | Regular paper |
| 3 | Packing Lists | `PackingLists` | — (print from Sheets, not PDF) | Regular paper |
| 4 | Delivery Labels | `DeliveryLabels` | `Printable_Labels_<timestamp>.pdf` | Avery 6240 label sheets |

PDFs are saved automatically to your **Google Drive**. Packing Lists are printed directly from Sheets (see Printing section).

---

## Prerequisites — The Source Sheet

Everything runs from the **`Deliveries-UPDATE HERE`** tab. That sheet must have the following column headers (exact spelling matters):

| Column | Description |
|---|---|
| `Name` | Recipient name |
| `Address` | Delivery address — format: `Street, RouteName, City, State` |
| `Phone` | Phone number |
| `routeDescription` | Route name — must match entries in `DeliveriesHelpTable` |
| `Driver` | Assigned driver name |
| `Conf for <date>` | Confirmation status — `YES` or `NO ANS` to include, `NO` or blank to skip |
| `diapers1(size)` | First diaper size (if any) |
| `diapers2(size)` | Second diaper size (if any) |
| `diapers3(size)` | Third diaper size (if any) |
| `wipes(qty)` | Wipes quantity |
| `dogFood(qty)` | Dog food quantity |
| `catFood(qty)` | Cat food quantity |
| `toiletries(qty)` | Toiletries quantity |
| `fem hygiene(qty)` | Fem hygiene quantity |
| `# packs` | Food pack count |
| `# eggs` | Egg count |
| `# milk` | Milk count |
| `Special Item Requests` | Free-text special items |
| `Special Delivery Instructions` | Free-text delivery notes |
| `Delivery Map Links` | Hyperlink to the route map (used when `AUTO_GENERATE_MAP_LINKS = false`) |

> Only rows where `Conf for` equals `YES` or contains `NO ANS` are included in any deliverable. Rows marked `NO` or left blank are skipped.

---

## Running the Automation

### Menus

Three custom menus appear when the sheet opens:

**📦 Distribution Day Tasks**

| Item | What it does |
|---|---|
| 1. Filter Data for Map (Col I) | Filters source sheet to rows with a map link |
| 2. Clear Map Data Filter | Removes that filter |
| 3. Make/Update Driver Route Tabs | Starts batched route sheet generation (see below) |
| 4. Make/Update NaviDash | Rebuilds the Navigator Dashboard tab |
| 5. Make/Update Packing Lists Tab | Rebuilds the PackingLists tab |
| 6. Make/Update Delivery Labels Tab | Rebuilds the DeliveryLabels tab (prompts for label style) |
| 7. Export Route Tabs → PDF (run 3 first) | Exports existing Route_ tabs to PDF — does NOT regenerate |
| 8. Make NaviDash then PDF | Rebuilds NaviDash and exports PDF |
| 9. Make Delivery Labels Tab then PDF | Prompts for label style, generates sheet and exports PDF |

**📦 Label Generator**

Generates and exports a label PDF directly in a chosen style without the interactive prompt.

**📦 Sheet Tools**

| Item | What it does |
|---|---|
| Delete Route Sheets | Deletes all `Route_` prefixed tabs |
| ⛔ Cancel Route Tab Generation | Stops an in-progress batched generation and clears all state |

---

### Recommended Workflow — Distribution Day

Route tab generation and PDF export are now **two separate steps** (items 3 and 7). This is intentional — generation takes ~8 minutes across batches and PDF export needs its own fresh execution budget.

```
Item 3  →  wait for "All route tabs complete!" (takes ~8 min, runs automatically)
Item 7  →  export Route Tabs PDF → start print job
Item 4  →  NaviDash → Item 8 → PDF → print
Item 5  →  Packing Lists → print from Sheets (see below)
Item 6  →  Labels → Item 9 → PDF → print on Avery 6240
```

---

### Batched Route Tab Generation

Driver route sheets take ~14 seconds per route and cannot complete for 27 routes inside Apps Script's 6-minute execution limit. The solution is automatic batching:

- **Item 3** processes routes 1–8 immediately, then schedules a trigger to continue
- Each subsequent batch fires ~30 seconds after the previous one finishes
- A `✅ Complete: <timestamp>` stamp appears in column G of the `Driver_Deliveries` index tab when all routes are done
- The Executions panel in Apps Script also shows each batch completing

**Do not run Item 3 again while a batch job is in progress.** If something goes wrong, use **⛔ Cancel Route Tab Generation** in Sheet Tools to clear all state, then run Item 3 fresh.

> **Authorization note:** The first time you run Item 3, Apps Script will request permission to manage triggers (`script.scriptapp` scope). Accept the permission prompt or the batching will not work. This only happens once.

---

### Packing Lists — No PDF

The Packing Lists tab is intentionally not exported to PDF by the script. Page breaks across route tables cannot be set programmatically, so the layout must be configured manually before printing. See the Printing section below.

---

### Label Style Options

When prompted (items 6, 9, or Label Generator menu):

| # | Style | Row 2 of each label |
|---|---|---|
| 1 | Driver Names | Assigned driver's name |
| 2 | Full Route Names | Full `routeDescription` value |
| 3 | Abbreviated Route Codes | Short code (see table below) |
| 4 | No Route Info | Blank |

Labels within each route are sorted by special-items complexity: plain deliveries first, then pet food, then diapers/wipes, then hygiene, then any delivery with special requests last.

---

### Abbreviated Route Codes

| Route | Code |
|---|---|
| Azalea | `AZALEA` |
| Chula Vista East | `CHVST E 1` |
| Chula Vista West | `CHVST W 1` / `CHVST W 2` / `CHVST W 3` |
| Downtown | `DTOWN 1` / `DTOWN 2` |
| El Cajon | `ELCAJO 1` / `ELCAJO 2` / `ELCAJO 3` |
| Encanto / Lemon Grove | `ENCNTO 1` / `ENCNTO 2` |
| Grantville / College East | `GRNTVI` |
| Lake Murray / La Mesa | `LMRAY 1` / `LMRAY 2` |
| Logan Heights | `LOGANH 1` / `LOGANH 2` |
| Mountain View / National City | `MTNVCY 1` / `MTNVCY 2` |
| Normal Heights | `NORMHT` |
| North Coast | `NCOAST 1` / `NCOAST 2` |
| Teralta | `TERALT 1` / `TERALT 2` |
| Volunteers | `VOLNTR` |
| Walking Delivery | `WLKDLV 1` / `WLKDLV 2` |

> To add a new route: open `Code.js`, find `ROUTE_MAP` inside `abbreviateRoute()`, and add one line: `"Exact Route Name": "NEWCOD"`. If the new code ends in N, S, E, or W and has a 5-letter prefix, it will be treated as a directional code — choose a different ending letter to avoid that.

---

## The Navigator Dashboard (NaviDash)

The NaviDash tab is the coordinator and runners lead view. It contains:

- **Standard items totals** — packs, eggs, milk (confirmed-only)
- **Confirmation status** — YES / NO ANS / NO / Total counts
- **Special Items Pull List** — dog food, cat food, hygiene, and every diaper size as individual rows, with columns for On Hand / Given Out / Not on Hand tracking
- **Per-Route Special Items Breakdown** — same items subtotaled per route so the runner lead can stage piles by driver
- **Special Requests Flag List** — every delivery with a non-empty Special Item Requests field, sorted by route, with request text spanning columns D–I and a `☐` checkbox in column J

The pull list and flag list both include a lighter italic footer row repeating the column headers, so the table is readable across a page break when printed.

---

## Printing

### Driver Route Sheets

1. Run Item 3, wait for the completion stamp in `Driver_Deliveries` column G
2. Run Item 7 to export the PDF
3. Download from Google Drive, open in **Adobe Acrobat/Reader**
4. Scale: **Fit** — Print on both sides: **off**
5. Print **Walking Delivery pages a second time** — those teams sometimes split up stops

### Navigator Dashboard

1. Run Item 8 to generate and export PDF
2. Download from Google Drive, open and print
3. Portrait, single-sided

### Packing Lists

The Packing Lists **must be printed from Google Sheets**, not from a PDF:

1. Go to the **PackingLists** tab
2. **File → Print**
3. Scale: **Fit to Width** — Margins: Custom, all four sides **0.15**
4. Open **page break preview** — drag blue lines to fit 2–3 routes per page
5. Click **Confirm Breaks**
6. Under Headers and Footers, check **Page Numbers** and **Sheet Name**
7. Click **Next → Print**

Print page 1 two extra times — it contains the item summary referenced throughout packing.

### Delivery Labels

1. Run Item 9, download PDF from Drive
2. Open in **Adobe Acrobat/Reader** — load **Avery 6240** sheets
3. Scale: **Fit** — Print on both sides: **off**
4. ⚠️ Use Avery brand only — Office Depot labels don't stick to grocery bags reliably
5. The bottom-right label on every sheet is a page number / timestamp stamp, not a delivery label

---

## Key Toggles and Constants

These live at the top of `Code.js` and are safe to change:

| Constant | Default | What it does |
|---|---|---|
| `ROUTE_BATCH_SIZE` | `8` | Routes processed per batch execution |
| `AUTO_GENERATE_MAP_LINKS` | `false` | `true` = build map URLs from addresses automatically; `false` = use `Delivery Map Links` column |

---

## End of Process Checklist

- 📄 **Driver Route Sheets** — one copy per driver, extra copies of Walking Delivery pages
- 🗺️ **Navigator Dashboard** — one copy for the runners lead / coordinator
- 📋 **Packing Lists** — full set, plus two extra copies of page 1
- 🏷️ **Label sheets** — one full set on Avery 6240 stock

---

## Community & Solidarity

This project is **MIT licensed** — deliberately. Other mutual aids and community orgs can pick this up, adapt it, and decide for themselves how openly they share what they build. You set your own boundaries.

If your org adapts this and you want to compare notes, we'd love to hear about it.

🐙 **[github.com/ej9erfan](https://github.com/ej9erfan)**
💬 **[Open a Discussion](https://github.com/ej9erfan/WAWG-Distribution-Scripts/discussions)**

Solidarity.
