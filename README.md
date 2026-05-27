# Distribution Day Deliverables — README

This document explains how the Google Sheets automation works, what it produces, how to run it, and how to print everything correctly on Distribution Day.

---

## What Gets Generated

Running the automation produces **three deliverables**:

| # | Deliverable | File Name | Printed On |
|---|---|---|---|
| 1 | Driver Route Sheets | `Driver_Route_Sheets_<timestamp>.pdf` | Regular paper |
| 2 | Packing Lists | `Packing_Lists_<timestamp>.pdf` | Regular paper (printed from Sheets, not Adobe) |
| 3 | Delivery Labels | `Printable_Labels_<timestamp>.pdf` | Avery 6240 label sheets |

All three PDF files are saved automatically to your **Google Drive**.

---

## Prerequisites — The Source Sheet

Everything runs from the **`Deliveries-UPDATE HERE`** tab. That sheet must have the following column headers (exact spelling matters):

| Column | Description |
|---|---|
| `Name` | Recipient name |
| `Address` | Delivery address |
| `Phone` | Phone number |
| `routeDescription` | The route name — must match entries in `DeliveriesHelpTable` |
| `Driver` | Assigned driver name |
| `Conf for <date>` | Confirmation status — rows must say `YES` or contain `NO ANS` to be included |
| `diapers1(size)` | First diaper size (if any) |
| `diapers2(size)` | Second diaper size (if any) |
| `diapers3(size)` | Third diaper size (if any) |
| `wipes(qty)` | Wipes quantity |
| `dogFood(qty)` | Dog food quantity |
| `catFood(qty)` | Cat food quantity |
| `# packs` | Food pack count |
| `# eggs` | Egg count |
| `# milk` | Milk count |
| `Special Item Requests` | Free-text special items |
| `Special Delivery Instructions` | Free-text delivery notes |
| `Delivery Map Links` | Hyperlink to the route map |

> **Note:** Only rows where `Conf for` equals `YES` or contains `NO ANS` are included in any deliverable. Rows marked `NO` or left blank are skipped.

---

## Running the Automation

### Recommended Workflow — Serial, Manual, One Sitting

Each deliverable takes around five minutes to generate. The intended approach is to run them one at a time and overlap the print jobs:

1. Click **6. Make Driver Route Tabs then PDF** → while it runs, set up your printer
2. When the Driver Route Sheets PDF is in Drive, start that print job
3. While it prints, click **7. Make Packing Lists Tab then PDF**
4. When the Packing Lists PDF is ready, configure and print it (see Printing section below)
5. While it prints, click **8. Make Delivery Labels Tab then PDF**
6. Print the labels last

This keeps you moving through all three deliverables in one focused sitting without waiting for one phase to finish before starting the next.

> **Why not "Make All PDFs" in one click?** That function has been removed. Each deliverable now takes ~5 minutes on its own, which exceeds the Apps Script execution limit. More importantly, the Packing Lists require manual print settings that can't be automated, so the operator has to be present for each print job regardless.

---

### What Happens During Each Phase

**Phase 1 — Driver Route Tabs**
The script rebuilds the `DeliveriesHelpTable` (route → driver mapping), then creates one new sheet tab per route. Each tab contains a header with the driver name, route name, map link, item totals, and a stop-by-stop delivery table. The old route tabs are deleted first. This takes 30–45 seconds. Do not interrupt it.

**Phase 2 — Packing Lists**
The script generates the `PackingLists` tab with item summaries, confirmation counts, a diaper reference table, and a full section-by-section breakdown of every stop on every route.

**Phase 3 — Delivery Labels**
The script generates the `DeliveryLabels` tab formatted for Avery 6240 label sheets (30 labels per sheet, 10 rows × 3 columns). The bottom-right label on every sheet is a **page number + timestamp stamp** instead of a delivery label. Each label shows the recipient name, bag count (e.g. `BAG 2/3`), and optionally a driver name, route name, or abbreviated route code on the second line.

---

### The Slower Way — Generate Sheets and PDFs Individually

Use the **📦 Distribution Day Tasks** menu for step-by-step control:

| Item | What it does |
|---|---|
| 1. Filter Data for Map (Col I) | Filters the source sheet to show only rows with a map link, for use with route mapping tools |
| 2. Clear Map Data Filter | Removes that filter |
| 3. Make/Update Driver Route Tabs | Rebuilds the per-route driver sheets only (no PDF) |
| 4. Make/Update Packing Lists Tab | Rebuilds the PackingLists tab only (no PDF) |
| 5. Make/Update Delivery Labels Tab | Rebuilds the DeliveryLabels tab only — prompts you to choose a label style |
| 6. Make Driver Route Tabs then PDF | Rebuilds driver tabs and exports the PDF to Drive |
| 7. Make Packing Lists Tab then PDF | Rebuilds packing list and exports the PDF to Drive |
| 8. Make Delivery Labels Tab then PDF | Prompts for label style, then generates the sheet and exports the PDF |

---

### Label Style Options

When prompted for a label style (menu items 5 and 8, or the Label Generator menu), enter:

| # | Style | Row 2 of each label shows |
|---|---|---|
| 1 | Driver Names | The assigned driver's name |
| 2 | Full Route Names | The full `routeDescription` value |
| 3 | Abbreviated Route Codes | A short code (see table below) |
| 4 | No Route Info | Blank — name and bag count only |

You can also generate labels directly without a PDF prompt from the **📦 Label Generator** menu.

---

### Abbreviated Route Codes

When using style **3 — Abbreviated Route Codes**, the second row of each label shows a short code formatted as:

- **Standard:** `NCOAST 2` — 6 letters + space + number
- **Directional:** `CHVST W 1` — 5 letters + space + direction + space + number
- **No number** when only one route exists with that name (e.g. `AZALEA`, `NORMHT`, `GRNTVI`)

| Route | Code |
|---|---|
| Downtown | `DTOWN 1` / `DTOWN 2` |
| North Coast | `NCOAST 1` / `NCOAST 2` |
| Normal Heights | `NORMHT` |
| Azalea | `AZALEA` |
| Teralta | `TERALT 1` / `TERALT 2` |
| El Cajon | `ELCAJO 1` / `ELCAJO 2` / `ELCAJO 3` |
| Chula Vista East | `CHVST E 1` |
| Chula Vista West | `CHVST W 1` / `CHVST W 2` / `CHVST W 3` |
| Encanto / Lemon Grove | `ENCNTO 1` / `ENCNTO 2` |
| Mountain View / National City | `MTNVCY 1` / `MTNVCY 2` |
| Grantville / College East | `GRNTVI` |
| Lake Murray / La Mesa | `LMRAY 1` / `LMRAY 2` |
| Logan Heights | `LOGANH 1` / `LOGANH 2` |
| Walking Delivery | `WLKDLV 1` / `WLKDLV 2` |
| Volunteers | `VOLNTR` |

> **Adding a new route:** Open `Code.js`, find the `ROUTE_MAP` inside `abbreviateRoute()`, and add one line: `"Exact Route Name": "NEWCOD"`. If the new code ends in the letter N, S, E, or W and has a 5-letter prefix, it will be treated as a directional code — choose a different ending letter to avoid that.

---

## Printing

### Driver Route Sheets

1. Download `Driver_Route_Sheets_<timestamp>.pdf` from Google Drive
2. Open in **Adobe Acrobat/Reader** (not the browser)
3. Click the print icon
4. Set scale to **Fit** (not Actual Size)
5. Make sure **Print on both sides** is **off**
6. Click Print

**Print the Walking Delivery pages a second time** — print only the Walking Delivery route sheets again so there are two copies. Delivery teams sometimes split up those stops.

---

### Delivery Labels

1. Download `Printable_Labels_<timestamp>.pdf` from Google Drive
2. Open in **Adobe Acrobat/Reader**
3. Load **Avery 6240** label sheets into your printer
   - ⚠️ Use Avery brand. Office Depot brand labels do not stick to grocery bags reliably.
4. Set scale to **Fit** (not Actual Size)
5. Make sure **Print on both sides** is **off**
6. Click Print

---

### Packing Lists

The Packing Lists PDF **cannot** be printed from Adobe because Adobe won't let you set custom page breaks. Print directly from Google Sheets instead:

1. Go to the **PackingLists** tab in the Google Sheet
2. Click **File → Print**
3. Set **Scale** to **Fit to Width**
4. Set **Margins** to **Custom Numbers** and change all four values to **0.15**
5. Click into the **page break preview** — you will see blue dotted lines
6. Drag the blue lines to set page breaks so that 2–3 routes fit per page, with the timestamp/generation message visible on the first page
7. Once all page breaks look correct, click **Confirm Breaks** (upper right)
8. In the left panel, click **Headers and Footers** and check:
   - ✅ Page Numbers
   - ✅ Sheet Name
9. Click **Next**, then **Print**

**Print page 1 two extra times:**
After printing the full list, print just page 1 again twice (enter `1, 1` in the custom page range). The first page contains the item summary and is referenced frequently during packing.

---

## End of Process Checklist

When everything is done, you should have:

- 📄 **Driver Route Sheets** — one copy per driver, plus an extra copy of any Walking Delivery pages
- 📋 **Packing Lists** — full set, plus two extra copies of page 1
- 🏷️ **Label sheets** — one full printed set on Avery 6240 stock

Bring all of them to Distro!

---

## Running Order Reference

```
Menu item 6  →  print Driver Route Sheets
Menu item 7  →  configure + print Packing Lists
Menu item 8  →  print Delivery Labels
```

Start the next function as soon as the previous print job is queued. All three can be in progress simultaneously.
---

## Community & Solidarity

This project is licensed under **MIT** — deliberately. MIT was chosen so that other mutual aids and community orgs can pick this up, adapt it for their own context, and decide for themselves how openly they share what they build. No one is required to expose their volunteers, their data structures, or their internal processes. You set your own boundaries.

That said — if your org adapts this and you want to share what you've built or learned, we'd genuinely love to hear from it. Comparing notes across mutual aid networks makes all of us more effective.

🐙 **[github.com/ej9erfan](https://github.com/ej9erfan)**
💬 **[Open a Discussion](https://github.com/ej9erfan/WAWG-Distribution-Scripts/discussions)** on the repo — other mutual aids can see it too

Solidarity.
