## Remove `makeAllPDFs` — Serial Manual Execution Model

### Background

`makeAllPDFs` was originally introduced to let the operator kick off all three Distribution Day deliverables in one click and walk away while they generated. At the time, the entire process ran in roughly five minutes — within Apps Script's execution limit — and the result was three PDFs sitting in Google Drive ready to print.

Two things have changed since then:

1. **Each individual deliverable now takes approximately five minutes on its own.** The addition of driver route tabs, packing list formatting, and the full label generation pipeline means the three phases together far exceed the six-minute hard execution limit. Attempts to work around this with continuation triggers and `PropertiesService` state were implemented but proved too slow and too fragile to be reliable.

2. **The "walk away" workflow no longer matches how printing actually works.** The packing lists require very specific print settings that must be configured manually inside the Google Sheets print menu — this step cannot be automated and cannot be skipped. In practice, the operator has to be present for each print job anyway. The Driver Route Sheets and Delivery Labels PDFs also need to be downloaded and opened in Adobe Reader before printing with the correct scale settings.

### What This Changes

The serial manual model is now the intended workflow:

- Run **Make/Update Driver Route Tabs** → immediately start the print job → while it prints, run **Make Packing Lists** → configure and print → while it prints, run **Make Delivery Labels** → print.
- Each function can be triggered from the menu as soon as the previous print job is queued. The operator stays focused and moves through all three deliverables in one sitting without waiting for automation to finish before starting the next step.
- This approach also reduces the cognitive load of monitoring a background process and interpreting toast notifications or error states from a multi-phase trigger chain.

### Code Removed

- `makeAllPDFs()` — the entry point function
- `_makeAllPDFs_continue()` — the time-based trigger continuation handler
- `_clearContinuationTrigger()` — the trigger cleanup helper
- Menu item **9. Make All PDFs** from the Distribution Day Tasks menu
- All `PropertiesService` phase state management

### Nothing Else Changes

The three individual "make X then PDF" functions (`makeDriverRouteTabsThenPDF`, `makePackingListThenPDF`, `makeLabelsPDF`) are unchanged. The Label Generator menu is unchanged. Menu items 1–8 are unchanged.



