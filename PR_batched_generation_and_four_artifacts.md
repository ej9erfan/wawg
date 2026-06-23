## Batched Route Generation + Four-Artifact Architecture

### Background

This document replaces `PR_makeAllPDFs_removal.md`, which described an earlier attempt at continuation triggers that was removed. That removal note is now obsolete — a working continuation system has been implemented and is the current production approach.

---

### What Changed

**Driver route sheet generation is now batched across multiple executions.**

With 27 routes at ~14 seconds each (dominated by `insertSheet` server latency), the full generation run takes ~380 seconds — just over Apps Script's 6-minute hard limit. The solution uses `PropertiesService` to persist state and `ScriptApp.newTrigger` to chain executions automatically.

- `makeDriverRouteTabs` (menu item 3) processes the first 8 routes and schedules a 30-second trigger
- `continueDriverRouteTabs` picks up where the previous batch left off, processes the next 8, schedules another trigger if needed
- When the final batch completes, it writes the `Driver_Deliveries` index sheet and stamps a `✅ Complete: <timestamp>` cell in column G
- Total wall time: ~8 minutes across 4 executions, no user intervention needed between batches
- `cancelDriverRouteTabs` (Sheet Tools menu) is the emergency stop — clears all PropertiesService state and deletes pending triggers

**Route tab generation and PDF export are now separate menu items.**

`makeDriverRouteTabsThenPDF` (item 7) exports existing `Route_` tabs to PDF without regenerating them. This is intentional — generation consumes the full 6-minute budget and leaves no time for a PDF fetch in the same execution. Run item 3, wait for completion, then run item 7.

**Four deliverables, not three.**

The Navigator Dashboard (`NaviDash`) has been added as a fourth deliverable alongside driver route sheets, packing lists, and delivery labels. It is the coordinator / runners lead view and contains:
- Standard items totals and confirmation status
- Special items pull list (cross-route, with On Hand / Given Out / Not on Hand tracking columns)
- Per-route special items breakdown for staging by driver pile
- Special requests flag list with request text spanning D:I and a checkbox in J

**Stop sort order added.**

Both driver route sheets and packing lists now sort stops/deliveries within each route by special-items complexity: plain deliveries first, then pet food, then diapers/wipes, then hygiene, then special requests last. This supports human data parsing during prep and distribution.

**Address display cleaned up.**

The route name segment (e.g. `Teralta1`) embedded in address strings is stripped from the driver sheet stop display. Drivers see `7968 Highland Ave, San Diego, CA` instead of `7968 Highland Ave, Teralta1, San Diego, CA`.

---

### Why the Previous Continuation Attempt Failed (And Why This One Works)

The previous attempt (`_makeAllPDFs_continue`) was trying to chain three different deliverable functions across triggers — each with its own data loading, sheet setup, and state requirements. The state serialization was complex, the trigger timing was too aggressive, and any error in one phase silently corrupted the next.

The current implementation is scoped to one function (`makeDriverRouteTabs`) and one type of state (route index + accumulated summary rows). The batch function (`_routeBatch_runBatch`) is stateless and re-derives everything it needs from the source sheet on each execution. Only two small pieces of state persist: which route to start from, and the accumulated index rows. Both are simple JSON in PropertiesService. Triggers are cleaned up by function name, not just ID, so stale triggers from aborted runs don't accumulate.

---

### Nothing Removed

All three individual "make X then PDF" functions are present. The Label Generator menu is unchanged. The packing list is still printed from Sheets, not from a PDF, for the same reasons as before (manual page break configuration required).
