# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

Google Apps Script source for spreadsheets used to run WoW raid preloot/attrib tracking. Each top-level folder is a separate **clasp** project bound to a different Google Sheet — there is no shared build system, package manager, or bundler; these are plain `.gs`/`.js` files pushed straight to the Apps Script runtime (V8).

- `preloot/` — scripts bound to the "Préloot" sheet (script ID in `preloot/.clasp.json`). Active.
- `attribs/` — placeholder for scripts bound to the parsi_attribs sheet. Currently empty; clone into it the same way as `preloot/` when needed.

## Commands

All commands run from inside the relevant project folder (e.g. `cd preloot`), since each has its own `.clasp.json` with its own `scriptId`.

```bash
clasp login                 # one-time OAuth login (opens browser)
clasp pull                  # pull latest code from the bound Apps Script project
clasp push                  # push local changes to the bound Apps Script project
clasp open                  # open the project in the Apps Script web editor
clasp clone <scriptId> --rootDir .   # bind a new folder to an existing script project
```

There is no test runner, linter, or build step configured. `attendance.js` contains its own in-sheet smoke tests (`testRaidJsonGenerator`, `testLiveWarcraftLogsExample`) that must be run from the Apps Script editor or via the spreadsheet's custom menu (**Warcraft Logs → Test JSON generator**), not from a CLI.

## Architecture (`preloot/`)

Three scripts share one spreadsheet and communicate purely through sheet cells — there is no in-memory or module-level state passed between them. Understanding the data flow requires reading all three files together:

- **`ratio.js`** — loot-side processing. Reads raw RCLC/Gargul JSON loot export payloads from the `raid_logs` sheet (columns C/D), resolves each looter against the roster on `Ratio Présence/Loot (Préloot)` (via `normalizePlayerName`/`getRosterPlayerKey`, which strip realm suffixes and case-fold names — reuse these for any new player-matching code instead of comparing names directly), applies item coefficients (prio items = 0, tier-specific decay for Kara/Gruul/Magh after a hardcoded date), and writes results to two places: the `loot_details` sheet (per-player item breakdown) and column J of the ratio sheet (weighted loot totals). `writeJsonToCell` then base64-encodes two JSON summaries (RCLC ratio, Gargul ratio) into `Admin!C2` and `Admin!D2` for external import.
- **`attendance.js`** — attendance-side processing. Pulls raid logs from the Warcraft Logs API (OAuth client-credentials flow, token cached via `CacheService`), computes per-player attendance percentages, and separately parses the `raid_logs.attendees` JSON column to populate `presence_details` (one row per raid, one column per roster player) and roll up totals into column F of the ratio sheet. Reroll characters are mapped to mains via the `Rerolls` sheet (`getRerollPresenceMaps_`). `process_presences()` is the entry point that chains attendee-fetching → presence table build → ratio sheet update.
- **`onedit.js`** — a simple `onEdit` trigger: any edit to columns E/F (presence/bench) on the ratio sheet stamps a timestamp in `Admin!B1` and clears the cached import cells (`resetAdminImportRatioCell`, defined in `ratio.js`), invalidating the base64 exports until `writeJsonToCell` is rerun.

### Cross-file conventions

- **Roster identity**: always resolve player names through `getRosterPlayerKey()` (defined in `ratio.js`, used by both `ratio.js` and `attendance.js`) rather than raw string comparison — it normalizes case and strips the `-Thunderstrike` realm suffix, and is the only thing that keeps loot and attendance data joined to the same roster row.
- **Sheet names/cells are hardcoded** per-function (e.g. `"Ratio Présence/Loot (Préloot)"`, `Admin!C2`) rather than centralized constants — when renaming a sheet, grep across both `ratio.js` and `attendance.js`.
- **Config objects**: `attendance.js` centralizes its tunables in `WCL_CONFIG` (API URLs, column numbers, attendance threshold) at the top of the file — follow that pattern for new constants instead of inlining magic numbers.
- Script Properties (`WCL_CLIENT_ID`, `WCL_CLIENT_SECRET`) must be set in the Apps Script project (Project Settings → Script Properties) for `attendance.js`'s Warcraft Logs calls to work; they are never stored in this repo.
