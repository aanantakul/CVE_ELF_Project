## Brief
This repository contains Google Apps Script code that renders building "plan" and "elevation" views inside a Google Sheet by manipulating cell sizes, borders and backgrounds. The AI assistant should preserve the Apps Script runtime patterns and sheet-specific anchors while making changes.

## Big Picture
- **Runtime / platform:** Google Apps Script (SpreadsheetApp). Code is organized under `src/` as plain JS files intended for the Apps Script environment.
- **Major components:**
  - `src/drawing_cell.js`: Main drawing logic — reads inputs from `INPUT` sheet and writes a cell-based rendering into `STRUCTURE` sheet. Key exported functions: `drawPlanAsCells`, `drawElevationAsCells`, helpers and public commands: `updatePlanCellsFromInput`, `updateElevationCellsFromInput`, `updateBothCells`, `clearWholePlotSheet`.
  - `src/main.js`: Menu registration (`onOpen`) that adds the `ELF` menu. Keep function names matching menu item targets.
  - `src/appsscript.json`: (manifest) — check before changing file names/entrypoints.

## Data flow / why structure
- Inputs are read from specific cells on the `INPUT` sheet (see `DC.IN_SPAN_X`, `DC.IN_SPAN_Y`, `DC.IN_PXPM`, `DC.IN_COL_M`, `DC.IN_STOREYS` in `src/drawing_cell.js`). Values are CSV strings parsed by `csvNums_()` and then scaled to a grid using `toCells_()`.
- Rendering writes directly to the `STRUCTURE` sheet (constant `DC.SHEET_PLOTS`). The code uses row/column sizing, borders and background fill instead of images — this keeps the drawing editable in the spreadsheet and avoids external assets.

## Key files to read before editing
- `src/drawing_cell.js` — main source of truth for all drawing logic and the DC config block (anchors, input cell addresses, colors, `CELL_PX`).
- `src/main.js` — menu wiring: changing the menu labels or function names requires keeping consistency here.
- `src/appsscript.json` — confirm script manifest and any entrypoint/function visibility.

## Project-specific conventions and patterns
- Helper functions end with a trailing underscore (e.g. `csvNums_`, `ensureSheet_`, `safeRange_`) and are intended to be internal helpers.
- Public commands (invoked by menu or by users) are short camelCase functions without trailing underscore: `updatePlanCellsFromInput`, `clearWholePlotSheet`.
- Input cells are located in fixed addresses in `INPUT` sheet (D5..D9) as per `DC.IN_*` constants — do not change these addresses unless you update `DC` and document the migration.
- Use `ensureCapacity_` and `safeRange_` before writing ranges — these functions expand the sheet as needed and clamp row/col indices.

## Common tasks & examples
- To update the plan programmatically, call `updatePlanCellsFromInput()` (menu item labelled `Update Plan (cells)`).
- To clear and reset the drawing area call `clearWholePlotSheet()` — it resets content, formats and many rows/cols.
- Example: if you need to compute the number of grid cells from meters, reuse `toCells_(metersArray, pxPerMeter)` to keep consistent scaling.

## Editing rules for AI agents
- Preserve public function names used by the UI/menu. Renaming a public function requires updating `src/main.js` and the manifest if applicable.
- Preserve the `DC` config object keys and their semantics. If changing `CELL_PX` or anchor coordinates, update comments and verify on a sample sheet.
- Avoid heavy refactors that change side effects (sheet resizing, border setting) without manual verification in an actual Google Sheet; these operations are stateful and can be destructive in a live sheet.
- Keep Thai-language comments/strings as-is unless asked to translate; they are meaningful to maintainers and to end-users of the sheet.

## Build / run / debug
- There is no JS bundler or test harness in `package.json`. The code is intended to be deployed/run inside Google Apps Script.
- Typical workflows (not present here, but safe to suggest if you add tooling): use `clasp` to push/pull files between local repo and Apps Script. If you add `clasp`, add a README entry and `.clasp.json`.
- Quick manual test: open the target Google Spreadsheet that contains these scripts, use the Apps Script editor and run `onOpen` or the menu `ELF` → `Update Plan (cells)` to exercise code paths.

## What to avoid / gotchas
- Do not assume a build step. Pull requests that introduce transpilation or packaging must include CI updates and instructions.
- Many helpers call `ensureCapacity_` which expands sheet dimensions; avoid looping edits that cause repeated large sheet expansions.
- The project uses SpreadsheetApp APIs that run under Apps Script quotas. Keep algorithmic complexity reasonable for large inputs.

## Where to add tests / future improvements
- If adding unit tests, isolate pure functions like `csvNums_`, `toCells_`, and `alphaLabels_`. Avoid direct SpreadsheetApp calls in unit tests or mock them explicitly.

## If unsure — quick checklist for the agent
- Did you keep public function names used by the menu? (`updatePlanCellsFromInput`, etc.)
- Did you preserve `DC` anchors and input addresses or update them with clear, discoverable changes?
- Did you run a quick manual verification (Apps Script editor or `clasp`) against a copy of the spreadsheet after changes?

---
Please review these notes and tell me if you want the file to include additional instructions (e.g., a `clasp` workflow, unit-test examples, or CI hooks). 
