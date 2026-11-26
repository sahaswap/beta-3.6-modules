# Beta 3.6 - VBA modules

The VBA modules of the Beta 3.6 AML case tool, exported as text so changes
can be tracked. Code only: no workbooks, exports or transaction data belong
in this repository (see `.gitignore`).

## Importing into the workbook

In the VBA editor (Alt+F11): remove the old module first, then
**File > Import File**, then **Debug > Compile VBAProject**, then save the
workbook. A module imported over one with the same macro names stops with
"Ambiguous name detected".

## Main files

| File | What it is |
| --- | --- |
| `Module9.bas` | Export Trx File - the normal export behind the button (`Consolidated_AML_Workflow`). The ECM import sets `gEcmAutoMode` ("EN") and `gEcmAutoQuiet` for one run: no export-type picker and no "complete" message; the button itself is unchanged. |
| `modLargeExport.bas` | **Separate large-transaction export tool - not part of the normal Beta 3.6.1 tool.** Version 2026-11-24b (`Consolidated_AML_Workflow_Large`, `Consolidated_AML_Workflow_Large_Resume`, `Lookback_Pivots_From_Files`, `Lookback_Fix_Start_From_Files`), run from Developer > Macros in the large-export workbook. Legacy reads Excel and CSV files, keeps alerted transactions only, saves CP Selection / DeDupe / Non Alerted Transaction as separate files, a pivot-only main file, a Lookback Transactions file with its pivots, and deletes leftover temporary folders. The picker also offers Lookback Only, which makes just those Lookback files. |
| `clsLookbackButton.cls` | **Large-export tool only**, imported with `modLargeExport.bas`. Carries the click of the Lookback Only button that module adds to `frmExportMode` while it runs. |
| `modSetNonEscTemplate.bas` | One-time macro `SetNonEscalationTemplate`: writes the combined Individual / Entity Non-Escalation narrative to Template_DB and adds the `[Country Risk]` tag to Sheet7. |
| `NonEscalation_Template_Combined.txt` | Readable copy of that narrative template. |
| `Module2.bas` | OSDD Search. Reads its inputs from the OSDD Search Inputs panel on Sheet1 (V3:AB10, under a copy of the form's banner shape: a column per entity - customer, CP1-5 - with Negative News name, 3 additional addresses, telephone / email). `StyleOsddInputs` re-colours the panel after a theme change; `FitSheet1ToWindow` zooms Sheet1 to fit the window. `OsddFillNegNews` fills each Negative News cell with the name the search uses as soon as a name is typed or pasted in J13 / J18-J22; `OsddPaste` is Ctrl+V on Sheet1: plain text into the selected input cell, with the copy kept for more cells. A case copied by the ECM floater (`ecm-to-tool.js` in the screen-recorder repo) fills Sheet1 in one Ctrl+V: ECM Case, Alert ID, Customer Name and Address, country (ISO Code 2), Program Name matched to Program_Library, the customer's phone, email and other addresses, and counterparties 1-5 with their addresses and other addresses. **Import to tool** in the floater saves the case's transaction file and then the case itself (`<ECM>_AMLTOOL.txt`) to Downloads; and on the clipboard (marked `auto=1`); while the tool is open it looks at Downloads and the clipboard every 5 seconds (`EcmWatchStart`, started when the workbook opens - Workbook_Open's `FitSheet1ToWindow`, `Auto_Open` - and by Ctrl+V, stopped by `Auto_Close`; a case sent more than 30 minutes ago is left alone, the same case is taken in once) and takes the case in by itself, bringing the tool to the front (Alt+F8 > `EcmWatchStatus` shows whether it is running, the folder it watches and any case files there; a case taken in is left on the clipboard without its mark, and the same case already on Sheet1 is asked about) - no Ctrl+V - filling Sheet1 and going on (a case with a transaction file pasted with Ctrl+V does the same): `EcmFullCase` makes the case folder (`<Desktop>\<ECM>\Transaction Files` and `\Pivot`), takes the file from Downloads (waits while it downloads, asks for it if it isn't there), unzips a .zip into Transaction Files (Windows zip, else `tar`; a folder that already has an Excel file is used as it is), runs Export Trx File as EN Network (no picker, no "complete" message - Module9's `gEcmAutoMode` / `gEcmAutoQuiet`), then fills from ConsolidatedData: counterparties J18:J22 if blank, by the counterparty sampling rule (top counterparty of each alert rule by amount, at most 5, topped up to 3 from the alerted rows and then - when they have no more counterparties - from the EN export's Non Alerted Transaction sheet by dollar value, all sorted by dollar value; fewer than 3 names from the floater are topped up the same way and the rows re-sorted, each name's cells moving with it; alerted rows only; spellings of one name - AMAZON, AMAZ, AMZN, AMZON, AMAZON MKTPL - counted as one counterparty, amounts added, the biggest spelling written and the others listed; reference codes taken off the names - `AMAZON MKTPL*2K4L39RT2`, `Amazon.QXZKWPRTB`, `ABA123456789-<hash>-AMAZON` (routing number and hash joined with hyphens) -> `Amazon` - and `Amazon.com` / `Amazon.QXZKWPRTB` counted as one merchant), each counterparty's most frequent address in T and other addresses in X:Z (DR rows: Beneficiary, CR rows: Originator), and the customer's other addresses that differ from J14 in X14:Z14. Counterparties already on Sheet1 (picked by the floater's Read case, or typed) are kept, and the summary lists any the rule would pick from the exported transactions that Sheet1 doesn't have. Only blank cells are filled; nothing but the floater's own case file is deleted from Downloads. An automatic import that stops with an error says so, and its case file is kept as `.failed`. |
| `modInstallUpdate.bas` | One-time installer `InstallUpdate`: from a folder holding Module2, modTronLegacy, modThemeManager and modDesert (CRLF copies), backs up and replaces those modules, edits the single lines that change in the others (Reset, Sheet1's code, narrative / Rename / PDF Merge on 5 counterparties, ThisWorkbook), adds Sheet1's Negative News hook and ThisWorkbook's Ctrl+V key, then lays out Sheet1 (panel, theme picker hidden, row 23 emptied, the open case's OSDD inputs kept) and Sheet7 / Search Matrix. Safe to run again. Needs Trust access. Delete it afterwards. |
| `modLayoutDump.bas` | Diagnostic macro `DumpSheet1Layout`: writes Sheet1's layout - every shape's position, size and settings, column widths, row heights, merged areas; no cell values - to "Sheet1 layout.txt" on the Desktop and opens it. |
| `Module3.bas` | Generate Narrative (`ExportToWord`). |
| `Sheet1_CodeModule_UPDATED.txt` | Sheet1's sheet-module code, to paste into the sheet's code window. |

The theme modules (`modTronLegacy`, `modDesert`, `modThemeManager`,
`modThemeSearchMatrix`, `modThemeDebug`) may sit under other names
(Module11-17) in a given workbook.

### Themes in Beta 3.6.1

In 3.6.1 the Dark Blue theme is `Module13` and the theme manager is
`Module15`. The files here are those modules, newer:

- `modTronLegacy` is Module13 with every cell fill set through `PaintCells`
  (a leftover cross-hatch pattern, such as the greyed Escalation Scenario
  row, no longer washes the dark fill out), no message box on each theme
  switch, and the picker at V13.
- `modThemeManager` is Module15 with the Desert theme registered, the
  picker at V13 and the call to `StyleOsddInputs` after a theme change.
- `modDesert` is the Desert theme.

To use them: remove Module13 and Module15, import the three files, run
**Debug > Compile**, then `BuildThemePicker` once so the dropdown lists
Desert. Or let `modInstallUpdate` do all of it.

## Notes

- The `.frm` user forms are here without their `.frx` resource files, so
  they are for reading and comparing; re-importing a form needs its `.frx`.
- Files are stored byte for byte (see `.gitattributes`), all with Windows (CRLF)
  line endings - the VBA editor needs them to import a module.
