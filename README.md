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
| `Module9.bas` | Export Trx File - the normal export behind the button (`Consolidated_AML_Workflow`). |
| `modLargeExport.bas` | **Separate large-transaction export tool - not part of the normal Beta 3.6.1 tool.** Version 2026-10-29c (`Consolidated_AML_Workflow_Large`, `Consolidated_AML_Workflow_Large_Resume`), run from Developer > Macros in the large-export workbook. Legacy reads Excel and CSV files, keeps alerted transactions only, saves CP Selection / DeDupe / Non Alerted Transaction as separate files, a pivot-only main file, and deletes leftover temporary folders. |
| `modSetNonEscTemplate.bas` | One-time macro `SetNonEscalationTemplate`: writes the combined Individual / Entity Non-Escalation narrative to Template_DB and adds the `[Country Risk]` tag to Sheet7. |
| `NonEscalation_Template_Combined.txt` | Readable copy of that narrative template. |
| `Module2.bas` | OSDD Search. Reads its inputs from the OSDD Search Inputs panel on Sheet1 (V3:AB10: a column per entity - customer, CP1-5 - with Negative News name, 3 additional addresses, telephone, email). `StyleOsddInputs` re-colours the panel after a theme change; `FitSheet1ToWindow` zooms Sheet1 to fit the window. |
| `modInstallUpdate.bas` | One-time installer `InstallUpdate`: from a folder holding Module2, modTronLegacy, modThemeManager and modDesert (CRLF copies), backs up and replaces those modules, edits the single lines that change in the others (Reset, Sheet1's code, narrative / Rename / PDF Merge on 5 counterparties, ThisWorkbook), then lays out Sheet1 (panel, theme picker in V13, row 23 emptied) and Sheet7 / Search Matrix. Needs Trust access. Delete it afterwards. |
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
