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
| `Module2.bas` | OSDD Search. Reads its inputs from Sheet1 beside each entity (columns V:AB: Negative News name, additional address Yes/No and up to 3 addresses, telephone, email; other parties in V27). Run `SetupOsddInputs` once to lay those cells out; it also hides Counterparty 6 (the tool takes 5). `StyleOsddInputs` re-colours them after a theme change. |
| `Module3.bas` | Generate Narrative (`ExportToWord`). |
| `Sheet1_CodeModule_UPDATED.txt` | Sheet1's sheet-module code, to paste into the sheet's code window. |

The theme modules (`modTronLegacy`, `modDesert`, `modThemeManager`,
`modThemeSearchMatrix`, `modThemeDebug`) may sit under other names
(Module11-17) in a given workbook.

## Notes

- The `.frm` user forms are here without their `.frx` resource files, so
  they are for reading and comparing; re-importing a form needs its `.frx`.
- Files are stored byte for byte as exported (see `.gitattributes`). Most use Windows
  (CRLF) line endings; the theme modules and `Module19.bas` use LF.
