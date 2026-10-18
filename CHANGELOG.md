# Changelog

## Unreleased

- XLSX: sorting can now be saved in files with Excel content Sheets does not model, when the result is exact (cells take the original format records along; notes, hyperlinks, native tables, validation, merged cells and formulas referring to other rows stay blocked).
- XLSX: sheets added at the end are written as new parts of the original package; files without `styles.xml` get a minimal style table so formatting can be saved. Renaming, deleting and moving sheets stay blocked.
- XLSX protection messages follow the UI language (German/English).
- CSV: opening and saving keeps the file content: line endings, a missing final line break and a UTF-8 BOM are kept; leading zeros, dates, percentages, foreign decimal separators and long numbers stay text.
- CSV: Windows-1252 files (CSV exports of Excel on German/Western European Windows) are read and saved in their encoding instead of turning umlauts into replacement characters; text Windows-1252 cannot hold is saved as UTF-8 with BOM, with a one-time notice.
- CSV: agent tools for unopened files and the MCP server keep the delimiter, decimal comma, line endings, BOM and encoding of a CSV file. Before, any agent edit rewrote a semicolon CSV with commas and decimal points.
- CSV: a save always writes the file's data sheet, even when another sheet is active.
- XLSX: formula results stored in the file follow edits. Unchanged formulas whose result changes (also on other sheets) get the new cached value, new formulas get their computed value; only formulas with functions the engine knows, never an invented result. Programs that read cached values without calculating no longer see stale results.
- XLSX: an edit that only changes formula results on a protected sheet no longer blocks saving.
- Recursive LAMBDA formulas that never end stop after 1000 nested calls with #CALC! within milliseconds instead of freezing the editor for about half a second.
- CSV: a one-time, non-blocking notice when a save drops content CSV cannot hold (other sheets, charts, formulas, formatting, ...). "Save As" writes the format of the target file. Cell texts starting with `=`, `+`, `-` or `@` are written unchanged (see the CSV injection note in docs/xlsx-compatibility.md).
- Editor: a loading notice while large files are parsed; files that cannot be opened show the error in the editor and are never saved (no more "getFileData timed out").
- Conditional formatting dialog lists rules as readable text ("C2:C6 · Cell value > 6"); dialogs never grow taller than the editor and keep their buttons visible; the number format box names custom formats (e.g. currency).
- Pivot reports leave row/column combinations without records empty (like Excel) and widen narrow output columns.
- Fix: crashes on worksheets or style parts with empty or whitespace-only elements; formula evaluation in data validation no longer throws on runaway recursion.
- Much faster "current region" detection (sorting/charts from a single cell) for very large sheets.

## 1.0.2

- Fix Windows agent file paths: recognize UNC shares and normalize relative backslash paths, including remote workspaces.
- Preserve AltGr character input on Windows/Linux and IME composition instead of triggering editor shortcuts.
- Launch the VS Code host test through Code.exe on Windows, including installation paths with spaces.
- Run typechecks, tests and VSIX packaging on Windows, macOS and Linux in CI; document desktop OS support.

## 1.0.1

- Fix: installing an update from the update notice failed with "No Servers".

## 1.0.0

- Formulas: `LET`, `LAMBDA` (also via defined names), `MAP`, `BYROW`, `BYCOL`, `REDUCE`, `SCAN`, `MAKEARRAY`, `ISOMITTED`, `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `CHOOSEROWS`, `CHOOSECOLS`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `EXPAND`, `SORTBY`, `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, `IRR`, `XIRR`, `RATE`, `AGGREGATE` and `HYPERLINK`.
- Structured table references (`Table1[Col]`, `[@Col]`, `Table1[#All]`, ...), stored in Excel's qualified form.
- Find and replace, go to cell, zoom (25-400 %), hide/unhide rows and columns, paste values and paste special (formulas, formats, skip blanks, transpose).
- `SUBTOTAL` skips rows hidden by filters (and 101-111 also manually hidden rows); `AGGREGATE` can ignore errors and hidden rows.
- Data validation: dropdown lists, number/date/time/text length limits and custom formulas with input and error messages; read, written and patched in preserved XLSX files.
- Conditional formatting: data bars and all Excel icon sets.
- Notes and hyperlinks are shown; links and `HYPERLINK` formulas open with Ctrl/Cmd+click.
- `.xlsm` workbooks open in Sheets; macros are kept but never run.
- "Export Sheet as CSV…" command and an in-editor warning when changes could not be saved as XLSX without loss.
- New agent tools: `replace`, `set_hidden`, `set_data_validation`, `list_data_validations`; `copy_range` can transpose and skip blanks.
- Much faster reading and saving of large workbooks (e.g. reading 1M cells ~10 s instead of ~44 s, saving a one-cell edit ~15 s instead of ~53 s).
- Update notice: on startup Sheets checks GitHub Releases for a newer version, links to it and can download and install the release VSIX directly (setting `sheets.checkForUpdates`).
- Formula editing: Shift+arrow keys extend the pointed reference to a range.
- Fix: hidden checkbox rows in dialogs stayed visible.

## 0.2.3

- Explicitly associate XLSX files with Sheets as the default editor.
- Verify ordinary file opening through VS Code's editor resolver in the real host test.

## 0.2.2

- Preserve original XLSX packages and patch supported cell/style/layout changes instead of rebuilding unknown contents.
- Block saves that cannot safely retain unimplemented Excel features; preserve original formula caches and metadata.
- Keep original packages in JSON and hot-exit backups, including edits blocked from XLSX save.
- Reject detected external file changes before overwriting files.
- Preserve literal OOXML escape strings and control characters in new cell text.
- Add compatibility fixtures and regression checks for unknown parts, styles, formulas, backups and failed saves.

## 0.2.1

- Simplify the extension icon to a flat two-color table mark.
- Remove obsolete build files and add a scoped clean command and EditorConfig.
- Separate the Copilot adapter and handle cancellation and tool errors explicitly.
- Add a real VS Code LM tool registry/invocation test and Copilot setup documentation.

## 0.2.0

- Make the repository root a directly runnable and packageable VS Code extension.
- Add the Sheets icon, English/German metadata, an extension README and MIT license.
- Add F5 debugging, watch tasks, reproducible VSIX packaging and CI validation.
- Open XLSX workbooks and CSV/TSV files using the native Sheets custom editor.
- Support table highlighting, sortable/filterable table headers and cell formatting.
- Preserve tables and value filters through the VS Code open/edit/save lifecycle.
- Queue tool and save requests until the webview confirms workbook initialization.
- Provide formulas, dynamic arrays, conditional formatting, charts and pivot reports.
- Expose spreadsheet operations as language model tools and an extension API.
- Restrict AI writes to unopened files to trusted workspaces.

## 0.1.0

- Initial spreadsheet engine, canvas UI, web demo and VS Code custom editor prototype.
