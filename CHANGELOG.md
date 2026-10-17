# Changelog

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
