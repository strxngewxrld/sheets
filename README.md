# Sheets – Excel & CSV Editor

[Deutsch](README.de.md) · [User manual](docs/manual.en.md) · [Handbuch](docs/manual.de.md)

Open and edit Excel and CSV files directly in Visual Studio Code. Sheets provides formulas, formatting, table filters and charts without requiring Microsoft Excel. Spreadsheet processing runs locally.

![Sheets spreadsheet editor](docs/images/editor.png)

## Installation

Download a VSIX from [GitHub Releases](https://github.com/strxngewxrld/sheets/releases), then run **Extensions: Install from VSIX…** in VS Code. To build the package yourself:

```bash
npm ci
npm run package
code --install-extension artifacts/sheets-vscode-1.2.3.vsix
```

Requires VS Code Desktop 1.95 or later on Windows, macOS or Linux. Development and packaging require Node.js 22 or later. The VSIX bundles its runtime dependencies; vscode.dev is not supported.

Sheets checks GitHub Releases for updates at startup and offers installation when a VSIX is available. Disable this with `sheets.checkForUpdates`.

## Getting started

1. Run **Sheets: Getting Started** from the Command Palette to open the interactive walkthrough.
2. Open an `.xlsx` or `.xlsm` file, or run **Sheets: New Workbook** to create an XLSX file.
3. For `.csv` or `.tsv`, right-click the file in Explorer and choose **Open with Sheets**. Alternatively, use **Reopen Editor With… → Sheets (CSV)**.
4. Enter values and formulas such as `=SUM(A1:A2)`. Use English function names and commas in both UI languages.
5. Save with **Ctrl+S** / **Cmd+S**. Run **Sheets: Open User Manual** for detailed instructions.

The interface and walkthrough follow the VS Code language: English or German. The [user manual](docs/manual.en.md) includes a worked example, shortcuts and troubleshooting.

## Features

- Edit, copy, paste, fill and manage cells, rows, columns and sheets; hide rows and columns and freeze panes.
- Paste values, formulas or formats; skip blanks and transpose.
- Find and replace across a sheet or workbook, including formulas and regular expressions. Jump to cells, ranges or names; zoom from 25% to 400%.
- Apply number formats, fonts, colors, borders, alignment and table styles.
- Sort complete records and filter table columns by values with search, multiple selections and blanks. Use structured references such as `Table1[Revenue]`.
- Calculate formulas including `XLOOKUP`, `FILTER`, `LET`, `LAMBDA`, `MAP`, `REDUCE`, `VSTACK`, `TEXTSPLIT` and `XIRR`.
- Apply data validation, conditional formatting, color scales, data bars and icon sets.
- Create column, bar, line, area, pie, doughnut and scatter charts, and create or refresh pivot reports.
- Display notes and open hyperlinks with Ctrl/Cmd+click.
- Save Excel workbooks and export the active sheet as CSV/TSV. CSV saves retain the source delimiter and supported encoding.

## AI tools

In Copilot Chat, select **Agent**, then enable the Sheets tools under **Configure Tools**. Use a current VS Code version for agent support (extension tools require VS Code 1.99 or later). Sheets registers 45 `sheets_*` tools for compatible clients. See the [manual](docs/manual.en.md#ai-assistance) and [Copilot guide](docs/copilot.md).

Changes to open files appear in the editor and can be undone. `sheets.agentTools.headlessWrites` also permits direct saves to unopened files in trusted workspaces. Manual editing requires no AI account. Sheets itself does not send spreadsheet data to external AI services; your AI client may send tool results to its provider.

Other extensions can call the command interface:

```ts
await vscode.commands.executeCommand('sheets.executeTool', {
  name: 'read_range',
  file: '/path/to/workbook.xlsx',
  input: { range: 'A1:D10' },
});
```

## Formats and limitations

Supported: `.xlsx`, `.xlsm`, `.csv`, `.tsv`. Macros are preserved in XLSM but never executed. Legacy `.xls` is not supported. CSV/TSV store only cell values from one sheet, without formatting or charts.

Not every Excel feature is displayed or editable. Images and native pivot caches are preserved rather than rendered as editable objects; notes are read-only. Some built-in table styles are approximated in the editor while the original style stays in the file.

Saving uses the original XLSX package and retains unimplemented parts. An unchanged workbook is returned byte for byte. Edits that cannot be preserved safely are blocked before overwriting; the editor shows the reason above the grid. Recovery backups retain unsaved changes. See [XLSX compatibility and save protection (German)](docs/xlsx-compatibility.md).

## Development

```bash
npm ci
npm run build        # Typecheck and bundle the extension and webview
npm test             # Core, UI, agent and extension tests
npm run package      # Build, package and validate the VSIX
npm run check:package
```

Open the repository in VS Code and press **F5** to launch an Extension Development Host with the watch build. Reload the host after extension code changes; reopen the spreadsheet editor after webview changes. `npm run build:all` also builds the MCP server. See [DEVELOPMENT.md](DEVELOPMENT.md).

## License

[MIT](LICENSE). Bundled dependency licenses: [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
