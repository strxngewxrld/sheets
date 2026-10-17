# XLSX compatibility fixture

`excel-features.xlsx` was generated independently with openpyxl 3.1.5 for these regression tests; it is not a file produced by the Sheets serializer. It contains synthetic data and the repository's MIT-licensed Sheets icon.

The workbook includes an image, a chart, a cell comment, a hyperlink, list validation, conditional-formatting data bars, global and sheet-local names, print area/title/header/footer settings, outlines, a hidden column, a veryHidden sheet, and font/alignment/border properties not fully represented by Sheets.

The preservation test requires a no-op save to return the exact ZIP bytes. A value and style edit may change only the target worksheet, appended style resources and workbook calculation flags. Every other package entry must retain its exact payload.
