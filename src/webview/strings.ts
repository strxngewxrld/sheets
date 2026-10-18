/** Texts of the editor shell (loading, errors, save notices) shared by the host and the webview. */
const en = {
  loading: 'Loading {0}…',
  loadingDetail: 'Large files can take a few seconds.',
  openFailed: '{0} could not be opened.',
  openFailedDetail: 'The file was not changed. Saving is disabled for this editor until the file can be opened.',
  saveRefusedNotOpened: 'Sheets: {0} could not be opened and is therefore not saved. The file on disk was not changed.',
  csvLoss: 'Sheets: {0} is saved as {1}. This format stores only the values of the sheet "{2}"; {3} are not saved. Use "Save As…" with .xlsx to keep them.',
  csvUtf8Fallback: 'Sheets: {0} was saved as UTF-8 because its previous encoding Windows-1252 cannot represent {1}. Programs that expect Windows-1252 may show these characters incorrectly.',
  csvLossParts: { sheets: 'other sheets', charts: 'charts', formulas: 'formulas (only their results)', formats: 'formatting', tables: 'tables', rules: 'conditional formatting and data validation', merges: 'merged cells', notes: 'notes and hyperlinks', pivots: 'pivot definitions', names: 'defined names' } as Record<string, string>,
};

const de: typeof en = {
  loading: '{0} wird geladen…',
  loadingDetail: 'Große Dateien können einige Sekunden dauern.',
  openFailed: '{0} konnte nicht geöffnet werden.',
  openFailedDetail: 'Die Datei wurde nicht verändert. Speichern ist in diesem Editor deaktiviert, bis sich die Datei öffnen lässt.',
  saveRefusedNotOpened: 'Sheets: {0} konnte nicht geöffnet werden und wird deshalb nicht gespeichert. Die Datei auf dem Datenträger wurde nicht verändert.',
  csvLoss: 'Sheets: {0} wird als {1} gespeichert. Dieses Format speichert nur die Werte des Blatts „{2}“; {3} werden nicht gespeichert. Mit „Speichern unter…“ als .xlsx bleiben sie erhalten.',
  csvUtf8Fallback: 'Sheets: {0} wurde als UTF-8 gespeichert, weil die bisherige Kodierung Windows-1252 {1} nicht darstellen kann. Programme, die Windows-1252 erwarten, zeigen diese Zeichen eventuell falsch an.',
  csvLossParts: { sheets: 'weitere Blätter', charts: 'Diagramme', formulas: 'Formeln (nur ihre Ergebnisse)', formats: 'Formatierungen', tables: 'Tabellen', rules: 'bedingte Formatierung und Datenüberprüfung', merges: 'verbundene Zellen', notes: 'Notizen und Hyperlinks', pivots: 'Pivot-Definitionen', names: 'definierte Namen' },
};

export type ShellStrings = typeof en;

export function shellStrings(language: string | undefined): ShellStrings {
  return language?.toLowerCase().startsWith('de') ? de : en;
}

export function fill(template: string, ...values: string[]): string {
  return template.replace(/\{(\d)\}/g, (_, i: string) => values[Number(i)] ?? '');
}
