import type { SheetData } from '../sheet.js';
import { child, contents, editXml, nodeAttribute, prefix, protect, raw, setAttribute, setChild, xmlRoot, type XmlEdit } from './xlsx-preserve-xml.js';

/** Change size attributes only, retaining hidden/outline/style and all unrelated XML. */
export function patchLayout(xml: string, before: SheetData, after: SheetData): string {
  const changed = (a: Record<number, number> = {}, b: Record<number, number> = {}) => [...new Set([...Object.keys(a), ...Object.keys(b)])].map(Number).filter((i) => a[i] !== b[i]).sort((a, b) => a - b);
  const columns = changed(before.colWidths, after.colWidths);
  if (columns.length) {
    const root = xmlRoot(xml), p = prefix(root), cols = child(root, 'cols');
    const existing = cols?.children.filter((n) => n.name === 'col').map((node) => ({ node, min: Number(nodeAttribute(xml, node, 'min')) - 1, max: Number(nodeAttribute(xml, node, 'max')) - 1 })) ?? [];
    let last = -1;
    for (const c of existing) {
      if (!Number.isInteger(c.min) || !Number.isInteger(c.max) || c.min <= last || c.max < c.min) protect('Überlappende oder unsortierte Spaltendefinitionen.', 'Overlapping or unsorted column definitions.');
      last = c.max;
    }
    const make = (original: string, index: number, end = index, apply = true) => {
      const node = xmlRoot(original); let open = original.slice(0, node.openEnd);
      open = setAttribute(setAttribute(open, 'min', index + 1), 'max', end + 1);
      if (apply) {
        const value = after.colWidths?.[index];
        if (value !== undefined && (!Number.isFinite(value) || value < 0)) protect('Ungültige Spaltenbreite.', 'Invalid column width.');
        if (value === 0) open = setAttribute(open, 'hidden', 1);
        else {
          open = setAttribute(open, 'width', value === undefined ? undefined : Math.max(0, Math.round((value - 5) / 7 * 100) / 100));
          open = setAttribute(open, 'customWidth', value === undefined ? undefined : 1);
          if (before.colWidths?.[index] === 0) open = setAttribute(open, 'hidden', undefined);
        }
      }
      return editXml(original, [{ start: 0, end: node.openEnd, value: open }]);
    };
    const edits: XmlEdit[] = [], inserted = new Map<number, string>();
    for (const c of existing) {
      const within = columns.filter((i) => i >= c.min && i <= c.max);
      if (!within.length) continue;
      const boundaries = [...new Set([c.min, c.max + 1, ...within.flatMap((i) => [i, i + 1])])].sort((a, b) => a - b);
      if (boundaries.length > 2 && c.node.children.length) protect('Nicht teilbare erweiterte Spaltendefinition.', 'Extended column definition cannot be split.');
      const original = raw(xml, c.node);
      const value = boundaries.slice(0, -1).map((i, k) => make(original, i, boundaries[k + 1]! - 1, within.includes(i))).join('');
      edits.push({ start: c.node.start, end: c.node.end, value });
    }
    for (const index of columns.filter((i) => !existing.some((c) => i >= c.min && i <= c.max))) {
      if (after.colWidths?.[index] === undefined) continue;
      const position = existing.find((c) => c.min > index)?.node.start ?? existing.at(-1)?.node.end ?? cols?.openEnd ?? 0;
      inserted.set(position, (inserted.get(position) ?? '') + make(`<${p}col/>`, index));
    }
    if (!cols) xml = setChild(xml, 'cols', `<${p}cols>${[...inserted.values()].join('')}</${p}cols>`, ['sheetData']);
    else if (cols.closeStart === cols.openEnd && /\/\s*>$/.test(xml.slice(cols.start, cols.openEnd))) xml = editXml(xml, [{ start: cols.start, end: cols.end, value: contents(xml, cols, [...inserted.values()].join('')) }]);
    else xml = editXml(xml, [...edits, ...[...inserted].map(([start, value]) => ({ start, end: start, value }))]);
  }
  const rows = changed(before.rowHeights, after.rowHeights);
  if (rows.length) {
    const root = xmlRoot(xml), p = prefix(root), data = child(root, 'sheetData');
    if (!data) protect('sheetData fehlt.', 'sheetData is missing.');
    let rowIndex = -1;
    const existing = data.children.filter((n) => n.name === 'row').map((node) => {
      rowIndex = Number(nodeAttribute(xml, node, 'r') ?? rowIndex + 2) - 1; return { index: rowIndex, node };
    });
    const edits: XmlEdit[] = [], inserted = new Map<number, string>();
    for (const index of rows) {
      const old = existing.find((r) => r.index === index);
      if (!old && existing.some((r) => nodeAttribute(xml, r.node, 'r') === undefined)) protect('Neue Zeilen bei impliziten Original-Zeilenadressen.', 'New rows in a sheet with implicit row addresses.');
      if (!old && after.rowHeights?.[index] === undefined) continue;
      const original = old ? raw(xml, old.node) : `<${p}row r="${index + 1}"/>`;
      const root = xmlRoot(original); let open = original.slice(0, root.openEnd);
      const value = after.rowHeights?.[index];
      if (value !== undefined && (!Number.isFinite(value) || value < 0)) protect('Ungültige Zeilenhöhe.', 'Invalid row height.');
      if (value === 0) open = setAttribute(open, 'hidden', 1);
      else {
        open = setAttribute(open, 'ht', value === undefined ? undefined : Math.round(value * 75) / 100);
        open = setAttribute(open, 'customHeight', value === undefined ? undefined : 1);
        if (before.rowHeights?.[index] === 0) open = setAttribute(open, 'hidden', undefined);
      }
      const changed = editXml(original, [{ start: 0, end: root.openEnd, value: open }]);
      if (old) edits.push({ start: old.node.start, end: old.node.end, value: changed });
      else {
        const position = existing.find((r) => r.index > index)?.node.start ?? existing.at(-1)?.node.end ?? data.openEnd;
        inserted.set(position, (inserted.get(position) ?? '') + changed);
      }
    }
    if (data.closeStart === data.openEnd && /\/\s*>$/.test(xml.slice(data.start, data.openEnd))) xml = editXml(xml, [{ start: data.start, end: data.end, value: contents(xml, data, [...inserted.values()].join('')) }]);
    else xml = editXml(xml, [...edits, ...[...inserted].map(([start, value]) => ({ start, end: start, value }))]);
  }
  return xml;
}
