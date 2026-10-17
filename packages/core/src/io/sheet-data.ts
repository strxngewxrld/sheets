import { parseXml, type XNode } from './xml.js';

/**
 * Fast reader for the <sheetData> of a worksheet, which holds nearly all of
 * its bytes. It produces the same objects as the generic XML parser
 * (attributes as "@_name", text as strings or "#text") for rows and for cells
 * that only contain <v> and <f>; other cells are handed to the generic parser.
 * Returns null when the markup is unusual, so callers parse the whole sheet.
 */
export function scanSheetData(xml: string): { start: number; end: number; rows: XNode[] } | null {
  const open = /<(?:([\w.-]+):)?sheetData(?=[\s/>])[^>]*?(\/?)>/.exec(xml);
  if (!open) return null;
  const prefix = open[1] ? `${open[1]}:` : '';
  const start = open.index;
  if (open[2] === '/') return { start, end: start + open[0].length, rows: [] };
  const close = xml.indexOf(`</${prefix}sheetData>`, start);
  if (close < 0) return null;
  const body = xml.slice(start + open[0].length, close);
  if (body.includes('<!--') || body.includes('<![CDATA[') || body.includes('<?')) return null;
  const rows: XNode[] = [];
  const row = new RegExp(`<${esc(prefix)}row(?=[\\s/>])([^>]*?)(/?)>`, 'g');
  const rowClose = `</${prefix}row>`;
  let pos = 0;
  for (let m; (m = row.exec(body));) {
    if (body.slice(pos, m.index).trim()) return null; // unexpected content between rows
    const node: XNode = attributes(m[1]!);
    if (m[2] === '/') {
      rows.push(node);
      pos = row.lastIndex;
      continue;
    }
    const end = body.indexOf(rowClose, row.lastIndex);
    if (end < 0) return null;
    const cells = scanCells(body.slice(row.lastIndex, end), prefix);
    if (!cells) return null;
    if (cells.length) node.c = cells;
    rows.push(node);
    pos = end + rowClose.length;
    row.lastIndex = pos;
  }
  if (body.slice(pos).trim()) return null;
  return { start, end: close + `</${prefix}sheetData>`.length, rows };
}

function esc(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

const ATTRIBUTE = /([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** Attributes as the generic parser stores them (namespace prefixes removed, xmlns dropped). */
function attributes(text: string): XNode {
  const node: XNode = {};
  ATTRIBUTE.lastIndex = 0;
  for (let m; (m = ATTRIBUTE.exec(text));) {
    const name = m[1]!;
    if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
    node[`@_${name.slice(name.indexOf(':') + 1)}`] = m[2] ?? m[3] ?? '';
  }
  return node;
}

/** Element value as the generic parser returns it: text, or an object with attributes. */
function element(attrText: string, text: string | undefined): XNode {
  const node = attributes(attrText);
  if (!Object.keys(node).length) return text ?? '';
  if (text) node['#text'] = text;
  return node;
}

function scanCells(content: string, prefix: string): XNode[] | null {
  const cells: XNode[] = [];
  const cell = new RegExp(`<${esc(prefix)}c(?=[\\s/>])([^>]*?)(/?)>`, 'g');
  const cellClose = `</${prefix}c>`;
  const child = new RegExp(`<${esc(prefix)}(v|f)(?=[\\s/>])([^>]*?)(?:/>|>([^<]*)</${esc(prefix)}\\1>)`, 'gy');
  let pos = 0;
  for (let m; (m = cell.exec(content));) {
    if (content.slice(pos, m.index).trim()) return null;
    const node: XNode = attributes(m[1]!);
    if (m[2] === '/') {
      cells.push(node);
      pos = cell.lastIndex;
      continue;
    }
    const end = content.indexOf(cellClose, cell.lastIndex);
    if (end < 0) return null;
    const inner = content.slice(cell.lastIndex, end);
    let simple = true;
    let i = 0;
    child.lastIndex = 0;
    while (i < inner.length) {
      if (/\s/.test(inner[i]!)) { i++; continue; }
      child.lastIndex = i;
      const c = child.exec(inner);
      if (!c || node[c[1]!] !== undefined) { simple = false; break; }
      node[c[1]!] = c[1] === 'v' ? (c[3] ?? '') : element(c[2]!, c[3]);
      i = child.lastIndex;
    }
    if (simple) cells.push(node);
    else {
      // Inline strings, rich text and extensions: use the generic parser for this cell.
      const parsed = parseXml(`<${prefix}c${m[1]}>${inner}</${prefix}c>`)?.c;
      if (parsed === undefined) return null;
      cells.push(parsed);
    }
    pos = end + cellClose.length;
    cell.lastIndex = pos;
  }
  if (content.slice(pos).trim()) return null;
  return cells;
}
