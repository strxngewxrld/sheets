import { STYLE_KEYS, type CellStyle } from '../style.js';
import { toArgb } from './colors.js';
import { escapeXml } from './xml.js';
import { attribute, canonical, child, contents, editXml, nodeAttribute, prefix, protect, raw, setAttribute, setChild, xmlRoot } from './xlsx-preserve-xml.js';

const COLLECTIONS = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
const ITEMS: Record<string, string> = { numFmts: 'numFmt', fonts: 'font', fills: 'fill', borders: 'border', cellXfs: 'xf' };
const FONT_TAGS = { bold: 'b', italic: 'i', underline: 'u', strike: 'strike', fontName: 'name', fontSize: 'sz', color: 'color' } as const;

/** Append-only style resources; existing indices and unimplemented properties stay intact. */
export class PreservedStyles {
  private cache = new Map<string, number>();
  constructor(public xml: string) {}

  private item(collection: string, index: number): string {
    const root = xmlRoot(this.xml), group = child(root, collection);
    const item = group?.children.filter((n) => n.name === ITEMS[collection])[index];
    if (!item) protect(`Style-Referenz ${collection}[${index}] ist ungültig.`, `Invalid style reference ${collection}[${index}].`);
    return raw(this.xml, item);
  }

  private append(collection: string, value: string): number {
    const root = xmlRoot(this.xml), p = prefix(root), group = child(root, collection);
    const count = group?.children.filter((n) => n.name === ITEMS[collection]).length ?? 0;
    if (group) {
      const open = setAttribute(this.xml.slice(group.start, group.openEnd), 'count', count + 1);
      this.xml = editXml(this.xml, [{ start: group.start, end: group.end, value: contents(this.xml, group, this.xml.slice(group.openEnd, group.closeStart) + value, open) }]);
    } else {
      const later = COLLECTIONS.slice(COLLECTIONS.indexOf(collection) + 1);
      this.xml = setChild(this.xml, collection, `<${p}${collection} count="1">${value}</${p}${collection}>`, later);
    }
    return count;
  }

  private open(xml: string, updates: Record<string, string | number | undefined>): string {
    const root = xmlRoot(xml);
    let open = xml.slice(root.start, root.openEnd);
    for (const [name, value] of Object.entries(updates)) open = setAttribute(open, name, value);
    return editXml(xml, [{ start: root.start, end: root.openEnd, value: open }]);
  }

  /** Everything a cell format record resolves to; equal fingerprints format cells identically. */
  fingerprint(index: number): string {
    const xf = this.item('cellXfs', index);
    const open = xf.slice(0, xmlRoot(xf).openEnd);
    const id = attribute(open, 'numFmtId') ?? '0';
    const root = xmlRoot(this.xml);
    const numFmt = child(root, 'numFmts')?.children.find((n) => n.name === 'numFmt' && nodeAttribute(this.xml, n, 'numFmtId') === id);
    const part = (collection: string, attr: string) => this.item(collection, Number(attribute(open, attr) ?? 0));
    return [xf.replace(/^<[^\s/>]+/, ''), part('fonts', 'fontId'), part('fills', 'fillId'), part('borders', 'borderId'), numFmt ? nodeAttribute(this.xml, numFmt, 'formatCode') : `#${id}`].join('\u0000');
  }

  get(original: number, before: CellStyle = {}, after: CellStyle = {}): number {
    const changed = STYLE_KEYS.filter((key) => canonical(before[key]) !== canonical(after[key]));
    if (!changed.length) return original;
    for (const key of ['color', 'fill'] as const) if (changed.includes(key) && after[key] && !/^#[0-9a-f]{6}$/i.test(after[key]!)) protect('Ungültige Style-Farbe.', 'Invalid style color.');
    if (changed.includes('fontSize') && after.fontSize !== undefined && (!Number.isFinite(after.fontSize) || after.fontSize <= 0)) protect('Ungültige Schriftgröße.', 'Invalid font size.');
    const cacheKey = canonical([original, changed.map((key) => [key, after[key]])]);
    const known = this.cache.get(cacheKey); if (known !== undefined) return known;
    let xf = this.item('cellXfs', original);
    const p = prefix(xmlRoot(xf));
    const xfOpen = xf.slice(0, xmlRoot(xf).openEnd);
    const updates: Record<string, number | string | undefined> = {};
    if (changed.some((key) => key in FONT_TAGS)) {
      let font = this.item('fonts', Number(attribute(xfOpen, 'fontId') ?? 0));
      const fp = prefix(xmlRoot(font));
      for (const key of Object.keys(FONT_TAGS) as (keyof typeof FONT_TAGS)[]) {
        if (!changed.includes(key)) continue;
        const tag = FONT_TAGS[key], value = after[key];
        const childXml = value === undefined ? '' : typeof value === 'boolean'
          ? `<${fp}${tag}${value ? '' : ` val="${key === 'underline' ? 'none' : '0'}"`}/>`
          : `<${fp}${tag} ${key === 'color' ? 'rgb' : 'val'}="${escapeXml(key === 'color' ? toArgb(String(value)) : String(value))}"/>`;
        font = setChild(font, tag, childXml);
      }
      updates.fontId = this.append('fonts', font); updates.applyFont = 1;
    }
    if (changed.includes('fill')) {
      updates.fillId = after.fill ? this.append('fills', `<${p}fill><${p}patternFill patternType="solid"><${p}fgColor rgb="${toArgb(after.fill)}"/><${p}bgColor indexed="64"/></${p}patternFill></${p}fill>`) : 0;
      updates.applyFill = 1;
    }
    if (changed.includes('numFmt')) {
      let id = 0;
      if (after.numFmt && after.numFmt !== 'General') {
        const root = xmlRoot(this.xml), group = child(root, 'numFmts');
        const entries = group?.children.filter((n) => n.name === 'numFmt') ?? [];
        const existing = entries.find((n) => nodeAttribute(this.xml, n, 'formatCode') === after.numFmt);
        id = existing ? Number(nodeAttribute(this.xml, existing, 'numFmtId')) : Math.max(163, ...entries.map((n) => Number(nodeAttribute(this.xml, n, 'numFmtId')))) + 1;
        if (!Number.isInteger(id) || id < 0) protect('Ungültige Zahlenformat-Referenz.', 'Invalid number format reference.');
        if (!existing) this.append('numFmts', `<${p}numFmt numFmtId="${id}" formatCode="${escapeXml(after.numFmt)}"/>`);
      }
      updates.numFmtId = id; updates.applyNumberFormat = 1;
    }
    if (changed.some((key) => ['hAlign', 'vAlign', 'wrap'].includes(key))) {
      const node = child(xmlRoot(xf), 'alignment');
      let alignment = node ? raw(xf, node) : `<${p}alignment/>`;
      const changes: Record<string, string | number | undefined> = {};
      if (changed.includes('hAlign')) changes.horizontal = after.hAlign;
      if (changed.includes('vAlign')) changes.vertical = after.vAlign === 'middle' ? 'center' : after.vAlign;
      if (changed.includes('wrap')) changes.wrapText = after.wrap === undefined ? undefined : after.wrap ? 1 : 0;
      alignment = this.open(alignment, changes);
      xf = setChild(xf, 'alignment', alignment, ['protection', 'extLst']); updates.applyAlignment = 1;
    }
    if (changed.includes('border')) {
      let border = this.item('borders', Number(attribute(xfOpen, 'borderId') ?? 0));
      for (const side of ['left', 'right', 'top', 'bottom'] as const) {
        const a = before.border?.[side], b = after.border?.[side];
        if (canonical(a) === canonical(b)) continue;
        const node = child(xmlRoot(border), side);
        let sideXml = node ? raw(border, node) : `<${p}${side}/>`;
        sideXml = this.open(sideXml, { style: b?.style });
        if (a?.color !== b?.color) sideXml = setChild(sideXml, 'color', b?.color ? `<${p}color rgb="${toArgb(b.color)}"/>` : '');
        border = setChild(border, side, sideXml, ['diagonal', 'vertical', 'horizontal', 'start', 'end', 'extLst']);
      }
      updates.borderId = this.append('borders', border); updates.applyBorder = 1;
    }
    xf = this.open(xf, updates);
    const index = this.append('cellXfs', xf); this.cache.set(cacheKey, index); return index;
  }
}
