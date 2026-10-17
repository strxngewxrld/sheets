import { XMLParser } from 'fast-xml-parser';

/* eslint-disable @typescript-eslint/no-explicit-any */
export type XNode = any;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  processEntities: false,
  ignoreDeclaration: true,
  ignorePiTags: true,
});

export function parseXml(xml: string): XNode {
  return parser.parse(xml);
}

const ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_, e: string) => {
    if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    return ENTITIES[e]!;
  });
}

/** Decodes OOXML "_xHHHH_" escapes. */
export function decodeOoxmlEscapes(s: string): string {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)));
}

export function arr<T = XNode>(x: T | T[] | undefined | null): T[] {
  if (x === undefined || x === null) return [];
  return Array.isArray(x) ? x : [x];
}

export function text(n: XNode): string {
  if (n === undefined || n === null) return '';
  if (typeof n === 'object') {
    const t = n['#text'];
    return t === undefined ? '' : decodeEntities(String(t));
  }
  return decodeEntities(String(n));
}

export function attr(n: XNode, name: string): string | undefined {
  if (!n || typeof n !== 'object') return undefined;
  const v = n[`@_${name}`];
  return v === undefined ? undefined : decodeEntities(String(v));
}

export function numAttr(n: XNode, name: string): number | undefined {
  const v = attr(n, name);
  return v === undefined || v === '' ? undefined : Number(v);
}

/** Boolean OOXML flag element like <b/> or <b val="0"/>. */
export function flag(n: XNode, key: string): boolean {
  if (!n || typeof n !== 'object' || !(key in n)) return false;
  const v = attr(n[key], 'val');
  return v === undefined || (v !== '0' && v !== 'false');
}

/** Collects the text of all <t> elements below a node (rich text). */
export function collectText(n: XNode): string {
  if (n === undefined || n === null || typeof n !== 'object') return '';
  let out = '';
  for (const [k, v] of Object.entries(n)) {
    if (k.startsWith('@_')) continue;
    if (k === 't') out += arr(v).map(text).join('');
    else for (const child of arr(v)) out += collectText(child);
  }
  return out;
}

export function escapeXml(s: string): string {
  return s
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Resolves a relationship target relative to the directory of the source part. */
export function resolvePath(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const base = sourcePart.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop();
    else if (seg !== '.') base.push(seg);
  }
  return base.join('/');
}

export function relsPathFor(part: string): string {
  const dir = part.split('/').slice(0, -1).join('/');
  const file = part.split('/').pop();
  return `${dir ? dir + '/' : ''}_rels/${file}.rels`;
}
