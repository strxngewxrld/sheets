import { XMLValidator } from 'fast-xml-parser';
import { decodeEntities, escapeXml } from './xml.js';

export interface XmlSpan {
  name: string; qname: string;
  start: number; openEnd: number; closeStart: number; end: number;
  children: XmlSpan[];
}
export interface XmlEdit { start: number; end: number; value: string }

export function protect(reason: string): never {
  throw new Error(`XLSX-Schutz: ${reason} Diese Änderung kann noch nicht verlustfrei gespeichert werden. Die Originaldatei wurde nicht überschrieben.`);
}

/** Retain byte spans rather than parsing/rebuilding unknown XML. */
export function xmlRoot(xml: string): XmlSpan {
  if (/<!DOCTYPE/i.test(xml) || XMLValidator.validate(xml) !== true) protect('Nicht unterstützte oder ungültige XML-Struktur.');
  const stack: XmlSpan[] = [];
  let root: XmlSpan | undefined;
  let i = 0;
  while ((i = xml.indexOf('<', i)) !== -1) {
    let special: string | undefined;
    if (xml.startsWith('<!--', i)) special = '-->';
    else if (xml.startsWith('<![CDATA[', i)) special = ']]>';
    else if (xml.startsWith('<?', i)) special = '?>';
    if (special) { const end = xml.indexOf(special, i); if (end < 0) protect('XML-Abschluss fehlt.'); i = end + special.length; continue; }
    let end = i + 1, quote = '';
    for (; end < xml.length; end++) {
      const char = xml[end]!;
      if (quote) { if (char === quote) quote = ''; }
      else if (char === '"' || char === "'") quote = char;
      else if (char === '>') break;
    }
    const token = xml.slice(i, end + 1);
    const qname = /^<\/?([^\s/>]+)/.exec(token)?.[1];
    if (!qname) protect('XML-Elementname fehlt.');
    if (token.startsWith('</')) {
      const node = stack.pop();
      if (!node || node.qname !== qname) protect('XML-Verschachtelung ist ungültig.');
      node.closeStart = i; node.end = end + 1;
    } else {
      const node: XmlSpan = { name: qname.split(':').at(-1)!, qname, start: i, openEnd: end + 1, closeStart: end + 1, end: end + 1, children: [] };
      if (stack.length) stack.at(-1)!.children.push(node);
      else { if (root) protect('Mehrere XML-Wurzelelemente.'); root = node; }
      if (!/\/\s*>$/.test(token)) stack.push(node);
    }
    i = end + 1;
  }
  if (!root || stack.length) protect('Unvollständiges XML.');
  return root;
}

export function child(node: XmlSpan, name: string): XmlSpan | undefined {
  const found = node.children.filter((c) => c.name === name);
  if (found.length > 1) protect(`Doppeltes XML-Element ${name}.`);
  if (found.some((c) => c.qname !== prefix(node) + name)) protect(`Nicht unterstützte XML-Namensraumzuordnung bei ${name}.`);
  return found[0];
}
export function raw(xml: string, node: XmlSpan): string { return xml.slice(node.start, node.end); }
export function prefix(node: XmlSpan): string { return node.qname.includes(':') ? node.qname.slice(0, node.qname.indexOf(':') + 1) : ''; }
export function attribute(open: string, name: string): string | undefined {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = new RegExp(`\\s${escaped}\\s*=\\s*(["'])([\\s\\S]*?)\\1`).exec(open);
  return match ? decodeEntities(match[2]!) : undefined;
}
export function setAttribute(open: string, name: string, value: string | number | undefined): string {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`\\s${escaped}\\s*=\\s*(["'])([\\s\\S]*?)\\1`);
  const attr = value === undefined ? '' : ` ${name}="${escapeXml(String(value))}"`;
  if (re.test(open)) return open.replace(re, () => attr);
  return value === undefined ? open : open.replace(/\/?\s*>$/, (end) => attr + end);
}
export function nodeAttribute(xml: string, node: XmlSpan, name: string): string | undefined { return attribute(xml.slice(node.start, node.openEnd), name); }

export function editXml(xml: string, edits: XmlEdit[]): string {
  const sorted = [...edits].sort((a, b) => b.start - a.start || b.end - a.end);
  let limit = xml.length;
  for (const edit of sorted) {
    if (edit.end > limit || edit.start < 0 || edit.end < edit.start) protect('Überlappende XML-Änderungen.');
    xml = xml.slice(0, edit.start) + edit.value + xml.slice(edit.end);
    limit = edit.start;
  }
  return xml;
}

export function contents(xml: string, node: XmlSpan, body: string, open = xml.slice(node.start, node.openEnd)): string {
  return open.replace(/\/\s*>$/, '>') + body + `</${node.qname}>`;
}

/** Patch a single direct child while preserving all other XML, including attributes. */
export function setChild(xml: string, name: string, value: string, before: string[] = []): string {
  const root = xmlRoot(xml), found = child(root, name);
  if (found) return editXml(xml, [{ start: found.start, end: found.end, value }]);
  if (!value) return xml;
  if (root.closeStart === root.openEnd && /\/\s*>$/.test(xml.slice(root.start, root.openEnd))) return contents(xml, root, value);
  const position = root.children.find((n) => before.includes(n.name))?.start ?? root.closeStart;
  return editXml(xml, [{ start: position, end: position, value }]);
}

/**
 * Structural equality with the same semantics as `canonical(a) === canonical(b)`
 * (key order and undefined properties are ignored) without building strings.
 */
export function sameValue(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (Array.isArray(a) || Array.isArray(b)) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (!sameValue(a[i], b[i])) return false;
    return true;
  }
  if (a && b && typeof a === 'object' && typeof b === 'object') {
    const x = a as Record<string, unknown>, y = b as Record<string, unknown>;
    let count = 0;
    for (const key in x) {
      if (x[key] === undefined) continue;
      if (y[key] === undefined || !sameValue(x[key], y[key])) return false;
      count++;
    }
    for (const key in y) if (y[key] !== undefined) count--;
    return count === 0;
  }
  return (JSON.stringify(a) ?? 'null') === (JSON.stringify(b) ?? 'null');
}

export function canonical(value: unknown): string {
  if (Array.isArray(value)) return '[' + value.map(canonical).join(',') + ']';
  if (value && typeof value === 'object') return '{' + Object.entries(value).filter(([, v]) => v !== undefined).sort(([a], [b]) => a.localeCompare(b)).map(([k, v]) => JSON.stringify(k) + ':' + canonical(v)).join(',') + '}';
  return JSON.stringify(value) ?? 'null';
}
