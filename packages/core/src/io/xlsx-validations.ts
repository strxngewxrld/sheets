import { parseRange } from '../address.js';
import { addFutureFunctionPrefixes, stripFutureFunctionPrefixes } from '../formula/rewrite.js';
import { VALIDATION_OPERATORS, VALIDATION_TYPES, type DataValidation, type ValidationOperator, type ValidationType } from '../validation.js';
import { arr, attr, decodeOoxmlEscapes, encodeOoxmlEscapes, escapeXml, text, type XNode } from './xml.js';

/** Reads <dataValidations> and the Excel 2010 variant in extLst (lists from other sheets). */
export function readValidations(ws: XNode): DataValidation[] {
  const nodes = [
    ...arr(ws.dataValidations?.dataValidation).map((node) => ({ node, sqref: attr(node, 'sqref'), f1: node.formula1 === undefined ? undefined : text(node.formula1), f2: node.formula2 === undefined ? undefined : text(node.formula2) })),
    ...arr(ws.extLst?.ext).flatMap((ext) => arr(ext.dataValidations?.dataValidation)).map((node) => ({
      node, sqref: node.sqref === undefined ? undefined : text(node.sqref), f1: node.formula1?.f === undefined ? undefined : text(node.formula1.f), f2: node.formula2?.f === undefined ? undefined : text(node.formula2.f),
    })),
  ];
  const out: DataValidation[] = [];
  for (const { node, sqref, f1, f2 } of nodes) {
    const type = (attr(node, 'type') ?? 'none') as ValidationType;
    const range = (sqref ?? '').trim().split(/\s+/).filter((r) => parseRange(r) && !parseRange(r)!.sheet).join(' ');
    if (!range || !VALIDATION_TYPES.includes(type) || f1 === undefined || !f1.trim()) continue;
    const yes = (name: string) => ['1', 'true'].includes(attr(node, name) ?? '');
    const rule: DataValidation = { range, type, formula1: stripFutureFunctionPrefixes(f1) };
    const operator = attr(node, 'operator') as ValidationOperator | undefined;
    if (operator && VALIDATION_OPERATORS.includes(operator) && type !== 'list' && type !== 'custom') rule.operator = operator;
    if (f2 !== undefined && f2.trim()) rule.formula2 = stripFutureFunctionPrefixes(f2);
    if (!yes('allowBlank')) rule.allowBlank = false;
    if (yes('showDropDown')) rule.dropdown = false; // Excel's flag hides the dropdown.
    const style = attr(node, 'errorStyle');
    if (style === 'warning' || style === 'information') rule.errorStyle = style;
    if (!yes('showErrorMessage')) rule.showError = false;
    if (!yes('showInputMessage')) rule.showPrompt = false;
    for (const [key, name] of [['errorTitle', 'errorTitle'], ['error', 'error'], ['promptTitle', 'promptTitle'], ['prompt', 'prompt']] as const) {
      const value = attr(node, name);
      if (value) rule[key] = decodeOoxmlEscapes(value);
    }
    out.push(rule);
  }
  return out;
}

/** <dataValidations> element for a sheet (prefix for namespaced originals). */
export function validationsXml(rules: DataValidation[], p: string): string {
  const items = rules.map((rule) => {
    const attrs = [
      `type="${rule.type}"`,
      rule.type !== 'list' && rule.type !== 'custom' && rule.operator && rule.operator !== 'between' ? `operator="${rule.operator}"` : '',
      rule.errorStyle && rule.errorStyle !== 'stop' ? `errorStyle="${rule.errorStyle}"` : '',
      rule.allowBlank !== false ? 'allowBlank="1"' : '',
      rule.dropdown === false ? 'showDropDown="1"' : '',
      rule.showPrompt !== false ? 'showInputMessage="1"' : '',
      rule.showError !== false ? 'showErrorMessage="1"' : '',
      ...(['errorTitle', 'error', 'promptTitle', 'prompt'] as const).map((key) => (rule[key] ? `${key}="${escapeXml(encodeOoxmlEscapes(rule[key]!))}"` : '')),
      `sqref="${escapeXml(rule.range)}"`,
    ].filter(Boolean).join(' ');
    const formula = (name: string, f: string | undefined) => (f === undefined ? '' : `<${p}${name}>${escapeXml(addFutureFunctionPrefixes(f))}</${p}${name}>`);
    return `<${p}dataValidation ${attrs}>${formula('formula1', rule.formula1)}${formula('formula2', rule.formula2)}</${p}dataValidation>`;
  });
  return `<${p}dataValidations count="${rules.length}">${items.join('')}</${p}dataValidations>`;
}
