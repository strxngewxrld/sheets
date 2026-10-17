import { MAX_COLS, MAX_ROWS, nameToCol, splitSheetPrefix } from '../address.js';
import { errorFromCode } from '../values.js';
import { FormulaParseError, type BinaryOp, type Node, type RefPoint } from './ast.js';
import { tokenize, type Token } from './tokenizer.js';

const BINARY_PRECEDENCE: Record<string, number> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
  ':': 8,
};
const PERCENT_PRECEDENCE = 6;
const PREFIX_PRECEDENCE = 7;

/** Strips Excel's future-function prefixes (_xlfn., _xlws.). */
export function normalizeFunctionName(name: string): string {
  return name.toUpperCase().replace(/^(_XLFN\.)?(_XLWS\.)?/, '');
}

export function parseRefPoint(text: string): RefPoint {
  const m = /^(\$?)([A-Za-z]{0,3})(\$?)(\d*)$/.exec(text)!;
  return {
    col: m[2] ? nameToCol(m[2]) : 0,
    row: m[4] ? Number(m[4]) - 1 : 0,
    absCol: m[1] === '$',
    absRow: m[3] === '$',
  };
}

/** Converts a `ref` token text into a cell or range node. */
export function refTokenToNode(text: string): Node {
  const { sheet, ref } = splitSheetPrefix(text);
  const withSheet = <T extends object>(n: T): T => (sheet !== undefined ? { ...n, sheet } : n);
  const parts = ref.split(':');
  if (parts.length === 1) return withSheet({ type: 'cell' as const, ref: parseRefPoint(parts[0]!) });
  const [a, b] = parts as [string, string];
  if (/^\$?[A-Za-z]+$/.test(a)) {
    const start = parseRefPoint(a);
    const end = parseRefPoint(b);
    start.row = 0;
    end.row = MAX_ROWS - 1;
    start.absRow = end.absRow = true;
    return withSheet({ type: 'range' as const, start, end, kind: 'cols' as const });
  }
  if (/^\$?\d+$/.test(a)) {
    const start = parseRefPoint(a);
    const end = parseRefPoint(b);
    start.col = 0;
    end.col = MAX_COLS - 1;
    start.absCol = end.absCol = true;
    return withSheet({ type: 'range' as const, start, end, kind: 'rows' as const });
  }
  return withSheet({ type: 'range' as const, start: parseRefPoint(a), end: parseRefPoint(b), kind: 'cells' as const });
}

class Parser {
  private pos = 0;
  private readonly tokens: Token[];
  constructor(tokens: Token[]) {
    this.tokens = tokens.filter((t) => t.type !== 'ws');
  }

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private next(): Token {
    const t = this.tokens[this.pos++];
    if (!t) throw new FormulaParseError('Unexpected end of formula', this.lastEnd());
    return t;
  }

  private lastEnd(): number {
    return this.tokens.at(-1)?.end ?? 0;
  }

  private expect(type: Token['type']): Token {
    const t = this.next();
    if (t.type !== type) throw new FormulaParseError(`Expected ${type} but found "${t.text}"`, t.start);
    return t;
  }

  parse(): Node {
    if (this.tokens.length === 0) throw new FormulaParseError('Empty formula', 0);
    const node = this.expression(0);
    const extra = this.peek();
    if (extra) throw new FormulaParseError(`Unexpected "${extra.text}"`, extra.start);
    return node;
  }

  private expression(minPrec: number): Node {
    let left = this.prefix();
    for (;;) {
      const t = this.peek();
      if (!t || t.type !== 'op') break;
      if (t.text === '#') {
        this.pos++;
        left = { type: 'spill', arg: left };
        continue;
      }
      if (t.text === '%') {
        if (PERCENT_PRECEDENCE < minPrec) break;
        this.pos++;
        left = { type: 'percent', arg: left };
        continue;
      }
      const prec = BINARY_PRECEDENCE[t.text];
      if (prec === undefined || prec < minPrec) break;
      this.pos++;
      const right = this.expression(prec + 1);
      left = { type: 'binary', op: t.text as BinaryOp, left, right };
    }
    return left;
  }

  private prefix(): Node {
    const t = this.next();
    switch (t.type) {
      case 'number':
        return { type: 'number', value: Number(t.text) };
      case 'string':
        return { type: 'string', value: t.text.slice(1, -1).replace(/""/g, '"') };
      case 'bool':
        return { type: 'bool', value: t.text.toUpperCase() === 'TRUE' };
      case 'error':
        return { type: 'error', code: errorFromCode(t.text).code };
      case 'ref':
        return refTokenToNode(t.text);
      case 'name':
        return { type: 'name', name: t.text };
      case 'func':
        return this.call(t);
      case 'lparen': {
        const inner = this.expression(0);
        this.expect('rparen');
        return inner;
      }
      case 'lbrace':
        return this.arrayLiteral();
      case 'op':
        if (t.text === '-' || t.text === '+') {
          return { type: 'unary', op: t.text, arg: this.expression(PREFIX_PRECEDENCE) };
        }
        break;
    }
    throw new FormulaParseError(`Unexpected "${t.text}"`, t.start);
  }

  private call(nameTok: Token): Node {
    this.expect('lparen');
    const args: Node[] = [];
    const name = normalizeFunctionName(nameTok.text);
    if (this.peek()?.type === 'rparen') {
      this.pos++;
      return { type: 'func', name, args };
    }
    for (;;) {
      const t = this.peek();
      if (t && (t.type === 'comma' || t.type === 'rparen')) args.push({ type: 'missing' });
      else args.push(this.expression(0));
      const sep = this.next();
      if (sep.type === 'rparen') break;
      if (sep.type !== 'comma') throw new FormulaParseError(`Expected "," or ")" but found "${sep.text}"`, sep.start);
    }
    return name === 'ANCHORARRAY' && args.length === 1 ? { type: 'spill', arg: args[0]! } : { type: 'func', name, args };
  }

  private arrayLiteral(): Node {
    const rows: Node[][] = [[]];
    for (;;) {
      rows.at(-1)!.push(this.expression(0));
      const sep = this.next();
      if (sep.type === 'rbrace') break;
      if (sep.type === 'semicolon') rows.push([]);
      else if (sep.type !== 'comma') throw new FormulaParseError(`Unexpected "${sep.text}" in array`, sep.start);
    }
    return { type: 'array', rows };
  }
}

export function parseFormula(formula: string): Node {
  const src = formula.startsWith('=') ? formula.slice(1) : formula;
  return new Parser(tokenize(src)).parse();
}

/** Visits all nodes depth-first. */
export function walk(node: Node, visit: (n: Node) => void): void {
  visit(node);
  switch (node.type) {
    case 'func':
      node.args.forEach((a) => walk(a, visit));
      break;
    case 'unary':
    case 'percent':
    case 'spill':
      walk(node.arg, visit);
      break;
    case 'binary':
      walk(node.left, visit);
      walk(node.right, visit);
      break;
    case 'array':
      node.rows.forEach((r) => r.forEach((c) => walk(c, visit)));
      break;
  }
}
