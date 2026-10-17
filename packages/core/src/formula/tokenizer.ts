import { MAX_COLS, MAX_ROWS, nameToCol } from '../address.js';

export type TokenType =
  | 'number'
  | 'string'
  | 'bool'
  | 'error'
  | 'ref'
  | 'func'
  | 'name'
  | 'op'
  | 'lparen'
  | 'rparen'
  | 'comma'
  | 'semicolon'
  | 'lbrace'
  | 'rbrace'
  | 'ws'
  | 'unknown';

export interface Token {
  type: TokenType;
  text: string;
  start: number;
  end: number;
}

const SHEET_PREFIX = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)!`;
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
const COL = String.raw`\$?[A-Za-z]{1,3}`;
const ROW = String.raw`\$?\d+`;
const REF_RE = new RegExp(
  String.raw`^(?:${SHEET_PREFIX})?(?:${CELL}(?::${CELL})?|${COL}:${COL}|${ROW}:${ROW})`,
  'u',
);
const ERROR_RE = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|CIRC!|SPILL!|CALC!|GETTING_DATA)/i;
const NUMBER_RE = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
const IDENT_RE = /^[A-Za-z_\\À-￿][\w.À-￿]*/u;
const OPS = ['<=', '>=', '<>', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', ':'];

/** Validates that every cell/column/row inside a ref text is within sheet bounds. */
function refInBounds(text: string): boolean {
  const body = text.slice(text.lastIndexOf('!') + 1);
  for (const part of body.split(':')) {
    const m = /^\$?([A-Za-z]{0,3})\$?(\d*)$/.exec(part);
    if (!m) return false;
    if (m[1] && nameToCol(m[1]) >= MAX_COLS) return false;
    if (m[2] && (Number(m[2]) < 1 || Number(m[2]) > MAX_ROWS)) return false;
  }
  return true;
}

/**
 * Splits a formula (without the leading "=") into tokens. Whitespace is kept
 * as tokens so formulas can be rewritten losslessly.
 */
export function tokenize(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const push = (type: TokenType, len: number) => {
    tokens.push({ type, text: src.slice(i, i + len), start: i, end: i + len });
    i += len;
  };
  while (i < src.length) {
    const ch = src[i]!;
    const rest = src.slice(i);
    if (/\s/.test(ch)) {
      push('ws', /^\s+/.exec(rest)![0].length);
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') j += 2;
          else break;
        } else j++;
      }
      push('string', Math.min(j + 1, src.length) - i);
      continue;
    }
    if (ch === '#') {
      const m = ERROR_RE.exec(rest);
      push(m ? 'error' : 'op', m ? m[0].length : 1);
      continue;
    }
    // References (must be tried before numbers for row ranges like 1:3 and before identifiers).
    const prev = lastSignificant(tokens);
    const refM = REF_RE.exec(rest);
    if (refM && refInBounds(refM[0])) {
      const after = src[i + refM[0].length] ?? '';
      const isRowRangeOnly = /^\$?\d+:\$?\d+$/.test(refM[0]);
      const blockedByPrev = isRowRangeOnly && prev && (prev.type === 'number' || prev.type === 'ref');
      if (!/[\w.(À-￿]/u.test(after) && !blockedByPrev) {
        push('ref', refM[0].length);
        continue;
      }
    }
    const numM = NUMBER_RE.exec(rest);
    if (numM && /[\d.]/.test(ch)) {
      push('number', numM[0].length);
      continue;
    }
    if (ch === "'") {
      // Quoted sheet name followed by something we could not parse as a ref.
      const m = /^'(?:[^']|'')+'!/.exec(rest);
      push('unknown', m ? m[0].length : 1);
      continue;
    }
    const idM = IDENT_RE.exec(rest);
    if (idM) {
      const word = idM[0];
      const next = src.slice(i + word.length).match(/^\s*\(/);
      if (next) push('func', word.length);
      else if (/^(TRUE|FALSE)$/i.test(word)) push('bool', word.length);
      else push('name', word.length);
      continue;
    }
    const op = OPS.find((o) => rest.startsWith(o));
    if (op) {
      push('op', op.length);
      continue;
    }
    switch (ch) {
      case '(': push('lparen', 1); break;
      case ')': push('rparen', 1); break;
      case ',': push('comma', 1); break;
      case ';': push('semicolon', 1); break;
      case '{': push('lbrace', 1); break;
      case '}': push('rbrace', 1); break;
      default: push('unknown', 1);
    }
  }
  return tokens;
}

function lastSignificant(tokens: Token[]): Token | undefined {
  for (let k = tokens.length - 1; k >= 0; k--) if (tokens[k]!.type !== 'ws') return tokens[k];
  return undefined;
}
