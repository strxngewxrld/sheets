import type { ErrorCode } from '../values.js';

export interface RefPoint {
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
}

export type StructuredArea = '#All' | '#Data' | '#Headers' | '#Totals' | '#This Row';

/** Parsed structured table reference (Table1[[#Headers],[A]:[B]], [@Col], ...). */
export interface StructuredRef {
  /** Table name; omitted for references inside a table such as [@Col]. */
  table?: string;
  /** Row areas; empty means #Data. */
  areas: StructuredArea[];
  /** First and last column name; omitted means all columns. */
  columns?: [string, string];
}

export type Node =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'bool'; value: boolean }
  | { type: 'error'; code: ErrorCode }
  | { type: 'cell'; sheet?: string; ref: RefPoint }
  | { type: 'range'; sheet?: string; start: RefPoint; end: RefPoint; kind: 'cells' | 'cols' | 'rows' }
  | { type: 'name'; name: string }
  | ({ type: 'structured' } & StructuredRef)
  | { type: 'func'; name: string; args: Node[] }
  | { type: 'unary'; op: '-' | '+'; arg: Node }
  | { type: 'percent'; arg: Node }
  | { type: 'spill'; arg: Node }
  | { type: 'binary'; op: BinaryOp; left: Node; right: Node }
  | { type: 'array'; rows: Node[][] }
  | { type: 'missing' };

export type BinaryOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=' | ':';

export class FormulaParseError extends Error {
  readonly position: number;
  constructor(message: string, position: number) {
    super(message);
    this.position = position;
  }
}
