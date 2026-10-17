export type ErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#CIRC!'
  | '#SPILL!'
  | '#CALC!'
  | '#GETTING_DATA';

export const ERROR_CODES: readonly ErrorCode[] = [
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#CIRC!', '#SPILL!', '#CALC!', '#GETTING_DATA',
];

export class CellError {
  readonly code: ErrorCode;
  readonly message?: string;
  constructor(code: ErrorCode, message?: string) {
    this.code = code;
    if (message !== undefined) this.message = message;
  }
  toString(): string {
    return this.code;
  }
  toJSON(): { error: ErrorCode } {
    return { error: this.code };
  }
}

/** A computed cell value. `null` means empty. */
export type Scalar = number | string | boolean | null | CellError;

export const Err = {
  NULL: new CellError('#NULL!'),
  DIV0: new CellError('#DIV/0!'),
  VALUE: new CellError('#VALUE!'),
  REF: new CellError('#REF!'),
  NAME: new CellError('#NAME?'),
  NUM: new CellError('#NUM!'),
  NA: new CellError('#N/A'),
  CIRC: new CellError('#CIRC!'),
} as const;

export function isError(v: unknown): v is CellError {
  return v instanceof CellError;
}

export function errorFromCode(code: string): CellError {
  const c = code.toUpperCase();
  return new CellError((ERROR_CODES as readonly string[]).includes(c) ? (c as ErrorCode) : '#VALUE!');
}

/** Serializable representation of a scalar (errors become `{error: code}`). */
export type JsonScalar = number | string | boolean | null | { error: ErrorCode };

export function scalarToJson(v: Scalar): JsonScalar {
  return isError(v) ? { error: v.code } : v;
}

export function scalarFromJson(v: JsonScalar | undefined): Scalar {
  if (v === undefined) return null;
  if (v !== null && typeof v === 'object') return errorFromCode(v.error);
  return v;
}
