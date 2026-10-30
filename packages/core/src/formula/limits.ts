import { CellError } from '../values.js';

export const MAX_ARRAY_CELLS = 1_000_000;
/** Excel's limit for the length of a text value. */
export const MAX_TEXT_LENGTH = 32_767;

/**
 * Budgets of one formula and of one recalculation. Work units (evaluated nodes,
 * LAMBDA calls, materialized cells and text characters) bound memory; the time
 * limits bound CPU, as units of different work take very different time.
 * Files from others must not be able to block the editor or the extension host.
 * Mutable for tests.
 */
export const calcLimits = {
  cellWork: 10_000_000,
  passWork: 100_000_000,
  cellMs: 2_000,
  passMs: 5_000,
};

/**
 * Thrown when a calculation budget is used up. It is not a CellError, so that
 * functions cannot turn it into a value and continue; the cell gets #CALC!.
 */
export class CalcLimitError extends Error {}

let cellWork = 0;
let passWork = 0;
let cellDeadline = Infinity;
let passDeadline = Infinity;
let passDepth = 0;
let charges = 0;
let unchecked = 0;

/** Counts work against the budgets of the current formula and recalculation. */
export function chargeWork(units: number): void {
  cellWork += units;
  passWork += units;
  if (cellWork > calcLimits.cellWork) throw new CalcLimitError('Formula exceeds the calculation limit');
  if (passWork > calcLimits.passWork) throw new CalcLimitError('Workbook exceeds the calculation limit');
  // Reading the clock on every charge would cost more than most of the work.
  unchecked += units;
  if ((++charges & 0x3fff) !== 0 && unchecked < 0x100000) return;
  unchecked = 0;
  const now = performance.now();
  if (now > cellDeadline) throw new CalcLimitError('Formula exceeds the calculation time limit');
  if (now > passDeadline) throw new CalcLimitError('Workbook exceeds the calculation time limit');
}

/** Runs one formula with its own budget; formulas computed meanwhile as dependencies get theirs. */
export function withCellBudget<T>(fn: () => T): T {
  const outerWork = cellWork;
  const outerDeadline = cellDeadline;
  cellWork = 0;
  cellDeadline = performance.now() + calcLimits.cellMs;
  try {
    return fn();
  } finally {
    cellWork = outerWork;
    cellDeadline = outerDeadline;
  }
}

/** Runs a recalculation; nested calls share the budget of the outermost one. */
export function withPassBudget<T>(fn: () => T): T {
  if (passDepth++ === 0) {
    passWork = 0;
    passDeadline = performance.now() + calcLimits.passMs;
  }
  try {
    return fn();
  } finally {
    if (--passDepth === 0) passDeadline = Infinity;
  }
}

/** Check before allocating a dense formula result, including intermediate matrices. */
export function checkMatrixSize(rows: number, cols: number): void {
  if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows < 0 || cols < 0 ||
      rows > MAX_ARRAY_CELLS || cols > MAX_ARRAY_CELLS || rows * cols > MAX_ARRAY_CELLS) {
    throw new CellError('#CALC!', 'Array exceeds the cell limit');
  }
  chargeWork(rows * cols);
}

/** Check before building a text of `length` characters. */
export function checkTextLength(length: number): void {
  if (!(length <= MAX_TEXT_LENGTH)) throw new CellError('#VALUE!', 'Text exceeds 32,767 characters');
  chargeWork(length);
}
