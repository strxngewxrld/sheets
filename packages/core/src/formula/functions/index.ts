import { registerFunctions } from '../registry.js';
import { dateFunctions } from './date.js';
import { financialFunctions } from './financial.js';
import { arrayFunctions } from './array.js';
import { logicalFunctions } from './logical.js';
import { lookupFunctions } from './lookup.js';
import { mathFunctions } from './math.js';
import { textFunctions } from './text.js';

let registered = false;

/** Registers all built-in functions (idempotent). */
export function registerBuiltinFunctions(): void {
  if (registered) return;
  registered = true;
  registerFunctions([...mathFunctions, ...financialFunctions, ...arrayFunctions, ...logicalFunctions, ...lookupFunctions, ...textFunctions, ...dateFunctions]);
}

export { makeCriteria, wildcardToRegExp } from './helpers.js';
