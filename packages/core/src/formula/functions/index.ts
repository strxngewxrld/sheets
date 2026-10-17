import { registerFunctions } from '../registry.js';
import { dateFunctions } from './date.js';
import { logicalFunctions } from './logical.js';
import { lookupFunctions } from './lookup.js';
import { mathFunctions } from './math.js';
import { textFunctions } from './text.js';

let registered = false;

/** Registers all built-in functions (idempotent). */
export function registerBuiltinFunctions(): void {
  if (registered) return;
  registered = true;
  registerFunctions([...mathFunctions, ...logicalFunctions, ...lookupFunctions, ...textFunctions, ...dateFunctions]);
}

export { makeCriteria, wildcardToRegExp } from './helpers.js';
