import { registerFunctions } from '../registry.js';
import { dateFunctions } from './date.js';
import { engineeringFunctions } from './engineering.js';
import { financialFunctions } from './financial.js';
import { arrayFunctions } from './array.js';
import { lambdaFunctions } from './lambda.js';
import { logicalFunctions } from './logical.js';
import { lookupFunctions } from './lookup.js';
import { mathFunctions } from './math.js';
import { statisticsFunctions } from './statistics.js';
import { textFunctions } from './text.js';

let registered = false;

/** Registers all built-in functions (idempotent). */
export function registerBuiltinFunctions(): void {
  if (registered) return;
  registered = true;
  registerFunctions([...mathFunctions, ...statisticsFunctions, ...engineeringFunctions, ...financialFunctions, ...arrayFunctions, ...logicalFunctions, ...lambdaFunctions, ...lookupFunctions, ...textFunctions, ...dateFunctions]);
}

export { makeCriteria, wildcardToRegExp } from './helpers.js';
