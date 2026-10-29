import { RE2JS } from 're2js';

/** Untrusted patterns must never run in the native backtracking engine. */
export function compileSafeRegex(source: string, flags = 0): RE2JS {
  if (source.length > 4096) throw new Error('Regular expression exceeds the 4096-character limit.');
  return RE2JS.compile(RE2JS.translateRegExp(source), flags);
}

export { RE2JS };
