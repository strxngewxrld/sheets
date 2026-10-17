export type BorderLineStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair';

export interface BorderSide {
  style: BorderLineStyle;
  /** CSS hex color, e.g. "#000000". */
  color?: string;
}

export interface CellBorders {
  top?: BorderSide;
  right?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
}

export type HAlign = 'left' | 'center' | 'right' | 'justify';
export type VAlign = 'top' | 'middle' | 'bottom';

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontName?: string;
  /** Font size in points. */
  fontSize?: number;
  /** Font color as CSS hex "#RRGGBB". */
  color?: string;
  /** Background fill as CSS hex "#RRGGBB". */
  fill?: string;
  hAlign?: HAlign;
  vAlign?: VAlign;
  wrap?: boolean;
  /** Excel number format code, e.g. "#,##0.00" or "dd.mm.yyyy". */
  numFmt?: string;
  border?: CellBorders;
}

/** A style patch: `null` removes a property. */
export type StylePatch = { [K in keyof CellStyle]?: CellStyle[K] | null };

export const STYLE_KEYS: readonly (keyof CellStyle)[] = [
  'bold', 'italic', 'underline', 'strike', 'fontName', 'fontSize', 'color', 'fill', 'hAlign', 'vAlign', 'wrap', 'numFmt', 'border',
];

export function applyStylePatch(base: CellStyle | undefined, patch: StylePatch): CellStyle | undefined {
  const out: Record<string, unknown> = { ...(base ?? {}) };
  for (const [k, v] of Object.entries(patch)) {
    if (v === null || v === undefined || v === '') delete out[k];
    else out[k] = v;
  }
  return Object.keys(out).length ? (out as CellStyle) : undefined;
}

export function stylesEqual(a: CellStyle | undefined, b: CellStyle | undefined): boolean {
  return JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});
}

/** Stable key used to deduplicate styles (e.g. for xlsx export). */
export function styleKey(s: CellStyle | undefined): string {
  if (!s) return '';
  const sorted: Record<string, unknown> = {};
  for (const k of STYLE_KEYS) if (s[k] !== undefined) sorted[k] = s[k];
  return JSON.stringify(sorted);
}
