import type { SpacingConfig } from './models.js';

export interface ThemeConfig {
  name: string;
  description: string;
  packages: string[];
  captionPosition: 'top' | 'bottom';
  fontSize: string | null;
  columnSep: string;
  spacing: Required<SpacingConfig>;
}

export type RawThemeConfig = {
  name?: unknown;
  description?: unknown;
  packages?: unknown;
  column_sep?: unknown;
  font_size?: unknown;
  caption_position?: unknown;
  spacing?: Record<string, unknown> | null;
};

function normalizeSpacing(raw: Record<string, unknown> | null | undefined): Required<SpacingConfig> {
  const spacing = raw ?? {};
  return {
    tabcolsep: typeof spacing.tabcolsep === 'string' ? spacing.tabcolsep : null,
    arraystretch: typeof spacing.arraystretch === 'string' ? spacing.arraystretch : null,
    heavyrulewidth: typeof spacing.heavyrulewidth === 'string' ? spacing.heavyrulewidth : '1.0pt',
    lightrulewidth: typeof spacing.lightrulewidth === 'string' ? spacing.lightrulewidth : '0.5pt',
    arrayrulewidth: typeof spacing.arrayrulewidth === 'string' ? spacing.arrayrulewidth : '0.5pt',
    aboverulesep: typeof spacing.aboverulesep === 'string' ? spacing.aboverulesep : '0pt',
    belowrulesep: typeof spacing.belowrulesep === 'string' ? spacing.belowrulesep : '0pt',
  };
}

export function toThemeConfig(raw: RawThemeConfig | null | undefined, fallbackName: string): ThemeConfig {
  if (!raw || typeof raw.name !== 'string' || raw.name.trim() === '') {
    throw new Error(`Theme '${fallbackName}' has invalid config`);
  }

  return {
    name: raw.name,
    description: typeof raw.description === 'string' ? raw.description : '',
    packages: Array.isArray(raw.packages)
      ? raw.packages.filter((pkg): pkg is string => typeof pkg === 'string')
      : [],
    columnSep: typeof raw.column_sep === 'string' ? raw.column_sep : '1em',
    fontSize: typeof raw.font_size === 'string' ? raw.font_size : null,
    captionPosition: raw.caption_position === 'bottom' ? 'bottom' : 'top',
    spacing: normalizeSpacing(raw.spacing),
  };
}
