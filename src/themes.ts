export interface ThemeConfig {
  name: string;
  packages: string[];
  captionPosition: 'top' | 'bottom';
  fontSize: string | null;
  columnSep: string;
}

const THEMES: Record<string, ThemeConfig> = {
  three_line: {
    name: 'three_line',
    packages: ['booktabs', 'multirow', 'xcolor', 'diagbox', 'makecell', 'adjustbox', 'amssymb', 'pifont'],
    captionPosition: 'top',
    fontSize: null,
    columnSep: '1em',
  },
  // Fallback style for simple layouts, kept for compatibility/extension.
  simple: {
    name: 'simple',
    packages: ['booktabs', 'multirow', 'xcolor'],
    captionPosition: 'top',
    fontSize: null,
    columnSep: '1em',
  },
};

export const DEFAULT_THEME = 'three_line';

export function listThemes(): string[] {
  return Object.keys(THEMES).sort((a, b) => a.localeCompare(b));
}

export function getTheme(name: string | undefined | null): ThemeConfig {
  const n = (name == null || !String(name).trim()) ? DEFAULT_THEME : String(name).trim();
  const theme = THEMES[n];
  if (!theme) {
    throw new Error(`Theme '${n}' not found. Available: ${listThemes().join(', ')}`);
  }
  return theme;
}
