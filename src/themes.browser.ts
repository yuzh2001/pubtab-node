import type { RawThemeConfig, ThemeConfig } from './theme-schema.js';
import { toThemeConfig } from './theme-schema.js';
import { THEME_PRESETS } from './generated/theme-presets.js';

const DEFAULT_THEME = 'three_line';
const THEME_CACHE = new Map<string, ThemeConfig>();

type ThemePresetMap = Record<string, RawThemeConfig>;

const themePresets = THEME_PRESETS as ThemePresetMap;

export { DEFAULT_THEME };

export function listThemes(): string[] {
  return Object.keys(themePresets).sort((a, b) => a.localeCompare(b));
}

export function getTheme(name: string | undefined | null): ThemeConfig {
  const resolvedName = (name == null || !String(name).trim()) ? DEFAULT_THEME : String(name).trim();
  const cached = THEME_CACHE.get(resolvedName);
  if (cached) {
    return {
      ...cached,
      packages: [...cached.packages],
      spacing: { ...cached.spacing },
    };
  }

  const raw = themePresets[resolvedName];
  if (!raw) {
    throw new Error(`Theme '${resolvedName}' not found. Available: ${listThemes().join(', ')}`);
  }

  const theme = toThemeConfig(raw, resolvedName);
  THEME_CACHE.set(resolvedName, theme);
  return {
    ...theme,
    packages: [...theme.packages],
    spacing: { ...theme.spacing },
  };
}
