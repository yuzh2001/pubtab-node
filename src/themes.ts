import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import YAML from 'yaml';

import type { RawThemeConfig, ThemeConfig } from './theme-schema.js';
import { toThemeConfig } from './theme-schema.js';

const DEFAULT_THEME = 'three_line';
const THEMES_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'themes');
const THEME_CACHE = new Map<string, ThemeConfig>();

function loadThemeConfig(name: string): ThemeConfig {
  const configPath = path.join(THEMES_DIR, name, 'config.yaml');
  if (!fs.existsSync(configPath)) {
    throw new Error(`Theme '${name}' not found. Available: ${listThemes().join(', ')}`);
  }

  const raw = YAML.parse(fs.readFileSync(configPath, 'utf8')) as RawThemeConfig | null;
  return toThemeConfig(raw, name);
}

export { DEFAULT_THEME };

export function listThemes(): string[] {
  if (!fs.existsSync(THEMES_DIR)) return [];
  return fs.readdirSync(THEMES_DIR, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && fs.existsSync(path.join(THEMES_DIR, entry.name, 'config.yaml')))
    .map((entry) => entry.name)
    .sort((a, b) => a.localeCompare(b));
}

export function getTheme(name: string | undefined | null): ThemeConfig {
  const resolvedName = (name == null || !String(name).trim()) ? DEFAULT_THEME : String(name).trim();
  const cached = THEME_CACHE.get(resolvedName);
  const theme = cached ?? loadThemeConfig(resolvedName);
  if (!cached) THEME_CACHE.set(resolvedName, theme);
  return {
    ...theme,
    packages: [...theme.packages],
    spacing: { ...theme.spacing },
  };
}
