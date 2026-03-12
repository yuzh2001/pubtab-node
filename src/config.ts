import fs from 'node:fs/promises';

export type ConfigRecord = Record<string, unknown>;

const KEY_VALUE_ERROR = '配置解析失败';
const ROOT_TYPE_ERROR = 'Config YAML root must be a mapping';

function stripYamlComment(line: string): string {
  let inSingle = false;
  let inDouble = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '\'' && !inDouble) {
      inSingle = !inSingle;
      continue;
    }
    if (ch === '"' && !inSingle) {
      inDouble = !inDouble;
      continue;
    }
    if (ch === '#' && !inSingle && !inDouble) return line.slice(0, i);
  }
  return line;
}

function unquoteYamlValue(raw: string): string {
  const s = raw.trim();
  if ((s.startsWith('"') && s.endsWith('"')) || (s.startsWith('\'') && s.endsWith('\''))) {
    return s.slice(1, -1);
  }
  return s;
}

function parseYamlValue(raw: string): unknown {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  const unquoted = unquoteYamlValue(trimmed);
  if (unquoted === '') return '';
  if (/^(~|null)$/iu.test(unquoted)) return null;
  if (/^(true|false)$/iu.test(unquoted)) return unquoted.toLowerCase() === 'true';
  if (/^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/u.test(unquoted)) return Number(unquoted);
  return unquoted;
}

export function parseYamlSimple(content: string): ConfigRecord {
  const result: ConfigRecord = {};
  const lines = content.split(/\r?\n/gu);

  for (const line of lines) {
    const trimmedLine = stripYamlComment(line).trim();
    if (!trimmedLine) continue;
    const idx = trimmedLine.indexOf(':');
    if (idx < 0) throw new Error(`${KEY_VALUE_ERROR}：非法行 ${trimmedLine}`);

    const key = trimmedLine.slice(0, idx).trim();
    const value = trimmedLine.slice(idx + 1).trim();
    if (!key) throw new Error(`${KEY_VALUE_ERROR}：非法键 ${trimmedLine}`);

    result[key] = parseYamlValue(value);
  }

  return result;
}

function isMappingValue(parsed: unknown): parsed is ConfigRecord {
  return parsed != null
    && typeof parsed === 'object'
    && !Array.isArray(parsed);
}

export async function loadConfig(path: string): Promise<[ConfigRecord, null]> {
  const raw = await fs.readFile(path, 'utf8');
  const trimmed = raw.trim();
  if (!trimmed) {
    return [{}, null];
  }

  if (/^\s*-/m.test(raw)) {
    throw new Error(`${ROOT_TYPE_ERROR}; mapping object required.`);
  }

  const parsed: unknown = parseYamlSimple(raw);
  if (!isMappingValue(parsed)) {
    throw new Error(`${ROOT_TYPE_ERROR}; mapping object required.`);
  }
  return [parsed, null];
}
