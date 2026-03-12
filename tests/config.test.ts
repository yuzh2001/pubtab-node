import { describe, it, expect } from 'vitest';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';

import { loadConfig } from '../src/config.js';

async function mkTmpDir(prefix: string): Promise<string> {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

describe('loadConfig migration', () => {
  it('空配置返回空参数', async () => {
    const dir = await mkTmpDir('pubtab-ts-config-empty-');
    const cfg = path.join(dir, 'empty.yaml');
    await fs.writeFile(cfg, '', 'utf8');
    const [kwargs, formatter] = await loadConfig(cfg);
    expect(kwargs).toEqual({});
    expect(formatter).toBeNull();
  });

  it('非映射根抛出错误', async () => {
    const dir = await mkTmpDir('pubtab-ts-config-bad-');
    const cfg = path.join(dir, 'bad.yaml');
    await fs.writeFile(cfg, '- a\n- b\n', 'utf8');
    await expect(loadConfig(cfg)).rejects.toThrow(/mapping/i);
  });
});
