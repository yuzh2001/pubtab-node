import { describe, expect, it } from 'vitest';
import { spawn } from 'node:child_process';
import path from 'node:path';

function run(command: string, args: string[], cwd: string): Promise<{ code: number | null; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(command, args, {
      cwd,
      stdio: ['ignore', 'ignore', 'pipe'],
    });

    let stderr = '';
    child.stderr.on('data', (chunk) => {
      stderr += String(chunk);
    });

    child.on('close', (code) => {
      resolve({ code, stderr });
    });
  });
}

describe('playground build', () => {
  it('playground 可以完成生产构建', async () => {
    const repoRoot = path.resolve(import.meta.dirname, '..', '..');
    const result = await run('pnpm', ['build:playground'], repoRoot);

    expect(result.code, result.stderr).toBe(0);
  });
});
