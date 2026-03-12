import { describe, it, expect } from 'vitest';
import path from 'node:path';
import fs from 'node:fs/promises';
import os from 'node:os';
import ExcelJS from 'exceljs';

import { xlsx2tex } from '../src/index.js';

async function mkTmpDir(prefix: string): Promise<string> {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

async function writeWorkbook(dir: string, filename: string): Promise<string> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S1');
  ws.getCell('A1').value = 'A';
  ws.getCell('B1').value = 'B';
  const out = path.join(dir, filename);
  await wb.xlsx.writeFile(out);
  return out;
}

describe('xlsx2tex theme', () => {
  it('默认 theme 为 three_line，默认会包含完整主题包提示', async () => {
    const dir = await mkTmpDir('pubtab-ts-theme-default-');
    const xlsxPath = await writeWorkbook(dir, 'input.xlsx');
    const outTex = path.join(dir, 'out.tex');

    const tex = await xlsx2tex(xlsxPath, outTex);

    expect(tex).toContain('% \\usepackage{booktabs}');
    expect(tex).toContain('% \\usepackage{multirow}');
    expect(tex).toContain('% \\usepackage[table]{xcolor}');
    expect(tex).toContain('% \\usepackage{diagbox}');
    expect(tex).toContain('% \\usepackage{makecell}');
    expect(tex).toContain('% \\usepackage{adjustbox}');
    expect(tex).toContain('% \\usepackage{amssymb}');
    expect(tex).toContain('% \\usepackage{pifont}');
  });

  it('theme simple 会只输出简化包提示', async () => {
    const dir = await mkTmpDir('pubtab-ts-theme-simple-');
    const xlsxPath = await writeWorkbook(dir, 'input.xlsx');
    const outTex = path.join(dir, 'out.tex');

    const tex = await xlsx2tex(xlsxPath, outTex, { theme: 'simple' });

    expect(tex).toContain('% \\usepackage{booktabs}');
    expect(tex).toContain('% \\usepackage{multirow}');
    expect(tex).toContain('% \\usepackage[table]{xcolor}');
    expect(tex).not.toContain('% \\usepackage{diagbox}');
    expect(tex).not.toContain('% \\usepackage{makecell}');
    expect(tex).not.toContain('% \\usepackage{adjustbox}');
    expect(tex).not.toContain('% \\usepackage{amssymb}');
    expect(tex).not.toContain('% \\usepackage{pifont}');
  });
});
