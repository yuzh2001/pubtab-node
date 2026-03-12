import { readFile } from 'node:fs/promises';

import table1Tex from '../../tests/fixtures/table1.tex?raw';
import table4Tex from '../../tests/fixtures/table4.tex?raw';

export async function loadFixtureXlsx(relativePath: string): Promise<ArrayBuffer> {
  const bytes = await readFile(new URL(relativePath, import.meta.url));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

export const playgroundFixtures = {
  table1Tex,
  table4Tex,
  table1XlsxPath: '../../tests/fixtures/table1.xlsx',
  table4XlsxPath: '../../tests/fixtures/table4.xlsx',
};
