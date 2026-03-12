import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';

import { texToTableResult, xlsxToTableResult } from '../../src/browser.js';
import { browserFixtures, loadFixtureXlsx } from './fixtures.js';

describe('browser fixture: table result', () => {
  it('xlsx 结果包含 TanStack 列和无 placeholder 的 header/body 渲染行', async () => {
    const buffer = await loadFixtureXlsx(browserFixtures.table4XlsxPath);
    const result = await xlsxToTableResult(buffer, { headerRows: 'auto' });

    expect(result.columns.length).toBe(result.table.numCols);
    expect(result.leafColumnIds.length).toBe(result.table.numCols);
    expect(result.data.length).toBe(result.table.numRows - result.table.headerRows);
    expect(result.headerRows.length).toBe(result.table.headerRows);
    expect(result.headerRows[0]?.cells[0]?.colSpan).toBeGreaterThanOrEqual(1);
  });

  it('tex 结果保留 header span 和 body cell 原点信息', async () => {
    const result = await texToTableResult(browserFixtures.table1Tex);
    const firstHeader = result.headerRows[0]?.cells[0];
    const firstBodyCell = result.bodyRows[0]?.cells[0];

    expect(firstHeader).toBeTruthy();
    expect(typeof firstHeader?.text).toBe('string');
    expect(firstHeader?.rowSpan).toBeGreaterThanOrEqual(1);
    expect(firstBodyCell).toBeTruthy();
    expect(firstBodyCell?.originRowIndex).toBeGreaterThanOrEqual(result.table.headerRows);
    expect(firstBodyCell?.originColIndex).toBeGreaterThanOrEqual(0);
  });

  it('xlsx auto headerRows 会把横向分组下的子标题保留在 headerRows 中', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S1');

    ws.getCell('A1').value = 'Backbones';
    ws.getCell('B1').value = 'Methods';
    ws.getCell('C1').value = 'Rubric Grading';
    ws.getCell('G1').value = 'Verifiability';
    ws.getCell('I1').value = 'ROUGE-1';

    ws.mergeCells('A1:A2');
    ws.mergeCells('B1:B2');
    ws.mergeCells('C1:F1');
    ws.mergeCells('G1:H1');
    ws.mergeCells('I1:I2');

    ws.getCell('C2').value = 'Relevance';
    ws.getCell('D2').value = 'Breadth';
    ws.getCell('E2').value = 'Interest Level';
    ws.getCell('F2').value = 'Organization';
    ws.getCell('G2').value = 'CR';
    ws.getCell('H2').value = 'CP';

    ws.getCell('A3').value = 'Qwen-2.5';
    ws.mergeCells('A3:A7');
    ws.getCell('B3').value = 'oRAG';
    ws.getCell('C3').value = '3.91';
    ws.getCell('D3').value = '4.08';
    ws.getCell('E3').value = '3.96';
    ws.getCell('F3').value = '3.63';
    ws.getCell('G3').value = '00.00';
    ws.getCell('H3').value = '00.00';
    ws.getCell('I3').value = '00.00';

    const raw = await wb.xlsx.writeBuffer();
    const buffer = raw instanceof ArrayBuffer
      ? raw
      : raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength);

    const result = await xlsxToTableResult(buffer, { headerRows: 'auto' });

    expect(result.headerRows).toHaveLength(2);
    expect(result.headerRows[1]?.cells.map((cell) => cell.text)).toEqual([
      'Relevance',
      'Breadth',
      'Interest Level',
      'Organization',
      'CR',
      'CP',
    ]);
    expect(result.bodyRows[0]?.cells[0]?.text).toBe('Qwen-2.5');
    expect(result.bodyRows[0]?.cells[0]?.rowSpan).toBe(5);
  });
});
