import type {
  Cell,
  TableColumnSchema,
  TableData,
  TableDataRow,
  TableRenderCell,
  TableRenderRow,
  TableResult,
  TableViewModel,
} from '../models.js';

function cellText(cell: Cell): string {
  if (cell.richSegments && cell.richSegments.length > 0) {
    return cell.richSegments.map((seg) => seg[0]).join('');
  }
  if (cell.style.diagbox && cell.style.diagbox.length >= 2) {
    return `${cell.style.diagbox[0]} / ${cell.style.diagbox[1]}`;
  }
  return String(cell.value ?? '');
}

function buildCoverageMap(table: TableData): Map<string, { row: number; col: number }> {
  const covered = new Map<string, { row: number; col: number }>();
  for (let rowIndex = 0; rowIndex < table.cells.length; rowIndex += 1) {
    const row = table.cells[rowIndex];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cell = row[colIndex];
      const rowSpan = Math.max(1, cell.rowspan || 1);
      const colSpan = Math.max(1, cell.colspan || 1);
      if (rowSpan === 1 && colSpan === 1) continue;
      for (let r = rowIndex; r < rowIndex + rowSpan; r += 1) {
        for (let c = colIndex; c < colIndex + colSpan; c += 1) {
          if (r === rowIndex && c === colIndex) continue;
          covered.set(`${r},${c}`, { row: rowIndex, col: colIndex });
        }
      }
    }
  }
  return covered;
}

function leafColumnId(index: number): string {
  return `col_${index}`;
}

function inferLeafHeader(table: TableData, coverage: Map<string, { row: number; col: number }>, colIndex: number): string {
  for (let rowIndex = table.headerRows - 1; rowIndex >= 0; rowIndex -= 1) {
    const coveredBy = coverage.get(`${rowIndex},${colIndex}`);
    const sourceRowIndex = coveredBy?.row ?? rowIndex;
    const sourceColIndex = coveredBy?.col ?? colIndex;
    const cell = table.cells[sourceRowIndex]?.[sourceColIndex];
    if (!cell) continue;
    const text = cellText(cell).trim();
    if (text) return text;
  }
  return `C${colIndex + 1}`;
}

function toRenderRow(
  table: TableData,
  coverage: Map<string, { row: number; col: number }>,
  rowIndex: number,
  section: 'header' | 'body',
): TableRenderRow {
  const rawRow = table.cells[rowIndex] ?? [];
  const cells: TableRenderCell[] = [];
  const values: unknown[] = [];

  for (let colIndex = 0; colIndex < table.numCols; colIndex += 1) {
    const raw = rawRow[colIndex] ?? { value: '', style: {}, rowspan: 1, colspan: 1, richSegments: null };
    values.push(raw.value);

    if (coverage.has(`${rowIndex},${colIndex}`)) continue;

    cells.push({
      id: `r${rowIndex}c${colIndex}`,
      rowIndex,
      colIndex,
      columnId: leafColumnId(colIndex),
      value: raw.value,
      text: cellText(raw),
      style: raw.style,
      richSegments: raw.richSegments ?? null,
      rowSpan: Math.max(1, raw.rowspan || 1),
      colSpan: Math.max(1, raw.colspan || 1),
      originRowIndex: rowIndex,
      originColIndex: colIndex,
      section,
    });
  }

  return {
    id: `row_${rowIndex}`,
    index: rowIndex,
    section,
    cells,
    values,
  };
}

function toDataRow(row: TableRenderRow, leafColumnIds: string[]): TableDataRow {
  const values: Record<string, unknown> = {};
  for (let index = 0; index < leafColumnIds.length; index += 1) {
    values[leafColumnIds[index]] = row.values[index] ?? '';
  }

  return {
    id: row.id,
    rowIndex: row.index,
    values,
    cells: row.cells,
  };
}

export function tableToViewModel(table: TableData): TableViewModel {
  const coverage = buildCoverageMap(table);
  const leafColumnIds = Array.from({ length: table.numCols }, (_, index) => leafColumnId(index));
  const columns: TableColumnSchema[] = leafColumnIds.map((id, index) => ({
    id,
    accessorKey: id,
    header: inferLeafHeader(table, coverage, index),
    index,
  }));

  const headerRows: TableRenderRow[] = [];
  const bodyRows: TableRenderRow[] = [];
  for (let rowIndex = 0; rowIndex < table.numRows; rowIndex += 1) {
    const section = rowIndex < table.headerRows ? 'header' : 'body';
    const row = toRenderRow(table, coverage, rowIndex, section);
    if (section === 'header') headerRows.push(row);
    else bodyRows.push(row);
  }

  return {
    columns,
    leafColumnIds,
    data: bodyRows.map((row) => toDataRow(row, leafColumnIds)),
    headerRows,
    bodyRows,
    headerDepth: headerRows.length,
    size: {
      rows: table.numRows,
      cols: table.numCols,
      bodyRows: bodyRows.length,
    },
  };
}

export function tableToResult(table: TableData): TableResult {
  return {
    ...tableToViewModel(table),
    table,
  };
}
