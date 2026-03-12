import type { Cell, RenderOptions, RichSegment, TableData } from './models.js';
import { hexToLatexColor, latexEscape } from './utils.js';
import { getTheme } from './themes.js';

const SUBSCRIPT_CHAR_MAP: Record<string, string> = {
  'α': '\\\\alpha',
  'β': '\\\\beta',
  'γ': '\\\\gamma',
  'δ': '\\\\delta',
  'ε': '\\\\epsilon',
  'θ': '\\\\theta',
  'λ': '\\\\lambda',
  'μ': '\\\\mu',
  'π': '\\\\pi',
  'σ': '\\\\sigma',
  'φ': '\\\\phi',
  'ω': '\\\\omega',
  'Γ': '\\\\Gamma',
  'Δ': '\\\\Delta',
  'Θ': '\\\\Theta',
  'Ω': '\\\\Omega',
};

function normalizeUnicodeSubscript(text: string): string {
  return text.replace(/_([^\s\\{])/gu, (_m, ch) => {
    const escaped = SUBSCRIPT_CHAR_MAP[ch] ?? ch;
    return `$_{${escaped}}`;
  });
}

function cellToLatex(cell: Cell): string {
  const style = cell.style ?? {};
  let text = style.rawLatex ? String(cell.value ?? '') : latexEscape(cell.value ?? '');
  if (!style.rawLatex) {
    text = normalizeUnicodeSubscript(text);
  }

  if (!style.rawLatex && cell.richSegments && cell.richSegments.length > 1) {
    const parts: string[] = [];
    for (const seg of cell.richSegments as RichSegment[]) {
      const [segText, segColor, segBold, segItalic, segUnderline] = seg;
      let s = latexEscape(segText);
      if (segBold) s = `\\textbf{${s}}`;
      if (segItalic) s = `\\textit{${s}}`;
      if (segUnderline) s = `\\underline{${s}}`;
      if (segColor) s = `\\textcolor[RGB]{${hexToLatexColor(segColor)}}{${s}}`;
      parts.push(s);
    }
    text = parts.join('');
  }

  if (!style.rawLatex && text.includes('\n')) {
    text = `\\makecell{${text.replace(/\n/g, '\\\\')}}`;
  }

  if (!style.rawLatex && style.diagbox && style.diagbox.length >= 2) {
    text = `\\diagbox{${latexEscape(style.diagbox[0])}}{${latexEscape(style.diagbox[1])}}`;
  }

  if (!style.rawLatex && !(cell.richSegments && cell.richSegments.length > 1)) {
    if (style.bold) text = `\\textbf{${text}}`;
    if (style.italic) text = `\\textit{${text}}`;
    if (style.underline) text = `\\underline{${text}}`;
    if (style.color) text = `\\textcolor[RGB]{${hexToLatexColor(style.color)}}{${text}}`;
    if (style.bgColor && cell.colspan <= 1) {
      text = `\\cellcolor[RGB]{${hexToLatexColor(style.bgColor)}}{${text}}`;
    }
  }

  if (!style.rawLatex && style.rotation) {
    if (cell.rowspan > 1) {
      // Match pubtab-python: omit origin=c for multirow to avoid bottomrule overflow.
      text = `\\rotatebox{${style.rotation}}{${text}}`;
    } else {
      text = `\\rotatebox[origin=c]{${style.rotation}}{${text}}`;
    }
  }

  if (cell.rowspan > 1) {
    text = `\\multirow{${cell.rowspan}}{*}{${text}}`;
  }

  if (cell.colspan > 1) {
    const align = style.alignment?.[0] ?? 'c';
    if (style.bgColor) {
      text = `\\multicolumn{${cell.colspan}}{>{\\columncolor[RGB]{${hexToLatexColor(style.bgColor)}}}${align}}{${text}}`;
    } else {
      text = `\\multicolumn{${cell.colspan}}{${align}}{${text}}`;
    }
  }

  return text;
}

function buildPackageHints(table: TableData, opts: RenderOptions): string {
  const theme = getTheme(opts.theme);
  const base = [...theme.packages];
  const needGraphicx = Boolean(opts.resizebox) || table.cells.some((r) => r.some((c) => (c.style.rotation ?? 0) !== 0));
  if (needGraphicx) base.push('graphicx');
  const seen = new Set<string>();
  const unique: string[] = [];
  for (const pkg of base) {
    if (!seen.has(pkg)) {
      seen.add(pkg);
      unique.push(pkg);
    }
  }
  const packages = unique;

  const lines = ['% Theme package hints for this table (add in your preamble):'];
  for (const pkg of packages) {
    if (pkg === 'xcolor') {
      lines.push('% \\usepackage[table]{xcolor}');
    } else {
      lines.push(`% \\usepackage{${pkg}}`);
    }
  }
  lines.push('');
  return lines.join('\n');
}

function cellHasPayload(cell: Cell): boolean {
  return (cell.value !== '' && cell.value !== null && cell.value !== undefined)
    || (cell.richSegments?.length ?? 0) > 0
    || (cell.style.diagbox?.length ?? 0) > 0;
}

function rowIsSection(row: Cell[], numCols: number): boolean {
  if (!row.length) return false;
  const first = row[0];
  if (first.colspan >= numCols) return true;
  if (first.colspan >= numCols - 1 && row.slice(1).every((cell) => !cellHasPayload(cell))) {
    return true;
  }

  const payloadIdx = row.reduce<number[]>((acc, cell, idx) => {
    if (cellHasPayload(cell)) acc.push(idx);
    return acc;
  }, []);

  if (payloadIdx.length === 1) {
    const idx = payloadIdx[0];
    const sec = row[idx];
    return sec.colspan > 1 && idx + sec.colspan >= (numCols - 1);
  }

  if (payloadIdx.length === 2) {
    const [a, b] = payloadIdx;
    for (const [labelIdx, secIdx] of [[a, b], [b, a]] as Array<[number, number]>) {
      const label = row[labelIdx];
      const sec = row[secIdx];
      if (label.rowspan > 1 && sec.colspan > 1 && (secIdx + sec.colspan) >= (numCols - 1)) {
        return true;
      }
    }
  }

  return false;
}

function hasActiveFirstColMultirow(bodyCells: Cell[][], bodyRowIdx: number): boolean {
  for (let r = 0; r < bodyRowIdx; r += 1) {
    const row = bodyCells[r];
    if (!row[0]) continue;
    const cell = row[0];
    if (cell.rowspan > 1 && (r + cell.rowspan) > bodyRowIdx) return true;
  }

  if (bodyRowIdx < bodyCells.length && bodyCells[bodyRowIdx][0]?.rowspan > 1) return true;
  return false;
}

function sectionSeparator(bodyCells: Cell[][], bodyRowIdx: number, numCols: number): string {
  if (numCols > 1 && hasActiveFirstColMultirow(bodyCells, bodyRowIdx)) {
    return `\\cmidrule(lr){2-${numCols}}`;
  }
  return '\\midrule';
}

function normalizeGroupSeparators(
  gs: TableData['groupSeparators'],
): Record<number, string | string[]> {
  if (Array.isArray(gs)) {
    const out: Record<number, string> = {};
    for (const idx of gs) out[idx] = '\\midrule';
    return out;
  }
  return gs || {};
}

function renderCells(row: Cell[]): string {
  const rendered: string[] = [];
  for (let c = 0; c < row.length; ) {
    const cell = row[c];
    rendered.push(cellToLatex(cell));
    c += Math.max(1, cell.colspan || 1);
  }
  return `${rendered.join(' & ')} \\\\`;
}

export function render(table: TableData, opts: RenderOptions = {}): string {
  const theme = getTheme(opts.theme);
  const colSpec = opts.colSpec ?? 'c'.repeat(Math.max(1, table.numCols));
  const lines: string[] = [];
  lines.push(buildPackageHints(table, opts));

  const env = opts.spanColumns ? 'table*' : 'table';
  lines.push(`\\begin{${env}}[${opts.position ?? 'htbp'}]`);
  lines.push('\\centering');
  lines.push('\\setlength{\\heavyrulewidth}{1.0pt}');
  const captionPosition = theme.captionPosition ?? 'top';
  if (opts.caption && captionPosition === 'top') {
    lines.push(`\\caption{${latexEscape(opts.caption)}}`);
  }
  if (opts.label) lines.push(`\\label{${latexEscape(opts.label)}}`);
  lines.push(`\\begin{tabular}{${colSpec}}`);
  lines.push('\\toprule');

  const bodyCells = table.cells.slice(table.headerRows);
  const groupSeparators = normalizeGroupSeparators(table.groupSeparators);
  const bodyRowsWithSep: string[] = [];

  for (let i = 0; i < bodyCells.length; i += 1) {
    const row = bodyCells[i];
    const isSection = rowIsSection(row, table.numCols);
    const sep = isSection ? sectionSeparator(bodyCells, i, table.numCols) : '\\midrule';

    if (isSection && i > 0 && (bodyRowsWithSep.length === 0 || bodyRowsWithSep[bodyRowsWithSep.length - 1] !== sep)) {
      bodyRowsWithSep.push(sep);
    }

    bodyRowsWithSep.push(renderCells(row));

    const absoluteIdx = table.headerRows + i;
    const hasCustomSeparator = Object.prototype.hasOwnProperty.call(groupSeparators, absoluteIdx);
    if (isSection && !hasCustomSeparator) {
      bodyRowsWithSep.push(sep);
    }

    if (hasCustomSeparator) {
      const custom = groupSeparators[absoluteIdx];
      if (isSection && custom === '\\midrule') {
        bodyRowsWithSep.push(sep);
      } else if (Array.isArray(custom)) {
        bodyRowsWithSep.push(...custom);
      } else {
        bodyRowsWithSep.push(custom);
      }
    }
  }

  for (let r = 0; r < table.cells.length; r += 1) {
    const row = table.cells[r];
    if (r < table.headerRows) {
      lines.push(renderCells(row));
      if (r === table.headerRows - 1) lines.push('\\midrule');
      continue;
    }
    if (r === table.headerRows) {
      lines.push(...bodyRowsWithSep);
    }
    break;
  }

  lines.push('\\bottomrule');
  lines.push('\\end{tabular}');
  if (opts.resizebox) {
    lines.push(`% resizebox hint: \\resizebox{${opts.resizebox}}{!}{...}`);
  }
  if (opts.caption && captionPosition === 'bottom') {
    lines.push(`\\caption{${latexEscape(opts.caption)}}`);
  }
  lines.push(`\\end{${env}}`);
  lines.push('');
  return lines.join('\n');
}
