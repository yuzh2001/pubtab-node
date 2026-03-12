# pubtab-node

![Node >=18](https://img.shields.io/badge/node-%3E%3D18-2f6f3e)
![License ISC](https://img.shields.io/badge/license-ISC-1f4b99)
![TypeScript](https://img.shields.io/badge/language-TypeScript-3178c6)

[`pubtab`](https://github.com/Galaxy-Dawn/pubtab) is a feature-rich, well-tested tool for converting between LaTeX tables and Excel spreadsheets.

`pubtab-node` is a TypeScript port of the original Python `pubtab`, providing two-way conversion between Excel workbooks and LaTeX `tabular`.

The Node version removes the preview feature from the Python version, while keeping most core capabilities:

- `.xlsx -> .tex`
- `.tex -> .xlsx`
- basic CLI
- original `three_line` theme support
- programmable API

## Installation

### Install the CLI globally

```bash
pnpm add -g pubtab-node
```

After installation:

```bash
pubtab-node --help
```

### Install as a library

```bash
pnpm add pubtab-node
```

## CLI Usage

### Excel to LaTeX

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex
```

Export by sheet:

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex --sheet 0
pubtab-node xlsx2tex table.xlsx out/table.tex --sheet Sheet1
```

Add caption / label / position:

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex \
  --caption "My Table" \
  --label "tab:my_table" \
  --position htbp
```

### LaTeX to Excel

```bash
pubtab-node tex2xlsx table.tex out/table.xlsx
```

### Use a YAML config file

```yaml
sheet: 1
caption: Example Table
label: tab:example
position: htbp
theme: three_line
headerRows: auto
spacing:
  tabcolsep: 2.4pt
  arraystretch: 1.0
```

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex --config pubtab.yml
```

Explicit CLI flags override fields with the same name in the config file.

The config loader accepts both camelCase and the original Python-style snake_case keys, for example:

- `headerRows` / `header_rows`
- `fontSize` / `font_size`
- `colSpec` / `col_spec`
- `headerSep` / `header_sep`
- `spanColumns` / `span_columns`

### CLI help

```bash
pubtab-node --help
```

Currently supported top-level commands:

- `pubtab-node xlsx2tex`
- `pubtab-node tex2xlsx`

## API Usage

It is recommended to wrap LaTeX strings with `String.raw` to avoid escape issues.

```ts
// Backslashes inside String.raw do not need to be escaped.
const noNeedToUseEscapes = String.raw`\begin{something}`;
const latexString = String.raw`
    \begin{tabular}{cc}
    A & B\\
    \end{tabular}\
`;

// Not recommended: backslashes must be escaped manually.
const latexStringNotRecommended =
  '\\begin{tabular}{cc}A & B\\\\ \\end{tabular}';
```

```ts
import { xlsx2tex, texToExcel, readTex, render } from 'pubtab-node';

await xlsx2tex('table.xlsx', 'out/table.tex');
await texToExcel('table.tex', 'out/table.xlsx');

const table = readTex(String.raw`\begin{tabular}{cc}A & B\\ \end{tabular}`);
const tex = render(table);
```

Main exported APIs:

- `xlsx2tex(input, output, options?)`
- `texToExcel(input, output)`
- `render(table, options?)`
- `readTex(tex)`
- `readTexAll(tex)`

## Current Capabilities

### Supported

- read `.xlsx` and output `.tex`
- read `.tex` and output `.xlsx`
- load the original `three_line` theme config from `themes/three_line/config.yaml`
- export all worksheets when `sheet` is not specified
- export a single sheet when requested
- batch conversion for directory inputs
- `output` supports both file paths and directory paths
- preserve the basic semantics of merged cells
- automatic or explicit `headerRows`
- `caption`, `label`, `position`, `resizebox`, `colSpec`
- core style round-trip for bold, italic, underline, text color, background color, rotation, and rich text runs

### Not included yet

- preview pipeline
- `.xls` input
- full parity for every original edge-case LaTeX layout behavior
- stronger TeX fault tolerance
- full support for expanding `definecolor` / `newcommand` macros
- writing multiple tables from a single `.tex` file into multiple sheets

## Gaps Compared with the Original `pubtab`

This project is a TypeScript port of the original Python `pubtab`, but it is currently "core usable" rather than fully aligned.

The main remaining gaps are:

- no `preview` support yet
- `.xls` is not supported yet
- a few edge-case LaTeX layout behaviors are still being aligned with the original implementation
- tolerance for malformed or unusual TeX input is still being improved
- some original test scenarios have not been migrated one by one yet

## Testing and Development

Development environment:

```bash
git clone https://github.com/Galaxy-Dawn/pubtab .pubtab-python # local reference clone of the original Python project; not included in the npm package
pnpm i
pnpm test
pnpm build
```

This repository is primarily managed with `pnpm`. Prefer `pnpm i`, `pnpm test`, and `pnpm build` for local development and release verification.

This repository includes fixtures, round-trip tests, CLI config tests, and compatibility tests. The published package ships `dist/` and `themes/`, but does not ship the test directories.

## Repository Notes

- main implementation: `src/`
- runtime theme configs: `themes/`
- tests: `tests/`
- `tests/fixtures` contains samples used for round-trip comparisons
- `./.pubtab-python` is a local reference clone of the original Python project and is not included in the npm package

## Acknowledgements

The design goals, behavior-alignment direction, and part of the test migration work in this project all directly benefit from the original Python `pubtab`.

Thanks to the original project for providing a clear abstraction of the problem domain, a useful behavioral reference, and many valuable test fixtures that make it possible to keep pushing this conversion workflow forward in the TypeScript ecosystem.
