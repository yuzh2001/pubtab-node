# pubtab-node

![Node >=18](https://img.shields.io/badge/node-%3E%3D18-2f6f3e)
![License ISC](https://img.shields.io/badge/license-ISC-1f4b99)
![TypeScript](https://img.shields.io/badge/language-TypeScript-3178c6)

`pubtab` https://github.com/Galaxy-Dawn/pubtab 是一个功能丰富、测试完整的LaTeX表格与excel双向转换工具。

`pubtab-node` 是 Python 原版 `pubtab` 的 TypeScript 复刻，用来在 Excel 表格和 LaTeX `tabular` 之间做双向转换。

node版本去除了python版本中的preview功能，支持绝大部分的核心能力：

- `.xlsx -> .tex`
- `.tex -> .xlsx`
- 基础 CLI
- theme
- 可编程 API

## 安装

### 全局安装 CLI

```bash
pnpm add -g pubtab-node
```

安装后可直接使用：

```bash
pubtab-node --help
```

### 作为库安装

```bash
pnpm add pubtab-node
```

## CLI 使用

### Excel 转 LaTeX

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex
```

按 sheet 导出：

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex --sheet 0
pubtab-node xlsx2tex table.xlsx out/table.tex --sheet Sheet1
```

补充 caption / label / position：

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex \
  --caption "My Table" \
  --label "tab:my_table" \
  --position htbp
```

### LaTeX 转 Excel

```bash
pubtab-node tex2xlsx table.tex out/table.xlsx
```

### 使用 YAML 配置

```yaml
sheet: 1
caption: Example Table
label: tab:example
position: htbp
theme: three_line
headerRows: auto
```

```bash
pubtab-node xlsx2tex table.xlsx out/table.tex --config pubtab.yml
```

命令行显式参数会覆盖配置文件中的同名项。

### CLI 帮助

```bash
pubtab-node --help
```

当前支持的主命令：

- `pubtab-node xlsx2tex`
- `pubtab-node tex2xlsx`

## API 使用

推荐使用 String.raw 语法来包裹LaTeX代码，避免转义问题。

```ts
// String.raw内的反斜杠不需要转义
const noNeedToUseEscapes = String.raw`\begin{something}`;
const latexString = String.raw`
    \begin{tabular}{cc}
    A & B\\
    \end{tabular}\
`;

// 不推荐，需要手动转义反斜杠
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

当前导出的主要 API：

- `xlsx2tex(input, output, options?)`
- `texToExcel(input, output)`
- `render(table, options?)`
- `readTex(tex)`
- `readTexAll(tex)`

## 当前能力

### 已支持

- 读取 `.xlsx` 并输出 `.tex`
- 读取 `.tex` 并输出 `.xlsx`
- 未指定 `sheet` 时导出全部工作表
- 指定单个 sheet 导出
- 目录输入批量转换
- `output` 同时支持文件路径和目录路径
- 合并单元格的基础语义保持
- 自动或显式 `headerRows`
- `caption`、`label`、`position`、`resizebox`、`colSpec`
- 样式的核心 round-trip：粗体、斜体、下划线、文字颜色、背景色、旋转、富文本片段

### 当前不包含

- preview 管线
- `.xls` 输入
- 原版完整主题系统与全部渲染细节
- 更强的 TeX 容错解析
- `definecolor` / `newcommand` 宏展开的完整支持
- 单个 `.tex` 多表写入多 sheet

## 与原版 pubtab 的差距

这个项目是对 Python 原版 `pubtab` 的 TypeScript 复刻，但当前仍然是“核心可用”而不是“完全对齐”。

现阶段最主要的差距有：

- 还没有 `preview` 能力
- `.xls` 暂未支持
- 主题系统和排版细节还没有完全对齐原版
- 对异常 TeX 输入的容错能力仍在继续补齐
- 部分原版测试场景还没有逐项迁移

## 测试与开发

开发环境：

```bash
git clone https://github.com/Galaxy-Dawn/pubtab .pubtab-python # 本地参考用的 Python 原版副本，不会进入 npm 发布包
pnpm i
pnpm test
pnpm build
```

这个仓库现在主要使用 `pnpm` 管理，日常开发和发布前校验也应优先使用 `pnpm i`、`pnpm test`、`pnpm build`。

当前仓库包含 fixtures、round-trip、CLI 配置与兼容性测试；发布包本身只包含 `dist`、`README.md` 和 `package.json`，不会带上测试目录。

## 仓库说明

- 主要实现位于 `src/`
- 测试位于 `tests/`
- `tests/fixtures` 包含用于 round-trip 对比的样例
- `./.pubtab-python` 是本地参考用的 Python 原版副本，不会进入 npm 发布包

## 致谢

本项目的设计目标、行为对齐方向和部分测试迁移工作，都直接受益于 Python 原版 `pubtab`。

感谢原版项目提供了清晰的问题域抽象、可参考的行为定义，以及大量有价值的测试样例，使 `pubtab-node` 能够在 TypeScript 生态里继续推进这条转换链路。
