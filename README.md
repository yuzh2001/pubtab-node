# pubtab-node

![Node >=18](https://img.shields.io/badge/node-%3E%3D18-2f6f3e)
![License ISC](https://img.shields.io/badge/license-ISC-1f4b99)
![TypeScript](https://img.shields.io/badge/language-TypeScript-3178c6)

`pubtab-node` 是一个面向 Node.js / TypeScript 的表格转换工具，用来在 Excel 表格和 LaTeX `tabular` 之间做双向转换。

适合把论文表格、报告表格或标注表格接入 Node.js 工作流，用 CLI 批处理，或者在代码里直接调用。

当前版本优先提供一套可用、可测、可发布的核心能力：

- `.xlsx -> .tex`
- `.tex -> .xlsx`
- 基础 CLI
- 可编程 API
- 已覆盖一批 round-trip 与兼容性测试

它的目标不是在第一版就完全复刻原版 `pubtab` 的全部行为，而是先把最常用、最稳定的转换链路落下来，再逐步补齐与原版的差距。

## 适用场景

- 需要把 `.xlsx` 批量转成 LaTeX 表格
- 需要把已有的 LaTeX `tabular` 回写成 `.xlsx`
- 需要在 Node.js 脚本或服务里嵌入转换能力
- 希望通过配置文件或 CLI 参数稳定复现输出

## 暂不适用

- 依赖完整 preview 管线
- 需要与 Python 原版在排版细节上完全一致
- 需要处理大量异常、破损或高度定制的 TeX 输入
- 需要 `.xls` 老格式支持

## 安装

### 全局安装 CLI

```bash
npm i -g pubtab-node
```

安装后可直接使用：

```bash
pubtab-node --help
```

### 作为库安装

```bash
npm i pubtab-node
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

```ts
import { xlsx2tex, texToExcel, readTex, render } from 'pubtab-node';

await xlsx2tex('table.xlsx', 'out/table.tex');
await texToExcel('table.tex', 'out/table.xlsx');

const table = readTex('\\begin{tabular}{cc}A & B\\\\\\end{tabular}');
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

如果你的目标是：

- 在 Node.js 环境里完成 Excel / LaTeX 的基础双向转换
- 通过 CLI 或 API 集成到自己的流程里
- 依赖一套相对轻量、可测试、可扩展的实现

那么当前版本已经适合使用。

如果你的目标是：

- 完整替代原版 `pubtab`
- 依赖 preview、复杂主题、极强的 TeX 容错
- 要求输出格式与原版几乎逐字符一致

那么目前还不适合直接视为完全替代品。

## 测试与开发

开发环境：

```bash
npm i
npm test
npm run build
```

当前仓库包含 fixtures、round-trip、CLI 配置与兼容性测试；发布包本身只包含 `dist`、`README.md` 和 `package.json`，不会带上测试目录。

## 仓库说明

- 主要实现位于 `src/`
- 测试位于 `tests/`
- `tests/fixtures` 包含用于 round-trip 对比的样例
- `./.pubtab-python` 是本地参考用的 Python 原版副本，不会进入 npm 发布包

## 致谢

本项目的设计目标、行为对齐方向和部分测试迁移工作，都直接受益于 Python 原版 `pubtab`。

感谢原版项目提供了清晰的问题域抽象、可参考的行为定义，以及大量有价值的测试样例，使 `pubtab-node` 能够在 TypeScript 生态里继续推进这条转换链路。
