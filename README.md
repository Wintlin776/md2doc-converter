# md2doc

把 **Markdown 转成 PDF 或 Word（.docx）** 的本地命令行工具，专为中文、数学公式与化学式排版优化。从剪贴板读取 Markdown，一条命令生成排版精良的文档并保存到桌面。

A local CLI tool that converts **Markdown to PDF or Word (.docx)**, tuned for Chinese typesetting, math formulas, and chemistry. It reads Markdown from your clipboard and produces a nicely typeset document on your Desktop with a single command.

---

## ✨ 最新亮点 / Latest Highlights

- **三线表 / Three-line tables** —— 表格统一渲染为学术三线表（无竖线），更专业。
  Tables render as clean academic three-line tables (no vertical rules).
- **跨页不丢数据 / Page-break friendly** —— 长表格自动跨页、表头每页重复，内容自动换行不被裁剪。
  Long tables flow across pages with repeated headers; content wraps instead of being clipped.
- **货币 vs 公式 / Currency vs. math** —— 自动区分 `$120` 货币与 `$x^2$` 公式，不再因美元符号破坏表格。
  Distinguishes currency like `$120` from math like `$x^2$`, so dollar signs no longer break tables.
- **去 AI 寒暄 / Strip AI preamble** —— 检测首条分隔线前的 AI 寒暄开场白，询问后一键删除。
  Detects an AI greeting before the first horizontal rule and offers to remove it.

---

## 🧩 核心功能 / Core Features

- **中文友好 / Chinese-friendly**：PDF 可选苹方现代风或 Noto 学术衬线风；Word 用宋体 + Times New Roman。
  PDF offers a modern (PingFang) or academic serif (Noto) theme; Word uses SimSun + Times New Roman.
- **数学公式 / Math**：支持行内 `$...$` 与块级 `$$...$$`；PDF 用 MathJax，Word 转原生可编辑 OMML。
  Inline `$...$` and block `$$...$$`; MathJax for PDF, native editable OMML for Word.
- **化学式 / Chemistry**：支持 mhchem 的 `\ce{}` 语法。 Supports mhchem `\ce{}` syntax.
- **代码高亮 / Code highlighting**：围栏代码块带样式渲染。 Styled fenced code blocks.

---

## 📦 安装 / Installation

```bash
pip install markdown playwright python-docx latex2mathml lxml
playwright install chromium

# 可选：离线渲染 PDF 公式 / Optional: offline PDF formula rendering
npm install -g mathjax
```

---

## 🚀 使用 / Usage

1. 复制一段 Markdown 到剪贴板。 Copy some Markdown to your clipboard.
2. 运行 / Run:
   ```bash
   python3 md2doc.py
   ```
3. 按提示选择格式、字体主题、文件名。 Follow the prompts: format, font theme, filename.
4. 文档保存到桌面并自动定位。 The document is saved to your Desktop and revealed automatically.

---

## 🖥️ 环境要求 / Requirements

- Python 3.10+
- macOS（`pbpaste`）或 / or Windows（`tkinter`）
- PDF 生成需要 Playwright Chromium / Playwright Chromium for PDF

---

## 📄 许可 / License

个人使用，欢迎自取自改。 Personal use — feel free to copy and modify.
