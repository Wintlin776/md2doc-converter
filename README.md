# md2doc

把 **Markdown 转成 PDF 或 Word（.docx）** 的本地命令行工具，专为中文、数学公式与化学式排版优化。

从剪贴板读取 Markdown，一条命令生成排版精良的文档并保存到桌面。

## ✨ 特性

- **中文友好** —— PDF 可选「苹方 + Helvetica Neue」现代风或「Noto Serif CJK SC + Times New Roman」学术衬线风；Word 默认宋体 + Times New Roman。
- **数学公式** —— 支持 LaTeX，行内 `$...$` 与块级 `$$...$$`。PDF 用 MathJax 渲染；Word 转为原生 OMML（可在 Word 里直接编辑）。
- **有机化学式** —— 支持 mhchem 的 `\ce{}` 语法。
- **三线表** —— 表格统一渲染为学术三线表（无竖线），支持**跨页表头自动重复**，并自动换行防止超宽数据丢失。
- **货币识别** —— 自动区分货币符号 `$120` 与行内公式 `$x^2$`，不会把货币误当公式而破坏表格。
- **代码高亮** —— 围栏代码块带样式渲染。

## 📦 安装依赖

```bash
pip install markdown playwright python-docx latex2mathml lxml
playwright install chromium

# 可选：离线渲染 PDF 公式（不装则联网用 CDN）
npm install -g mathjax
```

> 公式转 Word 时，若检测到本机安装了 Microsoft Word（含 `MML2OMML.XSL`），会优先用它生成兼容性最佳的原生 OMML；否则使用内置的 LaTeX→OMML 生成器。

## 🚀 使用

1. 复制一段 Markdown 到剪贴板
2. 运行脚本：

   ```bash
   python3 md2doc.py
   ```

3. 按提示选择：

   ```
   输出格式：1. PDF   2. Word
   PDF 字体主题：1. 苹方   2. Noto 衬线
   输出文件名
   ```

4. 文档自动保存到**桌面**，并在访达 / 资源管理器中高亮显示。

## 🖥️ 环境要求

- Python 3.10+
- macOS（`pbpaste` 读剪贴板）或 Windows（`tkinter` 读剪贴板）
- PDF 生成需要 Playwright Chromium

## 📝 支持的 Markdown 语法

标题、段落、**粗体** / *斜体*、`行内代码`、围栏代码块、有序 / 无序列表、引用、水平线、表格、行内 / 块级数学公式、化学式（`\ce{}`）。

## 📄 许可

个人使用，欢迎自取自改。
