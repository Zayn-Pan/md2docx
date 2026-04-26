# md2docx

Markdown 转 Word (Docx) 桌面工具 / A lightweight desktop tool for converting Markdown to Word.

[English](#english) | [中文](#中文)

---

## English

A lightweight desktop tool for converting Markdown files to Word documents (.docx).

### Features

- **Full conversion** - Headings, bold, italic, strikethrough, lists, code blocks, tables, images, links, blockquotes, horizontal rules
- **GUI** - Click and convert, no command line required
- **Single exe** - Windows portable executable, no Python needed
- **Auto image paths** - Handles relative image paths automatically

### Supported Markdown

| Element | Supported |
|---------|-----------|
| Headings (H1-H6) | Yes |
| Bold / Italic / Strikethrough | Yes |
| Bullet / Numbered Lists | Yes |
| Code Blocks / Inline Code | Yes |
| Tables | Yes |
| Images | Yes |
| Links | Yes |
| Blockquotes | Yes |
| Horizontal Rules | Yes |

### Quick Start

**Run the packaged .exe (Windows)**
```
dist/Md2Docx.exe
```

**Run from source**
```bash
pip install -r requirements.txt
python main.py
```

### Development

```bash
pip install markdown python-docx Pillow beautifulsoup4
pip install pyinstaller
python -m PyInstaller --onefile --windowed --name "Md2Docx" main.py
```

Output: `dist/Md2Docx.exe`

### Project Structure
```
md2docx/
- main.py              # Entry point
- converter.py         # Core conversion logic
- gui.py              # GUI interface
- requirements.txt    # Python dependencies
- test.md / test.docx # Test samples
```

---

## 中文

轻量级桌面工具，将 Markdown 文件转换为 Word 文档 (.docx)。

### 功能特性

- **完整转换** — 标题、加粗、斜体、删除线、列表、代码块、表格、图片、链接、引用、水平线
- **GUI 界面** — 点击即用，无需命令行
- **单文件 exe** — Windows 便携版，无需 Python 环境
- **图片路径自适应** — 自动处理相对路径图片

### 支持的 Markdown 元素

| 元素 | 支持 |
|------|------|
| 标题 (H1-H6) | Yes |
| 加粗 / 斜体 / 删除线 | Yes |
| 无序 / 有序列表 | Yes |
| 代码块 / 行内代码 | Yes |
| 表格 | Yes |
| 图片 | Yes |
| 链接 | Yes |
| 引用块 | Yes |
| 水平线 | Yes |

### 快速开始

**运行打包版本（Windows）**
```
dist/Md2Docx.exe
```

**从源码运行**
```bash
pip install -r requirements.txt
python main.py
```

### 开发说明

```bash
pip install markdown python-docx Pillow beautifulsoup4
pip install pyinstaller
python -m PyInstaller --onefile --windowed --name "Md2Docx" main.py
```

### 项目结构
```
md2docx/
- main.py              # 程序入口
- converter.py         # 转换核心逻辑
- gui.py              # GUI 界面
- requirements.txt     # Python 依赖
- test.md / test.docx  # 测试示例
```

---

## License

MIT License