# md2docx

Markdown 转 Word (Docx) 桌面工具 — 轻量、简洁、无依赖烦恼。

## 功能特性

- **完整转换** — 标题、加粗、斜体、删除线、列表、代码块、表格、图片、链接、引用、水平线
- **GUI 界面** — 点击即用，无需命令行
- **轻量打包** — 生成单个 .exe 文件，Windows 即开即用
- **图片路径自适应** — 自动处理相对路径图片

## 支持的 Markdown 元素

| 元素 | 支持 |
|------|------|
| 标题 (H1-H6) | ✅ |
| 加粗 / 斜体 / 删除线 | ✅ |
| 无序 / 有序列表 | ✅ |
| 代码块 / 行内代码 | ✅ |
| 表格 | ✅ |
| 图片 | ✅ |
| 链接 | ✅ |
| 引用块 | ✅ |
| 水平线 | ✅ |

## 快速开始

### 运行打包版本（Windows）

```bash
# 双击 dist/Md2Docx.exe 即可
```

### 从源码运行

```bash
pip install -r requirements.txt
python main.py
```

## 项目结构

```
md2docx/
├── main.py              # 程序入口
├── converter.py          # 转换核心逻辑
├── gui.py               # GUI 界面
├── requirements.txt     # Python 依赖
└── test.md / test.docx  # 测试示例
```

## 开发说明

### 安装依赖

```bash
pip install markdown python-docx Pillow beautifulsoup4
```

### 打包为 exe

```bash
pip install pyinstaller
python -m PyInstaller --onefile --windowed --name "Md2Docx" main.py
```

打包产物：`dist/Md2Docx.exe`

## 截图预览

运行后界面简洁：选择一个 `.md` 文件 → 点击转换 → 自动生成同名的 `.docx` 文件。

## License

MIT License