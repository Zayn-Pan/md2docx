# md2docx

A lightweight desktop tool for converting Markdown files to Word documents (.docx).

## Features

- **Full conversion** - Headings, bold, italic, strikethrough, lists, code blocks, tables, images, links, blockquotes, horizontal rules
- **GUI** - Click and convert, no command line required
- **Single exe** - Windows portable executable, no Python needed
- **Auto image paths** - Handles relative image paths automatically

## Supported Markdown

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

## Quick Start

### Run the packaged .exe (Windows)

    dist/Md2Docx.exe

### Run from source

    pip install -r requirements.txt
    python main.py

## Project Structure

    md2docx/
    - main.py              # Entry point
    - converter.py          # Core conversion logic
    - gui.py               # GUI interface
    - requirements.txt     # Python dependencies
    - test.md / test.docx  # Test samples

## Development

### Install dependencies

    pip install markdown python-docx Pillow beautifulsoup4

### Build .exe with PyInstaller

    pip install pyinstaller
    python -m PyInstaller --onefile --windowed --name "Md2Docx" main.py

Output: `dist/Md2Docx.exe`

## License

MIT License
