# MD2DOCX

**Markdown to Word Converter with LaTeX Formula Support**

Convert Markdown documents to Word (.docx) with all mathematical formulas converted to native Word equation editor format (OMML).

[中文文档](README_CN.md)

## Features

- **LaTeX Formula Conversion** - Block formulas (`$$...$$`) and inline formulas (`$...$`) are converted to Word's native equation editor format
- **Smart Math Detection** - Automatically detects bare mathematical expressions in text (Greek letters, subscripts, superscripts)
- **Chinese Document Support** - SimSun/SimHei fonts with proper encoding
- **Full Markdown Support** - Headings, tables, bullet points, bold text
- **Multiple Usage Methods** - GUI, drag-and-drop, command line

## Installation

### Option 1: Download Pre-built Application

Download from the [Releases](../../releases) page:
- **Windows**: `MD2DOCX.exe` - Double-click to run
- **macOS**: `MD2DOCX-Mac.zip` - Unzip and drag `MD2DOCX.app` to Applications

No Python installation required.

### Option 2: Install via pip

```bash
pip install python-docx lxml latex2mathml
```

Then download `md2docx.py` and run it directly.

### Option 3: Clone Repository

```bash
git clone https://github.com/proecheng/md2docx.git
cd md2docx
pip install -r requirements.txt
python md2docx.py
```

## Usage

### GUI Mode (Double-click)

1. Double-click `MD2DOCX.exe` (or run `python md2docx.py`)
2. Select a `.md` file in the dialog
3. The converted `.docx` will be saved in the same directory

### Drag-and-Drop

Drag any `.md` file onto `MD2DOCX.exe` - the converted file will be created automatically.

### Command Line

```bash
python md2docx.py document.md
# or
MD2DOCX.exe document.md
```

## Supported Formula Syntax

### Block Formulas

```markdown
$$
\frac{a}{b} + \sum_{i=1}^{n} x_i
$$
```

### Inline Formulas

```markdown
The formula $E = mc^2$ is famous.
```

### Bare Math Expressions

The converter automatically detects:
- Greek letters: α, β, γ, θ, π, etc.
- Subscript variables: `x_t`, `h_i`, `W_k`
- Superscript variables: `W^T`, `x^2`
- Function notation: `f(x)`, `π_θ(a|s)`

## Technical Architecture

```
Markdown (.md)
     │
     ▼
[Parse Markdown] ──► Headings, Tables, Bullets, Paragraphs
     │
     ▼
[Detect Formulas] ──► Block ($$...$$) / Inline ($...$) / Bare math
     │
     ▼
[LaTeX Preprocessing] ──► Handle \text{}, Chinese, custom functions
     │
     ▼
latex2mathml.converter.convert()
     │
     ▼
MathML (XML)
     │
     ▼
[Custom Python Converter] ──► 20+ element handlers (mfrac, msub, msup, etc.)
     │
     ▼
OMML (Office Math Markup Language)
     │
     ▼
python-docx ──► Embed formulas in Word document
     │
     ▼
Word (.docx) with native equation editor formulas
```

### Why Not XSLT?

Microsoft provides `MML2OMML.xsl` for MathML→OMML conversion, but it requires XSLT 2.0. Python's `lxml` library only supports XSLT 1.0. This project implements a pure Python MathML→OMML converter as a workaround, supporting 20+ MathML elements.

## Supported MathML Elements

| Element | Description | OMML Equivalent |
|---------|-------------|-----------------|
| `mfrac` | Fractions | `m:f` |
| `msqrt` | Square root | `m:rad` |
| `mroot` | N-th root | `m:rad` |
| `msup` | Superscript | `m:sSup` |
| `msub` | Subscript | `m:sSub` |
| `msubsup` | Sub+superscript | `m:sSubSup` |
| `munder` | Under | `m:limLow` / `m:nary` |
| `mover` | Over | `m:limUpp` / `m:acc` |
| `munderover` | Under+over | `m:nary` |
| `mfenced` | Parentheses | `m:d` |
| `mtable` | Matrix/Table | `m:m` |

## Build from Source

### Windows
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "MD2DOCX" md2docx.py
```

### macOS
```bash
pip3 install pyinstaller
# or run:
chmod +x build_mac.sh && ./build_mac.sh
```

The output will be in the `dist/` folder.

Alternatively, push a version tag to trigger the [GitHub Actions workflow](.github/workflows/build.yml) which builds for both platforms automatically:
```bash
git tag v1.0.1
git push origin v1.0.1
```

## Dependencies

- `python-docx` >= 0.8.11
- `lxml` >= 4.9.0
- `latex2mathml` >= 3.0.0
- `tkinter` (included with Python)

## License

MIT License - see [LICENSE](LICENSE) file.

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

## Acknowledgments

- [python-docx](https://python-docx.readthedocs.io/) for Word document manipulation
- [latex2mathml](https://github.com/roniemartinez/latex2mathml) for LaTeX to MathML conversion
