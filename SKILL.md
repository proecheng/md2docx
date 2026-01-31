# MD2DOCX Skill - Markdown to Word Converter

Convert Markdown documents to Word (.docx) with LaTeX formulas converted to native Word equation editor format (OMML).

## Description

This skill enables conversion of Markdown documents containing LaTeX mathematical formulas into Microsoft Word documents. The formulas are converted to Word's native equation editor format (OMML), making them fully editable in Word.

## When to Use

Use this skill when:
- Converting Markdown documents with LaTeX formulas to Word format
- Creating professional Word documents from technical markdown with equations
- Preserving mathematical formula editability in Word output
- Processing documents with Chinese text and mathematical content

## Capabilities

### Formula Conversion
- **Block formulas** (`$$...$$`) → Centered OMML equations
- **Inline formulas** (`$...$`) → Inline OMML equations
- **Bare math expressions** → Auto-detected and converted (Greek letters, subscripts, etc.)

### Document Elements
- Headings (H1-H3)
- Tables with proper formatting
- Bullet points
- Bold text
- Chinese/English mixed content

### Supported LaTeX Commands
- Fractions: `\frac{a}{b}`
- Subscripts/Superscripts: `x_i`, `x^2`, `x_i^j`
- Roots: `\sqrt{x}`, `\sqrt[n]{x}`
- Sums/Products: `\sum_{i=1}^{n}`, `\prod`
- Greek letters: `\alpha`, `\beta`, `\theta`, etc.
- Integrals: `\int_a^b`
- Matrices: `\begin{matrix}...\end{matrix}`
- Functions: `\sin`, `\cos`, `\log`, `\operatorname{...}`

## Technical Approach

```
LaTeX → MathML (latex2mathml) → OMML (custom Python) → Word (python-docx)
```

Key innovation: Pure Python MathML→OMML converter (20+ element handlers) that bypasses the XSLT 2.0 requirement of Microsoft's official converter.

## Usage

### Python API

```python
from md2docx import convert_md_to_docx

# Convert file
stats, output_path = convert_md_to_docx("document.md")
print(f"Converted: {stats['block']} block formulas, {stats['inline']} inline formulas")

# Convert with custom output path
stats, output_path = convert_md_to_docx("input.md", "output.docx")
```

### Command Line

```bash
python md2docx.py document.md
```

### GUI

```bash
python md2docx.py  # Opens file selection dialog
```

## Key Functions

### `convert_md_to_docx(input_file, output_file=None)`
Main conversion function. Returns (stats_dict, output_path).

### `latex_to_omml(latex_str)`
Convert LaTeX string to OMML element. Returns lxml Element or None.

### `mathml_to_omml(mathml_element)`
Convert MathML element tree to OMML. Core converter function.

### `identify_math_in_text(text)`
Detect mathematical expressions in plain text. Returns list of (start, end, expr) tuples.

### `text_to_latex(text)`
Convert detected math text to LaTeX format.

## Dependencies

```
python-docx>=0.8.11
lxml>=4.9.0
latex2mathml>=3.0.0
```

## Building Standalone Executables

When packaging with PyInstaller, you **must** include the `latex2mathml` data file:

### Windows
```batch
for /f "delims=" %%i in ('python -c "import latex2mathml; import os; print(os.path.dirname(latex2mathml.__file__))"') do set LATEX2MATHML_PATH=%%i
pyinstaller --onefile --windowed --name "MD2DOCX" --add-data "%LATEX2MATHML_PATH%\unimathsymbols.txt;latex2mathml" md2docx.py
```

### macOS
```bash
LATEX2MATHML_PATH=$(python3 -c "import latex2mathml; import os; print(os.path.dirname(latex2mathml.__file__))")
pyinstaller --onefile --windowed --name "MD2DOCX" \
    --add-data "${LATEX2MATHML_PATH}/unimathsymbols.txt:latex2mathml" \
    md2docx.py
```

**Important:** Without this data file, formula conversion will silently fail. The `--add-data` separator is `;` on Windows and `:` on macOS.

## Example

**Input (Markdown):**
```markdown
# Physics Formula

The famous equation $$E = mc^2$$ describes mass-energy equivalence.

The momentum is given by $p = mv$, where $m$ is mass and $v$ is velocity.
```

**Output (Word):**
- Heading "Physics Formula" in 黑体 22pt
- Block equation E = mc² (centered, OMML format)
- Paragraph with inline OMML equations for p = mv, m, and v

## Repository

https://github.com/proecheng/md2docx

## License

MIT License
