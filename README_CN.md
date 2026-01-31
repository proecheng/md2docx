# MD2DOCX

**Markdown 转 Word 转换器（支持 LaTeX 公式）**

将 Markdown 文档转换为 Word (.docx)，所有数学公式自动转换为 Word 公式编辑器格式 (OMML)。

[English Documentation](README.md)

## 功能特点

- **LaTeX 公式转换** - 块级公式 (`$$...$$`) 和行内公式 (`$...$`) 转换为 Word 原生公式编辑器格式
- **智能数学识别** - 自动识别文本中的裸数学表达式（希腊字母、下标、上标）
- **中文文档支持** - 宋体/黑体字体，正确的编码处理
- **完整 Markdown 支持** - 标题、表格、项目符号、加粗文本
- **多种使用方式** - 图形界面、拖放、命令行

## 安装方式

### 方式一：下载预编译 EXE（Windows 推荐）

从 [Releases](../../releases) 页面下载 `MD2DOCX.exe`，无需安装，双击即可使用。

### 方式二：pip 安装依赖

```bash
pip install python-docx lxml latex2mathml
```

然后下载 `md2docx.py` 直接运行。

### 方式三：克隆仓库

```bash
git clone https://github.com/YOUR_USERNAME/md2docx.git
cd md2docx
pip install -r requirements.txt
python md2docx.py
```

## 使用方法

### 图形界面（双击运行）

1. 双击 `MD2DOCX.exe`（或运行 `python md2docx.py`）
2. 在弹出对话框中选择 `.md` 文件
3. 转换后的 `.docx` 文件保存在同一目录

### 拖放方式

将任意 `.md` 文件拖放到 `MD2DOCX.exe` 图标上，自动在原目录生成同名 `.docx` 文件。

### 命令行

```bash
python md2docx.py 文档.md
# 或
MD2DOCX.exe 文档.md
```

## 支持的公式语法

### 块级公式

```markdown
$$
\frac{a}{b} + \sum_{i=1}^{n} x_i
$$
```

### 行内公式

```markdown
著名的公式 $E = mc^2$ 描述了质能关系。
```

### 裸数学表达式

转换器会自动识别以下内容：
- 希腊字母：α, β, γ, θ, π 等
- 下标变量：`x_t`, `h_i`, `W_k`
- 上标变量：`W^T`, `x^2`
- 函数表示：`f(x)`, `π_θ(a|s)`

## 技术架构

```
Markdown (.md)
     │
     ▼
[解析 Markdown] ──► 标题、表格、项目符号、段落
     │
     ▼
[识别公式] ──► 块级 ($$...$$) / 行内 ($...$) / 裸数学
     │
     ▼
[LaTeX 预处理] ──► 处理 \text{}、中文、自定义函数
     │
     ▼
latex2mathml.converter.convert()
     │
     ▼
MathML (XML)
     │
     ▼
[自定义 Python 转换器] ──► 20+ 元素处理器 (mfrac, msub, msup 等)
     │
     ▼
OMML (Office Math Markup Language)
     │
     ▼
python-docx ──► 嵌入 Word 文档
     │
     ▼
Word (.docx) 原生公式编辑器格式
```

### 为什么不用 XSLT？

微软提供了 `MML2OMML.xsl` 用于 MathML→OMML 转换，但它需要 XSLT 2.0，而 Python 的 `lxml` 库只支持 XSLT 1.0。本项目实现了纯 Python 的 MathML→OMML 转换器，支持 20+ 种 MathML 元素。

## 支持的 MathML 元素

| 元素 | 描述 | OMML 对应 |
|------|------|-----------|
| `mfrac` | 分数 | `m:f` |
| `msqrt` | 平方根 | `m:rad` |
| `mroot` | n次方根 | `m:rad` |
| `msup` | 上标 | `m:sSup` |
| `msub` | 下标 | `m:sSub` |
| `msubsup` | 上下标 | `m:sSubSup` |
| `munder` | 下方 | `m:limLow` / `m:nary` |
| `mover` | 上方 | `m:limUpp` / `m:acc` |
| `munderover` | 上下方 | `m:nary` |
| `mfenced` | 括号 | `m:d` |
| `mtable` | 矩阵/表格 | `m:m` |

## 从源码构建 EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "MD2DOCX" md2docx.py
```

生成的可执行文件在 `dist/` 目录下。

## 依赖项

- `python-docx` >= 0.8.11
- `lxml` >= 4.9.0
- `latex2mathml` >= 3.0.0
- `tkinter`（Python 自带）

## 许可证

MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。

## 贡献

欢迎提交 Issue 和 Pull Request！

## 致谢

- [python-docx](https://python-docx.readthedocs.io/) - Word 文档操作库
- [latex2mathml](https://github.com/roniemartinez/latex2mathml) - LaTeX 到 MathML 转换
