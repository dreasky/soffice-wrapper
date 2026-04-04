# soffice_wrapper

基于 LibreOffice（soffice）的封装库，将 Office 文档转换为 PDF。

## 安装

```bash
# 作为 git submodule
git submodule add https://github.com/dreasky/soffice_wrapper.git libs/soffice_wrapper
pip install -e libs/soffice_wrapper

# 直接从 GitHub 安装
pip install git+https://github.com/dreasky/soffice_wrapper.git
```

## 前置依赖

需要安装 LibreOffice 并确保 `soffice` 命令可用：

- 下载：<https://www.libreoffice.org/download/>
- 安装后将 LibreOffice 的 `program` 目录加入系统 `PATH`

验证：

```bash
soffice --version
```

## 使用

```python
from soffice_wrapper import SofficeWrapper, SOFFICE_EXTENSIONS
from pathlib import Path

wrapper = SofficeWrapper()

# 单文件转换
output_pdf = wrapper.convert(Path("report.docx"), output_dir=Path("output"))

# 批量转换（过滤支持的格式）
files = [f for f in Path("docs").iterdir() if f.suffix.lower() in SOFFICE_EXTENSIONS]
for f in files:
    wrapper.convert(f, output_dir=Path("output"))
```

## 支持的格式

`SOFFICE_EXTENSIONS` 包含：`.doc` `.docx` `.ppt` `.pptx` `.odt` `.odp`

LibreOffice 实际支持更多格式，详见 <https://documentation.libreoffice.org/>。
