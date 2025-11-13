# DOCX to PDF Converter

一个将 DOCX 文件转换为 PDF 并保留原始格式的 Python 脚本。

## 功能特点

- ✅ 保留原始文档格式（字体、样式、布局等）
- ✅ 支持单个文件转换
- ✅ 支持批量转换整个目录
- ✅ 自定义输出路径
- ✅ 自动创建输出目录
- ✅ 详细的错误处理和进度提示

## 安装依赖

### 方法 1: 使用脚本自动安装
```bash
python docx_to_pdf_converter.py --install
```

### 方法 2: 手动安装
```bash
pip install -r requirements.txt
```

### 方法 3: 单独安装
```bash
pip install docx2pdf python-docx reportlab
```

## 使用方法

### 1. 转换单个文件
```bash
python docx_to_pdf_converter.py document.docx
```

### 2. 指定输出路径
```bash
python docx_to_pdf_converter.py document.docx -o output.pdf
```

### 3. 批量转换目录中的所有 DOCX 文件
```bash
python docx_to_pdf_converter.py -d ./documents
```

### 4. 批量转换到指定输出目录
```bash
python docx_to_pdf_converter.py -d ./documents -o ./pdfs
```

## 命令行参数

- `input`: 输入的 DOCX 文件或目录路径
- `-o, --output`: 输出的 PDF 文件或目录路径
- `-d, --directory`: 将输入视为包含 DOCX 文件的目录
- `--install`: 安装所需的依赖包
- `-h, --help`: 显示帮助信息

## 系统要求

- Python 3.6+
- Microsoft Word（docx2pdf 依赖 Word 进行转换）
- 或者 LibreOffice（作为 Word 的替代方案）

## 注意事项

1. **Windows 用户**: 需要 Microsoft Word 安装在系统上
2. **macOS 用户**: 需要 Microsoft Word for Mac
3. **Linux 用户**: 可以使用 LibreOffice 作为替代

## 错误排查

### 常见问题

1. **"docx2pdf package not found"**
   - 解决方案: 运行 `python docx_to_pdf_converter.py --install`

2. **"Microsoft Word not found"**
   - 解决方案: 安装 Microsoft Word 或 LibreOffice

3. **"Permission denied"**
   - 解决方案: 检查文件权限，确保有读写权限

4. **"No DOCX files found"**
   - 解决方案: 确保目录中包含 .docx 文件

## 示例

```bash
# 基本用法
python docx_to_pdf_converter.py my_document.docx

# 批量转换
python docx_to_pdf_converter.py -d ./word_documents -o ./pdf_output

# 安装依赖
python docx_to_pdf_converter.py --install
```

转换完成后，PDF 文件将保持与原始 DOCX 文件相同的格式，包括：
- 字体样式和大小
- 段落格式
- 表格结构
- 图片位置
- 页面布局
- 标题和样式