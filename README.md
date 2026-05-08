# Duties & VAT Extract

批量提取海关 PDF 文件中的：

- MRN（进口条码编码）
- Customs Duties（关税）
- VAT 税额

并自动汇总生成 Excel 表格。

---

# 功能特点

- 批量处理整个文件夹中的 PDF
- 自动提取：
  - MRN
  - Customs Duties
  - VAT (PVA)
- 自动导出 Excel 汇总文件
- 支持调试模式（Debug）
- 基于正则表达式解析
- 使用 pdfplumber 提取 PDF 文本

---

# 项目结构

```text
Duties-VAT_extract/
│
├── batch_extract.py
├── README.md
├── requirements.txt
└── sample_pdf/
```

---

# 安装依赖

建议使用 Python 3.10+

安装依赖：

```bash
pip install pandas pdfplumber openpyxl
```

或者：

```bash
pip install -r requirements.txt
```

---

# 使用方法

## 基础用法

```bash
python batch_extract.py --folder ./pdf_folder --output result.xlsx
```

---

## 开启调试模式

```bash
python batch_extract.py --folder ./pdf_folder --output result.xlsx --debug
```

调试模式会打印：

- PDF提取文本片段
- 税种匹配过程
- 正则命中情况

方便排查 PDF 格式问题。

---

# 输出结果

程序会生成：

```text
result.xlsx
```

Excel 包含：

| 文件名 | MRN | 关税 | VAT税 |
|---|---|---|---|
| xxx.pdf | 24GBXXXX | 123.45 | 456.78 |

---

# 技术实现

项目主要使用：

- Python
- pdfplumber
- pandas
- openpyxl
- 正则表达式（Regex）

核心逻辑：

1. 提取 PDF 文本
2. 正则匹配 MRN
3. 提取 Customs Duties
4. 提取 VAT (PVA)
5. 汇总生成 Excel

---

# 已知限制

当前版本仅支持：

- 可复制文本的 PDF

暂不支持：

- 扫描版 PDF
- OCR 图片识别

未来可扩展：

- OCR（Tesseract）
- 多线程处理
- GUI 图形界面
- 自动分类归档

---

# 示例命令

```bash
python batch_extract.py \
  --folder ./customs_pdf \
  --output customs_summary.xlsx
```

---

# License

MIT License
