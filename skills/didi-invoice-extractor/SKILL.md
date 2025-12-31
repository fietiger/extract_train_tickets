---
name: didi-invoice-extractor
description: 专门用于处理“滴滴出行电子发票”PDF文件，提取发票关键信息（开票日期、金额、购买方/销售方名称及识别号）并汇总为Excel文件的技能。
---

# 滴滴电子发票处理技能

此技能用于自动化提取滴滴出行电子发票中的结构化数据。它能够识别 PDF 中的开票信息，并生成统一的 Excel 汇总表。

## 使用场景

- 当用户需要将多个“滴滴电子发票”PDF 文件汇总到 Excel 时。
- 需要提取发票中的金额、日期、购方和销方信息进行报销核对时。

## 核心功能

1. **信息提取**：自动识别 PDF 中的以下字段：
   - 文件名
   - 开票日期
   - 价税合计金额
   - 购买方名称及纳税人识别号
   - 销售方名称及纳税人识别号
2. **格式优化**：生成的 Excel 会自动调整列宽，并设置美观的表头样式。

## 工作流程

1. **定位文件**：在指定的目录中查找包含“发票”关键字的 PDF 文件。
2. **执行脚本**：调用内置的 `scripts/extract_didi_invoices.py` 进行处理。
3. **生成结果**：在目标位置生成 `.xlsx` 汇总表。

## 使用指南

直接调用 `scripts/extract_didi_invoices.py` 脚本，并传入输入目录和输出路径：

```bash
python scripts/extract_didi_invoices.py <input_directory> <output_excel_path>
```

### 依赖项

- `pdfplumber`：用于解析 PDF 文本。
- `pandas`：用于数据管理。
- `openpyxl`：用于生成带样式的 Excel。
