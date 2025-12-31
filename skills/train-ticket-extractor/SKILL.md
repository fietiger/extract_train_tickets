---
name: train-ticket-extractor
description: 此技能用于从中国铁路电子发票（PDF格式）中提取关键行程信息。当用户需要汇总、分析或整理火车票信息并输出为 Excel (XLSX) 文件时，应使用此技能。它支持递归搜索文件夹，能够提取发票号码、乘车日期、车次、出发站、到达站、乘客姓名、席别、价格等信息，并自动去重和格式化（防止长数字显示错误）。
---

# 火车票信息提取 (Train Ticket Extractor)

## 概述

此技能通过 Python 脚本 `extract_train_tickets.py` 自动化处理 PDF 格式的火车票电子发票。它能够深入子文件夹查找所有 PDF 文件，解析其中的文本内容，并将提取到的结构化数据保存为 XLSX 文件。

## 使用场景

- 财务报销汇总：批量提取多张火车票信息。
- 行程整理：从发票文件夹中提取行程详情。
- 格式化输出：将 PDF 内容转换为可编辑、可计算的 Excel 表格。

## 核心功能

1. **递归搜索**：自动查找指定目录及其所有子目录下的 `.pdf` 文件。
2. **字段提取**：
   - 文件名 (`filename`)
   - 发票号码 (`invoice_number`) - 自动格式化为 `="编号"` 模式
   - 乘车日期 (`date`) - 转换为 `YYYY-MM-DD` 格式
   - 车次 (`train_number`)
   - 出发/到达站及路线 (`departure_station`, `arrival_station`, `route`)
   - 出发时间 (`departure_time`)
   - 乘客姓名 (`passenger_name`)
   - 席位信息 (`seat_type`, `seat_number`)
   - 价格 (`price`) - 自动格式化为 `=价格` 模式
3. **数据去重**：基于发票号码自动过滤重复文件。
4. **Excel 兼容性**：输出带有公式保护的 XLSX 文件，确保长数字（如发票号）在 Excel 中显示正确。

## 使用方法

### 运行脚本

直接调用脚本并指定目标文件夹路径和输出文件名：

```bash
python scripts/extract_train_tickets.py <目标文件夹路径> [输出文件名]
```

- **目标文件夹路径**：可选，默认为当前工作目录。
- **输出文件名**：可选，默认为 `火车票汇总信息表.xlsx`。

如果未指定路径，默认处理当前工作目录。


### 输出文件

脚本将在当前目录下生成 `火车票汇总信息表.xlsx`。

## 依赖库

此技能依赖以下 Python 库：
- `pdfplumber`: 用于解析 PDF 文本。
- `pandas`: 用于数据处理和生成 Excel。
- `openpyxl`: `pandas` 生成 XLSX 所需的引擎。

## 资源说明

### scripts/
- `extract_train_tickets.py`: 核心提取逻辑脚本。
