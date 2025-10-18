# 火车票信息提取工具

这是一个用于从PDF格式的火车票中自动提取票据信息的Python工具。

## 功能特性

- 🚄 **自动提取火车票信息**：从PDF文件中提取日期、车次、路线、时间、乘客、座位、价格等信息
- 🧾 **发票号码识别**：自动识别并提取发票号码
- 🔄 **智能去重**：基于发票号码自动去除重复的票据记录
- 📊 **CSV导出**：将提取的数据保存为CSV格式，方便在Excel中查看和处理
- 🔢 **Excel兼容**：发票号码使用特殊格式防止Excel将其识别为数字

## 系统要求

- Python 3.6+
- pdfplumber库

## 安装依赖

```bash
pip install pdfplumber
```

## 使用方法

### 1. 准备PDF文件

将需要处理的火车票PDF文件放在同一个文件夹中。

### 2. 运行脚本

```bash
python extract_train_tickets.py
```

脚本会自动：
- 扫描当前目录下的所有PDF文件
- 逐个处理每个PDF文件
- 提取票据信息
- 去除重复记录
- 生成CSV文件

### 3. 查看结果

处理完成后会生成 `train_tickets_extracted.csv` 文件，包含以下字段：

| 字段名 | 说明 |
|--------|------|
| filename | PDF文件名 |
| invoice_number | 发票号码（格式：="数字"） |
| date | 乘车日期 |
| train_number | 车次 |
| departure_station | 出发站 |
| arrival_station | 到达站 |
| departure_time | 出发时间 |
| arrival_time | 到达时间 |
| passenger_name | 乘客姓名 |
| seat_type | 座位类型 |
| seat_number | 座位号 |
| price | 票价 |

## 输出示例

```
Processing PDF files in: /path/to/your/folder
Found 4 PDF files to process...
Processing: ticket_001.pdf
  - Extracted data for ticket_001.pdf
Processing: ticket_002.pdf
  - Extracted data for ticket_002.pdf
...

==================================================
EXTRACTION SUMMARY
==================================================

Ticket 1: ticket_001.pdf
  Invoice Number: 12345678901234567890
  Date: 2024年03月15日
  Train: G123
  Route: 北京南 → 上海虹桥
  Time: 08:30 -
  Passenger: 张三
  Seat: 二等座 05车12A号
  Price: ¥553.00

...

Data successfully saved to train_tickets_extracted.csv
Total unique tickets: 4 (removed 0 duplicates)
```

## 特殊功能说明

### 去重机制

- **主要依据**：发票号码
- **备用依据**：如果没有发票号码，则使用文件名
- **处理逻辑**：相同发票号码的票据只保留一条记录

### Excel兼容性

发票号码在CSV中使用 `="数字"` 格式保存，这样可以：
- 防止Excel将长数字转换为科学计数法
- 保持发票号码的完整性
- 确保在Excel中正确显示为文本

## 注意事项

1. **PDF格式要求**：目前支持标准格式的中国铁路电子客票
2. **文件编码**：CSV文件使用UTF-8编码，确保中文正确显示
3. **重复处理**：如果同一张票有多个PDF文件，只会保留一条记录
4. **错误处理**：如果某个PDF无法解析，会跳过该文件并继续处理其他文件

## 故障排除

### 常见问题

**Q: 运行时提示"Permission denied"错误**
A: 请确保CSV文件没有被Excel或其他程序打开

**Q: 提取的信息不完整**
A: 请检查PDF文件是否为标准格式的火车票，某些特殊格式可能无法正确识别

**Q: 中文显示乱码**
A: 请确保使用支持UTF-8编码的文本编辑器或Excel打开CSV文件

## 技术支持

如果遇到问题或需要功能改进，请检查：
1. Python版本是否符合要求
2. 依赖库是否正确安装
3. PDF文件格式是否标准

## 更新日志

- **v1.0**: 基础功能实现
- **v1.1**: 添加发票号码提取功能
- **v1.2**: 实现基于发票号码的去重功能
- **v1.3**: 优化Excel兼容性，发票号码格式化