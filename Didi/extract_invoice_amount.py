import PyPDF2
import re
import sys
import os
import csv
from pathlib import Path

def extract_invoice_info(pdf_path):
    """
    从滴滴电子发票 PDF 中提取详细信息
    
    Returns:
        dict: 包含文件名、金额、日期、销售方、购买方的字典
    """
    info = {
        '文件名': os.path.basename(pdf_path),
        '开票日期': '未找到',
        '金额': 0.0,
        '销售方名称': '未找到',
        '销售方识别号': '未找到',
        '购买方名称': '未找到',
        '购买方识别号': '未找到'
    }

    if not os.path.exists(pdf_path):
        return info

    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            full_text = ""
            for page in reader.pages:
                full_text += page.extract_text() + "\n"
            
            lines = [line.strip() for line in full_text.split('\n') if line.strip()]
            
            # 1. 提取金额 (价税合计)
            amount_patterns = [
                r'价税合计.*?(\d+(?:\.\d+)?)¥',
                r'价税合计.*?(\d+(?:\.\d+)?)',
                r'小写.*?(\d+(?:\.\d+)?)',
                r'¥\s*(\d+(?:\.\d+)?)'
            ]
            for line in lines:
                if '价税合计' in line:
                    for pattern in amount_patterns:
                        match = re.search(pattern, line)
                        if match:
                            info['金额'] = float(match.group(1))
                            break
                    if info['金额'] > 0: break

            # 2. 提取日期
            # 格式: 开票日期 :2025年12月29日
            date_match = re.search(r'开票日期\s*[:：]\s*(\d{4}年\d{1,2}月\d{1,2}日)', full_text)
            if date_match:
                info['开票日期'] = date_match.group(1)

            # 3. 提取购买方和销售方信息
            # 滴滴发票的文本提取结果比较碎，需要根据上下文逻辑提取
            
            # 获取所有名称和识别号
            all_names = re.findall(r'名称\s*[:：]\s*([^\s\n]+)', full_text)
            all_ids = re.findall(r'纳税人识别号\s*[:：]\s*([A-Z0-9]+)', full_text)

            # 购买方固定信息
            BUYER_ID = "91440300MA5F1W6866"
            BUYER_NAME_KEYWORD = "宝链"
            
            # 逻辑：先在所有提取到的 ID 中找购买方 ID
            if BUYER_ID in all_ids:
                info['购买方识别号'] = BUYER_ID
            
            # 在所有提取到的名称中找购买方名称
            for name in all_names:
                if BUYER_NAME_KEYWORD in name:
                    info['购买方名称'] = re.split(r'统一社会|纳税人|识别号', name)[0]
                    break
            
            # 销售方信息：排除掉购买方后的第一个
            for tax_id in all_ids:
                if tax_id != BUYER_ID:
                    info['销售方识别号'] = tax_id
                    break
            
            for name in all_names:
                # 排除包含购买方关键字的名称
                clean_name = re.split(r'统一社会|纳税人|识别号', name)[0]
                if BUYER_NAME_KEYWORD not in clean_name and clean_name != '未找到':
                    info['销售方名称'] = clean_name
                    break

            # 针对 B.pdf 这种特殊连写情况的补丁
            if info['销售方名称'] == '未找到' or info['销售方识别号'] == '未找到':
                special = re.search(r'纳税人识别号\s*[:：]\s*([A-Z0-9]+)名称\s*[:：]\s*([^\s\n]+)', full_text)
                if special:
                    found_id = special.group(1)
                    found_name = re.split(r'统一社会|纳税人|识别号', special.group(2))[0]
                    
                    if found_id == BUYER_ID or BUYER_NAME_KEYWORD in found_name:
                        # 这是购买方，更新购买方信息
                        info['购买方识别号'] = found_id
                        info['购买方名称'] = found_name
                    else:
                        # 这是销售方
                        info['销售方识别号'] = found_id
                        info['销售方名称'] = found_name

            return info
    except Exception as e:
        print(f"处理文件 {pdf_path} 时出错: {e}")
        return info

def main():
    # 确定目标文件
    if len(sys.argv) > 1:
        target_files = sys.argv[1:]
    else:
        didi_dir = Path('Didi')
        if didi_dir.exists():
            target_files = list(didi_dir.glob('滴滴电子发票*.pdf'))
        else:
            target_files = list(Path('.').glob('滴滴电子发票*.pdf'))

    if not target_files:
        print("未找到滴滴电子发票文件。")
        return

    all_data = []
    for file_path in target_files:
        print(f"正在处理: {os.path.basename(file_path)}...")
        data = extract_invoice_info(str(file_path))
        all_data.append(data)

    # 保存为 CSV
    output_file = 'didi_invoices_extracted.csv'
    fieldnames = ['文件名', '开票日期', '金额', '购买方名称', '购买方识别号', '销售方名称', '销售方识别号']
    
    try:
        # 在写入 CSV 之前，对识别号进行特殊处理，防止 Excel 显示为科学计数法
        # 这种处理方式是使用 Excel 公式格式，例如 ="91440300MA5F1W6866"
        formatted_data = []
        for row in all_data:
            new_row = row.copy()
            if new_row['购买方识别号'] != '未找到':
                new_row['购买方识别号'] = f'="{new_row["购买方识别号"]}"'
            if new_row['销售方识别号'] != '未找到':
                new_row['销售方识别号'] = f'="{new_row["销售方识别号"]}"'
            formatted_data.append(new_row)

        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for row in formatted_data:
                writer.writerow(row)
        print(f"\n提取完成！结果已保存至: {output_file}")
        
        # 打印简要统计
        total = sum(d['金额'] for d in all_data)
        print(f"处理文件数: {len(all_data)}")
        print(f"总计金额: {total:.2f}")
        
    except Exception as e:
        print(f"保存 CSV 时出错: {e}")

if __name__ == "__main__":
    main()
