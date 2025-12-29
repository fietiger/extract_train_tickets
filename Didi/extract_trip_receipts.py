import os
import csv
import sys
from pathlib import Path


def check_dependencies():
    """检查必要的依赖库"""
    missing_deps = []
    
    try:
        import PyPDF2
    except ImportError:
        missing_deps.append("PyPDF2")
        
    try:
        from PIL import Image
    except ImportError:
        missing_deps.append("Pillow")
        
    try:
        import pytesseract
    except ImportError:
        missing_deps.append("pytesseract")
    
    if missing_deps:
        print(f"警告: 缺少以下依赖库: {', '.join(missing_deps)}")
        print("请运行以下命令安装: pip install " + " ".join(missing_deps))
        print("对于tesseract OCR引擎，您还需要单独安装Tesseract-OCR软件")
        return False
    
    return True


def extract_text_from_pdf(pdf_path):
    """从PDF文件中提取文本"""
    try:
        import PyPDF2
    except ImportError:
        print("警告: 未安装PyPDF2库，无法处理PDF文件")
        return ""
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        print(f"处理PDF文件 {pdf_path} 时出错: {e}")
        return ""


def extract_text_from_image(image_path):
    """从图像文件中提取文本（OCR）"""
    try:
        from PIL import Image
        import pytesseract
    except ImportError:
        print("警告: 未安装Pillow或pytesseract库，无法处理图像文件")
        return ""
    
    try:
        image = Image.open(image_path)
        # 尝试使用中文+英文识别
        text = pytesseract.image_to_string(image, lang='chi_sim+eng')
        return text
    except Exception as e:
        print(f"处理图像文件 {image_path} 时出错: {e}")
        return ""


def extract_content_from_file(file_path):
    """根据文件扩展名选择适当的提取方法"""
    file_path = Path(file_path)
    extension = file_path.suffix.lower()
    
    if extension == '.pdf':
        content = extract_text_from_pdf(file_path)
    elif extension in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
        content = extract_text_from_image(file_path)
    else:
        # 假设是文本文件
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            # 尝试其他编码
            try:
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()
            except UnicodeDecodeError:
                print(f"无法读取文件 {file_path}，编码不支持")
                return ""
        except Exception as e:
            print(f"读取文件 {file_path} 时出错: {e}")
            return ""
    
    # 清理提取的文本
    return clean_extracted_text(content)


def is_trip_data_line(line):
    """检查一行是否是行程数据行（以序号开头）"""
    import re
    # 检查是否以数字开头且后面跟着车型关键词或其一部分
    # B.pdf 中数字和文字可能连在一起，将 \s+ 改为 \s*
    match = re.match(r'^(\d+)\s*(.*)', line)
    if match:
        seq = match.group(1)
        rest = match.group(2)
        # 即使只包含车型的一部分也认为是开始
        trip_keywords = ['滴滴特快', '特惠快车', '惊喜特价', '快车', '专车', '出租车', '滴滴特', '特惠快', '惊喜特', '特快', '惊喜']
        if any(keyword in rest for keyword in trip_keywords):
            return True
        # 或者直接跟着时间
        if re.search(r'\d{2}-\d{2}\s+\d{2}:\d{2}', rest):
            return True
    return False


def extract_trip_data(text, filename):
    """从文本中提取行程相关数据"""
    import re
    lines = text.split('\n')
    extracted_data = []
    
    # 查找表格头部的索引
    start_index = -1
    for i, line in enumerate(lines):
        if '序号' in line and '车型' in line and '上车时间' in line:
            start_index = i
            break
            
    if start_index == -1:
        return []

    # 提取表格内容行
    table_lines = lines[start_index + 1:]
    current_trip = None
    
    for line in table_lines:
        line = line.strip()
        if not line or '页码' in line or '合计' in line:
            if '合计' in line and current_trip:
                extracted_data.append(current_trip)
                current_trip = None
            continue
            
        # 检查是否是新行的开始（以数字开头）
        if re.match(r'^\d+\s+', line):
            if current_trip:
                extracted_data.append(current_trip)
            current_trip = line
        elif current_trip:
            # 合并到当前行。智能处理空格：如果连接处是两个中文字符，则不加空格
            if current_trip and line:
                last_char = current_trip[-1]
                first_char = line[0]
                # 判断是否都是中文
                is_last_zh = '\u4e00' <= last_char <= '\u9fff'
                is_first_zh = '\u4e00' <= first_char <= '\u9fff'
                
                if is_last_zh and is_first_zh:
                    current_trip += line
                else:
                    current_trip += " " + line
            else:
                current_trip += line
            
    if current_trip:
        extracted_data.append(current_trip)
        
    final_data = []
    for trip_line in extracted_data:
        parsed_data = parse_trip_data(trip_line.strip(), filename)
        if len(parsed_data) == 8:
            final_data.append([filename] + parsed_data)
            
    return final_data


def parse_trip_data(trip_line, filename):
    """解析行程数据行，将其分割成特定字段"""
    import re
    
    # 清理行：将多个空格替换为一个，处理常见的跨行断词
    trip_line = re.sub(r'\s+', ' ', trip_line).strip()
    
    # 1. 提取序号
    seq_match = re.match(r'^(\d+)\s*', trip_line)
    if not seq_match:
        return []
    seq_num = seq_match.group(1)
    remaining = trip_line[seq_match.end():].strip()
    
    # 2. 提取车型和上车时间
    # 滴滴行程单格式：[车型][MM-DD HH:mm]
    # 时间格式：\d{2}-\d{2}\s+\d{2}:\d{2}
    time_pattern = r'(\d{2}-\d{2}\s+\d{2}:\d{2})'
    time_match = re.search(time_pattern, remaining)
    
    if not time_match:
        return []
    
    raw_vehicle = remaining[:time_match.start()].strip()
    # 车型可能包含空格（跨行合并导致的），清理掉
    vehicle_type = raw_vehicle.replace(' ', '')
    # 常见车型修正
    vehicle_fixes = {
        '特惠快车': ['特惠快车', '特惠快', '特惠'],
        '惊喜特价': ['惊喜特价', '惊喜特', '惊喜'],
        '滴滴特快': ['滴滴特快', '滴滴特', '特快']
    }
    for correct, wrongs in vehicle_fixes.items():
        if any(w in vehicle_type for w in wrongs):
            vehicle_type = correct
            break

    time_str_raw = time_match.group(1)
    remaining = remaining[time_match.end():].strip()
    
    # 3. 处理时间格式
    date_part, clock_part = time_str_raw.split(' ')
    month, day = date_part.split('-')
    hour, minute = clock_part.split(':')
    time_str = f"2025/{int(month)}/{int(day)} {int(hour)}:{minute}"
    
    # 4. 提取城市和星期
    # 格式通常是: 周X 城市名 ... 或 周X城市名 ...
    # 移除 OCR 可能产生的 "周 X"、"周, X" 等干扰
    remaining = re.sub(r'周[,，\s]*[一二三四五六日]\s*', '', remaining).strip()
    
    # 检查是否包含城市（支持 "城市名 市" 或 "城市名市"）
    city = ""
    city_match = re.search(r'([\u4e00-\u9fa5]+?)\s*市', remaining)
    if city_match:
        city = city_match.group(1) + "市"
        # 移除城市及它前面的内容
        remaining = remaining[city_match.end():].strip()
    
    # 清理起点前面的残余字符（如 城市名、市、逗号等）
    remaining = re.sub(r'^[,，\s]*', '', remaining)
    remaining = re.sub(r'^[\u4e00-\u9fa5]{2,3}?\s*市', '', remaining).strip()
    remaining = re.sub(r'^[,，\s]*', '', remaining)
    
    # 4. 提取里程和金额 (在末尾)
    # 格式通常是: [里程] [金额]
    # 里程和金额中间可能有空格，也可能紧贴在一起
    # 里程可能包含一位小数，金额固定两位小数
    # 我们从右往左寻找金额，然后再寻找里程
    amount_match = re.search(r'(\d+\.?\d*)\s*(\d+\.\d{2})$', remaining)
    distance = ""
    amount = ""
    if amount_match:
        distance = amount_match.group(1)
        amount = amount_match.group(2)
        remaining = remaining[:amount_match.start()].strip()
    
    # 5. 分割起点和终点
    # 移除可能存在的 OCR 冗余前缀（如“周日”、“周一”等）
    remaining = re.sub(r'^.*?[周星期][一二三四五六日]\s*', '', remaining).strip()
    
    # 特殊处理 B.pdf 中这种括号被截断的情况：如果起点或终点包含未闭合的左括号，且另一部分以右括号开头
    # 但更通用的做法是先清理一次 OCR 常见的括号错误
    remaining = remaining.replace('\n', '')
    
    # 寻找分割点
    bar_indices = [m.start() for m in re.finditer(r'\|', remaining)]
    
    start_location = ""
    end_location = ""
    
    if len(bar_indices) >= 2:
        # 方案 A: 寻找两个 "|" 之间的分界（空格 + 行政区划）
        # B.pdf 中可能没有空格，只有行政区划，如 "...旁)光谷|..."
        # 我们寻找紧贴在另一个地点前的行政区划名
        split_match = re.search(r'([\s\)]+)([\u4e00-\u9fa5]{2,5}[市区县]|[\u4e00-\u9fa5]{2,5})\|', remaining)
        if split_match:
            # 找到最后一个匹配（最靠近终点的那个）
            all_matches = list(re.finditer(r'([\s\)]+)([\u4e00-\u9fa5]{2,5}[市区县]|[\u4e00-\u9fa5]{2,5})\|', remaining))
            last_match = all_matches[-1]
            split_index = last_match.start(2)
            start_location = remaining[:split_index].strip()
            end_location = remaining[split_index:].strip()
        else:
            # 方案 B: 强制在第二个 "|" 前的行政区名处分割
            second_loc_match = re.search(r'([\u4e00-\u9fa5]{2,5}[市区县]|[\u4e00-\u9fa5]{2,5})\|', remaining[bar_indices[0]+1:])
            if second_loc_match:
                split_point = bar_indices[0] + 1 + second_loc_match.start()
                start_location = remaining[:split_point].strip()
                end_location = remaining[split_point:].strip()
            else:
                # 方案 C: 尝试寻找最后一个明显的空格
                last_space = remaining.rfind(' ', 0, bar_indices[-1])
                if last_space != -1:
                    start_location = remaining[:last_space].strip()
                    end_location = remaining[last_space:].strip()
                else:
                    parts = remaining.split(' ', 1)
                    start_location = parts[0]
                    end_location = parts[1] if len(parts) > 1 else ""
    elif len(bar_indices) == 1:
        # 只有一个 "|"，说明起点或终点中有一个没有 "|"
        # B.pdf 第2笔: 光谷物联港-西门(北大荒信息有限公司武汉研发中心旁)光谷|中国铁建·梧桐苑1期-南门
        # 这种情况通常是 [地点1] [行政区]| [地点2]
        split_match = re.search(r'([^\s\|]+?)([\s\)]+)([\u4e00-\u9fa5]{2,5}[市区县]|[\u4e00-\u9fa5]{2,5})\|', remaining)
        if split_match:
            # 在行政区（group 3）前分割
            split_index = split_match.start(3)
            start_location = remaining[:split_index].strip()
            end_location = remaining[split_index:].strip()
        else:
            # 找 "|" 后面的第一个空格
            split_index = remaining.find(' ', bar_indices[0])
            if split_index != -1:
                start_location = remaining[:split_index].strip()
                end_location = remaining[split_index:].strip()
            else:
                parts = remaining.split(' ', 1)
                start_location = parts[0]
                end_location = parts[1] if len(parts) > 1 else ""
    else:
        # 没有 "|"，按第一个空格分割
        parts = remaining.split(' ', 1)
        start_location = parts[0]
        end_location = parts[1] if len(parts) > 1 else ""
    
    # 清理每个地点内部的 OCR 错误
    def clean_loc(loc):
        if not loc: return ""
        # 移除换行
        loc = loc.replace('\n', '')
        # 合并中文字符之间的空格
        loc = re.sub(r'([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])', r'\1\2', loc)
        # 合并括号前后的空格
        loc = re.sub(r'\( +', '(', loc)
        loc = re.sub(r' + \)', ')', loc)
        loc = re.sub(r'（ +', '（', loc)
        loc = re.sub(r' + ）', '）', loc)
        # 特殊处理：如果地点以右括号开头，且前面一部分以左括号结尾，则这可能是分割错误
        # 在这里我们主要清理地点内部的冗余
        return loc.strip()

    start_location = clean_loc(start_location)
    end_location = clean_loc(end_location)
    
    # 再次检查起点和终点的分割是否合理（处理 B.pdf 中括号截断导致的分割错误）
    if start_location.endswith('(') or start_location.endswith('（'):
        if end_location.startswith(')') or end_location.startswith('）'):
            # 这说明分割点选在了括号中间，需要向后找下一个分割点
            pass # 这种极端情况暂不处理，先尝试更简单的逻辑
    
    # 移除地点末尾的冗余（如 OCR 误识别的逗号或由于分割不当留下的部分）
    start_location = start_location.rstrip(' ,，')
    end_location = end_location.rstrip(' ,，')
    
    return [seq_num, vehicle_type, time_str, city, start_location, end_location, distance, amount]


def clean_extracted_text(text):
    """清理提取的文本，处理可能的格式问题"""
    import re
    
    # 将多个连续空格替换为单个空格
    text = re.sub(r' +', ' ', text)
    
    # 按行分割文本
    lines = text.split('\n')
    
    # 首先查找表格区域并处理表格数据
    in_table_section = False
    table_header_keywords = ['序号', '车型', '上车时间', '城市', '起点', '终点', '里程', '金额', '备注']
    
    processed_lines = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        
        # 检查是否进入表格区域
        if any(keyword in line for keyword in table_header_keywords):
            in_table_section = True
            processed_lines.append(line)
            i += 1
        elif in_table_section and is_new_table_row(line):
            # 这是一个新的表格行，尝试合并可能被错误分割的行
            current_line = line
            j = i + 1
            while j < len(lines):
                next_line = lines[j].strip()
                if not next_line:
                    j += 1
                    continue
                
                # 检查下一行是否可能是当前行的延续
                # 判断标准：不是新的表格行，且包含表格数据特征
                if not is_new_table_row(next_line) and looks_like_table_continuation(next_line):
                    current_line += " " + next_line
                    j += 1
                else:
                    # 如果下一行包含结束标识（如页码），则停止合并
                    if '页码' in next_line or '合计' in next_line:
                        break
                    else:
                        # 不是表格延续，也不是结束标识，可能是另一个表格行的开始
                        break
            
            processed_lines.append(current_line)
            i = j
        elif '页码' in line or any(keyword in line for keyword in ['申请日期', '行程起止日期', '姓名', '工号', '部门']):
            # 遇到页码或非表格信息，退出表格区域
            in_table_section = False
            processed_lines.append(line)
            i += 1
        else:
            # 不在表格区域内或非表格行
            processed_lines.append(line)
            i += 1
    
    return '\n'.join(processed_lines)


def looks_like_table_continuation(line):
    """检查一行是否可能是表格数据的延续"""
    # 检查是否包含时间格式（如 11-09 08:25）
    import re
    time_pattern = r'\d{2}-\d{2}\s+\d{2}:\d{2}'
    if re.search(time_pattern, line):
        return True
    
    # 检查是否包含金额格式（如 19.90）
    amount_pattern = r'\d+\.\d{2}'
    if re.search(amount_pattern, line):
        return True
    
    # 检查是否包含公里数格式（如 4.5）
    km_pattern = r'\d+\.\d+'
    if re.search(km_pattern, line):
        return True
    
    # 检查是否包含城市、地点等关键词
    location_keywords = ['市', '区', '|', '站', '门', '路', '村', '园', '中心']
    if any(keyword in line for keyword in location_keywords):
        return True
    
    return False





def is_new_table_row(line):
    """检查一行是否是新的表格行（如包含序号开头）"""
    import re
    # 检查是否以数字开头（可能是新的序号）
    if re.match(r'^\d+\s+', line):
        # 检查是否包含车型等表格特征，以确认是否是新行
        table_keywords = ['滴滴特快', '特惠快车', '惊喜特价', '快车', '专车', '出租车']
        if any(keyword in line for keyword in table_keywords):
            return True
    
    return False


def process_trip_receipts(input_dir, output_csv):
    """处理所有行程报销单文件并生成CSV"""
    input_path = Path(input_dir)
    all_data = []
    
    # 只处理匹配"滴滴出行行程报销单*.pdf"模式的文件
    import fnmatch
    
    # 检查是否有匹配的文件
    found_files = False
    for file_path in input_path.rglob('*'):
        if (file_path.is_file() and 
            file_path.suffix.lower() == '.pdf' and 
            fnmatch.fnmatch(file_path.name, '滴滴出行行程报销单*.pdf')):
            found_files = True
            print(f"发现文件: {file_path.name} (路径: {file_path})")
    
    if not found_files:
        print(f"在目录 {input_dir} 中未找到匹配 '滴滴出行行程报销单*.pdf' 模式的PDF文件")
        return
    
    # 遍历输入目录中的所有文件
    for file_path in input_path.rglob('*'):
        if (file_path.is_file() and 
            file_path.suffix.lower() == '.pdf' and 
            fnmatch.fnmatch(file_path.name, '滴滴出行行程报销单*.pdf')):
            print(f"处理文件: {file_path.name}")
            
            # 提取文件内容
            content = extract_content_from_file(file_path)
            
            if content:
                # 提取行程数据
                trip_data = extract_trip_data(content, file_path.name)
                
                # 添加到总数据列表
                all_data.extend(trip_data)
            else:
                print(f"警告: 无法从文件 {file_path.name} 中提取内容")
    
    if all_data:
        # 写入CSV文件，使用UTF-8 BOM编码
        with open(output_csv, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            # 写入表头 - 检查数据是否已解析为多列
            if all_data and len(all_data[0]) == 9:  # [文件名, 序号, 车型, 时间, 城市, 起点, 终点, 里程, 金额]
                writer.writerow(['文件名', '序号', '车型', '上车时间', '城市', '起点', '终点', '里程', '金额'])
            else:
                # 如果数据未解析，使用原始格式
                writer.writerow(['文件名', '内容'])
            # 写入数据
            writer.writerows(all_data)
        
        print(f"处理完成！共提取 {len(all_data)} 行数据，保存到 {output_csv}")
    else:
        print("没有提取到任何数据，请检查文件格式和依赖库")


def main():
    print("滴滴出行行程报销单内容提取工具")
    print("=" * 40)
    
    # 检查依赖
    if not check_dependencies():
        print("\n缺少必要的依赖库，程序无法正常运行。")
        return
    
    # 设置默认输入目录和输出文件路径
    input_directory = "."  # 当前目录
    output_file = "trip_receipts.csv"
    
    # 检查输入目录是否存在
    if not os.path.exists(input_directory):
        print(f"错误: 输入目录不存在 - {input_directory}")
        return
    
    # 检查是否是目录
    if not os.path.isdir(input_directory):
        print(f"错误: 指定路径不是目录 - {input_directory}")
        return
    
    # 处理报销单
    try:
        process_trip_receipts(input_directory, output_file)
        print("\n任务完成！")
    except Exception as e:
        print(f"程序执行出错: {e}")


if __name__ == "__main__":
    main()