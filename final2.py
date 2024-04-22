import pandas as pd
import re
from openai import OpenAI
from pathlib import Path
import datetime
from dateutil import parser
import os
import numpy as np


def extract_excel(file_path, api_key):
# 用chatgpt提取相关信息，并询问摘要是否和AI相关，输出
    # 初始化OpenAI客户端
    client = OpenAI(api_key=api_key)

    # 读取Excel文件到DataFrame
    df = pd.read_excel(file_path)

    # 选择'Content'列进行处理
    input_texts = df['Content'].tolist()

    # 创建空列表来存储提取的信息，仅在列不存在时创建
    extracted_data = {}
    columns_to_extract = ['Date', 'Time', 'Speaker', 'Venue', 'Affiliation', 'Abstract', 'Notes', 'AI']
    for col in columns_to_extract:
        if col not in df.columns:
            extracted_data[col] = []
    # 遍历'Content'列的每个项
    for text in input_texts:
        # 构建API请求
        question = """Below is the information about a web seminar.Please extract related information from the content and output it in the following format:
        1. Date: Exact date of the seminar, Just output the exact date of the seminar, in the following format 03 October 2018, if no date found, output blank
        2. Time: The exact time of the seminar, such as 10pm to 11pm
        3. Speaker: Just output the exact name of the speaker, such as ABC 
        4. Venue: The exact venue of the seminar
        5. Affiliation: The exact affiliation of the speaker
        6. Notes: In the information about a seminar, other things worth noting include, such as Registration Required: https://url-for-registration.org/ 
        Zoom only: https://zoom-url/ 
        7. Abstract: The exact abstract of the seminar.
        8. AI: Based on the abstract part you extracted, determine if the seminar is related to AI. If it is, output "Yes"; if not, output "No", such as Yes.


        If any item is missing in the content, then output None."""

        prompt = question + " " + str(text)
        response = client.chat.completions.create(model="gpt-4-turbo-preview",
                                                  messages=[
                                                      {"role": "system", "content": "你是一个助手。"},
                                                      {"role": "user", "content": prompt}
                                                  ],
                                                  temperature=0.5,
                                                  max_tokens=1000)

        # 解析响应的文本
        gpt_output = response.choices[0].message.content.strip()
        # 提取各个信息
        date = time = speaker = venue = affiliation = abstract = notes = ai = None
        if 'Date' not in df.columns:
            date_match = re.search(r'Date: (.*)', gpt_output)
            date = date_match.group(1) if date_match else 'None'
            extracted_data['Date'].append(date)

        if 'Time' not in df.columns:
            time_match = re.search(r'Time: (.*)', gpt_output)
            time = time_match.group(1) if time_match else 'None'
            extracted_data['Time'].append(time)

        if 'Speaker' not in df.columns:
            speaker_match = re.search(r'Speaker: (.*)', gpt_output)
            speaker = speaker_match.group(1) if speaker_match else 'None'
            extracted_data['Speaker'].append(speaker)

        if 'Venue' not in df.columns:
            venue_match = re.search(r'Venue: (.*)', gpt_output)
            venue = venue_match.group(1) if venue_match else 'None'
            extracted_data['Venue'].append(venue)

        if 'Affiliation' not in df.columns:
            affiliation_match = re.search(r'Affiliation: (.*)', gpt_output)
            affiliation = affiliation_match.group(1) if affiliation_match else 'None'
            extracted_data['Affiliation'].append(affiliation)

        if 'Notes' not in df.columns:
            Notes_match = re.search(r'Notes: (.*)', gpt_output, re.DOTALL)
            Notes_match = re.search(r'Notes: (.*?)(?=7|$)', gpt_output, re.DOTALL)
            Notes = Notes_match.group(1).strip() if Notes_match else 'None'
            extracted_data['Notes'].append(Notes)

        if 'Abstract' not in df.columns:
            abstract_match = re.search(r'Abstract: (.*?)8', gpt_output, re.DOTALL)
            abstract = abstract_match.group(1).strip() if abstract_match else 'None'
            extracted_data['Abstract'].append(abstract)

        if 'AI' not in df.columns:
            AI_match = re.search(r'AI: (.*)', gpt_output, re.DOTALL)
            AI = AI_match.group(1).strip() if AI_match else 'None'
            extracted_data['AI'].append(AI)
    # 把提取的信息添加到DataFrame
    for key, value in extracted_data.items():
        df[key] = value

    # 将更新后的DataFrame保存为新的Excel文件
    updated_file_path = file_path.replace('.xlsx', '_2.xlsx')
    df.to_excel(updated_file_path, index=False)

    return updated_file_path


def process_excel(excel_files, column_order, column_to_delete, output_dir=''):
    """
    读取Excel文件，删除指定的列，按照给定的顺序排列其他列，并输出到新的Excel文件。

    参数:
    excel_files: 文件路径列表。
    column_order: 指定新Excel文件中列的顺序列表。
    column_to_delete: 需要删除的列名。
    output_dir: 输出文件的目录路径。如果为空，则在当前路径输出。
    返回:
    无
    """
    # 创建输出目录的Path对象
    output_dir_path = Path(output_dir)
    # 确保输出目录存在
    output_dir_path.mkdir(parents=True, exist_ok=True)

    # 初始化一个空的列表用于收集每个文件处理后的DataFrame
    dataframes = []

    for file_name in excel_files:
        # 读取每个Excel文件到DataFrame
        df = pd.read_excel(file_name)

        # 如果指定列存在，则删除
        if column_to_delete in df.columns:
            df.drop(columns=[column_to_delete], inplace=True)

        # 排除不在COLUMN_ORDER中的列，避免索引错误
        df = df.reindex(columns=column_order)

        # 保存排列后的DataFrame到新的Excel文件
        sorted_file_name = output_dir_path / f"sorted_{Path(file_name).name}"
        df.to_excel(sorted_file_name, index=False)

        # 添加当前处理过的DataFrame到列表中
        dataframes.append(df)

    # 使用pandas.concat来合并所有的DataFrame
    combined_df = pd.concat(dataframes, ignore_index=True)

    # 保存汇总的DataFrame到一个Excel文件
    combined_sorted_file_path = output_dir_path / 'combined_sorted.xlsx'
    combined_df.to_excel(combined_sorted_file_path, index=False)



api_key = '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'
file_paths =['/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_112/1.xlsx',
             '/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_113/3.xlsx',
             '/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_115/4.xlsx',
             '/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_116/5.xlsx']

for file_path in file_paths:
    new_file_path = extract_excel(file_path, api_key)
    print(f"处理后的文件保存在：{new_file_path}")

Column_order = ['Published', 'Date', 'Time', 'Venue', 'Speaker', 'Affiliation', 'Title', 'Series',
                'Abstract', 'Note', 'AI', 'Host', 'Link']  # 你要保留的列的新顺序
Column_to_delete = 'Content'  # 你要删除的列名字
Excel_files = ['/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_112/1_2.xlsx',
               '/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_113/3_2.xlsx',
               '/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_115/4_2.xlsx',
               '/Users/dengyuchen/Downloads/EasySpider_MacOS/Data/Task_116/5_2.xlsx']  # Excel文件列表
Output_dir = '/Users/dengyuchen/Desktop/爬虫/final'  # 输出文件的目录
process_excel(Excel_files, Column_order, Column_to_delete, Output_dir)
# 读取Excel文件
Output_dir = Path('/Users/dengyuchen/Desktop/爬虫/final')
file_path = Output_dir / 'combined_sorted.xlsx'
df = pd.read_excel(file_path)
# 用来尝试解析日期的函数
def try_parse_date(date_str):
    try:
        return parser.parse(date_str, ignoretz=True)
    except (ValueError, TypeError):
        return None

# 应用该函数并创建一个新的“Parsed Date”列用于排序和格式化
df['Parsed Date'] = df['Date'].apply(lambda d: try_parse_date(d))
# 分割 DataFrame 为两部分：可以解析日期的和无法解析日期的
date_rows = df.dropna(subset=['Parsed Date'])
non_date_rows = df[df['Parsed Date'].isna()]
# 对可以解析日期的部分进行排序
date_rows = date_rows.sort_values(by='Parsed Date', ascending=False)
# 更新原始的日期列，将其格式化为指定格式
date_rows['Date'] = date_rows['Parsed Date'].apply(lambda d: d.strftime('%d/%m/%Y'))
# 将未能解析的行附加到末尾
result_df = pd.concat([date_rows, non_date_rows], ignore_index=True)
# 删除“Parsed Date”辅助列
# result_df = result_df.drop(columns=['Parsed Date'])
# result_df = result_df.assign(DateModified=np.nan)
# 输出文件
now = datetime.datetime.now()
filename = now.strftime("%Y-%m-%d-%H-%M-%S-raw-seminar-list.xlsx")
# 设置完整的文件路径
filepath = os.path.join('/Users/dengyuchen/Desktop/爬虫', filename)
result_df.to_excel(filepath, index=False)