import pandas as pd
from pathlib import Path

def remove_prefix_from_column(df, column_name, prefix, replace_once=False):
    """
    在指定列中去除指定的前缀。

    参数:
    df: pandas DataFrame，操作的数据框。
    column_name: str，需要处理的列名。
    prefix: str，需要去除的前缀。
    replace_once: bool，如果为True，只替换最前面的一个匹配项；如果为False，替换所有匹配项。

    返回:
    修改后的DataFrame。
    """
    if column_name in df.columns:
        if replace_once:
            # 只替换最前面的一个匹配项
            df[column_name] = df[column_name].apply(lambda x: x.replace(prefix, '', 1).strip() if pd.notnull(x) else x)
        else:
            # 替换所有匹配项
            df[column_name] = df[column_name].str.replace(prefix, '', regex=False)
    else:
        print(f"'{column_name}'列不存在于DataFrame中。")
    return df

def process_excel_columns(file_path, output_dir=''):
    """
    读取Excel文件，处理特定列，然后保存到新的Excel文件，允许指定输出目录。

    参数:
    file_path: str，待处理的Excel文件路径。
    output_dir: str，处理后的文件输出目录路径。如果为空，则在文件的当前目录输出。

    返回:
    无
    """
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 处理'Speaker'列和'Title'列
    df = remove_prefix_from_column(df, 'Speaker', 'Speaker:', replace_once=False)
    df = remove_prefix_from_column(df, 'Title', 'Title', replace_once=True)

    # 创建输出目录的Path对象，确保输出目录存在
    output_dir_path = Path(output_dir)
    output_dir_path.mkdir(parents=True, exist_ok=True)

    # 构建输出文件的完整路径
    output_file_path = output_dir_path / f'modified_{Path(file_path).name}'

    # 保存更改回一个新文件
    df.to_excel(output_file_path, index=False)

    print(f'完成处理，修改后的文件已保存为: {output_file_path}')

# 示例用法
file_path = '/Users/dengyuchen/Desktop/爬虫/2024-04-17-08-50-21-raw-seminar-list.xlsx'
output_dir = '/Users/dengyuchen/Desktop/爬虫'
process_excel_columns(file_path, output_dir)