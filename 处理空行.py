import pandas as pd

# 读取Excel文件
df = pd.read_excel('/Users/dengyuchen/Downloads/2024-04-17-08-50-21-raw-seminar-list.xlsm')

# 将每个单元格内的换行符替换为空格
df = df.apply(lambda x: x.map(lambda y: y.replace('\n', ' ') if isinstance(y, str) else y))
# 使用split函数分割"Speaker"列的每个单元格的内容，并取出第一个逗号前面的内容
df['Speaker'] = df['Speaker'].apply(lambda x: x.split(',')[0] if isinstance(x, str) else x)

# 保存修改后的数据到指定目录的新Excel文件
df.to_excel('/Users/dengyuchen/Downloads/new_2024-04-17-08-50-21-raw-seminar-list.xlsx', index=False)

