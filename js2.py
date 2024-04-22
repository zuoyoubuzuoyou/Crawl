import pandas as pd


def convert_excel_to_js(file_path, output_file):
    # 加载Excel文件
    data = pd.read_excel(file_path)

    # 初始化JavaScript数组字符串
    js_array = "const seminars = [\n"

    for index, row in data.iterrows():
        # 为每一行数据创建JavaScript对象
        seminar_js = f"  {{\n"
        seminar_js += f"    name: '{row['Title']}',\n"
        seminar_js += f"    dateTime: '{row['Date']} {row['Time']}',\n"
        seminar_js += f"    location: '{row['Venue']}',\n"
        seminar_js += f"    speaker: '{row['Speaker']}',\n"
        seminar_js += f"    speakerLocation: '{row['Affiliation']}',\n"
        seminar_js += f"    host: '{row['Host']}',\n"
        seminar_js += f"    link: '{row['Link']}',\n"
        seminar_js += f"  }},\n"

        # 将生成的JavaScript对象添加到数组字符串中
        js_array += seminar_js

    # 结束数组定义
    js_array += "]\n"

    # 将生成的JavaScript代码写入到指定的文件
    with open(output_file, 'w') as file:
        file.write(js_array)


# 指定文件路径和输出文件路径
file_path = '/Users/dengyuchen/Downloads/new_2024-04-17-08-50-21-raw-seminar-list.xlsx'
output_file = '/Users/dengyuchen/Desktop/seminars.js'
# 调用函数
convert_excel_to_js(file_path, output_file)