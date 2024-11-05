import pandas as pd
from datetime import datetime
import os

# 表格路径设置
form_folder = './form'
input_filename = '11月5号补发单号.csv'
# 使用 datetime.now().strftime('%Y-%m-%d_%H%M%S') 来生成包含日期和时间的文件名
output_filename = f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}_补发物流单号整理.xlsx"

# 使用 os.path.join 组合路径和文件名
input_file = os.path.join(form_folder, input_filename)
output_file = os.path.join(form_folder, output_filename)

# 读取表格
encodings = ['utf-8', 'GBK', 'latin1', 'ISO-8859-1']
for encoding in encodings:
    try:
        df = pd.read_csv(input_file, encoding=encoding)
        break  # 如果成功读取，跳出循环
    except UnicodeDecodeError:
        continue  # 如果读取失败，尝试下一个编码

# 确保 df 不是 None，如果所有编码都失败，则抛出异常
if df is None:
    raise UnicodeDecodeError("都无法读取文件，请检查文件编码")

# 转换 "物流单号" 列为整数类型，保证无科学计数法和小数点
df['物流单号'] = df['物流单号'].apply(lambda x: str(int(x)) if pd.notnull(x) else x)

# 筛选两种店铺名称
shop1 = '余猫旗舰店-天猫'
shop2 = '潮洁居家日用旗舰店-天猫'

df_shop1 = df[df['店铺名称'] == shop1][['物流单号', '店铺名称']]
df_shop2 = df[df['店铺名称'] == shop2][['物流单号', '店铺名称']]

# 将筛选结果写入新的 Excel 文件
with pd.ExcelWriter(output_file) as writer:
    df_shop1.to_excel(writer, sheet_name=shop1, index=False)
    df_shop2.to_excel(writer, sheet_name=shop2, index=False)

print(f"文件已成功创建：{output_file}")