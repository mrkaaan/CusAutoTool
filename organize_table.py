import pandas as pd
from datetime import datetime
import os

# 原始表格处理 筛选制定店铺名称的物流单号
def process_original_table(
    input_filename, # 处理表格名称
    form_folder='./form',    # 表格路径设置
):
    # 精确日期文件名
    output_filename = f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}_补发物流单号.xlsx"

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
    return output_filename


# 二次处理表格 清洗原始单号
# input_filename = 'ERP二次导出表格.csv'
def process_original_number(input_filename, shop_name, form_folder = './form'):
    output_filename = f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}_{shop_name}_补发单号.xlsx"
  
    # 使用 os.path.join 组合路径和文件名
    input_file = os.path.join(form_folder, input_filename)
    output_file = os.path.join(form_folder, output_filename)

    # 读取表格
    encodings = ['utf-8', 'GBK', 'latin1', 'ISO-8859-1']
    df = None
    for encoding in encodings:
        try:
            df = pd.read_csv(input_file, encoding=encoding)
            break  # 如果成功读取，跳出循环
        except UnicodeDecodeError:
            continue  # 如果读取失败，尝试下一个编码

    # 确保 df 不是 None，如果所有编码都失败，则抛出异常
    if df is None:
        raise UnicodeDecodeError("都无法读取文件，请检查文件编码")

    # 清理“原始单号”列
    df["原始单号"] = (
        df["原始单号"]
        .astype(str)  # 转为字符串类型
        .str.replace(r"-\d+$", "", regex=True)  # 去除斜杠及其后面的数字
        .str.replace(r"^[\'\"]", "", regex=True)  # 去掉以单引号或双引号开头的字符
        .str.replace(r"[=“”\"\'']", "", regex=True)  # 去除指定符号
        .apply(lambda x: str(int(x)) if x.isdigit() else x)  # 确保只包含数字，去掉无效字符
    )
    # 转换 "物流单号" 列为整数类型，保证无科学计数法和小数点
    # df['物流单号'] = df['物流单号'].apply(lambda x: str(int(x)) if pd.notnull(x) else x)

    # 将处理后的结果写入新的 Excel 文件
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, index=False)

    print(f"文件已成功创建：{output_file}")
    return output_filename

if __name__ == '__main__':
    '''
    表格介绍

    企业微信中发的原始补发单号表格.csv 
        有部分问题 原始单号列 因为科学计数法以及单元格格式 导致后四位为0 同时物流单号的单元格格式也不对
        原始单号无法还原使用 需要使用物流单号去ERP中搜索补发单号
        所以该表格需经过process_original_table整理 形成新的表格: 日期_补发物流单号.xlsx
    日期_补发物流单号.xlsx
        该表格中包含两个sheet 一个是店铺1 一个是店铺2
        需要手动进入该表格中,复制物流单号到ERP中筛选店铺导出表格进行二次处理: 日期_店铺_ERP二次导出表格.csv
    日期_店铺_ERP二次导出表格.csv
        改标因为是筛选店铺 所以会有多个 即几个店铺有几个表 
        为了方便操作 需要手动重命名为: 日期_店铺_ERP二次导出表格.xlsx
        但是该表格还无法直接使用 因为原始单号包含有引号、等号等特殊字符，以及部分行无原始单号 需要把这些行补上客户网名
        然后调用process_original_number进行二次处理 再次形成新的表格: 日期_店铺_补发单号.xlsx
    日期_店铺_补发单号.xlsx
        该表格可以直接使用 到gui.py中调用notification_reissue,第二个参数传入表名 进行自动补发
    '''
    '''
    need change
        测试发现 最终的表格还是会有部分无原始单号、无用户名 需要手动用订单编号去补充上才能自动化
        
        尽可能自动化操作
    '''


    # 首次处理 群中表格 筛选制定店铺名称的物流单号 一个表有两个sheet 结果文件拿去ERP搜索后导出为新的表格
    # process_original_table('11月6日补发单号.csv')

    # 二次处理 ERP导出表格 清洗新表格的原始单号 结果文件执行自动化操作
    # 有多少个店铺 就调用多少次 
    processed_df = process_original_number('2024-11-07_余猫_ERP二次导出表格.csv','余猫') # 日期_店铺_ERP二次导出表格
    processed_df = process_original_number('2024-11-07_潮洁_ERP二次导出表格.csv','潮洁') # 日期_店铺_ERP二次导出表格
