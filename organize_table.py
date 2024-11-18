import pandas as pd
from datetime import datetime
import os
import pyperclip

# 修改后的表格处理函数
def process_original_table(
    input_filename,  # 处理表格名称
    form_folder='./form'  # 表格路径设置
):
    # 生成精确日期的文件名
    output_filename = f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}_首次处理.xlsx"
    form_folder += f"/{datetime.now().strftime('%Y-%m-%d')}"

    # 创建文件夹（如果不存在）
    if not os.path.exists(form_folder):
        os.makedirs(form_folder)

    # 使用 os.path.join 组合路径和文件名
    input_file = os.path.join(form_folder, input_filename)
    output_file = os.path.join(form_folder, output_filename)

    # 读取表格，尝试多种编码
    encodings = ['utf-8', 'GBK', 'latin1', 'ISO-8859-1']
    df = None
    for encoding in encodings:
        try:
            df = pd.read_csv(input_file, encoding=encoding)
            break  # 成功读取后跳出循环
        except UnicodeDecodeError:
            continue  # 尝试下一个编码

    # 确保 df 不是 None
    if df is None:
        raise UnicodeDecodeError("无法读取文件，请检查文件编码")

    # 获取所有的店铺名称
    shops = df['店铺名称'].unique()

    # 存储所有订单编号
    all_order_numbers = []

    # 存储每个店铺的订单编号
    shop_order_numbers = {}

    # 将每个店铺的数据写入单独的工作表
    with pd.ExcelWriter(output_file) as writer:
        for shop in shops:
            # 选择当前店铺的数据
            df_shop = df[df['店铺名称'] == shop]
            # 写入工作表，工作表名为店铺名称
            df_shop.to_excel(writer, sheet_name=shop, index=False)
            
            # 提取当前店铺的订单编号
            shop_orders = df_shop['订单编号'].dropna().astype(str).tolist()
            shop_order_numbers[shop] = shop_orders
            all_order_numbers.extend(shop_orders)

    # 将所有订单编号放入剪贴板中，保留换行符
    all_order_numbers_str = '\n'.join(all_order_numbers)
    pyperclip.copy(all_order_numbers_str)

    print(f"文件已成功创建：{output_file}")
    print(f"所有订单编号已复制到剪贴板：\n{all_order_numbers_str}")

    return output_filename, all_order_numbers, shop_order_numbers

# 二次处理表格 清洗原始单号
def process_original_number(input_filename, form_folder='./form'):
    # 生成精确日期的文件名
    output_filename = f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}_补发单号.xlsx"
    form_folder += f"/{datetime.now().strftime('%Y-%m-%d')}"

    # 创建文件夹（如果不存在）
    if not os.path.exists(form_folder):
        os.makedirs(form_folder)

    # 使用 os.path.join 组合路径和文件名
    input_file = os.path.join(form_folder, input_filename)
    output_file = os.path.join(form_folder, output_filename)

    # 读取表格，尝试多种编码
    encodings = ['utf-8', 'GBK', 'latin1', 'ISO-8859-1']
    df = None
    for encoding in encodings:
        try:
            df = pd.read_csv(input_file, encoding=encoding)
            break  # 成功读取后跳出循环
        except UnicodeDecodeError:
            continue  # 尝试下一个编码

    # 确保 df 不是 None
    if df is None:
        raise UnicodeDecodeError("无法读取文件，请检查文件编码")

    # 填充空值
    df["原始单号"].fillna('', inplace=True)
    df["物流单号"].fillna('', inplace=True)

    # 清理“原始单号”列
    df["原始单号"] = (
        df["原始单号"]
        .astype(str)  # 转为字符串类型
        .str.replace(r"^[\'\"]", "", regex=True)  # 去掉以单引号或双引号开头的字符
        .str.replace(r"[=“”\"\'']", "", regex=True)  # 去除指定符号
        .str.replace(r"-\d+$", "", regex=True)  # 去除斜杠及其后面的数字
        # .apply(lambda x: re.sub(r"\.0$", "", x))  # 去除小数点后的0
    )

    # 清理“物流单号”列
    df["物流单号"] = (
        df["物流单号"]
        .astype(str)  # 转为字符串类型
        .str.replace(r"^[\'\"]", "", regex=True)  # 去掉以单引号或双引号开头的字符
        .str.replace(r"[=“”\"\'']", "", regex=True)  # 去除指定符号
        .str.replace(r"-\d+$", "", regex=True)  # 去除斜杠及其后面的数字
        # .apply(lambda x: re.sub(r"\.0$", "", x))  # 去除小数点后的0
        #  .apply(lambda x: str(int(x)) if x.isdigit() else x)  # 确保只包含数字，去掉无效字符
    )
    
    # 转换 "物流单号" 列为整数类型，保证无科学计数法和小数点
    # df['物流单号'] = df['物流单号'].apply(lambda x: str(int(x)) if pd.notnull(x) else x)

    # 获取所有的店铺名称
    shops = df['店铺名称'].unique()

    # 将每个店铺的数据写入单独的工作表
    with pd.ExcelWriter(output_file) as writer:
        if len(shops) == 1:
            # 如果只有一个店铺，使用默认的 sheet 名
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        else:
            # 如果有多个店铺，使用不同的 sheet 名来区分不同店铺的数据
            for shop in shops:
                df_shop = df[df['店铺名称'] == shop]
                df_shop.to_excel(writer, sheet_name=shop, index=False)

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
    # processed_df = process_original_number('2024-11-07_余猫_ERP二次导出表格 - test.csv') # 日期_店铺_ERP二次导出表格
    # processed_df = process_original_number('2024-11-07_潮洁_ERP二次导出表格.csv') # 日期_店铺_ERP二次导出表格

    # process_original_table('11月6日补发单号.csv')
