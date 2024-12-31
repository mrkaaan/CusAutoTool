import pandas as pd
import datetime
import os
import pyperclip
from utils import show_toast

# 读取csv文件
def read_csv(input_file):
    # 读取表格，尝试多种编码
    encodings = ['utf-8', 'GBK', 'latin1', 'ISO-8859-1']
    df = None
    for encoding in encodings:
        try:
            df = pd.read_csv(input_file, encoding=encoding)
            break
        except UnicodeDecodeError:
            pass
    return df

# 读取excel文件
def read_excel(input_file, dtype=None):
    # 读取表格，尝试多种编码
    encodings = ['utf-8', 'GBK', 'latin1', 'ISO-8859-1']
    df = None
    for encoding in encodings:
        try:
            if not dtype:
                df = pd.read_excel(input_file, dtype=dtype, encoding=encoding)
            else:
                df = pd.read_excel(input_file, encoding=encoding)
            break
        except UnicodeDecodeError:
            pass
    return df

# 处理表格
def process_table(input_filename, form_folder='../form'):
    file_format = input_filename.split('.')[-1]
    # 生成精确日期的文件名
    output_filename = f"{datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')}_处理结果.xlsx"
    form_folder += f"/{datetime.datetime.now().strftime('%Y-%m-%d')}"

    # 创建文件夹（如果不存在）
    if not os.path.exists(form_folder):
        os.makedirs(form_folder)

    # 使用 os.path.join 组合路径和文件名
    input_file = os.path.join(form_folder, input_filename)
    output_file = os.path.join(form_folder, output_filename)

    # 输入文件不存在
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"文件不存在：{input_file}")

    # 读取表格，尝试多种编码
    if file_format == 'csv':
        df = read_csv(input_file)
    elif file_format in ['xlsx', 'xls']:
        # df = read_excel(input_file, None)
        df = pd.read_excel(input_file)
    else:
        raise ValueError(f"不支持的文件格式：{file_format}")

    # 确保 df 不是 None
    if df is None:
        raise UnicodeDecodeError("无法读取文件，请检查文件编码")

    # 初始化存储所有订单编号的集合
    all_order_numbers_set = set()
    shop_order_numbers = {}

    # 创建 ExcelWriter 对象
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        if isinstance(df, pd.DataFrame):
            # 单个 sheet 的情况
            sheets = {'Sheet1': df}
        else:
            # 多个 sheet 的情况
            sheets = df

        # 检查是否已经存在“全部店铺”sheet
        if '全部店铺' in sheets:
            all_shops_df = sheets.pop('全部店铺')
            all_order_numbers_set.update(all_shops_df['订单编号'].dropna().astype(str).tolist())
            shop_order_numbers['全部店铺'] = all_order_numbers_set.copy()

        # 处理每个 sheet
        for sheet_name, sheet_df in sheets.items():
            print(sheet_name)
            # 填充空值
            sheet_df["原始单号"].fillna('', inplace=True)
            sheet_df["物流单号"].fillna('', inplace=True)

            # 清理“原始单号”列
            sheet_df["原始单号"] = (
                sheet_df["原始单号"]
                .astype(str)  # 转为字符串类型
                .str.replace(r"^[\'\"]", "", regex=True)  # 去掉以单引号或双引号开头的字符
                .str.replace(r"[=“”\"\'']", "", regex=True)  # 去除指定符号
                .str.replace(r"-\d+$", "", regex=True)  # 去除斜杠及其后面的数字
            )

            # 清理“物流单号”列
            sheet_df["物流单号"] = (
                sheet_df["物流单号"]
                .astype(str)  # 转为字符串类型
                .str.replace(r"^[\'\"]", "", regex=True)  # 去掉以单引号或双引号开头的字符
                .str.replace(r"[=“”\"\'']", "", regex=True)  # 去除指定符号
                .str.replace(r"-\d+$", "", regex=True)  # 去除斜杠及其后面的数字
                .str.replace(r"\.0$", "", regex=True)  # 去除末尾的 ".0"
            )

            # 获取所有的店铺名称
            shops = sheet_df['店铺名称'].unique()
            print(sheet_name)

            # 如果是单个 sheet 且只有一个默认名称，按店铺名称分割数据
            if sheet_name == 'Sheet1' or  sheet_name == '天猫' and len(shops) > 1:
                for shop in shops:
                    # 选择当前店铺的数据
                    df_shop = sheet_df[sheet_df['店铺名称'] == shop]
                    # 写入工作表，工作表名为店铺名称
                    df_shop.to_excel(writer, sheet_name=shop, index=False)
                    
                    # 提取当前店铺的订单编号
                    # 判断有无订单编号列
                    if '订单编号' in df_shop.columns:
                        shop_orders = df_shop['订单编号'].dropna().astype(str).tolist()
                    else:
                        shop_orders = df_shop['原始单号'].dropna().astype(str).tolist()
                    shop_order_numbers[shop] = shop_orders
                    all_order_numbers_set.update(shop_orders)
                continue

            # 如果 sheet 名是店铺名，检查是否有其他店铺的数据
            if sheet_name in shops:
                # 检查是否有其他店铺的数据
                other_shops = [shop for shop in shops if shop != sheet_name]
                if other_shops:
                    # 移动其他店铺的数据到正确的 sheet
                    for other_shop in other_shops:
                        other_shop_df = sheet_df[sheet_df['店铺名称'] == other_shop]
                        if other_shop not in sheets:
                            sheets[other_shop] = other_shop_df
                        else:
                            sheets[other_shop] = pd.concat([sheets[other_shop], other_shop_df])
                        sheet_df = sheet_df[sheet_df['店铺名称'] != other_shop]
                        shops = sheet_df['店铺名称'].unique()

            # 将当前 sheet 的数据写入
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 提取当前 sheet 的订单编号
            sheet_orders = sheet_df['订单编号'].dropna().astype(str).tolist()
            all_order_numbers_set.update(sheet_orders)
            shop_order_numbers[sheet_name] = sheet_orders

        # 创建“全部店铺”sheet
        all_shops_df = pd.concat(sheets.values(), ignore_index=True)
        all_shops_df.to_excel(writer, sheet_name='全部店铺', index=False)

    # 将所有订单编号放入剪贴板中，保留换行符
    all_order_numbers = list(all_order_numbers_set)
    all_order_numbers_str = '\n'.join(all_order_numbers)
    pyperclip.copy(all_order_numbers_str)

    print(f"文件已成功创建：{output_file}")
    print(f"所有订单编号已复制到剪贴板")
    show_toast("提醒", f"文件已成功创建：{output_file}\n所有订单编号已复制到剪贴板")

    return output_filename, all_order_numbers, shop_order_numbers

# 测试
if __name__ == "__main__":
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

    process_table('11月17日天猫补发单号.csv')