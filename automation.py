from WinGUI import WinGUI

# 标准库
import math       # 提供数学函数，例如三角函数、对数、幂运算等
import os         # 提供与操作系统交互的功能，如文件和目录管理
import shutil     # 提供高级的文件和目录操作，如复制、移动和删除
import time       # 提供时间相关的功能，如延迟、时间戳等

# 第三方库
import cv2 as cv       # OpenCV库，用于图像和视频处理
import keyboard        # 提供键盘事件处理的功能
import pyautogui       # 提供屏幕自动化控制，如鼠标点击、键盘输入、截图等
import win32con        # 包含Windows API常量，用于与Windows系统交互
import win32gui        # 提供与Windows GUI（图形用户界面）交互的功能
import pyperclip       # 处理剪贴板内容
import openpyxl

from PIL import ImageGrab     # 从PIL库导入ImageGrab模块，用于截图
from loguru import logger     # 引入loguru库，用于简便的日志记录

# 表格操作
import pandas as pd

# 循环执行 直到出现标志或者手动终止 需要修改
def running_loop(window_name, cycle_number=-1):
    """
    执行自动化程序：检测窗口状态，按循环次数或按键终止。
    
    :param window_name: 应用窗口的名称
    :param cycle_number: 循环次数，-1 表示无限循环 直到手动终止
    """
    exit_flag = False

    def on_key_event(event):
        # 当按键被按下时触发事件，如果按下 'q' 键，则终止循环
        nonlocal exit_flag
        if event.name == 'q':
            logger.info(f"END | terminated by user, windoe name: {window_name}") # 被用户终止
            exit_flag = True

    keyboard.on_press(on_key_event)  # 设置按键监听

    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    logger.info(f'START | window name: {window_name}')  # 记录窗口名称
    
    cycle_count = 0 # 初始化循环计数器
    while not exit_flag:
        try:
            if is_loop_over(app, 'icon.png'): # 检测是否结束
                logger.info(f"Cycle {cycle_count} is finished")  # 记录当前循环结束
                if cycle_number > 0 and cycle_count >= cycle_number:  # 检查是否达到设定的循环次数
                    logger.info(f"finished {cycle_count} cycles!")  # 记录完成循环次数
                    return
                

                cycle_count += 1  # 循环计数加一
        except Exception as err:
            logger.info(err)  # 记录异常信息

        
        time.sleep(0.5)   # 每次循环暂停1秒

# 判断循环是否结束 需要修改
def is_loop_over(app, icon):
    """
    检测指定图标是否出现在窗口中，以判断循环是否结束
    
    :param app: WinGUI 实例，提供窗口和图标检测功能
    :return: 如果测试结束则返回 True，否则返回 False
    """
    valid = app.check_icon(icon)  # 检测标志
    return valid

# 千牛 执行一次备注操作
def run_once_remarks_by_qianniu(window_name):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    # logger.info(f'START | window name: {window_name}')  # 记录窗口名称

    try:
        # app.get_app_screenshot()
        # 点击备注
        find_button_remarks = app.click_icon('Button_Remarks.png',0.5,1.0,0.4,1.0)
        if not find_button_remarks:
            logger.info(f"END | not find Button_Remarks.png, windoe name: {window_name}") # 停止记录
            return
        # 点击红色备注
        app.click_icon('Button_RedFlag.png',0.4,1.0,0.4,1.0)
        # 向下移动到输入框并点击
        app.rel_remove_and_click(0, 150)
        time.sleep(0.1)
        # 获取输入框当前的内容
        keyboard.press_and_release('ctrl+a')  
        keyboard.press_and_release('ctrl+c')  # 模拟按下并释放
        current_text = pyperclip.paste()      # 获取剪贴板中的文本
        time.sleep(0.1)
        entry_text = ''
        # 判断输入框是否有内容
        print('------')
        print(current_text)
        if current_text.strip():  # 如果字符串不为空
            entry_text = f'{current_text}\n已登记补发' if '已登记' in current_text else '已登记补发'
        else:
            entry_text = '已登记补发'
        print(entry_text)
        # 判断有无按下附加信息按钮
        # app.check_icon('button_additional_information.png'):
        time.sleep(0.1)
        # 输入
        keyboard.write(entry_text)
        # 点击确认
        # app.click_icon('Button_Confirm_Remarks.png',0.8,1.0,0.8,1.0)
        # 点击取消
        # app.click_icon('Button_Cancel_Remarks.png',0.8,1.0,0.8,1.0)

        # logger.info(f"END | terminated by program, windoe name: {window_name}") # 停止记录
    except Exception as err:
        logger.info(err)  # 记录异常信息

# 千牛 执行一次取消标记操作
def run_once_unmark_by_qianniu(window_name):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    # logger.info(f'START | window name: {window_name}')  # 记录窗口名称

    try:
        # app.get_app_screenshot()
        # 取消标记
        local_x, local_y, is_find = app.locate_icon('button_selected_session_annotation.png',0,0.4,0,1.0)
        if is_find:
            app.move_and_click(local_x, local_y, 'right')
            time.sleep(0.1)
            app.click_icon('button_cancel_annotations.png',0,0.4,0.2,1.0)
        else:
            logger.info(f"END | not find button_selected_session_annotation.png, windoe name: {window_name}") # 停止记录

        # logger.info(f"END | terminated by program, windoe name: {window_name}") # 停止记录
    except Exception as err:
        logger.info(err)  # 记录异常信息

# 万店通
# 订单管理界面 order_management_interface
# 历史订单界面 historical_order_interface
# 手工建单界面 manual_order_creation_interface

# 选中的订单    current_order
# 复制补发订单   copy_reissue_order
# 重复创建确认   repeat_creation_confirmation
# 是    repeat_creation_confirmation_yes

# 成交日期       transaction_date

# 已有订单   existing_orders

# 添加单品   add_single_item
# 货品名称    goods_name
# 编码    goods_code
# 保存   save_add_item 

# 客服备注   customer_service_remarks
def run_once_by_erp(window_name):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

    try:
        # 点击订单管理
        app.click_icon('order_management_interface.png',0.5,1.0,0.5,1.0)
        # 右键选中订单
        app.click_icon('current_order.png',0.5,1.0,0.5,1.0,'right')
        # 选择复制补发订单
        app.click_icon('copy_reissue_order.png',0.5,1.0,0.5,1.0)
        # 判断是否提示
        is_find = app.check_icon('repeat_creation_confirmation.png',0.5,1.0,0.5,1.0)
        if is_find:
            # 点击是
            app.click_icon('repeat_creation_confirmation_yes.png',0.5,1.0,0.5,1.0)
        # 延迟
        time.sleep(0.5)

        # 子窗口-创建补发订单
        child_window_reissue_order = WinGUI('创建补发订单')
        # 设置今天日期
        trans_data_icon_local_x, trans_data_icon_local_y, trans_data_icon_is_find = child_window_reissue_order.locate_icon('transaction_date.png',0.5,1.0,0.5,1.0)
        if trans_data_icon_is_find:
            # 向右移动并点击呼出日期下拉框
            child_window_reissue_order.move_and_click(trans_data_icon_local_x+50, trans_data_icon_local_y)
            # 相对向右下角移动点击今天
            child_window_reissue_order.rel_remove_and_click(50, 50)

        # 循环检测指定位置是否有1来判断是否已经存在商品 有则向右轻微移动到商品条目 双击删除商品 保证删除至没有商品
        while True:
            existence_orders_local_x, existing_orders_local_y, is_find_existence = child_window_reissue_order.locate_icon('existing_orders.png',0.5,1.0,0.5,1.0)
            if is_find_existence:
                # 向右轻微移动到商品条目 双击删除商品
                child_window_reissue_order.move_and_click(existence_orders_local_x+10, existing_orders_local_y)
                child_window_reissue_order.move_and_click(existence_orders_local_x+10, existing_orders_local_y)
                # 延迟
                time.sleep(0.5)
            else:
                break
        
        # 点击添加单品
        child_window_reissue_order.click_icon('add_single_item.png',0.5,1.0,0.5,1.0)
        # 子窗口-添加单品
        child_window_add_single_item = WinGUI('添加单品')
        # 定位到货品名称 并移动到输入框
        goods_name_local_x, goods_name_local_y, goods_name_local_is_find = child_window_add_single_item.locate_icon('goods_name.png',0.5,1.0,0.5,1.0)
        if goods_name_local_is_find:
            # 向下移动并点击输入框
            child_window_add_single_item.rel_remove_and_click(goods_name_local_x, goods_name_local_y+8)
            # 输入货品名称 备注
            keyboard.write('备注')
            # 回车
            keyboard.press_and_release('enter')

            # 双击指定编码商品添加
            goods_code_local_x, goods_code_local_y, goods_code_local_is_find = child_window_add_single_item.locate_icon('goods_code.png',0.5,1.0,0.5,1.0)
            if goods_code_local_is_find:
                # 双击商品名添加
                child_window_add_single_item.move_and_click(goods_code_local_x, goods_code_local_y)
                child_window_add_single_item.move_and_click(goods_code_local_x, goods_code_local_y)
                # 延迟
                time.sleep(0.5)
            
            # 点击保存
            child_window_add_single_item.click_icon('save_add_item.png',0.5,1.0,0.5,1.0)

        # 回到 子窗口-创建补发订单
        # 定位到客服备注 并移动到输入框
        customer_service_remarks_local_x, customer_service_remarks_local_y, customer_service_remarks_local_is_find = child_window_reissue_order.locate_icon('customer_service_remarks.png',0.5,1.0,0.5,1.0)
        if customer_service_remarks_local_is_find:
            # 向右移动点击输入框
            child_window_reissue_order.move_and_click(customer_service_remarks_local_x+10, customer_service_remarks_local_y)
            # 输入客服备注：补发
            keyboard.write('补发')
    except Exception as err:
        logger.info(err)



# 测试
def run_test(window_name):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    try:
        print()
    except Exception as err:
        logger.info(err)  # 记录异常信息

# 通知补发单号
def notification_reissue(window_name, table_name, shop_name=None, form_folder = './form'):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    try:
        table_file = os.path.join(form_folder, table_name)
        df = pd.read_excel(table_file, dtype={'原始单号': str})
        column_names = df.columns.tolist()  # 获取列名列表
        # 检查是否存在"是否通知"列，如果不存在则添加
        if '是否通知' not in column_names:
            df['是否通知'] = 0
            # 将更改写回 Excel 文件
            with pd.ExcelWriter(table_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, index=False)
            print("列 '是否通知' 已添加到表格中。")
        
        # 定义一个退出标志
        exit_flag = False
        
        # 定义按键监听事件
        def on_key_event(event):
            nonlocal exit_flag
            if event.name == 'q':
                logger.info("terminated by user")
                exit_flag = True
        
        keyboard.on_press(on_key_event)  # 设置按键监听

        # 限制 DataFrame 到前两行
        # df_subset = df.iloc
        # 逐行处理DataFrame
        for index, row in df.iterrows():
            if exit_flag:
                break  # 如果接收到退出信号，则终止循环
                
            # app.click_icon(shop_name_icon)
            # 检查是否已经通知
            if row['是否通知'] == 1:
                print('当前用户已通知')
                continue  # 如果已经通知，则跳过当前行
            # 获取原始单号和物流单号

            # 循环起始等待
            time.sleep(0.1)

            original_number = row['原始单号']
            logistics_number = row['物流单号']
            print(original_number)
            print(logistics_number)

            # original_number=123456789789

            app.get_app_screenshot()

            # app.move_and_click(750, 500)
            # time.sleep(0.5)
            # 模拟按下 alt+W 快捷键打开目标软件
            # keyboard.press_and_release('alt+c')
            # 等待软件响应
            # 模拟按下 ctrl+F 打开搜索功能
            # keyboard.press_and_release('ctrl+i')
            # time.sleep(0.5)
            # keyboard.press_and_release('ctrl+f')
            # time.sleep(0.5)

            # 容错 点击搜索框
            app.click_icon('button_search_cus.png',0,0.3,0,0.3)

            time.sleep(0.1)
            # 清除
            keyboard.press_and_release('ctrl+a')  
            keyboard.press_and_release('backspace')
            # 将 original_number 的内容输入到搜索框中
            # pyautogui.typewrite(original_number)
            # 将中文字符串复制到剪贴板
            pyperclip.copy(original_number)
            keyboard.press_and_release('ctrl+v') 
            
            # 等待搜索结果响应
            time.sleep(0.2)
            _, __, not_find_cus = app.locate_icon('not_find_customer.png',0, 0.4,0,0.6)
            # 判断搜索结果
            if not_find_cus:
                print(f"未搜索到结果，跳过 {original_number}")
                df.at[index, '是否通知'] = 0
                continue  # 未搜索到直接continue下一个
            else:
              print('搜索到指定用户，即将发送通知...')
            
            # 模拟按下回车键进入指定用户的聊天窗口
            keyboard.press_and_release('enter')
            time.sleep(0.2)

            # 按下 ctrl+J 定位到输入框中
            keyboard.press_and_release('ctrl+i')
            time.sleep(0.2)

            # 调用 app.locate_icon 传入图片名称，找到是否有某个图片存在
            # x, y, is_find = app.locate_icon('input_box_icon.png')
            # 判断 is_find 是否为 True
            # if not is_find:
            #     logger.err("无法定位到输入框，程序终止")
            #     break  # 无法定位到直接终止程序
            # 如果为 True，则相对向下移动 200 个像素后点击
            # app.remove_and_click(x, y + 200)

            # 清除
            keyboard.press_and_release('ctrl+a')  
            keyboard.press_and_release('backspace')
            time.sleep(0.2)
            # 将 亲 + logistics_number + 这是您的补发单号 请注意查收 这段内容输入到输入框中
            message = f"亲 {logistics_number} 这是您的补发单号 请注意查收"
            # 将中文字符串复制到剪贴板
            pyperclip.copy(message)
            # 使用 pyautogui.typewrite 粘贴剪贴板内容
            keyboard.press_and_release('ctrl+v')  
            # pyautogui.typewrite(message, paste=True)
            time.sleep(0.2)

            # 模拟按下回车发送消息
            keyboard.press_and_release('enter')
            # 将当前行的"是否通知"标记为1
            df.at[index, '是否通知'] = 1

            # 将更改写回 Excel 文件
            with pd.ExcelWriter(table_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, index=False)
                
            # 循环结束暂停
            time.sleep(0.3)
    except Exception as err:
        logger.info(err)

    keyboard.unhook_all()  # 移除所有按键监听
    return df
