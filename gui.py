from WinGUI import WinGUI
import organize_table as tb

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

from PIL import ImageGrab     # 从PIL库导入ImageGrab模块，用于截图
from loguru import logger     # 引入loguru库，用于简便的日志记录

# 表格操作
import pandas as pd
from datetime import datetime

# json操作
import json

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
            if is_loop_over(app): # 检测是否结束
                logger.info(f"Cycle {cycle_count} is finished")  # 记录当前循环结束
                if cycle_number > 0 and cycle_count >= cycle_number:  # 检查是否达到设定的循环次数
                    logger.info(f"finished {cycle_count} cycles!")  # 记录完成循环次数
                    return
                

                cycle_count += 1  # 循环计数加一
        except Exception as err:
            logger.info(err)  # 记录异常信息

        
        time.sleep(0.5)   # 每次循环暂停1秒

# 判断循环是否结束 需要修改
def is_loop_over(app):
    """
    检测指定图标是否出现在窗口中，以判断循环是否结束
    
    :param app: WinGUI 实例，提供窗口和图标检测功能
    :return: 如果测试结束则返回 True，否则返回 False
    """
    valid1, _, _ = app.check_icon("running_1.png")  # 检测标志
    if valid1:
        return False
    
    return not valid1   # 如果标志不存在，返回 True 表示测试结束

# 执行一次备注操作
def run_once_beizhu(window_name):
    """
    执行自动化程序：检测窗口状态，执行操作。
    
    :param window_name: 应用窗口的名称
    """

    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    logger.info(f'START | window name: {window_name}')  # 记录窗口名称

    try:
        # app.get_app_screenshot()
        # 点击备注
        find_button_remarks = app.click_icon('Button_Remarks.png',0.5,1.0,0.4,1.0)
        # 点击红色备注
        app.click_icon('Button_RedFlag.png',0.4,1.0,0.4,1.0)
        # 向下移动到输入框并点击
        app.rel_remove_and_click(0, 150)

        # 获取输入框当前的内容
        keyboard.press_and_release('ctrl+a')  
        keyboard.press_and_release('ctrl+c')  # 模拟按下并释放
        current_text = pyperclip.paste()      # 获取剪贴板中的文本

        # 判断输入框是否有内容
        print('------')
        print(current_text)
        if current_text.strip():  # 如果字符串不为空
            if '已登记' in current_text:  # 如果内容包含“已登记”
                pyautogui.typewrite(['right'])  # 回车后输入
                pyautogui.typewrite(['enter'])  # 回车后输入
            else:  # 如果内容不包含“已登记”
                keyboard.press_and_release('ctrl+a')  
                keyboard.press_and_release('backspace')
        else:  # 如果无内容
            pass  # 不需要做任何操作，直接输入
        
        # 判断有无按下附加信息按钮
        # app.check_icon('button_additional_information.png'):

        # 输入
        keyboard.write('已登记补发')
        # 点击确认
        app.click_icon('Button_Confirm_Remarks.png',0.8,1.0,0.8,1.0)

        # 取消标记位置
        app.click_icon('button_selected_session_annotation.png',0,0.4,0.2,1.0,'right')
        app.click_icon('button_cancel_annotations.png',0,0.5,0.2,1.0)

        logger.info(f"END | terminated by program, windoe name: {window_name}") # 停止记录
    except Exception as err:
        logger.info(err)  # 记录异常信息

# 测试
def run_test(window_name):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    try:
        print()
    except Exception as err:
        logger.info(err)  # 记录异常信息
    

# 模拟搜索结果
def simulate_search_result(original_number):
    # 模拟搜索结果判断，实际应用中应替换为具体的搜索逻辑
    # 这里随机返回True或False来模拟是否搜索到结果
    import random
    return random.choice([True, False])

# 通知补发单号
def notification_reissue(window_name, table_name, form_folder = './form'):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    try:

        table_file = os.path.join(form_folder, table_name)
        df = pd.read_excel(table_file)  # 读取 Excel 文件
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
        # 逐行处理DataFrame
        for index, row in df.iterrows():
            if exit_flag:
                break  # 如果接收到退出信号，则终止循环
                
            # 检查是否已经通知
            if row['是否通知'] == 1:
                continue  # 如果已经通知，则跳过当前行
            # 获取原始单号和物流单号

            time.sleep(0.5)

            original_number = row['原始单号']
            logistics_number = row['物流单号']
            print(original_number)
            print(logistics_number)

            app.get_app_screenshot()

            app.move_and_click(750, 500)
            time.sleep(0.5)
            
            # 模拟按下 alt+W 快捷键打开目标软件
            # keyboard.press_and_release('alt+c')
            # 等待软件响应
            # 模拟按下 ctrl+F 打开搜索功能
            keyboard.press_and_release('ctrl+i')
            time.sleep(0.5)
            keyboard.press_and_release('ctrl+f')


            # 等待搜索框出现
            time.sleep(0.5)
            # 清除
            keyboard.press_and_release('ctrl+a')  
            keyboard.press_and_release('backspace')

            # 将 original_number 的内容输入到搜索框中
            # pyautogui.typewrite(original_number)

            # 将中文字符串复制到剪贴板
            pyperclip.copy(original_number)
            time.sleep(0.1)
            keyboard.press_and_release('ctrl+v') 
            # 等待聊天窗口响应
            time.sleep(0.7)


            _, __, is_find_cumtomer = app.locate_icon('not_find_customer.png')
            # 判断搜索结果
            if not is_find_cumtomer:
                print(f"未搜索到结果，跳过 {original_number}")
                df.at[index, '是否通知'] = 0
                continue  # 未搜索到直接continue下一个
            
            time.sleep(0.7)
            # 模拟按下回车键进入指定用户的聊天窗口
            keyboard.press_and_release('enter')
            # 按下 ctrl+J 定位到输入框中
            keyboard.press_and_release('ctrl+i')

            # 调用 app.locate_icon 传入图片名称，找到是否有某个图片存在
            # x, y, is_find = app.locate_icon('input_box_icon.png')
            # 判断 is_find 是否为 True
            # if not is_find:
            #     logger.err("无法定位到输入框，程序终止")
            #     break  # 无法定位到直接终止程序
            # 如果为 True，则相对向下移动 200 个像素后点击
            # app.remove_and_click(x, y + 200)

            time.sleep(0.5)
            # 清除
            keyboard.press_and_release('ctrl+a')  
            keyboard.press_and_release('backspace')
            time.sleep(0.5)
            # 将 亲 + logistics_number + 这是您的补发单号 请注意查收 这段内容输入到输入框中
            message = f"亲 {logistics_number} 这是您的补发单号 请注意查收"
            print(message)
            time.sleep(0.5)

            # 将中文字符串复制到剪贴板
            pyperclip.copy(message)
            time.sleep(0.5)
            # 使用 pyautogui.typewrite 粘贴剪贴板内容
            # keyboard.press_and_release('ctrl+v')  

            # pyautogui.typewrite(message, paste=True)
            # time.sleep(0.2)

            # 模拟按下回车发送消息
            # keyboard.press_and_release('enter')

            # 将当前行的"是否通知"标记为1
            # df.at[index, '是否通知'] = 1

            # 将更改写回 Excel 文件
            # with pd.ExcelWriter(table_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # df.to_excel(writer, index=False)
            time.sleep(0.3)
    except Exception as err:
        logger.info(err)

    keyboard.unhook_all()  # 移除所有按键监听
    return df

# 存储句柄
def save_handle(name, handle):
    handles = load_handles()
    handles[name] = handle
    with open('handles.json', 'w') as f:
        json.dump(handles, f)

# 加载句柄
def load_handle(name):
    handles = load_handles()
    return handles.get(name)

# 关闭句柄
def load_handles():
    if os.path.exists('handles.json'):
        with open('handles.json', 'r') as f:
            return json.load(f)
    else:
        return {}

# 打开指定软件
def open_sof(name):
    handle = load_handle(name)
    if handle is None:
        handle = win32gui.FindWindow(0, name)
        if handle == 0:
            return None
        save_handle(name, handle)
    else:
        win32gui.SendMessage(handle, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)  # 发送消息以确保窗口恢复显示
        time.sleep(0.2) # 延迟1秒，等待窗口恢复
    
        win32gui.ShowWindow(handle, True)          # 使用win32gui库使窗口可见
        win32gui.SetForegroundWindow(handle)       # 使用win32gui库将窗口置于前台(焦点)
        win32gui.SendMessage(handle, win32con.SC_MAXIMIZE, 0) # 最大化窗口
        time.sleep(0.2)                              # 使用time库延迟1秒
        
        pos = win32gui.GetWindowRect(handle) # 窗口位置

        print(f'窗口位置 {pos}') # 打印窗口位置

# 快捷键
def auto_key(window_name):
    # 点击添加备注
    keyboard.add_hotkey('ctrl+shift+1', lambda: run_once_beizhu(window_name)) 
    # 保持脚本运行，直到按下退出快捷键
    keyboard.wait('esc')  # 按下 Esc 退出监听
    os.remove('handles.json')  # 清空句柄文件


if __name__ == "__main__":
    pyautogui.FAILSAFE = False  # 关闭 pyautogui 的故障保护机制
    pyautogui.PAUSE = 0.1  # 设置 pyautogui 的默认操作延时 如移动和点击的间隔

    logger.add("dev.log", rotation="10 MB")  # 设置日志文件轮换

    # ---------- 配置参数 -------------
    # original_folder = "C:/Users/Public/Documents/Data"  # 源文件夹路径
    # target_folder = r"C:\Users\Joey\Desktop\data"  # 目标文件夹路径
    # suffix_list = []  # 要移动的文件后缀列表
    # cycle_number = -1  # 循环次数，-1 表示无限循环

    window_name = r"千牛接待台"  # 窗口名称
    # ------------------------------------

    # 模拟键盘
    # keyboard.press_and_release('ctrl+shift+esc')  # 模拟按下并释放 Ctrl+Shift+Esc 组合键
    # 输入
    # keyboard.write('Hello, World!')  # 
    # ------------------------------------
    
    # 快捷键启动
    # esc退出
    # shift+ctrl+1 添加备注
    # auto_key(window_name)

    # 测试 目前为空
    # run_test(window_name)

    # 通知补发单号
    # 这里的表格必须经过格式化，有整理过后的原始单号以及物流单号
    # change 单号-2
    # change 多少个补发 多少个没有补发 全部补发了的提示
    # not_find_customer 未搜到处理办法
    # notification_reissue(window_name, '2024-11-06_220841_余猫_补发单号.xlsx')

    # 打开以启动软件
    open_sof('ToDesk')