from WinGUI import WinGUI
import utils as ul
import os   
import time     
import datetime
import keyboard   
import pyautogui   
import pyperclip   
import pandas as pd
from utils import show_toast
from loguru import logger  
import tkinter as tk
import json
from pathlib import Path
import random
import threading
import configparser


rule_json_path = r'../config/product_rules.json'

# 定义全局变量来存储坐标信息
coordinates = {}
coordinate_json_path = r'../config/coordinates.json'

def read_coordinate_by_key(key, reissue=True):
    global coordinates, coordinate_json_path
    if not coordinates:
        load_coordinates_from_json(coordinate_json_path, reissue)

    # if reissue:
    #     key_name = 'coordinates_by_reissue'
    # else:
    #     key_name = 'coordinates'
    # print(coordinates.keys())
    # 检查关键字段是否存在且非空
    # if not (key_name in coordinates):
    #     print(f"Key '{key_name}' does not exist in the coordinates.")
    #     return None
    # if coordinates[key_name] is None:
    #     print(f"Value for key '{key_name}' is None in the coordinates.")
    #     return None
    # specific_coordinates = coordinates[key_name]

    specific_coordinates = coordinates
    if key in specific_coordinates and specific_coordinates[key] is not None:
        position = specific_coordinates[key].get('position')
        if position is not None:
            return position
        else:
            print(f"{key}的坐标信息为空")
    else:
        print(f"{key}信息不存在")
        return None

def load_coordinates_from_json(file_path, reissue=True):
    """
    从指定路径加载JSON文件，并填充全局变量 `coordinates`。
    如果文件不存在或内容为空，则打印警告信息。
    """
    global coordinates
    if reissue:
        key_name = 'coordinates_by_reissue'
    else:
        key_name = 'coordinates'
    try:
        # 检查文件是否存在
        # if not os.path.exists(file_path):
        #     print("文件不存在")
        #     return None
        if not Path(file_path).is_file():
            print("文件不存在")
            return

        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)

        if not data or 'coordinates' not in data:
            print("Warning: JSON file is empty or does not contain 'coordinates'.")
            return

        coordinates = data[key_name]

        # 验证每个坐标点是否包含位置和描述
        for key, value in coordinates.items():
            if not all(k in value for k in ('position', 'description')):
                print(f"Warning: Invalid format for coordinate '{key}'.")
                coordinates[key] = {'position': None, 'description': None}

    except json.JSONDecodeError:
        print("文件不是有效的JSON格式")
    except Exception as e:
        print(f"An error occurred while loading coordinates: {e}")

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

def running_loop_test(cycle_number=-1):
    """
    :param cycle_number: 循环次数，-1 表示无限循环 直到手动终止
    """
    # 假设这里打开文件

    exit_flag = False

    def set_exit_flag():
        nonlocal exit_flag
        print(f"END | terminated by user")
        exit_flag = True

    # 设置组合键监听
    keyboard.add_hotkey('shift+ctrl+q', set_exit_flag)

    cycle_count = 0  # 初始化循环计数器
    try:
        while not exit_flag:
            print(f"Cycle {cycle_count} is finished")
            if cycle_number > 0 and cycle_count >= cycle_number:  # 检查是否达到设定的循环次数
                logger.info(f"finished {cycle_count} cycles!")  # 记录完成循环次数
                return

            print(f'do something {cycle_count}')
            time.sleep(0.3)
            
            print(f'do something {cycle_count}')
            time.sleep(0.3)
            
            print(f'do something {cycle_count}')
            time.sleep(0.3)
            
            print(f'do something {cycle_count}')
            time.sleep(0.3)    
            cycle_count += 1  # 循环计数加一


            # 省略一系列复杂操作
            # 操作末尾给当前数据打上处理完毕标记

            time.sleep(0.5)  # 每次循环暂停1秒

    except KeyboardInterrupt:
        print("检测到 Ctrl+C，正在退出...")
    except Exception as e:
        print(f"快捷键监听出错：{e}")
    finally:
        # 回写文件
        
        # 移除所有快捷键监听
        keyboard.unhook_all()
        print('退出监听')

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
def run_once_remarks_by_qianniu(window_name, click_remarks=True, unmark=True, unmark_mode=1):
    '''
        :param window_name: 应用窗口的名称
        :param unmark: 是否取消标记
        :param unmark_mode: 取消标记模式 仅在 unmark=True 时有效  1 使用快捷键取消标记 2 使用鼠标点击取消标记
    '''
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    # logger.info(f'START | window name: {window_name}')  # 记录窗口名称

    try:
        # app.get_app_screenshot()
        # 点击备注
        if click_remarks:
            button_remarks_x, button_remarks_y, find_button_remarks = app.locate_icon('Button_Remarks.png',0.7,1.0,0.2,0.9)
            if not find_button_remarks:
                logger.info(f"END | not find Button_Remarks.png, windoe name: {window_name}") # 停止记录
                return
            app.move_and_click(button_remarks_x, button_remarks_y)
        # 点击红色备注
        redflag_x, redflag_y, is_find_redflag = app.locate_icon('Button_RedFlag.png',0.5,1.0,0.6,1.0)
        if not is_find_redflag:
            logger.info(f"END | not find Button_RedFlag.png, windoe name: {window_name}") # 停止记录
            return
        app.move_and_click(redflag_x, redflag_y)
        # 向下移动到输入框并点击
        app.move_and_click(redflag_x, redflag_y+150)
        # 获取输入框当前的内容
        keyboard.press_and_release('ctrl+a')  
        # 按下右边取消选中
        keyboard.press_and_release('right')
        entry_text = '已登记补发'
        # 输入
        keyboard.write(entry_text)
        # 点击确认
        app.click_icon('Button_Confirm_Remarks.png',0.8,1.0,0.8,1.0)
        # 点击取消
        # app.click_icon('Button_Cancel_Remarks.png',0.8,1.0,0.8,1.0)
        # 取消备注
        if unmark:
            run_once_unmark_by_qianniu(window_name=None, mode=unmark_mode, app=app)
    except Exception as err:
        logger.info(err)  # 记录异常信息

# 千牛 执行一次取消标记操作
def run_once_unmark_by_qianniu(window_name, mode=1, app=None):
    '''
        :param window_name: 应用窗口的名称
        :param mode: 1 使用快捷键取消标记 2 使用鼠标点击取消标记
        :param app: WinGUI 实例  使用时默认为被其他函数调用时传入 window_name 这时为空
    '''
    # 创建 WinGUI 实例
    if not app:
        app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    else:
        # 如果传入 app 则使用传入的 app 实例 并截图
        app.get_app_screenshot()
    try:
        # app.get_app_screenshot()
        # 取消标记
        if mode == 1:
            time.sleep(0.1)
            keyboard.press_and_release('ctrl+i')
            time.sleep(0.1)
            # 按下三次 ctrl+w 取消标记
            keyboard.press_and_release('ctrl+w')
            time.sleep(0.1)
            keyboard.press_and_release('ctrl+w')
            time.sleep(0.1)
            keyboard.press_and_release('ctrl+w')
            time.sleep(0.1)
            keyboard.press_and_release('ctrl+w')
            time.sleep(0.1)
        elif mode == 2:
            # 千牛左侧列表有两个状态 正在接待列表 和  全部买家列表， 两个列表中被选中的用户的图案不同 判断在哪个列表中
            select_all_buyers = app.check_icon('select_all_buyers.png', 0, 0.3, 0.1, 0.4)
            if not select_all_buyers:
                local_x, local_y, is_find = app.locate_icon('button_selected_session_annotation.png',0,0.3,0.1,0.9)
                if is_find:
                    app.move_and_click(local_x, local_y, 'right')
                    # 按下上方向键
                    keyboard.press_and_release('up')
                    # 按下回车
                    keyboard.press_and_release('enter')
                    # time.sleep(0.1)
                    # button_cancel_x, button_cancel_y, is_find_buuton_canbel = app.locate_icon('button_cancel_annotations.png',0,0.4,0.2,1.0)
                    # if is_find_buuton_canbel:
                    #     app.move_and_click(button_cancel_x, button_cancel_y)
                else:
                    logger.info(f"END | not find button_selected_session_annotation.png, windoe name: {window_name}") # 停止记录
            else:
                # 找 button_selected_session_annotation.png
                local_other_x, local_other_y, is_find_other = app.locate_icon('button_selected_session_annotation_other.png',0,0.4,0,1.0)
                if is_find_other:
                    app.move_and_click(local_other_x, local_other_y, 'right')
                    # 按下上方向键
                    keyboard.press_and_release('up')
                    # 按下回车
                    keyboard.press_and_release('enter')
                    # time.sleep(0.1)
                    # button_cancel_x, button_cancel_y, is_find_buuton_canbel = app.locate_icon('button_cancel_annotations.png',0,0.4,0.2,1.0)
                    # if is_find_buuton_canbel:
                    #     app.move_and_click(button_cancel_x, button_cancel_y)
                else:
                    logger.info(f"END | not find button_selected_session_annotation.png, windoe name: {window_name}") # 停止记录

    except Exception as err:
        logger.info(err)  # 记录异常信息

def run_once_copy_username_by_qianniu(window_name):
    '''
        :param window_name: 应用窗口的名称
    '''
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

    try:
        local_x, local_y, is_find = app.locate_icon('button_selected_session_annotation.png',0,0.3,0.1,0.9)
        if is_find:
            app.move_and_click(local_x, local_y, 'right')
            # 按下上方向键
            keyboard.press_and_release('down')
            # 按下回车
            keyboard.press_and_release('enter')
            # time.sleep(0.1)
            # button_cancel_x, button_cancel_y, is_find_buuton_canbel = app.locate_icon('button_cancel_annotations.png',0,0.4,0.2,1.0)
            # if is_find_buuton_canbel:
            #     app.move_and_click(button_cancel_x, button_cancel_y)
        else:
            # 找 button_selected_session_annotation.png
            local_other_x, local_other_y, is_find_other = app.locate_icon('button_selected_session_annotation_other.png',0,0.4,0,1.0)
            if is_find_other:
                app.move_and_click(local_other_x, local_other_y, 'right')
                # 按下上方向键
                keyboard.press_and_release('down')
                # 按下回车
                keyboard.press_and_release('enter')
                # time.sleep(0.1)
                # button_cancel_x, button_cancel_y, is_find_buuton_canbel = app.locate_icon('button_cancel_annotations.png',0,0.4,0.2,1.0)
                # if is_find_buuton_canbel:
                #     app.move_and_click(button_cancel_x, button_cancel_y)
            else:
                logger.info(f"END | not find button_selected_session_annotation.png, windoe name: {window_name}") # 停止记录
    except Exception as err:
        logger.info(err)

# 测试
def run_test(window_name):
    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    try:
        print()
    except Exception as err:
        logger.info(err)  # 记录异常信息


def wait_a_moment_by_qianniu(window_name, mode=1, app=None):
    '''
        :param window_name: 应用窗口的名称
        :param mode: 1: 哨兵模式循环监视 2: 被调模式只执行一次
        :param app: WinGUI 实例，用于窗口操作
    '''
    
    # mode 只能1的时候直接调用，如果想要为2时在其他模式中调用，可能因为线程原因会影响 两个函数的终止快捷键都失效


    if not app:
        app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

    waiting_script = [
        '亲亲，抱歉呢，目前咨询人数较多，为了更好的给您解决问题，麻烦您稍等一下哦，先不用回复，小妹会尽快为您提供帮助',
        '亲亲，感谢您的咨询，当前咨询量较大，小妹正在逐一回复，您先不用回复，请您稍等片刻(づ￣3￣)づ╭❤～',
        '亲亲，您耐心等待一下，正在为您查找解决方案呢，会尽快给您一个明确的答复，先不用回复小妹呢亲',
        '非常理解您希望尽快解决问题的心情，正在为您查询相关的处理流程，您先不用回复，小妹会尽快为您解决问题',
        '亲亲，这里看下记录亲，您先稍等一下不用回复',
        '亲亲，当前客服系统正在进行数据同步，请您稍等片刻，不用回复，我们会尽快回到您的问题上来，感谢您的耐心',
        '亲亲，由于系统短暂维护，我们将在几分钟后恢复服务。请您耐心等待，小妹会尽快处理您的问题'
    ]
    global_random_number = 99

    exit_flag = False

    def set_exit_flag():
        nonlocal exit_flag
        print(f"END | terminated by user")
        exit_flag = True

    # 设置组合键监听
    keyboard.add_hotkey('shift+ctrl+a', set_exit_flag)
    cycle_count = 0  # 初始化循环计数器

    if mode == 1:
        print('哨兵模式自动回复启动')
        show_toast('提醒', '哨兵模式自动回复启动')

    try:
        while not exit_flag:
            print(f"当前监听新消息轮数: {cycle_count}")
            # 有新的信息提示 回复稍等
            _, __, is_find_new_message = app.locate_icon('new_message.png',0.3,1,0.1,1,1)
            if is_find_new_message:
                # 按下快捷键 Ctrl+E 呼出新消息界面
                keyboard.press_and_release('ctrl+e')
                time.sleep(0.2)
                # 按下快捷键Ctrl+i，聚焦到输入框
                keyboard.press_and_release('ctrl+i')
                time.sleep(0.2)

                # 随机等待消息 避免两次内容相同
                random_number_max = len(waiting_script) - 1
                random_number = random.randint(0, random_number_max)
                while random_number == 99 or random_number == global_random_number:
                    random_number = random.randint(0, random_number_max)
                    global_random_number = random_number

                # 输入随机等待消息
                pyperclip.copy(waiting_script[random_number])
                keyboard.press_and_release('ctrl+v')
                print(waiting_script[random_number])
                # 模拟按下回车键发送消息
                keyboard.press_and_release('enter')
                cycle_count += 1
                time.sleep(3)

                # 判断是否出现提示框 消息重复发送
                repeat_message_x, repeat_message_y, is_find_repeat_message = app.locate_icon('repeat_message.png',0.2,0.8,0.2,0.8,1)
                if is_find_repeat_message:
                    # 聚焦到输入框
                    app.move_and_click(repeat_message_x, repeat_message_y)
            if mode == 2:
                print('结束检测未回复消息')
                break
    except KeyboardInterrupt:
        print("检测到 Ctrl+C，正在退出...")
    except Exception as e:
        print(f"快捷键监听出错：{e}")
    finally:
        # 移除所有快捷键监听
        keyboard.unhook_all()
        print('退出监听')
        if mode == 1:
            show_toast('提醒', '哨兵模式自动回复退出')

# 定义一个退出标志
# exit_flag = False
exit_event = threading.Event()

# 定义按键监听事件
def set_exit_flag():
    # global exit_flag
    # exit_flag = True
    exit_event.set()
    print(f"END | terminated by user")
    show_toast('提醒', '程序终止')
    
keyboard.add_hotkey('shift+ctrl+q', set_exit_flag)

# 通知补发单号
# mode1 使用输入框通知 mode2 使用补发窗口通知
def notification_reissue(window_name, table_name, notic_shop_name, notic_mode=2, show_logistics=False, logistics_mode=1, use_today=None, test_mode=0, is_write=True, table_path='', form_folder='../form'):
    '''
        :param window_name: 应用窗口的名称
        :param table_name: 表单名称
        :param notic_shop_name: 店铺名称
        :param notic_mode: 通知模式 1：输入框通知 2：补发窗口按钮通知
        :param show_logistics: 是否显示物流公司 输入框通知模式下生效
        :param logistics_mode: 物流模式 1自动识别物流公司 2手动输入物流公司
        :param use_today: 是否使用今天日期作为路径 默认今天 指定则传入如 2024-11-27 注需要存在文件及路径
        :param test_mode: 测试模式 0：不测试 若测试则输入测试数量
        :param is_write: 是否写入数据 默认写入
        :param table_path: 表单路径 暂时用不到 预留位置 当前逻辑比较畸形避免处可以用于出错后续优化
        :param form_folder: 表单文件夹路径
    '''

    # global exit_flag

    def handle_shop_name(notic_shop_name):
        shop_name_icon = ''
        shop_name_icon_not_selected = ''
        # 避免店铺名称冲突
        if notic_shop_name == '团洁':
            notic_shop_name = '团洁旗舰'
        # 设置店铺名称图片名称 包含选中与未选中状态
        # '潮洁居家日用旗舰店-天猫', '余猫旗舰店-天猫', '团洁3504猫宁-天猫', '团洁旗舰店-天猫', '潮洁873猫宁-天猫'
        if notic_shop_name == '团洁旗舰':
            shop_name_icon = 'tuanjie_table_icon_selected.png'
            shop_name_icon_not_selected = 'tuanjie_table_icon_not_selected.png'
        elif notic_shop_name == '潮洁居家':
            shop_name_icon = 'chaojie_table_icon_selected.png'
            shop_name_icon_not_selected = 'chaojie_table_icon_not_selected.png'
        elif notic_shop_name == '余猫旗舰':
            shop_name_icon = 'yumao_table_icon_selected.png'
            shop_name_icon_not_selected = 'yumao_table_icon_not_selected.png'
        elif notic_shop_name == '猫宁3504':
            notic_shop_name = '3504猫宁'
            shop_name_icon = 'maoning_table_icon_selected.png'
            shop_name_icon_not_selected = 'maoning_table_icon_not_selected.png'
        elif notic_shop_name == '猫宁873':
            notic_shop_name = '873猫宁'
            shop_name_icon = 'maoning_table_icon_selected.png'
            shop_name_icon_not_selected = 'maoning_table_icon_not_selected.png'
        elif notic_shop_name == '音美旗舰':
            shop_name_icon = 'yinmei_table_icon_selected.png'
            shop_name_icon_not_selected = 'yinmei_table_icon_not_selected.png'
        elif notic_shop_name == 'lelodi':
            shop_name_icon = 'lelodi_table_icon_selected.png'
            shop_name_icon_not_selected = 'lelodi_table_icon_not_selected.png'
        elif notic_shop_name == 'yemo':
            shop_name_icon = 'yemo_table_icon_selected.png'
            shop_name_icon_not_selected = 'yemo_table_icon_not_selected.png'
        else:
            print(f"未知店铺名称：{notic_shop_name}")
            notic_shop_name = ''
        return notic_shop_name, shop_name_icon, shop_name_icon_not_selected

    # 打印参数信息 每条信息换行
    print(f"窗口名称：{window_name}\n表单名称：{table_name}\n店铺名称：{notic_shop_name}\n通知模式：{notic_mode}\n显示物流公司：{show_logistics}\n物流模式：{logistics_mode}\n是否使用今天日期作为路径：{use_today}\n测试模式：{test_mode}\n是否写入数据：{is_write}\n表单路径：{table_path}\n表单文件夹路径：{form_folder}")

    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    notic_shop_name, shop_name_icon, shop_name_icon_not_selected = handle_shop_name(notic_shop_name)

    if notic_shop_name == '':
        return
    
    try:
        # 组合表单路径
        # 如果使用今天日期 则进行组合
        if not use_today:
            form_folder += f"/{datetime.datetime.now().strftime('%Y-%m-%d')}"
        else:
            form_folder += f"/{use_today}"
        table_file = os.path.join(form_folder, table_name)
        table_file = r"{}".format(table_path)
        # 根据不同文件格式读取表格
        file_format = table_name.split('.')
        if 'xls' in file_format[1]:
            # 表格不存在提示
            if not os.path.exists(table_file):
                print(f"表单文件：{table_file} 不存在")
                show_toast('提醒', f'表单文件：{table_file} 不存在')
                return
            # 读取 Excel 文件
            df = pd.read_excel(table_file, sheet_name=None, dtype={'原始单号': str, '物流单号': str})
        elif 'csv' in file_format[1]:
            print('未处理 csv 文件')
            return
        else:
            print('未知格式')
            return
        
        # 获取所有sheet的名称
        sheet_names = df.keys()

        current_sheet_name = ''
        # 选定店铺名 循环判断notic_shop_name是否在sheet_names中
        for sheet_name in sheet_names:
            if notic_shop_name in sheet_name:
                print(f"选定表单：{sheet_name}")
                current_sheet_name = sheet_name
                break
        if not current_sheet_name:
            print(f"未找到 '{notic_shop_name}' 的表单，程序终止")
            return
        
        # 读取当前表单
        df_current_sheet = df[current_sheet_name]
        column_names = df_current_sheet.columns.tolist()  # 获取列名列表
        # print(column_names) # 打印列名列表

        # 检查是否存在"是否通知"列，如果不存在则添加
        if '是否通知' not in column_names:
            df_current_sheet['是否通知'] = 0
            # 将更改写回 Excel 文件
            with pd.ExcelWriter(table_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_current_sheet.to_excel(writer, index=False, sheet_name=current_sheet_name)
            print(f"列 '是否通知' 已添加到表格指定sheet '{current_sheet_name}' 中")
        else:
            print(f"列 '是否通知' 已存在于表格指定sheet '{current_sheet_name}' 中")
        

        # 点击店铺名称 分为两种情况 已经被选中和未被选中的状态
        is_find_shop_icon = app.click_icon(shop_name_icon, 0, 0.9, 0, 0.3)
        if not is_find_shop_icon:
            is_find_shop_icon = app.click_icon(shop_name_icon_not_selected, 0, 0.9, 0, 0.3)
        if not is_find_shop_icon:
            print(f"未找到店铺名称：{notic_shop_name}")
            return
        time.sleep(0.2)

        if test_mode:
            # 限制 DataFrame 到指定行数 用于测试 排除 '是否通知' 列中值为 0 的行
            # df_subset = df_current_sheet.head(2)
            # df_subset = df_current_sheet.iloc[:2]
            # df_subset = df_current_sheet.iloc
            print(f"当前sheet: {current_sheet_name}，测试模式，只处理前 {test_mode} 行")
            df_subset = df_current_sheet.loc[df_current_sheet['是否通知'] != 1, :].iloc[:test_mode]
        else:
            print(f"当前sheet: {current_sheet_name}，处理所有行")
            df_subset = df_current_sheet

        print(df_subset)
        
        waiting_script = [
            '亲亲，抱歉呢，目前咨询人数较多，为了更好的给您解决问题，麻烦您稍等一下哦，先不用回复，小妹会尽快为您提供帮助',
            '亲亲，感谢您的咨询，当前咨询量较大，小妹正在逐一回复，您先不用回复，请您稍等片刻(づ￣3￣)づ╭❤～',
            '亲亲，您耐心等待一下，正在为您查找解决方案呢，会尽快给您一个明确的答复，先不用回复小妹呢亲',
            '非常理解您希望尽快解决问题的心情，正在为您查询相关的处理流程，您先不用回复，小妹会尽快为您解决问题',
            '亲亲，这里看下记录亲，您先稍等一下不用回复',
            '亲亲，当前客服系统正在进行数据同步，请您稍等片刻，不用回复，我们会尽快回到您的问题上来，感谢您的耐心',
            '亲亲，由于系统短暂维护，我们将在几分钟后恢复服务。请您耐心等待，小妹会尽快处理您的问题'
        ]
        global_random_number = 99

        # 逐行处理DataFrame
        cycle_count  = 0
        # while not exit_flag:
        while not exit_event.is_set():
            if cycle_count >= len(df_subset):
                print('所有行已处理完毕')
                show_toast('提醒', '所有行已处理完毕')
                break
            
            #使用iloc根据索引获取当前行的数据
            row = df_subset.iloc[cycle_count]
            index = df_subset.index[cycle_count]  # 获取原始索引

            # 检查是否已经通知
            if row['是否通知'] == 1:
                print('当前用户已通知')
                cycle_count += 1
                continue  # 如果已经通知，则跳过当前行

            # 有新的信息提示 回复稍等
            is_find_new_message = True
            if is_find_new_message:
                _, __, is_find_new_message = app.locate_icon('new_message.png',0.3,1,0.1,1,1)
                if is_find_new_message:
                    print('有新的信息提示 回复稍等')
                    # 按下快捷键 Ctrl+E 呼出新消息界面
                    keyboard.press_and_release('ctrl+e')
                    time.sleep(0.2)
                    # 按下快捷键Ctrl+i，聚焦到输入框
                    keyboard.press_and_release('ctrl+i')
                    time.sleep(0.2)

                    # 随机等待消息 避免两次内容相同
                    random_number_max = len(waiting_script) - 1
                    random_number = random.randint(0, random_number_max)
                    while random_number == 99 or random_number == global_random_number:
                        random_number = random.randint(0, random_number_max)
                        global_random_number = random_number

                    # 输入随机等待消息
                    pyperclip.copy(waiting_script[random_number])
                    keyboard.press_and_release('ctrl+v')
                    print(waiting_script[random_number])
                    # 模拟按下回车键发送消息
                    keyboard.press_and_release('enter')
                    time.sleep(3)

                    # 判断是否出现提示框 消息重复发送
                    repeat_message_x, repeat_message_y, is_find_repeat_message = app.locate_icon('repeat_message.png',0.2,0.8,0.2,0.8,1)
                    if is_find_repeat_message:
                        # 聚焦到输入框
                        app.move_and_click(repeat_message_x, repeat_message_y)
                else:
                    print('没有新的信息提示')

            # 获取原始单号和物流单号
            original_number = row['原始单号']
            logistics_number = row['物流单号']
            print(f"原始单号：{original_number} 物流单号：{logistics_number}")
            # 原始单号为空 提示并跳过
            if not original_number:
                print(f"原始单号为空 跳过")
                # show_toast('提醒', f'原始单号为空 跳过')
                cycle_count += 1
                continue
            # 物流单号为空 提示并跳过
            if not logistics_number:
                print(f"物流单号为空 跳过")
                cycle_count += 1
                # show_toast('提醒', f'物流单号为空 跳过')
                continue

            # original_number=123456789789

            # app.get_app_screenshot()
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
            if exit_event.is_set():
                break
            # 点击搜索框
            app.click_icon('button_search_cus.png',0,0.3,0,0.3)
            time.sleep(0.1)

            # 清除搜索框
            keyboard.press_and_release('ctrl+a')  
            keyboard.press_and_release('backspace')

            # 将 original_number 的内容输入到搜索框中
            # pyautogui.typewrite(original_number)
            # 将中文字符串复制到剪贴板
            pyperclip.copy(original_number)
            keyboard.press_and_release('ctrl+v') 
            # 等待搜索结果响应
            time.sleep(0.2)

            if exit_event.is_set():
                break

            # 判断是否未找到
            _, __, not_find_cus = app.locate_icon('not_find_customer.png',0, 0.4, 0, 0.6)
            
            if exit_event.is_set():
                break

            if not_find_cus:
                print(f"未搜索到结果，跳过 {original_number}")
                cycle_count += 1
                continue  # 未搜索到直接continue下一个
            else:
              print('搜索到指定用户，即将发送通知...')
            
            if exit_event.is_set():
                break

            # 模拟按下回车键进入指定用户的聊天窗口
            keyboard.press_and_release('enter')
            time.sleep(0.3)

            if exit_event.is_set():
                break

            # 通知模式 1：输入框通知 2：补发窗口通知
            if notic_mode == 1:
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

                # 清除输入框
                keyboard.press_and_release('ctrl+a')  
                keyboard.press_and_release('backspace')

                time.sleep(0.2)

                if exit_event.is_set():
                    break

                # 获取快递公司
                if show_logistics:
                    logistics = ul.get_express_company(logistics_number)
                else:
                    logistics = ''
                message = f"亲 {logistics} {logistics_number} 这是您的补发单号 请注意查收"

                # 将中文字符串复制到剪贴板
                pyperclip.copy(message)
                # 使用 pyautogui.typewrite 粘贴剪贴板内容
                keyboard.press_and_release('ctrl+v')  
                # pyautogui.typewrite(message, paste=True)
                time.sleep(0.2)

                if exit_event.is_set():
                    break

                # 模拟按下回车发送消息
                keyboard.press_and_release('enter')

                if exit_event.is_set():
                    break
            elif notic_mode == 2:
                # 聚焦到智能客服 避免误触
                keyboard.press_and_release('ctrl+o')
                # 点击最近三月订单和安装服务 避免误触
                recent_orders_text_x, recent_orders_text_y, is_find_recent_orders_text = app.locate_icon('recent_orders_text.png', 0.6, 1, 0.2, 1)
                if exit_event.is_set():
                    break
                app.click_icon('recent_orders_text.png', 0.6, 1, 0.2, 1)
                if not is_find_recent_orders_text:
                    if exit_event.is_set():
                        break
                    app.click_icon('installation_services.png', 0.6, 1, 0.2, 1)
                # time.sleep(0.1)
                pyautogui.scroll(-100)
                if exit_event.is_set():
                    break

                # 尝试点击两次搜索订单按钮
                is_find_search_button = False 
                for i in range(3):
                    if exit_event.is_set():
                        break
                    search_button_x, search_button_y, is_find_search_button = app.locate_icon('search_order_button.png', 0.6, 1, 0.2, 1)
                    if exit_event.is_set():
                        break

                    if not is_find_search_button:
                        # 未找到搜索订单按钮 尝试点击以选中的按钮
                        selected_search_button_x, selected_search_button_y, is_find_selected_search_button = app.locate_icon('selected_search_order_button.png', 0.6, 1, 0.2, 1)
                        if not is_find_selected_search_button:
                            print(f'未找到搜索订单按钮，尝试滑动后再次查找...')
                            pyautogui.scroll(-160)
                        else:
                            print(f'当前搜索按钮已被点击直接执行下一步...')
                            is_find_search_button = True
                            break
                    else:
                        app.move_and_click(search_button_x, search_button_y)
                        break
                if not is_find_search_button:
                    print(f'未找到搜索订单按钮，尝试直接寻找搜索框...')
                # time.sleep(0.1)

                if exit_event.is_set():
                    break
                
                # 尝试点击两次搜索框
                is_find_search_text = False
                for i in range(3):
                    if exit_event.is_set():
                        break
                    search_text_x, search_text_y, is_find_search_text = app.locate_icon('search_order_text.png', 0.6, 1, 0.2, 1)
                    if exit_event.is_set():
                        break
                    if is_find_search_text:
                        print(f'找到搜索框，点击搜索框...')
                        break
                    print(f'未找到搜索框，尝试滑动后再次查找...')
                    pyautogui.scroll(-100)
                if exit_event.is_set():
                    break
                if not is_find_search_text:
                    print(f'未找到搜索框，尝试直接点击补发按钮...')
                else:
                    app.move_and_click(search_text_x+100, search_text_y)
                time.sleep(0.3)

                if exit_event.is_set():
                    break
                
                if is_find_search_text:
                    if exit_event.is_set():
                        break
                    # 保证输入框没有内容
                    keyboard.press_and_release('ctrl+a') 
                    keyboard.press_and_release('backspace') 
                    # 输入订单号
                    pyperclip.copy(original_number)
                    keyboard.press_and_release('ctrl+v') 
                    # time.sleep(0.1)
                    # 模拟按下回车搜索
                    keyboard.press_and_release('enter')
                    time.sleep(0.3)
                    # 使用 pyautogui 向下滚动鼠标滚轮
                    pyautogui.scroll(-300)
                    time.sleep(0.3)

                if exit_event.is_set():
                    break
                #  尝试点击两次补发按钮
                is_find_reissue_button = False
                for _ in range(3):
                    if exit_event.is_set():
                        break
                    reissue_button_x, reissue_button_y, is_find_reissue_button = app.locate_icon('reissue_button.png', 0.6, 1, 0.2, 1)
                    if exit_event.is_set():
                        break
                    if is_find_reissue_button:
                        print(f'找到补发按钮，点击补发按钮...')
                        break
                    print(f'未找补发按钮，尝试滑动后再次查找...')
                    pyautogui.scroll(-150)

                if not is_find_reissue_button:
                    if exit_event.is_set():
                        break
                    print(f'未找到补发按钮，跳过{original_number}')
                    cycle_count += 1
                    continue
                app.move_and_click(reissue_button_x, reissue_button_y, 'left')
                time.sleep(0.3)

                if exit_event.is_set():
                    break

                # 点击输入框
                add_logistics_number_x, add_logistics_number_y, is_find_add_logistics_number = app.locate_icon('add_logistics_number.png', 0.6, 1, 0.5, 1)
                
                if exit_event.is_set():
                    break
                
                if not is_find_add_logistics_number:
                    print(f'未找到添加物流单号提示文字，跳过{original_number}')
                    cycle_count += 1
                    continue
                app.move_and_click(add_logistics_number_x, add_logistics_number_y)

                if exit_event.is_set():
                    break
                # 将中文字符串复制到剪贴板
                pyperclip.copy(logistics_number)
                keyboard.press_and_release('ctrl+v') 
                keyboard.press_and_release('tab')
                time.sleep(0.5)

                if exit_event.is_set():
                    break
                
                # 手动输入物流公司
                if logistics_mode == 2:
                    print(f'手动输入快递公司模式..')
                    logistics = ul.get_express_company(logistics_number)
                    # 如果不为空 末尾加快递二字
                    if logistics:
                        logistics = f"{logistics}快递"
                        pyperclip.copy(logistics)
                        keyboard.press_and_release('ctrl+v')
                        keyboard.press_and_release('enter')

                    else:
                        logistics_mode = 1  # 自动识别物流公司
                # 自动识别物流
                if logistics_mode == 1:
                    print(f'自动输入快递模式...')
                    # 按下回车
                    keyboard.press_and_release('enter')
                time.sleep(0.1)

                if exit_event.is_set():
                    break
                # 点击确认补发按钮
                confirm_button_x, confirm_button_y, is_find_confirm_button = app.locate_icon('confirm_button.png', 0.6, 1, 0.2, 1)
                if not is_find_confirm_button:
                    print(f'未找到确认补发按钮，跳过{original_number}')
                    cycle_count += 1
                    continue
                else:
                    # 点击确认按钮
                    app.move_and_click(confirm_button_x, confirm_button_y)
                if exit_event.is_set():
                    break
                # change  判断是否弹出失败提示
            else:
                print(f"未知通知模式：{notic_mode}")
                show_toast('提示', '未知通知模式')
                return
                
            # 将当前行的"是否通知"标记为1
            print(f'已通知 {notic_shop_name} {original_number}')
            df_current_sheet.at[index, '是否通知'] = 1
            cycle_count += 1
            if exit_event.is_set():
                break
            # 循环结束暂停
            time.sleep(0.2)
    except KeyboardInterrupt:
        print("检测到 Ctrl+C，正在退出...")
    except Exception as e:
        print(f"通知程序异常：{e}")
        # 如果读取过程中出现异常，可能是文件被其他程序占用（例如已打开）
    except PermissionError:
        print("文件可能已被其他程序打开，请确保文件未被打开后再试。")
        show_toast('提示', '文件可能已被其他程序打开，请确保文件未被打开后再试。')
    finally:
        # 使用pyperclip库清空剪切板
        pyperclip.copy('')
        print('通知程序已退出')
        show_toast('提示', '通知程序已退出')
        # 将更改写回 Excel 文件
        with pd.ExcelWriter(table_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # 只写回当前处理的sheet
            df_current_sheet.to_excel(writer, sheet_name=current_sheet_name, index=False)

        keyboard.unhook_all()  # 移除所有按键监听
    return df


# 千年自动发送差价链接
def auto_send_price_link(window_name, product_number=1, mode=1):
    '''
        :param window_name: 窗口名称
        :param mode: 1 识别移动 2 相对移动
    '''
    try:
        app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        # 聚焦到智能客服 避免误触
        keyboard.press_and_release('ctrl+o')

        # 移动鼠标 1600 450
        pyautogui.moveTo(1600, 450, duration=0.1)

        # 滚动到顶部
        pyautogui.scroll(1000)

        # 寻找搜索商品图
        search_product_x, search_product_y, is_find_search_product = app.locate_icon('search_products.png', 0.6, 1, 0.2, 1)
        if not is_find_search_product:
            print('未找到搜索商品图标，寻找足迹按钮位置')
            footprint_x, footprint_y, is_find_footprint = app.locate_icon('footprint.png', 0.6, 1, 0.2, 1)
            if not is_find_footprint:
                print('未找到足迹图标，程序退出')
                return

        # 点击搜索商品图标
        if is_find_search_product:
            app.move_and_click(search_product_x, search_product_y)
        else:
            app.move_and_click(footprint_x+313, footprint_y+41)
            # 删除搜索框内容
            keyboard.press_and_release('ctrl+a')
            time.sleep(0.1)
            keyboard.press_and_release('backspace')
        time.sleep(0.1)

        # 搜索差价链接
        keyboard.write('差价')
        keyboard.press_and_release('enter')
        time.sleep(0.1)

        # 移动到邀请下单图标
        invite_order_x, invite_order_y, is_find_invite_order = app.locate_icon('invite_order.png', 0.6, 1, 0.2, 1)
        if not is_find_invite_order:
            print('未找到邀请下单图标，程序退出')
            return
        # 点击邀请下单
        app.move_and_click(invite_order_x, invite_order_y)
        time.sleep(0.1)

        # 点击输入数量
        if mode == 1:
            # 移动到输入数量图标
            product_quantity_x, product_quantity_y, is_find_product_quantity = app.locate_icon('product_quantity.png', 0.6, 1, 0.2, 1)
            if not is_find_product_quantity:
                print('未找到输入数量图标，程序退出')
                return
            # 点击输入数量
            app.move_and_click(product_quantity_x, product_quantity_y+35)
        elif mode == 2:
            # 使用相对移动点击输入数量 0,780
            app.move_and_click(search_product_x, search_product_y + 780)
        else:
            print(f'未知模式：{mode}')
            return
        time.sleep(0.1)

        # 输入数量
        keyboard.press_and_release('ctrl+a')
        keyboard.press_and_release('backspace')
        keyboard.write(product_number)
        keyboard.press_and_release('enter')
        time.sleep(0.1)

        # 点击邀请下单
        if mode == 1:
            # 移动到邀请下单图标
            send_invitation_x, send_invitation_y, is_find_send_invitation = app.locate_icon('send_invitation.png', 0.6, 1, 0.2, 1)
            if not is_find_send_invitation:
                print('未找到邀请下单图标，程序退出')
                return
            # 点击邀请下单
            app.move_and_click(send_invitation_x, send_invitation_y)
        elif mode == 2:
            # 使用相对移动点击邀请下单 280,235
            app.move_and_click(search_product_x + 280, search_product_y + 235)
        else:
            print(f'未知模式：{mode}')
            return
        time.sleep(0.1)

        keyboard.press_and_release('ctrl+i')
        keyboard.write('这里下单付款呢亲')
        keyboard.press_and_release('enter')

    except Exception as e:
        print(f"差价链接程序异常：{e}")


# 从剪切板获取数量
def get_clipboard_number(auto_copy=True):
    '''
        :param auto_copy: 是否自动复制到剪切板
        :return: 剪切板中的数字
    '''
    if auto_copy:
        # keyboard.press_and_release('ctrl+a')
        # keyboard.press_and_release('ctrl+x')
        pyautogui.hotkey('ctrl', 'a')  # 模拟按下并释放
        pyautogui.hotkey('ctrl', 'x')  # 模拟按下并释放
        time.sleep(0.1)
    clipboard_text = pyperclip.paste()
    if not clipboard_text.isdigit():
        print('剪切板中没有数字，程序退出')
        return 0
    return int(clipboard_text)

def handle_auto_send_price_link(window_name, mode=1):
    '''
        :param window_name: 窗口名称
        :param mode: 1 识别移动 2 相对移动
    '''
    number = get_clipboard_number()
    if number == 0:
        print('剪切板中没有数字，程序退出')
        return
    auto_send_price_link(window_name, str(number), mode)



def erp_select_today(window_name, app=None, reissue=True):
    '''
        :param window_name: 窗口名称
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        # 点击日期下拉框
        date_dropdown = read_coordinate_by_key('date_dropdown', reissue)
        if not date_dropdown:
            print('未找到日期下拉框，程序退出')
            return
        app.move_and_click(date_dropdown[0], date_dropdown[1])
        time.sleep(0.3)
        # 点击今天日期
        today_date = read_coordinate_by_key('today_date', reissue)
        if not today_date:
            print('未找到今天日期，程序退出')
            return
        app.move_and_click(today_date[0], today_date[1])
        # 获取当前的小时和分钟
        current_time_hour = datetime.datetime.now().hour
        current_time_minute = datetime.datetime.now().minute
        current_time_hour_pos = read_coordinate_by_key('current_time_hour', reissue)
        current_time_minute_pos = read_coordinate_by_key('current_time_minute', reissue)
        if not current_time_hour_pos or not current_time_minute_pos:
            print('未找到当前时间，程序退出')
            return
        # 点击小时输入小时
        app.move_and_click(current_time_hour_pos[0], current_time_hour_pos[1])
        keyboard.write(str(current_time_hour))
        # 点击分钟输入分钟
        app.move_and_click(current_time_minute_pos[0], current_time_minute_pos[1])
        keyboard.write(str(current_time_minute))

    except Exception as e:
        print(f"ERP选择今天日期异常：{e}")

def erp_clear_product(window_name, app=None, mode=5, reissue=True):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
        :param mode: 1 识别移动 非1之外的数字 直接循环点击次数
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        first_product = read_coordinate_by_key('first_product', reissue)
        if not first_product:
            print('未找到第一个商品，程序退出')
            return
        # 判断是否还有序号为1的商品
        if mode == 1:
            is_find_product = True
            while is_find_product:
                if is_find_product:
                    # 双击第一个商品以删除
                    app.move_and_click(first_product[0], first_product[1], 'left', 2)
                    time.sleep(0.1)
                else:
                    _, __, is_find_product = app.locate_icon('product_1.png', 0, 0.2, 0.2, 0.8)
                    time.sleep(0.1)
        else:
            for _ in range(mode):
                # 双击第一个商品以删除
                app.move_and_click(first_product[0], first_product[1])
                pyautogui.doubleClick()
    except Exception as e:
        print(f"ERP清空商品异常：{e}")

def erp_input_remarks(window_name, remarks='补发', app=None, reissue=True):
    '''
        :param window_name: 窗口名称
        :param remarks: 备注内容
        :param app: WinGUI 实例 默认为 None
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        # 点击备注输入框
        remarks_input = read_coordinate_by_key('remarks_input', reissue)
        if not remarks_input:
            print('未找到备注输入框，程序退出')
            return
        app.move_and_click(remarks_input[0], remarks_input[1], 'left')
        time.sleep(0.1)

        # 清空输入框
        keyboard.press_and_release('ctrl+a')
        time.sleep(0.1)
        keyboard.press_and_release('backspace')
        time.sleep(0.1)

        # 输入备注内容
        keyboard.write(remarks)

    except Exception as e:
        print(f"ERP输入备注异常：{e}")

def erp_choose_warehouse(window_name, warehouse='sz', app=None, reissue=True):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        # 点击仓库下拉框
        warehouse_dropdown = read_coordinate_by_key('warehouse_dropdown', reissue)
        if not warehouse_dropdown:
            print('未找到仓库下拉框，程序退出')
            return
        app.move_and_click(warehouse_dropdown[0], warehouse_dropdown[1])
        time.sleep(0.2)
        # 点击仓库
        if warehouse =='sz':
            warehouse_sz = read_coordinate_by_key('warehouse_sz', reissue)
            if not warehouse_sz:
                print('未找到深圳仓库，程序退出')
                return
            app.move_and_click(warehouse_sz[0], warehouse_sz[1])
        elif warehouse == 'cz':
            warehouse_cz = read_coordinate_by_key('warehouse_cz', reissue)
            if not warehouse_cz:
                print('未找到潮州仓库，程序退出')
                return
            app.move_and_click(warehouse_cz[0], warehouse_cz[1])

    except Exception as e:
        print(f"ERP选择仓库异常：{e}")


def erp_add_product(window_name, app=None, reissue=True):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        # 点击添加商品
        position_add_product_button = read_coordinate_by_key('add_product_button', reissue)
        if not position_add_product_button:
            return
        app.move_and_click(position_add_product_button[0], position_add_product_button[1])
        time.sleep(0.3)
        # 点击货品名称输入框
        position_product_name_input = read_coordinate_by_key('product_name_input', reissue)
        if not position_product_name_input:
            return
        app.move_and_click(position_product_name_input[0], position_product_name_input[1])
        time.sleep(0.1)
        keyboard.press_and_release('ctrl+a')
        keyboard.press_and_release('backspace')

    except Exception as e:
        print(f"ERP添加商品异常：{e}")

def erp_add_product_notes(window_name, local='sz', app=None):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        # 点击添加商品
        erp_add_product(window_name, app)
        # 输入备注内容
        if local =='sz':
            keyboard.write('（深圳仓）')
        else:
            keyboard.write('（潮州仓）')
        keyboard.press_and_release('enter')
        # 选中商品
        app.move_and_click(1000, 400, 'left', 2)
        # 保存添加备注
        app.move_and_click(1413, 787)

    except Exception as e:
        print(f"ERP添加备注异常：{e}")

def erp_add_specific_products(window_name, product_list=None, app=None):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
        :param product_list: 商品列表
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        # 点击添加商品
        erp_add_product(window_name, app)
        # 输入备注内容
        for product in product_list:
            keyboard.write(product)
            keyboard.press_and_release('enter')
            # 选中商品
            app.move_and_click(1000, 400, 'left', 2)
            time.sleep(0.1)
            if len(product_list) > 1:
                # 清空输入框 1237 290
                app.move_and_click(1237, 290)
                # 点击货品名称输入框
                app.move_and_click(1188, 288)
                time.sleep(0.1)
        # 保存添加备注
        app.move_and_click(1413, 787)

    except Exception as e:
        print(f"ERP添加特定商品异常：{e}")

# erp常用操作 选择今天日期 清空商品 选择仓库 添加商品备注 输入备注
def erp_common_action_1(window_name, choose_warehouse=False, warehouse='', remarks='', app=None):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作

        erp_select_today(window_name, app)
        erp_clear_product(window_name, app)
        if choose_warehouse:
            erp_choose_warehouse(window_name, warehouse, app)
        if warehouse != '':
            erp_add_product_notes(window_name,  warehouse, app)
        if remarks != '':
            erp_input_remarks(window_name, remarks, app)

    except Exception as e:
        print(f"ERP常用操作1异常：{e}")

# erp常用操作2 选择今天日期 清空商品 选择仓库 添加转接头  输入备注
def erp_common_action_2(window_name, warehouse='', product_list=[], remarks='', app=None):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
        if warehouse == '':
            print('请输入仓库')
            return
        erp_select_today(window_name, app)
        erp_clear_product(window_name, app)
        erp_choose_warehouse(window_name, warehouse, app)
        # 判断product_list列表为空
        if len(product_list) == 0:
            print('请输入商品')
            return
        print(product_list)
        erp_add_specific_products(window_name, product_list, app)
        if remarks != '':
            erp_input_remarks(window_name, remarks, app)

    except Exception as e:
        print(f"ERP常用操作1异常：{e}")

# erp常用操作3 选择今天日期 清空商品 选择仓库 添加转接头  输入备注
def erp_common_action_3(window_name, warehouse='', remarks='', app=None):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
    '''
    try:
        if not app:
            app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
        if warehouse == '':
            print('请输入仓库')
            return
        erp_select_today(window_name, app)
        erp_clear_product(window_name, app)
        erp_choose_warehouse(window_name, warehouse, app)
        if remarks != '':
            erp_input_remarks(window_name, remarks, app)
        # 点击添加商品
        erp_add_product(window_name, app)

    except Exception as e:
        print(f"ERP常用操作1异常：{e}")


# erp操作集合
def erp_action_collection(action_list=None):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
        :param action_list: 操作列表
    '''
    try:
        if action_list == None:
            print('请输入操作列表')
            return
        print(f'ERP操作集合：{action_list}')
        app = action_list.get('app', None)
        if not app:
            window_name = action_list.get('window_name', '')
            if window_name == '':
                print('请输入窗口名称')
                return
            else:
                app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
            
        # 获取action_list中的product_list
        product_list = action_list.get('product_items', [])
        if len(product_list) == 0:
            print('商品列表为空')

        warehouse = action_list.get('warehouse', '')
        if not warehouse:
            print('未选择仓库')
            return
        else:
            if warehouse == '深圳':
                warehouse ='sz'
            elif warehouse == '潮州':
                warehouse = 'cz'
            else:
                print('仓库名称错误')
                return
        # 确认模式
        reissuse_order = action_list.get('reissuse_order', False)

        print(f"ERP当前处理模式：{'补发模式' if not reissuse_order else '手工建单模式'}")

        # 选择今天
        if action_list.get('select_today', False):
            erp_select_today(window_name, app, reissuse_order)
        # 清空商品
        if action_list.get('clear_product', False):
            erp_clear_product(window_name, app, 5, reissuse_order)  
        # 选择仓库
        erp_choose_warehouse(window_name, warehouse, app, reissuse_order)
        # 添加备注
        remarks = action_list.get('remarks', ['补发'])
        for remark in remarks:
            erp_input_remarks(window_name, remark, app, reissuse_order)
        # 添加商品
        products = action_list.get('product_items', [])
        if len(products) > 0:
            erp_add_specific_products(window_name, products, app)
        
    except Exception as e:
        print(f"ERP操作集合异常：{e}")

config = configparser.ConfigParser()
config.read('../config/window_config.ini')
# 访问环境变量
window_open_mode = config['defaults'].get('WINDOW_OPEN_MODE')
# 尝试将 TRY_NUMBER 转换为整数
window_open_mode = int(window_open_mode)

tk_window = None  
tk_window_initialized = False # 标志 window 是否已经初始化

def on_close():
    global tk_window, tk_window_initialized
    tk_window.destroy()
    tk_window = None
    tk_window_initialized = False  # 窗口关闭后重置标志

# 定义一个回调函数，在主线程中调用create_window
def call_create_window():
    global tk_window
    if tk_window is not None and tk_window.winfo_exists():
        print("窗口已存在，不再创建")
        # 使用 wm_attributes 来确保窗口会暂时位于最顶层
        tk_window.attributes('-topmost', True)
        tk_window.after(50, lambda: tk_window.attributes('-topmost', False))
        tk_window.lift()  # 将窗口提升到最前面
        tk_window.after(50, tk_window.focus_force)  # 强制窗口获得焦点，并添加小延迟
    else:
        print("窗口不存在，创建新窗口")
        erp_aciton_box(mode=window_open_mode)
        # tk_window.after(0, create_window)  # 如果窗口不存在，创建新窗口

# erp创建输入框操作
def erp_aciton_box(mode=0):
    '''
        :param window_name: 窗口名称
        :param app: WinGUI 实例 默认为 None
    '''
    print(f"create_window mode: {mode} {'主程序模式' if mode == 0 else '被调用模式'}")

    global tk_window, tk_window_initialized
    if tk_window is not None and tk_window.winfo_exists():
        tk_window.lift()  # 如果窗口已经存在，将其提升到最前面
        return
    
    input_content = ''

    def on_confirm(entry):  
        global input_content, tk_window
        input_content = entry.get()
        entry.delete(0, tk.END)
        
        if input_content == '':
            print('输入内容为空')
            show_toast('提示', '输入内容为空')
        else:
            print(f'输入内容：{input_content}')
            if auto_close.get():
                tk_window.destroy()
                tk_window = None

            erp_handle_input_content(input_content, reissuse_order.get())

    def toggle_reissue_order():
        reissuse_order.set(not reissuse_order.get())  # 切换补发选项的状态
        print(f"补发选项: {'打开' if reissuse_order.get() else '关闭'}")
        show_toast('提示', f"补发选项: {'打开' if reissuse_order.get() else '关闭'}")

    def toggle_auto_close():
        auto_close.set(not auto_close.get())          # 切换自动关闭选项的状态
        print(f"自动关闭选项: {'打开' if auto_close.get() else '关闭'}")
        show_toast('提示', f"自动关闭选项: {'打开' if reissuse_order.get() else '关闭'}")

    tk_window = tk.Tk()

    # 设置标题
    tk_window.title('ERP_BOX')
    # 去掉标题栏
    # tk_window.overrideredirect(True)

    reissuse_order = tk.BooleanVar(value=True) # 选择是则为补发界面，不选择则为手工建单界面
    auto_close = tk.BooleanVar(value=True) # 选择是则点击确认后自动关闭窗口

    # 设置窗口大小
    window_width = 300
    window_height = 90

    # entry设置
    entry_font_fam = '楷体'
    entry_font_size = 11
    entry_color = '#000000'

    # button设置
    button_font_fam = '楷体'
    button_font_size = 10
    button_color = '#000000'

    # checkbox设置
    checkbox_font_fam = '楷体'
    checkbox_font_size = 6
    checkbox_color = '#000000'

    # 设置窗口位于屏幕的中间
    scnwidth, scnheight = tk_window.maxsize()
    x_offset = (scnwidth - window_width) / 2
    y_offset = (scnheight - window_height) / 2
    tk_window.geometry(f'{window_width}x{window_height}+{int(x_offset)}+{int(y_offset)}')

    # 禁止更改窗口大小
    tk_window.resizable(False, False)
    tk_window.update()  # 刷新窗口

    # 创建一个Frame来放置Entry和Buttons，确保它们可以正确地对齐
    frame = tk.Frame(tk_window)
    frame.place(relx=0.5, rely=0.5, anchor='center', relwidth=1, relheight=1)

    # Entry将填满整个窗口宽度，使用place而不是pack
    tk_entry = tk.Entry(frame, font=(entry_font_fam, entry_font_size), fg=entry_color)
    tk_entry.place(relx=0.5, rely=0.3, anchor='center', relwidth=0.9, relheight=0.4)

    # 确认按钮 
    tk_button_confirm = tk.Button(frame, text='确认', command=lambda: on_confirm(tk_entry), font=(button_font_fam, button_font_size), fg=button_color)
    tk_button_confirm.place(relx=0.15, rely=0.8, anchor='center')

    # 取消按钮
    tk_button_cancel = tk.Button(frame, text='取消', command=tk_window.destroy, font=(button_font_fam, button_font_size), fg=button_color)
    tk_button_cancel.place(relx=0.38, rely=0.8, anchor='center')

    # 添加两个复选框
    tk_check_reissue_order = tk.Checkbutton(frame, text='补发', variable=reissuse_order, font=(checkbox_font_fam, checkbox_font_size), fg=checkbox_color)
    tk_check_reissue_order.place(relx=0.7, rely=0.8, anchor='center')

    tk_check_auto_close = tk.Checkbutton(frame, text='自动关闭', variable=auto_close, font=(checkbox_font_fam, checkbox_font_size), fg=checkbox_color)
    tk_check_auto_close.place(relx=0.9, rely=0.8, anchor='center')

    # 绑定回车键到确认按钮
    tk_window.bind('<Return>', lambda event: on_confirm(tk_entry))

    # 绑定ctrl+回车到取消按钮
    tk_window.bind('<Control-Return>', lambda event: tk_window.destroy())

    # 绑定 Ctrl+1 到补发选项
    tk_window.bind('<Control-Key-1>', lambda event: toggle_reissue_order())

    # 绑定 Ctrl+2 到自动关闭选项
    tk_window.bind('<Control-Key-2>', lambda event: toggle_auto_close())

    tk_entry.focus_set()  # 设置Entry为焦点

    tk_window.protocol("WM_DELETE_WINDOW", on_close)  # 设置关闭窗口时的回调函数

    # 确保窗口初始化完成后，再执行以下操作
    tk_window.update_idletasks()  # 确保所有待处理的任务都被执行，包括GUI更新
    tk_window.lift()  # 尝试将窗口提升到最前面
    tk_window.attributes('-topmost', True)  # 暂时设置为最顶层
    tk_window.after(50, lambda: tk_window.attributes('-topmost', False))  # 然后恢复正常
    tk_window.focus_force()  # 强制窗口获得焦点
    tk_window.mainloop()


# erp处理输入框的内容
def erp_handle_input_content(input_content, reissuse_order=True):
    '''
        :param input_content: 输入内容
    '''
    try:
        # 处理输入内容
        if not input_content:
            print('输入内容为空')
            return
        # 去除前后的空格
        input_content = input_content.strip()
        input_content = input_content.split(' ')
        print(f"分割后输入内容：{input_content}")
        
        action_list = {
            'window_name': '旺店通ERP',  # 窗口名称
            'warehouse': None,  # 可以是 "Shenzhen" 或 "Chaozhou"
            'select_today': True,  # 可以是 True 或 False 默认True
            'clear_product': True,  # 可以是 True 或 False 默认True
            'reissue_order': reissuse_order,  # 可以是 True 或 False 选择是则为补发界面，不选择则为手工建单界面
            'remarks': ["补发"],  # 默认备注列表，
            'product_items': [],  # 用于存储产品的列表
            'undefined_product': [],  # 不存在的产品存储
        }

        # 加载映射规则
        mapping_rules = load_mapping_from_json(rule_json_path)
        # 执行验证和转换
        result = validate_and_convert(input_content, mapping_rules)
        if not result:
            print('输入内容不符合规则')
        else:
            # 处理结果
            action_list.update(result)
            # print(f"处理结果：{action_list}")
            erp_action_collection(action_list)
    except Exception as e:
        print(f"ERP处理输入框内容异常：{e}")


def load_mapping_from_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)

def validate_and_convert(input_content, mapping):
    action_list = {
        'product_items': [],  # 成功转换的产品列表
        'undefined_product': [],  # 不存在的产品列表
        'warehouse': None,  # 仓库名称
        'remarks': [],  # 备注列表
    }

    # 验证映射是否有效
    if not isinstance(mapping, dict) or not mapping.get('produdcts'):
        raise ValueError("Invalid mapping structure")
    for key, value in mapping['produdcts'].items():
        if not key or not value:
            raise ValueError(f"Empty key or value found for entry: {key}:{value}")
    
    # 所有转换头的key手动存一个list
    trans = mapping['trans']

    # 转换输入内容为标准产品名称
    converted_products = set()  # 使用集合去重
    found_warehouse = False # 是否找到仓库
    is_trans = False # 是否有转接头

    # 遍历输入的内容
    for item in input_content:
        found_product = False

        # 遍历映射表 寻找符合的产品 并判断是否为转接头
        for standard_name, aliases in mapping['produdcts'].items():
            if item in aliases:
                if standard_name in trans:
                    is_trans = True
                converted_products.add(standard_name)
                found_product = True
                break

        # 设置仓库
        # 有转接头内容则默认为深圳仓 有转接头还是指定深圳仓（找到一次就会停止）
        # 有转接头但是指定潮州仓 就用潮州的 并提示
        for warehouse_name, warehouse_aliases in mapping['warehouse'].items():
            if item in warehouse_aliases and not found_warehouse:
                action_list['warehouse'] = warehouse_name
                if is_trans and warehouse_name == '潮州': # 如果有转接头 并且指定了潮州仓 则提示
                    print('输入内容包含转接头，但仓库为潮州仓，请确认是否正确')
                    show_toast('提示', '输入内容包含转接头，但仓库为潮州仓，请确认是否正确')
                found_warehouse = True
                break

        if not found_warehouse:
            if is_trans: # 如果有转接头 则默认为深圳仓
                action_list['warehouse'] = '深圳'
            else:
                action_list['warehouse'] = None  # 如果没有找到匹配的仓库，设置为None

        # 判断是不是仓库 如果是则不存入undefined_product列表 否则存入undefined_product列表 
        if not found_product and not found_warehouse:
            action_list['undefined_product'].append(item)
    if is_trans:
        action_list['remarks'].append('补发金属转接头')
    else:
        action_list['remarks'].append('补发')
    action_list['product_items'] = list(converted_products)
    return action_list


# *几 坐标点

def win_key(num=2):
    # 确保num是整数
    if not isinstance(num, int):
        print("num必须是整数")
        return
    # 循环发送 Win + 1 至 Win + 2
    for i in range(1, num+2):
        if i == 10:
            i = '0'
        # keyboard.press_and_release(f'win+{i}')
        pyautogui.hotkey('win', str(i))
        # print(f'按下了win+{i}')
        # time.sleep(0.1)