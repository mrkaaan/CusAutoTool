from WinGUI import WinGUI
import organize_table as tb
import automation as au
import utils as ut
import keyboard_auto_file as kf

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

if __name__ == "__main__":

    # ---------- 配置参数 -------------
    pyautogui.FAILSAFE = False  # 关闭 pyautogui 的故障保护机制
    pyautogui.PAUSE = 0.1  # 设置 pyautogui 的默认操作延时 如移动和点击的间隔
    logger.add("dev.log", rotation="10 MB")  # 设置日志文件轮换
    window_name = r"千牛接待台"  # 窗口名称
    window_name_erp = r"旺店通ERP"
    
    # ---------- 快捷键启动 -------------
    hotkey_actions = [
        # {'key': 'alt+shift+e', 'func': ut.open_sof, 'args': ['旺店通ERP',265632,1,'CoolWindow']},
        {'key': 'alt+shift+e', 'func': ut.open_sof, 'args': ['旺店通ERP']}, # 打开制定软件
        {'key': 'ctrl+shift+o', 'func': au.run_once_remarks_by_qianniu, 'args': [window_name]}, # 添加备注 并 取消标记
        {'key': 'ctrl+shift+u', 'func': au.run_once_unmark_by_qianniu, 'args': [window_name]}, # 取消标记
        {'key': 'f2', 'func': kf.on_press_clipboard}, # 自动识别输入框并替换剪切板内容为图片或视频文件
        {'key': 'ctrl+space', 'func': kf.on_press_clipboard}, # 自动识别输入框并替换剪切板内容为图片或视频文件
        {'key': 'ctrl+shift+x', 'func': kf.update_clipboard}, # 默认剪切板为物流单号 拼凑为指定格式
    ]
    ut.auto_key_with_threads(hotkey_actions)


    # ---------- 通知补发单号 -------------
    notification_reissue_parameter = {
        'window_name': window_name, # 窗口名称
        'table_name': '2024-11-27_230453_处理结果.xlsx',  # 表格名称
        'notic_shop_name': '潮洁', # 通知店铺名称
        'use_today': '2024-11-27', # 使用日期 如果为空则使用当天日期
    }
    # au.notification_reissue(**notification_reissue_parameter)


    # ---------- 表格处理 -------------
    # tb.process_table('11月27日天猫补发单号.csv')

    