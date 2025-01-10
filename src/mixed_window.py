import os
import json
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk, messagebox
from datetime import datetime
import shutil
import re
import openpyxl
import configparser
import time
from utils import show_toast
import threading
import pyperclip

from organize_table import process_table
import auto_operation as au
from pyautogui import FailSafeException
from organize_table_window import choose_table, update_button_state
from notification_reissue_window import save_config

config = configparser.ConfigParser()
config.read('../config/window_config.ini')
# 访问环境变量
window_open_mode = config['defaults'].get('WINDOW_OPEN_MODE')
# 尝试将 TRY_NUMBER 转换为整数
window_open_mode = int(window_open_mode)

table_path_handle = ''
table_name_handle = ''
table_path_noti = ''
table_name_noti = ''

master = None
master_initialized = False  # 标志 window 是否已经初始化

DEFAULT_VALUES = {
    "defaults": {
        "window_name": "千牛接待台",
        "table_name": "xxx_处理结果.xlsx",
        "table_path": "",
        "notic_shop_name": "",
        "notic_mode": 2,
        "show_logistics": True,
        "logistics_mode": 1,
        "use_today": "",
        "test_mode": "0",
        "is_write": True
    },
    "last_used": {}
}

CONFIG_FILE = '../config/notic_config.json'

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        # 如果文件不存在，则创建一个默认的配置
        save_config(DEFAULT_VALUES)  # 保存默认配置
        return DEFAULT_VALUES
    
# 定义一个回调函数，在主线程中调用create_window
def call_create_window():
    global master
    if master is not None and master.winfo_exists():
        print("窗口已存在，不再创建")
        # 使用 wm_attributes 来确保窗口会暂时位于最顶层
        master.attributes('-topmost', True)
        master.after(50, lambda: master.attributes('-topmost', False))
        master.lift()  # 将窗口提升到最前面
        master.after(50, master.focus_force)  # 强制窗口获得焦点，并添加小延迟
    else:
        create_window(mode=window_open_mode)

def create_window(mode=0):
    '''
        :param mode: 0 当前文件启动 1 被其他程序调用启动 由于在utils中调用是非主线程，会导致窗口大小不同，设置了不同的窗口大小和位置参数
    '''
    print(f"create_window mode: {mode} {'主程序模式' if mode == 0 else '被调用模式'}")
    global master, master_initialized
    if master is not None and master.winfo_exists():
        master.lift()  # 如果窗口已经存在，将其提升到最前面
        return
    # 读取配置文件
    notic_config = load_config()

    # 获取全局表格变量
    global table_path_handle, table_name_handle, table_path_noti, table_name_noti

    master = tk.Tk()
    master.title("Handle Notic")

    # 设置窗口大小
    window_width = 400 if mode == 0 else 520
    window_height = 350 if mode == 0 else 350

    # 设置窗口位于屏幕的中间
    scnwidth, scnheight = master.maxsize()
    x_offset = (scnwidth - window_width) / 2
    y_offset = (scnheight - window_height) / 2
    master.geometry(f'{window_width}x{window_height}+{int(x_offset)}+{int(y_offset)}')