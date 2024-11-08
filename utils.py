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

# json操作
import json

import threading


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

# 打开句柄文件
def load_handles():
    if os.path.exists('handles.json'):
        with open('handles.json', 'r') as f:
            return json.load(f)
    else:
        return {}

# 打开指定软件
def open_sof(name, handle=None, manual=0, class_name=None):
    """
    打开指定软件窗口。
    :param name: 窗口名称
    :param handle: 句柄 (可选，手动模式下传入)
    :param manual: 手动模式标志 (0: 自动模式, 1: 手动模式)
    :param class_name: 窗口类名 (可选，作为备用查找方法)
    """
    if manual:
        # 手动模式启用，直接使用传入的句柄
        if not handle:
            print("手动模式下未提供句柄！")
            return None
        if not win32gui.IsWindow(handle):
            print(f"手动模式下提供的句柄 {handle} 无效，切换到自动模式...")
            manual = 0  # 自动切换到自动模式
        print(f"手动模式，指定句柄 {handle}。")
        save_handle(name, handle)  # 保存句柄到文件
        
    if not manual:
        # 自动模式：尝试通过文件加载或窗口名称查找句柄
        handle = load_handle(name)  # 从文件加载句柄
        if not handle or not win32gui.IsWindow(handle):
            print(f"文件中未找到有效句柄，尝试通过窗口名称 {name} 查找...")
            handle = win32gui.FindWindow(None, name)  # 通过窗口标题查找句柄
            if not handle and class_name:
                print(f"未通过窗口名称找到句柄，尝试使用类名 {class_name} 查找...")
                handle = win32gui.FindWindow(class_name, None)  # 通过类名查找句柄
            
            if not handle:
                print(f"未找到有效的窗口句柄，操作终止。")
                return None  # 所有查找方式均失败，终止运行
            
            save_handle(name, handle)  # 保存新找到的句柄
        else:
            print(f"文件中找到有效句柄：{handle}")

    try:
        win32gui.SendMessage(handle, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)  # 发送消息以确保窗口恢复显示
        time.sleep(0.2) # 延迟1秒，等待窗口恢复

        win32gui.ShowWindow(handle, True)          # 使用win32gui库使窗口可见
        win32gui.SetForegroundWindow(handle)       # 使用win32gui库将窗口置于前台(焦点)
        win32gui.SendMessage(handle, win32con.SC_MAXIMIZE, 0) # 最大化窗口
        time.sleep(0.2)                              # 使用time库延迟1秒
    except Exception as e:
        print(f"无法设置窗口为前台，错误信息：{e}")
        return None
        
    # pos = win32gui.GetWindowRect(handle) # 窗口位置
    # print(f'窗口位置 {pos}') # 打印窗口位置

def auto_key_with_threads(hotkeys):
    # 去除空格并移除重复的快捷键
    seen_keys = set()
    filtered_hotkeys = []
    for hotkey in hotkeys:
        clean_key = hotkey['key'].replace(" ", "")  # 去掉快捷键中多余的空格
        if clean_key in seen_keys:
            print(f"发现重复快捷键 {clean_key}，保留第一个绑定，移除后续重复绑定。")
            continue  # 跳过重复的快捷键
        seen_keys.add(clean_key)
        filtered_hotkeys.append({
            "key": clean_key,
            "func": hotkey['func'],
            "args": hotkey.get('args', [])
        })
    
    # 绑定快捷键
    def threaded_function(func, args):
        # 在独立线程中执行功能
        try:
            thread = threading.Thread(target=func, args=args)
            thread.daemon = True  # 守护线程，主线程结束后自动清理
            thread.start()
        except Exception as e:
            print(f"快捷键功能执行出错：{e}")
    
    for hotkey in filtered_hotkeys:
        keyboard.add_hotkey(
            hotkey['key'],
            lambda *args, f=hotkey['func'], a=hotkey['args']: threaded_function(f, a)
        )

    try:
        # 保持脚本运行，直到按下退出快捷键
        print("快捷键监听已启动，按下 Shift+Ctrl+E 退出。")
        keyboard.wait('shift+ctrl+e')  # 按下 Shift+Ctrl+E 退出监听
    finally:
        # 无论是否按下 Shift+Ctrl+E 都移除所有快捷键监听
        keyboard.unhook_all()
        # 清空句柄文件
        if os.path.exists('handles.json'):
            os.remove('handles.json')
        print('退出监听')
