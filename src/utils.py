import os
import time
import keyboard 
import win32con 
import win32gui
import json
from threading import Thread
import configparser
from pathlib import Path
import pyperclip
from pynput import mouse
from functools import partial
import pyautogui

from plyer import notification
# from win10toast import ToastNotifier

handles = '../config/handles.json'

# 读取配置文件
def read_config():
    config_path = Path(__file__).resolve().parent.parent / 'config' / 'config.ini'
    
    if not config_path.exists():
        raise FileNotFoundError("utils read_config | Config file not found.")
    
    config = configparser.ConfigParser()
    config.read(config_path)
    return config

# 获取配置文件的指定段落中的指定选项值
def get_config_option(section, option):
    '''
    :param section: 配置文件段落名称
    :param option: 配置文件选项名称
    :return: 配置文件选项值
    '''
    config = read_config()

    # 检查段落是否存在
    if section not in config:
        raise ValueError(f"utils get_config_option | Section '{section}' is not defined in the config.")
    
    # 获取选项值
    value = config.get(section, option, fallback=None)

    if not value:
        raise ValueError(f"utils get_config_option | Option '{option}' is not defined in the config section '{section}'.")
    
    return value
    
# 获取 .bat 文件的路径
def get_bat_path():
    """
    从配置文件中读取 .bat 文件的路径，并返回绝对路径。
    
    :return: .bat 文件的绝对路径
    """
   
    relative_path = get_config_option('paths', 'COPY_CLIPBOARD_PATH')
    
    # 构建绝对路径
    base_dir = Path(__file__).resolve().parent.parent
    copy_clipboard_path = base_dir / relative_path
    
    if not copy_clipboard_path.exists():
        raise FileNotFoundError(f"utils get_bat_path | Bat file {relative_path} not found at {copy_clipboard_path}.")
    
    return str(copy_clipboard_path)


# 写入 JSON 文件
def write_json(file_path, data, encoding=None):
    with open(file_path, 'w', encoding=encoding) as f:
        json.dump(data, f, indent=4)

# 读取 JSON 文件
def read_json(file_path, encoding=None):
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding=encoding) as f:
            return json.load(f)
    else:
        return {}
    
# 存储句柄到文件
def save_handle(name, handle, file_path=handles, encoding=None):
    handles = load_handles(file_path, encoding)
    handles[name] = handle
    write_json(file_path, handles, encoding)

# 加载句柄
def load_handle(name, file_path=handles, encoding=None):
    handles = load_handles(file_path, encoding)
    return handles.get(name)

# 打开句柄文件
def load_handles(file_path=handles, encoding=None):
    return read_json(file_path, encoding)


# 显示通知
def show_toast(title='title', message='message', timeout=0.2):
    notification.notify(
        title=title,
        message=message,
        app_name="提醒",
        timeout=timeout,
        toast=True
    )

# 显示 Windows 10 通知
# toaster = ToastNotifier()
# def show_toast_win10(message):
#     toaster.show_toast("提醒", message, duration=1, threaded=True)

# 打开指定软件
def open_sof(name, handle=None, class_name=None):
    """
    打开指定软件窗口。
    :param name: 窗口名称
    :param handle: 句柄 (可选，手动模式下传入)
    :param class_name: 窗口类名 (可选，作为备用查找方法)
    """
    if handle:
        # 手动模式启用，直接使用传入的句柄
        if not win32gui.IsWindow(handle):
            print(f"手动模式下提供的句柄 {handle} 无效，切换到自动模式...")
            handle = None  # 自动切换到自动模式
        print(f"手动模式，指定句柄 {handle}。")
        save_handle(name, handle)  # 保存句柄到文件
        
    if not handle:
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
        win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)  # 最大化窗口
        time.sleep(0.2)  # 等待窗口状态稳定

        # 尝试将窗口置于前台
        if not win32gui.SetForegroundWindow(handle):
            print("无法将窗口置于前台，尝试重试...")
            for _ in range(3):  # 重试3次
                time.sleep(0.1)  # 等待
                if win32gui.SetForegroundWindow(handle):
                    print("窗口已成功置于前台。")
                    break
            else:
                print("无法将窗口置于前台，操作失败。")

        # 将窗口置于前台并最大化
        # win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
        # win32gui.SetForegroundWindow(handle)
        # time.sleep(0.1)  # 延迟0.1秒，等待窗口变化完成

        #  确保窗口恢复显示
        # win32gui.SendMessage(handle, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
        # time.sleep(0.2)  # 延迟0.2秒，等待窗口恢复
        # # 将窗口置于前台
        # win32gui.SetForegroundWindow(handle)
        # time.sleep(0.1)  # 延迟0.1秒
        # # 最大化窗口
        # win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
        # time.sleep(0.1)  # 延迟0.1秒
    except Exception as e:
        print(f"无法设置窗口为前台，错误信息：{e}")
        return None
        
    # pos = win32gui.GetWindowRect(handle) # 窗口位置
    # print(f'窗口位置 {pos}') # 打印窗口位置
# 使用快捷键运行指定函数
def auto_key(hotkeys):
    # 去除空格并移除重复的快捷键
    seen_keys = set()
    filtered_hotkeys = []
    for hotkey in hotkeys:
        # clean_key = hotkey['key'].replace(" ", "")  # 去掉快捷键中多余的空格
        clean_key = hotkey['key']
        if clean_key in seen_keys:
            print(f"发现重复快捷键 {clean_key}，保留第一个绑定，移除后续重复绑定。")
            continue  # 跳过重复的快捷键
        seen_keys.add(clean_key)
        filtered_hotkeys.append({
            "key": clean_key,
            "func": hotkey['func'],
            "args": hotkey.get('args', []),
            "use_thread": hotkey.get('use_thread', False),  # 默认不使用线程
            "redo": hotkey.get('redo', False)  # 是否需要冗余调用，默认不需要
        })
    
    # 绑定快捷键
    # 绑定快捷键
    def threaded_function(func, args, use_thread, redo):
        # 在独立线程中执行功能
        try:
            def execute():
                func(*args)
                
            if use_thread:
                if redo:
                    # 如果需要冗余调用，则创建两次线程
                    thread1 = Thread(target=execute)
                    thread2 = Thread(target=execute)
                    thread1.daemon = True
                    thread2.daemon = True
                    thread1.start()
                    thread2.start()
                else:
                    # 否则只创建一次
                    thread = Thread(target=execute)
                    thread.daemon = True
                    thread.start()
            else:
                if redo:
                    # 直接调用两次函数
                    execute()
                    execute()
                else:
                    execute()
        except Exception as e:
            print(f"快捷键功能执行出错：{e}")

    # 定义一个函数来处理侧键事件
    def handle_side_button(button, func, args, use_thread, redo):
        if button == mouse.Button.x1:  # 检测侧键1
            print("Side Button 1 pressed")
            threaded_function(func, args, use_thread, redo)
        elif button == mouse.Button.x2:  # 检测侧键2
            print("Side Button 2 pressed")
            threaded_function(func, args, use_thread, redo)

    try:
        # if 'sideButton' in hotkey['key']:
        #     # 使用 pynput 绑定侧键
        #     listener = mouse.Listener(
        #         on_click=partial(handle_side_button, 
        #                         func=hotkey['func'], 
        #                         args=hotkey.get('args', []), 
        #                         use_thread=hotkey.get('use_thread', False), 
        #                         redo=hotkey.get('redo', False))
        #     )
        #     listener.start()
        # else:
        for hotkey in filtered_hotkeys:
            keyboard.add_hotkey(
                hotkey['key'],
                lambda *args, f=hotkey['func'], a=hotkey['args'], u=hotkey['use_thread'], r=hotkey['redo']: threaded_function(f, a, u, r)
            )
        # 保持脚本运行，直到按下退出快捷键
        print("快捷键监听已启动，按下 Shift+Ctrl+E 退出\n")
        keyboard.wait('shift+ctrl+e')  # 按下 Shift+Ctrl+E 退出监听


        # hotkeys_dict = {
        #     hotkey['key']: lambda: threaded_function(hotkey['func'], hotkey.get('args', []), hotkey.get('use_thread', False), hotkey.get('redo', False))
        #     for hotkey in filtered_hotkeys
        # }

        # with keyboard.GlobalHotKeys(hotkeys_dict) as listener:
        #     print("快捷键监听已启动，按下 Shift+Ctrl+E 退出\n")
        #     listener.join()
    except KeyboardInterrupt:
        print("检测到 Ctrl+C，正在退出...")
    except Exception as e:
        print(f"快捷键监听出错：{e}")
    finally:
        # 无论是否按下 Shift+Ctrl+E 都移除所有快捷键监听
        # keyboard.unhook_all()
        # 清空句柄文件
        if os.path.exists(handles):
            os.remove(handles)
        print('退出监听')



def get_express_company(tracking_number):
    if tracking_number.startswith('4'):
        return "韵达"
    elif tracking_number.startswith('77'):
        return "申通"
    elif tracking_number.startswith('SF'):
        return "顺丰"
    elif tracking_number.startswith('JT'):
        return "极兔"
    elif tracking_number.startswith('788') or tracking_number.startswith('784') or tracking_number.startswith('756'):
        return "中通"
    else:
        return ""

def update_clipboard_express_company():
    # 读取剪切板的内容 大写
    tracking_number = pyperclip.paste().strip().upper()

    if not tracking_number:
        print("剪切板内容为空，退出更新")
        return
    tracking_number_len = len(tracking_number)
    # 判断是否为物流单号
    if tracking_number_len < 10 or tracking_number_len > 15:
        print("剪切板内容长度不符合要求，退出更新")
        show_toast('提醒', '剪切板内容长度不符合要求，退出更新')
        return
    # 获取快递公司
    express_company = get_express_company(tracking_number)

    # 构建新的剪切板内容
    if express_company:
        new_content = f"{express_company} 快递单号 {tracking_number} 亲亲这个是您的补发单号哈 注意查收~"
    else:
        new_content = f"快递单号 {tracking_number} 亲亲这个是您的补发单号哈 注意查收~"

    # 将新的内容覆盖到剪切板
    pyperclip.copy(new_content)

    print(f"剪切板内容已更新为：{new_content}")

# 移动鼠标
def move_mouse(x, y):
    pyautogui.moveTo(x, y)