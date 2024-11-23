import pyperclip
import subprocess
import threading
import json
import time
import configparser
import keyboard
from plyer import notification

config = configparser.ConfigParser()
config.read('config.ini')

# 访问环境变量
bat_file_path = config['default']['BAT_FILE_PATH']
hotstrings_file_name = config['default']['HOT_FILE_NAME']

print(f"Batch file path: {bat_file_path}")  # 打印批处理文件路径
print(f"Hotstrings file name: {hotstrings_file_name}")  # 打印热字符串文件名

# 定义要监听的热字符串及其对应的批处理文件路径
with open(hotstrings_file_name, 'r', encoding='utf-8') as f:
    hotstrings = json.load(f)

# 将热字符串存储在一个集合中，以便更快地查找匹配项
hotstring_set = set(hotstrings.keys())

last_checked_time = 0
CHECK_INTERVAL = 0.5  # 检查间隔时间（秒）

previous_clipboard_content = ""

def show_toast(title, message, timeout=1):
    notification.notify(
        title=title,
        message=message,
        app_name="提醒",
        timeout=timeout,
        toast=True
    )

# 定义一个函数来处理热字符串
def on_press_clipboard(check_interval=None, check_duplicate=None):
    # mode_prompt = f"Starting listener...\nCheck interval: {check_interval}\nCheck duplicate: {check_duplicate}\n"
    # print(mode_prompt)
    # show_toast("提醒", mode_prompt, timeout=0.3)

    global last_checked_time, previous_clipboard_content

    current_time = time.time()
    if not check_interval or current_time - last_checked_time > CHECK_INTERVAL:
        last_checked_time = current_time
        # 获取当前复制的文本
        current_text = pyperclip.paste().replace(" ", "")  # 去掉多余的空格
        if not check_duplicate or current_text != previous_clipboard_content:
            previous_clipboard_content = current_text
            # 检查当前输入的文本是否是热字符串
            for hotstring in hotstring_set:
                if current_text.endswith(hotstring):
                    file_path = hotstrings[hotstring]
                    # 在独立线程中执行批处理文件
                    threading.Thread(target=execute_bat, args=(bat_file_path, file_path)).start()
                    show_toast("提醒", f"已执行 {file_path}", timeout=0.3)
                    print(f"Found hotstring: {hotstring}, executing batch file with argument: {file_path}")

def execute_bat(bat_file_path, file_path):
    try:
        result = subprocess.run([bat_file_path, file_path], check=True, capture_output=True, text=True)
        # print(f"Batch file executed successfully: {result.stdout}")
    except FileNotFoundError:
        print(f"Error: The batch file was not found at {bat_file_path}.")
    except subprocess.CalledProcessError as e:
        print(f"Error executing batch file: {e}, Output: {e.output}")

def clear_clipboard_content():
    global previous_clipboard_content
    previous_clipboard_content = ""
    print("Clipboard content cleared.")
    show_toast("提醒", "剪贴板内容已清除", timeout=0.3)

def start_listener(check_interval=None, check_duplicate=None, clear_on_combo=None):
    # 根据三个形参提示启用关闭了什么功能
    # mode_prompt = f"Starting listener...\nCheck interval: {check_interval}\nCheck duplicate: {check_duplicate}\nClear on combo: {clear_on_combo}"
    # print(mode_prompt)
    # show_toast("提醒", mode_prompt, timeout=0.3)

    # 绑定快捷键 Ctrl + Space
    # keyboard.add_hotkey('ctrl+space', on_press_clipboard())
    keyboard.add_hotkey('ctrl+space', lambda *args, f=on_press_clipboard, a=[check_interval, check_duplicate]: f(*a))

    # 绑定快捷键 Ctrl + Shift + Space
    if clear_on_combo:
        keyboard.add_hotkey('ctrl+shift+space', clear_clipboard_content)

def stop_listener():
    print('Stopping listener...')
    keyboard.unhook_all()
    print('Listener stopped')

if __name__ == '__main__':
    try:
        start_listener()
        while True:
            pass
    except KeyboardInterrupt:
        print('Exiting...')
    finally:
        stop_listener()