import pyperclip
import subprocess
import threading
import json
import time
import configparser
import keyboard

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

# 定义一个函数来处理热字符串
def on_press_clipboard():
    global last_checked_time, previous_clipboard_content

    current_time = time.time()
    if current_time - last_checked_time > CHECK_INTERVAL:
        last_checked_time = current_time
        # 获取当前复制的文本
        current_text = pyperclip.paste()
        if current_text != previous_clipboard_content:
            previous_clipboard_content = current_text
            # 检查当前输入的文本是否是热字符串
            for hotstring in hotstring_set:
                if current_text.endswith(hotstring):
                    file_path = hotstrings[hotstring]
                    print(f"Found hotstring: {hotstring}, executing batch file with argument: {file_path}")
                    # 在独立线程中执行批处理文件
                    threading.Thread(target=execute_bat, args=(bat_file_path, file_path)).start()

def execute_bat(bat_file_path, file_path):
    try:
        result = subprocess.run([bat_file_path, file_path], check=True, capture_output=True, text=True)
        print(f"Batch file executed successfully: {result.stdout}")
    except FileNotFoundError:
        print(f"Error: The batch file was not found at {bat_file_path}.")
    except subprocess.CalledProcessError as e:
        print(f"Error executing batch file: {e}, Output: {e.output}")

def clear_clipboard_content():
    global previous_clipboard_content
    previous_clipboard_content = ""
    print("Clipboard content cleared.")

def start_listener():
    print('Start listener')
    # 绑定快捷键 Ctrl + Space
    keyboard.add_hotkey('ctrl+space', on_press_clipboard)
    # 绑定快捷键 Ctrl + Shift + Space
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