import pyperclip
from pynput import keyboard
import subprocess
import threading
import json
import time
import configparser

config = configparser.ConfigParser()
config.read('config.ini')

# 访问环境变量
bat_file_path = config['default']['BAT_FILE_PATH']
hotstrings_file_name = config['default']['HOT_FILE_NAME']

# 定义要监听的热字符串及其对应的批处理文件路径
with open(hotstrings_file_name, 'r', encoding='utf-8') as f:
    hotstrings = json.load(f)


# 将热字符串存储在一个集合中，以便更快地查找匹配项
hotstring_set = set(hotstrings.keys())


last_checked_time = 0
CHECK_INTERVAL = 0.5  # 检查间隔时间（秒）

previous_clipboard_content = ""
ctrl_pressed = False
shift_pressed = False  # 检测Shift键是否被按下

# 定义一个函数来处理热字符串
def on_press_clipboard(key):
    try:
        global last_checked_time, previous_clipboard_content, ctrl_pressed, shift_pressed

        # 检查是否按下了Ctrl或Shift键
        if key == keyboard.Key.ctrl_l or key == keyboard.Key.ctrl_r:
            ctrl_pressed = True
        elif key == keyboard.Key.shift or key == keyboard.Key.shift_r:
            shift_pressed = True
        elif key == keyboard.Key.space and ctrl_pressed and shift_pressed:
            # 清除previous_clipboard_content变量
            previous_clipboard_content = ""
            print("Clipboard content cleared.")
        elif key == keyboard.Key.space and ctrl_pressed:
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
                            # execute_bat(bat_file_path, file_path)
                            # 在独立线程中执行批处理文件
                            threading.Thread(target=execute_bat, args=(bat_file_path, file_path)).start()
                            print(f"Executed {file_path}")
    except AttributeError:
        pass

def execute_bat(bat_file_path, file_path):
    try:
        result = subprocess.run([bat_file_path, file_path], check=True, capture_output=True, text=True)
        print(f"Batch file executed successfully: {result.stdout}")
    except subprocess.CalledProcessError as e:
        print(f"Error executing batch file: {e}, Output: {e.output}")

def on_release(key):
    global ctrl_pressed, shift_pressed

    if key == keyboard.Key.ctrl_l or key == keyboard.Key.ctrl_r:
        ctrl_pressed = False
    elif key == keyboard.Key.shift or key == keyboard.Key.shift_r:
        shift_pressed = False

listener = None

def start_listener():
    global listener
    print('Start listener')
    listener = keyboard.Listener(on_press=on_press_clipboard, on_release=on_release)
    listener.start()

def stop_listener():
    global listener
    if listener:
        listener.stop()
        listener.join()
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