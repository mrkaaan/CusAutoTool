import pyperclip
from pynput import keyboard
import subprocess
import threading
import re
import json
import time
import configparser
import os

config = configparser.ConfigParser()
config.read('config.ini')

# 访问环境变量
bat_file_path = config['default']['BAT_FILE_PATH']

# 定义要监听的热字符串及其对应的批处理文件路径
with open('hotstrings.json', 'r', encoding='utf-8') as f:
    hotstrings = json.load(f)

# 用于跟踪按键状态的字典
keys_pressed = {
    keyboard.Key.shift_l: False,
    keyboard.Key.ctrl_r: False,
    keyboard.KeyCode.from_char('e'): False
}

last_checked_time = 0
CHECK_INTERVAL = 0.5  # 检查间隔时间（秒）

# 用于构建用户输入的字符串
user_input = []

# 定义一个函数来处理热字符串
def on_press_clipboard(key):
    try:
        global last_checked_time

        # 检查是否按下了空格键
        if key == keyboard.Key.space:
            current_time = time.time()
            if current_time - last_checked_time > CHECK_INTERVAL:
                last_checked_time = current_time
                # 获取当前复制的文本
                current_text = pyperclip.paste()
                # 检查当前输入的文本是否是热字符串
                for hotstring, file_path in hotstrings.items():
                    if current_text.endswith(hotstring):
                        # 在独立线程中执行批处理文件
                        # threading.Thread(target=execute_bat, args=(bat_file_path, file_path)).start()

                        execute_bat(bat_file_path, file_path)
                        # subprocess.run([bat_file_path, file_path], check=True)
                        print(f"Executed {file_path}")
    except AttributeError:
        pass

    # 跟踪按键状态
    if key in keys_pressed:
        keys_pressed[key] = True

def on_press_entry(key):
    global user_input
    try:
        if key == keyboard.Key.space:
            current_text = ''.join(user_input).strip()  # 去除前后空格
            print(f"Space pressed\nCurrent input: '{current_text}' (Length: {len(current_text)})")
            
            for hotstring, file_path in hotstrings.items():
                cleaned_hotstring = re.sub(r'\s+', '', hotstring)  # 去除热字符串中的所有空白字符
                cleaned_current_text = re.sub(r'\s+', '', current_text)  # 去除当前输入中的所有空白字符
                print(f"Checking against: '{cleaned_hotstring}' (Length: {len(cleaned_hotstring)})")
                if cleaned_current_text == cleaned_hotstring:
                    subprocess.run([bat_file_path, file_path], check=True)
                    print(f"Executed {file_path}")
                    user_input = []  # 清空用户输入
                    break  # 找到匹配的热字符串后跳出循环
            
            user_input = []  # 确保无论是否找到匹配的热字符串都清空输入
        elif hasattr(key, 'char') and key.char is not None:
            # 检查是否有修饰键（如 Ctrl、Alt、Shift）被按下
            if any(keys_pressed.get(modifier_key, False) for modifier_key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]):
                print(f"Ignored key event due to modifier: {key}")
            else:
                user_input.append(key.char)
                print(f"Key pressed: '{key.char}'")  # 打印每个按键事件
        else:
            print(f"Ignored key event: {key}")  # 打印忽略的按键事件
        
        # 跟踪按键状态
        if key in keys_pressed:
            keys_pressed[key] = True
            
    except AttributeError:
        pass

def on_release_combination(key):
    if key in keys_pressed:
        keys_pressed[key] = False
    # 检查特定组合键是否被释放
    if all(not value for value in keys_pressed.values()):
        print('特定组合键释放，退出监听')
        return False

def on_release_esc(key):
    if key == keyboard.Key.esc:
        # 停止监听
        print('退出监听')
        return False

def execute_bat(bat_file_path, file_path):
    # 执行对应的批处理文件
    subprocess.run([bat_file_path, file_path], check=True)

listener = None

def start_listener():
    global listener
    print('Start listener')
    listener = keyboard.Listener(on_press=on_press_clipboard, on_release=on_release_combination)
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