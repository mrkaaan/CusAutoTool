import pyperclip
from pynput import keyboard
import subprocess
import threading

bat_file_path = r'D:\project\auto_customer\auto_file.bat'

# 定义要监听的热字符串及其对应的批处理文件路径
hotstrings = {
    'test1': 'D:\\test.png',
    'test2': r'C:\Users\Kan\Documents\Captura\2024-11-15-11-50-10.mp4'
}

# 用于跟踪按键状态的字典 用于函数 on_press_clipboard 和 on_release_combination
keys_pressed = {
    keyboard.Key.shift_l: False,
    keyboard.Key.ctrl_r: False,
    keyboard.KeyCode.from_char('e'): False
}

# 用于构建用户输入的字符串 用户函数 on_press_entry
user_input = []

# 定义一个函数来处理热字符串
def on_press_clipboard(key):
    try:
        # 检查是否按下了空格键
        if key == keyboard.Key.space:
            # 获取当前复制的文本
            current_text = pyperclip.paste()
            # 检查当前输入的文本是否是热字符串
            for hotstring, file_path in hotstrings.items():
                if current_text.endswith(hotstring):
                    # 在独立线程中执行批处理文件
                    threading.Thread(target=execute_bat, args=(bat_file_path, file_path)).start()
                    # subprocess.run([bat_file_path, file_path], check=True)
                    print(f"Executed {file_path}")
    except AttributeError:
        pass

    # 跟踪按键状态
    if key in keys_pressed:
        keys_pressed[key] = True

# 定义一个函数来处理热字符串
def on_press_entry(key):
    global user_input
    try:
        # 将按键添加到用户输入字符串
        if key == keyboard.Key.space:
            # 获取当前输入的文本
            current_text = ''.join(user_input)
            # 检查当前输入的文本是否是热字符串
            for hotstring, file_path in hotstrings.items():
                if current_text == hotstring:
                    # 执行对应的批处理文件
                    subprocess.run([bat_file_path, file_path], check=True)
                    print(f"Executed {file_path}")
                    # 清空用户输入
                    user_input = []
                    break
        elif hasattr(key, 'char') and key.char is not None:
            user_input.append(key.char)
    except AttributeError:
        pass

def on_release_combination(key):
    if key in keys_pressed:
        keys_pressed[key] = False

def on_release_esc(key):
    if key == keyboard.Key.esc:
        # 停止监听
        print('退出监听')
        return False

def execute_bat(bat_file_path, file_path):
    # 执行对应的批处理文件
    subprocess.run([bat_file_path, file_path], check=True)

# 监听键盘
# with keyboard.Listener(on_press=on_press_clipboard, on_release=on_release_combination) as listener:
def start_listener():
    print('Start listener')
    with keyboard.Listener(on_press=on_press_entry, on_release=on_release_combination) as listener:
        listener.join()


# # 在独立线程中启动监听
# listener_thread = threading.Thread(target=start_listener)
# listener_thread.start()

# # 保持主线程运行，以便可以响应 Ctrl+C
# try:
#     while listener_thread.is_alive():
#         pass
# except KeyboardInterrupt:
#     print('退出监听')
#     listener_thread.join()

start_listener()