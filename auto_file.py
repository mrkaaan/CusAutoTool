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

# 定义一个函数来处理热字符串
def on_press(key):
    try:
        # 检查是否按下了空格键
        if key == keyboard.Key.space:
            # 获取当前输入的文本
            current_text = pyperclip.paste()
            # 检查当前输入的文本是否是热字符串
            for hotstring, file_path in hotstrings.items():
                if current_text.endswith(hotstring):
                    # 在独立线程中执行批处理文件
                    threading.Thread(target=execute_bat, args=(bat_file_path, file_path)).start()
                    print(f"Executed {file_path}")
    except AttributeError:
        pass

def execute_bat(bat_file_path, file_path):
    # 执行对应的批处理文件
    subprocess.run([bat_file_path, file_path], check=True)

# 定义一个函数来处理退出快捷键
def on_release(key):
    if key == keyboard.Key.shift_l and keyboard.Controller().pressed(keyboard.Key.ctrl_l) and keyboard.Controller().pressed(keyboard.Key.e):
        # 停止监听
        print('退出监听')
        return False

# 监听键盘
print("快捷文件监听已启动，按下 Shift+Ctrl+E 退出。")
with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()