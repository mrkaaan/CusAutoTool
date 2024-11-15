import pyperclip
from pynput import keyboard
import subprocess

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
                    # 执行对应的批处理文件
                    subprocess.run([bat_file_path, file_path], check=True)
                    print(f"Executed {file_path}")
    except AttributeError:
        pass

# 监听键盘
with keyboard.Listener(on_press=on_press) as listener:
    print("快捷文件监听已启动")
    listener.join()