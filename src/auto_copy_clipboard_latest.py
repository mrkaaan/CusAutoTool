import pyperclip
import subprocess
import threading
import time
import keyboard
import pyautogui
from utils import show_toast, read_json
from config import setup_bat_path, setup_hot_file_name

# 读取配置文件
bat_file_path = setup_bat_path()
if not bat_file_path:
    print("未设置批处理文件路径，请在 config.py 中设置")
    exit()
hotstrings_file_name = setup_hot_file_name()
if not hotstrings_file_name:
    print("未设置热键配置文件，请在 config.py 中设置")
    exit()
hotstrings_path = f"../config/{hotstrings_file_name}"
# 定义要监听的热字符串及其对应的批处理文件路径
hotstrings = read_json(hotstrings_path, 'utf-8')
if not hotstrings:
    print(f"热键配置文件 {hotstrings_file_name} 为空，请检查文件内容")
    exit()

hotstrings_set = set(hotstrings.keys())

print(f"Batch file path: {bat_file_path}")  # 打印批处理文件路径
print(f"Hotstrings file name: {hotstrings_file_name}")  # 打印热字符串文件名

last_checked_time = 0
CHECK_INTERVAL = 0.5  # 检查间隔时间（秒）

previous_clipboard_content = ""


# 定义一个函数来处理热字符串
def on_press_clipboard(specify_filename='', auto_copy=True, check_interval=True, check_duplicate=False):
    '''
        :param auto_copy: 是否自动选择当前输入框内容 默认True
        :param check_interval: 是否检查间隔 默认False
        :param check_duplicate: 是否检查重复 默认False
    '''
    # mode_prompt = f"Starting listener...\nCheck interval: {check_interval}\nCheck duplicate: {check_duplicate}\n"
    # print(mode_prompt)
    # show_toast("提醒", mode_prompt)
    print('Starting listener...')
    global last_checked_time, previous_clipboard_content

    current_time = time.time()
    current_text = ""

    if auto_copy:
        # keyboard.press_and_release('ctrl+a')  
        # keyboard.press_and_release('ctrl+x')

        pyautogui.hotkey('ctrl', 'a')  # 模拟按下并释放
        pyautogui.hotkey('ctrl', 'x')  # 模拟按下并释放

    current_text = pyperclip.paste().replace(" ", "").lower()  # 获取剪贴板中的文本 去掉多余的空格

        # 空文本限制
    if not current_text:
        print(f"Empty text: {current_text}")
        show_toast("提醒", f"空文本: {current_text}")
        return
    
    # 限制文本长度
    if len(current_text) > 15:
        print(f"Text length exceeds 15 characters")
        show_toast("提醒", f"文本长度超过 15 个字符")
        if auto_copy:
            keyboard.press_and_release('ctrl+z')  # 模拟按下并释放
        return
    
    time.sleep(0.1)  # 等待复制操作完成

    is_find_hotstring = False

    # 检查间隔
    if not check_interval or current_time - last_checked_time > CHECK_INTERVAL:
        last_checked_time = current_time

        # 检查剪切板是否重复
        if not check_duplicate or current_text != previous_clipboard_content:
            previous_clipboard_content = current_text

            if specify_filename:
                current_text = specify_filename

            # 检查当前输入的文本是否是热字符串
            for hotstring in hotstrings_set:
                if current_text.endswith(hotstring):
                    is_find_hotstring = True
                    file_path = hotstrings[hotstring]

                    # 在独立线程中执行批处理文件
                    threading.Thread(target=execute_bat, args=(bat_file_path, file_path)).start()
                    show_toast("提醒", f"已执行 {file_path}")
                    print(f"Found hotstring: {hotstring}, executing batch file with argument: {file_path}")
                    break
            if not is_find_hotstring:
                print(f"No hotstring found: {current_text}")
                show_toast("提醒", f"未找到字符串: '{current_text}'")
        
# 定义一个函数来执行批处理文件
def execute_bat(bat_file_path, file_path):
    try:
        result = subprocess.run([bat_file_path, file_path], check=True, capture_output=True, text=True)
        # print(f"Batch file executed successfully: {result.stdout}")
    except FileNotFoundError:
        print(f"Error: The batch file was not found at {bat_file_path}.")
    except subprocess.CalledProcessError as e:
        print(f"Error executing batch file: {e}, Output: {e.output}")

# 清空剪切板 仅在当前文件启动时生效
def clear_clipboard_content():
    global previous_clipboard_content
    previous_clipboard_content = ""
    print("Clipboard content cleared.")
    show_toast("提醒", "剪贴板内容已清除")

# 清空剪切板 全局
def clear_clipboard():
    # 使用pyperclip库清空剪切板
    pyperclip.copy('')
    print("Clipboard content cleared.")
    show_toast("提醒", "剪贴板内容已清除")

# 仅在运行当前文件时生效
def start_listener(check_interval=None, check_duplicate=None, clear_on_combo=None):
    '''
        :param check_interval: 是否检查间隔 默认None
        :param check_duplicate: 是否检查重复 默认None
        :param clear_on_combo: 是否按下 Ctrl + Shift + Space 后清除剪切板内容 默认None
    '''
    # 根据三个形参提示启用关闭了什么功能
    mode_prompt = f"Starting listener...\nCheck interval: {check_interval}\nCheck duplicate: {check_duplicate}\nClear on combo: {clear_on_combo}"
    print(mode_prompt)
    show_toast("提醒", mode_prompt)

    # 绑定快捷键 Ctrl + Space
    keyboard.add_hotkey('ctrl+space', on_press_clipboard())

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



