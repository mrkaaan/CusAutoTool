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

from PIL import ImageGrab     # 从PIL库导入ImageGrab模块，用于截图
from loguru import logger     # 引入loguru库，用于简便的日志记录

class WinGUI:
    """
    WinGUI类用于控制Windows GUI界面，包括窗口截图、鼠标点击等操作。

    Attributes:
        window_name (str): 窗口名称，用于识别目标窗口。
        work_screen_path (str): 保存工作区域截图的路径。
        app_screen_path (str): 保存应用窗口截图的路径。
        img_folder_path (str): 图标文件的目录路径。
    """
    def __init__(self, window_name):
        """
        初始化WinGUI类。
        
        Parameters:
            window_name (str): 目标窗口的名称，用于定位窗口句柄。
        """
        self.window_name = window_name
        self.work_screen_path = "./temp/work_screen.png"
        self.app_screen_path = "./temp/app.png"
        self.img_folder_path = "./image" 

    # 需要 self.window_name
    def get_app_screenshot(self):
        """
        前置并获取指定窗口的截图并保存为app_screen_path。

        Returns:
            tuple: 返回窗口的坐标位置（left, top, right, bottom）。

        使用库：
            - `win32gui`: `ShowWindow`用于显示窗口，`SetForegroundWindow`将窗口置前。
            - `ImageGrab` (PIL库): `grab`方法用于截图。
        """
        # 获取窗口句柄和位置
        original_position, handle = get_window_pos(self.window_name)

        # 将窗口置前并高亮显示
        win32gui.ShowWindow(handle, True)          # 使用win32gui库使窗口可见
        win32gui.SetForegroundWindow(handle)       # 使用win32gui库将窗口置于前台(焦点)
        win32gui.SendMessage(handle, win32con.SC_MAXIMIZE, 0)
        time.sleep(0.2)                              # 使用time库延迟1秒

        print(f'窗口位置 {win32gui.GetWindowRect(handle)}') # 打印窗口位置

        # 截图并保存
        img_ready = ImageGrab.grab(original_position)  # 使用ImageGrab库截取指定区域
        img_ready.save(self.app_screen_path)           # 保存截图到指定路径
        return original_position
   
    def get_workscreen_screenshot(self):
        """
        获取整个工作屏幕的截图并保存为work_screen_path。

        Returns:
            Image: 返回截取的屏幕图像对象。

        使用库：
            - `ImageGrab` (PIL库): `grab`方法用于截图。
        """
        screenshot = ImageGrab.grab()                  # 使用ImageGrab库截取整个屏幕
        
        # 保存截图
        if screenshot:
            screenshot.save(self.work_screen_path)     # 保存截图到指定路径
            return screenshot
        return None


    def move_and_click(self, x, y, direction='left'):
        """
        移动鼠标到指定坐标并点击。

        Parameters:
            x (int): 鼠标点击的x坐标。
            y (int): 鼠标点击的y坐标。
            direction: 鼠标左右键 默认left

        使用库：
            - `pyautogui`: `moveTo`和`click`方法用于移动和点击。
            - `time`: 延时操作。
        """
        pyautogui.moveTo(x, y)                         # 使用pyautogui库移动鼠标
        pyautogui.click(x, y, button=direction)                          # 使用pyautogui库点击鼠标
        time.sleep(0.1)                                # 使用time库延迟1秒
    
    # 相对移动
    def rel_remove_and_click(self, x_add, y_add):
        """
        移动鼠标到指定坐标并点击。

        Parameters:
            x (int): 鼠标相对移动并点击的x坐标。
            y (int): 鼠标相对移动并点击的y坐标。

        使用库：
            - `pyautogui`: `moveTo`和`click`方法用于移动和点击。
            - `time`: 延时操作。
        """
        # 获取当前位置
        current_x, current_y = pyautogui.position()
        pyautogui.move(current_x + x_add, current_y + y_add)                         # 使用pyautogui库移动鼠标
        pyautogui.click()                          # 使用pyautogui库点击鼠标
        time.sleep(0.1)  

    # 需要 icon_path
    # 调用 locate_icon 和 move_and_click
    def click_icon(self, icon_path,x_start_ratio=0,x_end_ratio=1.0,y_start_ratio=0,y_end_ratio=1.0,direction='left'):
        """
        定位并点击指定的图标。

        Parameters:
            icon_path (str): 图标的文件路径，用于定位图标位置。

        使用库：
            - 调用 `locate_icon` 方法定位图标。
            - `move_and_click`: 调用该方法移动鼠标并点击。
        """
        x, y, is_find = self.locate_icon(icon_path,x_start_ratio,x_end_ratio,y_start_ratio,y_end_ratio) # 调用locate_icon方法获取图标在屏幕上的位置
        self.move_and_click(x, y, direction)                      # 使用move_and_click方法点击图标
        return is_find


    # 需要 img_name、self.img_folder_path 构建图标文件路径
    # 调用 get_app_screenshot 获取应用窗口截图
    # 需要 self.app_screen_path 获取应用窗口截图路径
    def locate_icon(
        self,
        img_name,
        x_start_ratio=0,
        x_end_ratio=1.0,
        y_start_ratio=0,
        y_end_ratio=1.0,
        try_number=3
    ):
        """
        查找目标图标的中心坐标，多个相似图标可通过指定搜索区域来定位

        Parameters:
            img_name (str): 图标文件名。
            x_start_ratio (float): x轴搜索起始比例（0到1之间），默认为0。
            x_end_ratio (float): x轴搜索结束比例（0到1之间），默认为1.0。
            y_start_ratio (float): y轴搜索起始比例（0到1之间），默认为0。
            y_end_ratio (float): y轴搜索结束比例（0到1之间），默认为1.0。
            try_number (int): 尝试查找的次数，默认为3次 因为应用程序界面刷新需要时间

        Returns:
            tuple: 目标图标中心的坐标 (result_x, result_y)。如果未找到，返回 (-1, -1)。
        
        使用库：
            - `os`: `path.join`用于组合文件路径。
            - `cv2`: `imread`读取图像，`matchTemplate`进行模板匹配，`minMaxLoc`获取匹配结果的最大相似位置。
            - `math`: `floor`和`ceil`用于确定搜索区域。
            - `logger`: 记录匹配相似度低的信息。
        """
        # return the coordinates of the center point of the target icon
        print()
        # 打印图标文件名
        logger.info(f'定位模板位置 {img_name} ')
        # 构造图标文件路径
        obj_path = os.path.join(self.img_folder_path, img_name)

        # 初始化返回值为未找到坐标
        result_x, result_y = -1, -1
        for i in range(try_number):
            print(f"第{i + 1}次查找")
            # print("mouse: ", pyautogui.position())

            # 获取应用窗口的截图及其起始坐标(:2代表左上角坐标)
            x_init, y_init = self.get_app_screenshot()[:2]
            source = cv.imread(self.app_screen_path)    # 使用cv2库读取截图

            # 获取截图的宽高和颜色深度(RGB)
            h, w, d = source.shape
            # print(f"original size: (width {w}, height {h}")

            # 计算裁剪区域
            # x direction: left to right
            # y direction: top to bottom
            x_start = math.floor(w * x_start_ratio)
            x_end = math.ceil(w * x_end_ratio)
            y_start = math.floor(h * y_start_ratio)
            y_end = math.floor(h * y_end_ratio)
            # print(f"crop_location: {x_start, y_start, x_end, y_end}")
            
             # 裁剪指定区域的图像
            source = source[y_start : y_end + 1, x_start : x_end + 1]   # +1 用于包含 x_end 和 y_end 边界，以确保边界像素点也在截取区域内
            # print(f"crop size: width {source.shape[1]}, height {source.shape[0]}")
            cv.imwrite(self.app_screen_path, source)    # 保存裁剪后的图像
            
            # 读取图标模板
            template = cv.imread(obj_path)              # 使用cv2库读取模板图像
            # print(f"template_size: {template.shape}")

            # 进行模板匹配
            result = cv.matchTemplate(source, template, cv.TM_CCOEFF_NORMED)
            similarity = cv.minMaxLoc(result)[1]       # 获取最大相似度
            # minMaxLoc(result) 返回匹配结果矩阵result的四个值：最小相似度、最大相似度、最小值位置 (x, y)、最大值位置 (x, y)

            # 如果相似度低于阈值，记录信息
            is_find = True
            if similarity < 0.90:
                logger.info(f"not find low similarity: {img_name}")   # 使用loguru库记录信息  低相似值
                logger.info(f'maximum value position: {cv.minMaxLoc(result)[3]}') # 最大值的位置
                is_find = False
            else:
                # 获取匹配位置并计算图标中心坐标(左上角坐标)
                pos_start = cv.minMaxLoc(result)[3]
                result_x = (
                    x_init + x_start + int(pos_start[0]) + int(template.shape[1] / 2) # 窗口左上角 + 裁剪区域左上角 + 图标左上角 + 图标宽度的一半
                )
                result_y = (
                    y_init + y_start + int(pos_start[1]) + int(template.shape[0] / 2) # 窗口左上角 + 裁剪区域左上角 + 图标左上角 + 图标高度的一半
                )
                break

        # 图标在整个屏幕上的中心坐标
        return result_x, result_y, is_find

    # 调用 locate_icon 没用了
    def check_icon(
        self,
        img_name,
        x_start_ratio=0,
        x_end_ratio=1.0,
        y_start_ratio=0,
        y_end_ratio=1.0,
    ):
        """
        检查图标是否存在于指定区域中。

        Parameters:
            img_name (str): 图标文件名。
            x_start_ratio (float): x轴搜索起始比例，默认为0。
            x_end_ratio (float): x轴搜索结束比例，默认为1.0。
            y_start_ratio (float): y轴搜索起始比例，默认为0。
            y_end_ratio (float): y轴搜索结束比例，默认为1.0。

        Returns:
            tuple: 返回 (是否找到, x坐标, y坐标)。
        """
        # 调用 locate_icon 方法获取图标坐标
        x, y, is_find = self.locate_icon(
            img_name,
            x_start_ratio,
            x_end_ratio,
            y_start_ratio,
            y_end_ratio,
            try_number=3,
        )
        # 如果坐标无效，则返回未找到状态
        if x < 0 or y < 0:
            return False
        return True

# 需要窗口名称 从而获取窗口位置和句柄
def get_window_pos(name):
    """
    获取指定名称窗口的位置和句柄。

    Parameters:
        name (str): 窗口的名称。

    Returns:
        tuple: 窗口的坐标 (left, top, right, bottom) 和句柄。如果未找到窗口，则返回 None。

    使用库：
        - `win32gui`: `FindWindow`用于根据名称查找窗口句柄，`GetWindowRect`获取窗口位置，`SendMessage`控制窗口状态。
        - `win32con`: 提供 Windows API 常量，例如 `WM_SYSCOMMAND` 和 `SC_RESTORE`。
        - `time`: `sleep`用于等待窗口恢复显示。
    """
    handle = win32gui.FindWindow(0, name)   # 获取窗口句柄 忽略窗口的类名只根据窗口标题来查找窗口
    if handle == 0:
        return None
    else:
        # 发送消息以确保窗口恢复显示
        win32gui.SendMessage(handle, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)  
        # 发送“恢复”命令 (SC_RESTORE)，确保窗口在非最小化状态。若窗口已经最小化将它恢复为正常大小，但不会将窗口置于前台或焦点
        time.sleep(0.2)                                  # 延迟1秒，等待窗口恢复
        return win32gui.GetWindowRect(handle), handle  # 返回窗口位置和句柄

def move_files(original_folder, target_folder, suffix_list=[]):
    """
    将 original_folder 文件夹中指定后缀的文件移动到 target_folder 的新子文件夹中。
    
    :param original_folder: 要移动文件的来源文件夹路径
    :param target_folder: 文件移动的目标文件夹路径
    :param suffix_list: 文件后缀列表，只移动符合这些后缀的文件
    """
    data_list = []
    dir_list = os.listdir(original_folder)  # 获取 original_folder 中的文件列表
    if len(suffix_list) == 0:   # 如果没有指定后缀，移动所有文件
        data_list = dir_list
    else:
        # 遍历文件名并根据后缀列表筛选
        for file_name in dir_list:
            for suffix in suffix_list:
                if file_name.endswith(suffix):
                    data_list.append(file_name)
                    # print(file_name)
    
    if len(data_list) == 0: # 如果没有符合条件的文件，则退出
        return

    cur_time = time.strftime("%Y-%m-%d_%H_%M_%S")  # 生成当前时间作为文件夹名称
    logger.info(cur_time)  # 使用 loguru 记录当前时间
    target_folder_new = os.path.join(target_folder, cur_time)  # 创建新的目标文件夹路径
    os.mkdir(target_folder_new)  # 创建带时间戳的新文件夹
    for file_name in data_list:
        source = os.path.join(original_folder, file_name)  # 文件的完整源路径
        destination = os.path.join(target_folder_new, file_name)  # 文件的目标路径
        shutil.move(source, destination)  # 使用 shutil.move 将文件移动到新位置

# need change
def running_program(window_name, original_folder, target_folder, cycle_number=-1, suffix_list=[]):
    """
    执行自动化程序：检测窗口状态，按循环次数或按键终止，并移动符合条件的文件。
    
    :param window_name: 应用窗口的名称
    :param original_folder: 源文件夹路径
    :param target_folder: 目标文件夹路径
    :param cycle_number: 循环次数，-1 表示无限循环 直到手动终止
    :param suffix_list: 文件后缀列表，只移动符合这些后缀的文件
    """
    exit_flag = False

    def on_key_event(event):
        # 当按键被按下时触发事件，如果按下 'q' 键，则终止循环
        nonlocal exit_flag
        if event.name == 'q':
            logger.info(f"END | terminated by user, windoe name: {window_name}") # 被用户终止
            exit_flag = True

    keyboard.on_press(on_key_event)  # 设置按键监听

    app = WinGUI(window_name)  # 创建 WinGUI 实例，用于窗口操作
    logger.info(f'START | window name: {window_name}')  # 记录窗口名称
    
    cycle_count = 0 # 初始化循环计数器
    while not exit_flag:
        try:
            if is_test_over(app): # 检测是否结束
                move_files(original_folder, target_folder, suffix_list)  # 移动文件
                logger.info(f"Cycle {cycle_count} is finished")  # 记录当前循环结束
                if cycle_number > 0 and cycle_count >= cycle_number:  # 检查是否达到设定的循环次数
                    logger.info(f"finished {cycle_count} cycles!")  # 记录完成循环次数
                    return
                cycle_count += 1  # 循环计数加一
                # 此处可以调用 app 的其他函数执行操作
                # write your operations
                # usually use app functions

                # operation end
        except Exception as err:
            logger.info(err)  # 记录异常信息

        # 处理异常情况
        # process abnormal cases
        try:
            print("test")
        except:
            print()
        
        time.sleep(0.5)   # 每次循环暂停1秒
