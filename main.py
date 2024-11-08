from WinGUI import WinGUI
import organize_table as tb
import automation as au
import utils as ut

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
import pyperclip       # 处理剪贴板内容

from PIL import ImageGrab     # 从PIL库导入ImageGrab模块，用于截图
from loguru import logger     # 引入loguru库，用于简便的日志记录

if __name__ == "__main__":

    # ---------- 配置参数 -------------
    pyautogui.FAILSAFE = False  # 关闭 pyautogui 的故障保护机制
    pyautogui.PAUSE = 0.1  # 设置 pyautogui 的默认操作延时 如移动和点击的间隔
    logger.add("dev.log", rotation="10 MB")  # 设置日志文件轮换
    # original_folder = "C:/Users/Public/Documents/Data"  # 源文件夹路径
    # target_folder = r"C:\Users\Joey\Desktop\data"  # 目标文件夹路径
    # suffix_list = []  # 要移动的文件后缀列表
    # cycle_number = -1  # 循环次数，-1 表示无限循环
    window_name = r"千牛接待台"  # 窗口名称
    
    # ---------- 快捷键启动 -------------
    '''
        按下 esc 退出
        按下 shift+ctrl+1 添加备注
    '''
    # auto_key()
    # hotkey_actions = [
    #     {'key': 'ctrl+r', 'func': ut.open_sof, 'args': ['旺店通ERP',394853,1,'CoolWindow']}
    # ]
    # ut.auto_key_with_threads(hotkey_actions)

    # ---------- 测试 -------------
    # au.run_test(window_name)
    # ut.open_sof('ToDesk')

    # ---------- 通知补发单号 -------------
    '''
        表格必须经过格式化，有整理过后的原始单号以及物流单号

        change 
            0 自动切换店铺

            1 添加功能 在直径结束后 显示当前表格有多少个需要通知补发 已经通知了多少个 多少个没有通知补发

            2 检查功能 is_find_cus 的未搜到处理办法 提示并跳过 有没有问题没有问题则下一个问题

            3 修改功能 能否修改监听停止按钮 但用一个q容易不起作用 改为使用组合键的 ctrl+shift+q监听停止

            4 修改功能 
            是否可以保证在按下停止的案件后 如果程序执行了发送通知 则必定在回写表格后才终止
            避免在已经通知但是暂未回写表格这个中间状态程序终止导致表格未更新 下次启动会重复通知

    '''
    au.notification_reissue(window_name, '2024-11-08_140753_余猫_补发单号 - test.xlsx')

    # ---------- 表格处理 -------------
    # 首次处理 群中表格 筛选制定店铺名称的物流单号 一个表有两个sheet 结果文件拿去ERP搜索后导出为新的表格
    # tb.process_original_table('11月6日补发单号.csv')

    # 二次处理 ERP导出表格 清洗新表格的原始单号 结果文件执行自动化操作
    # 有多少个店铺 就调用多少次 
    # tb.process_original_number('2024-11-07_余猫_ERP二次导出表格 - test.csv','余猫') # 日期_店铺_ERP二次导出表格
    # tb.process_original_number('2024-11-07_潮洁_ERP二次导出表格.csv','潮洁') # 日期_店铺_ERP二次导出表格
