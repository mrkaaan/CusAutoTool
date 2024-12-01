import argparse
import pyautogui
from loguru import logger
import utils as ut

def setup_arguments():
    parser = argparse.ArgumentParser(description="Main script and testing.")
    parser.add_argument('--test', action='store_true', help='Run tests')
    return parser.parse_args()

def setup_pyautogui():
    pyautogui.FAILSAFE = False  # 关闭 pyautogui 的故障保护机制
    pyautogui.PAUSE = 0.1  # 设置 pyautogui 的默认操作延时 如移动和点击的间隔

def setup_logging():
    logger.add("dev.log", rotation="10 MB")  # 设置日志文件轮换

def setup_bat_path():
    """
    从配置文件中读取 .bat 文件的路径，并返回绝对路径。

    :return: .bat 文件的绝对路径
    """
    try:
        bat_file_path = ut.get_bat_path()
        return bat_file_path
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        return None
    
def setup_hot_file_name():
    """
    从配置文件中读取热键文件名，并返回文件名。

    :return: 热键文件名
    """
    try:
        hot_file_name = ut.get_config_option('names', 'HOT_FILE_NAME')
        return hot_file_name
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        return None