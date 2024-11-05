# screen_position.py
import pyautogui

def calculate_screen_percentage():
    # 获取屏幕大小
    screen_width, screen_height = pyautogui.size()
    print(f"你的屏幕大小为 {screen_width} * {screen_height}")

    # 获取用户输入的坐标
    user_input = input("请输入图像位置坐标(空格分隔):")
    try:
        x, y = map(int, user_input.split(' '))
    except ValueError:
        print("输入格式错误，请确保输入的是两个整数，用逗号分隔。")
        return

    # 计算百分比位置
    x_percentage = (x / screen_width)
    y_percentage = (y / screen_height)

    # 输出结果
    print(f"您的坐标位于 x轴 {x_percentage:.2f}的位置，y轴 {y_percentage:.2f}的位置")

if __name__ == "__main__":
    calculate_screen_percentage()