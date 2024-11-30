import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import threading
from datetime import datetime
import pyperclip
import shutil

from organize_table import process_table

table_path = ''
table_name = ''
root = None
root_initialized = False  # 标志 root 是否已经初始化

# 桌面的表格
def choose_table(enter, open_desktop=True):
    '''
        :param enter: 输入框
        :param open_desktop: 是否打开桌面目录 默认为True 若为False 则打开当前项目根目录form文件
    '''
    global table_path
    global table_name

    # 设置默认打开目录为桌面
    default_path = os.path.join(os.path.expanduser("~"), "Desktop")

    # 打开目录选择文件
    if not open_desktop:
        default_path = f"./form"
    file_path_temp = filedialog.askopenfilename(
        title="选择文件",
        initialdir=default_path,  # 默认打开桌面
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    # 如果文件存在
    if file_path_temp:
        # 文本框追加新的文本
        # current_text = or_entry.get()  # 获取当前的文本
        enter.delete(0, tk.END)  # 清除输入框中的所有文本
        table_path = file_path_temp
        table_name = os.path.basename(file_path_temp)
        enter.insert(0, table_path)  # 插入新的组合文本

        # 获取当前日期的年月日
        today = datetime.now().strftime('%Y-%m-%d')
        target_dir = f"./form/{today}"

        # 在项目目录下./from下创建当天日期文件夹
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)

         # 组合目标文件路径
        target_file_path = os.path.join(target_dir, table_name)

        if os.path.exists(target_file_path):
            print(f"文件已存在当然日期文件夹中：{target_file_path}")
        else:
            # 复制文件到目标路径
            try:
                shutil.copy(file_path_temp, target_file_path)
                print(f"文件已成功复制到 {target_file_path}")
            except Exception as e:
                print(f"文件复制失败：{e}")


# 检测Entry内容变化
def update_button_state(entry_var, button, *args):
    if entry_var.get() == "":
        button.config(state='disabled', cursor="")
    else:
        button.config(state='normal', cursor="hand2")

# 化简店铺名
# 店铺：潮洁居家日用旗舰店-天猫 -> 潮洁
# 店铺：余猫旗舰店-天猫 -> 余猫
# 店铺：团洁3504猫宁-天猫 -> 猫宁3504
# 店铺：团洁旗舰店-天猫 -> 团洁
# 店铺：潮洁873猫宁-天猫 -> 猫宁873
def simplify_shop_name(shop_name):
    if "团洁旗舰" in shop_name:
        return "团洁旗舰"
    elif "潮洁居家" in shop_name:
        return "潮洁居家"
    elif "余猫" in shop_name:
        return "余猫旗舰"
    elif "3504" in shop_name:
        return "猫宁3504"
    elif "873" in shop_name:
        return "猫宁873"
    else:
        return shop_name


def create_window(mode=0):    # 主窗口
    '''
        :param mode: 0 当前文件启动 1 被其他程序调用启动 由于在utils中调用是非主线程，会导致窗口大小不同，设置了不同的窗口大小和位置参数
    '''
    print(f"create_window mode: {mode} {'主程序模式' if mode == 0 else '被调用模式'}")
    global root, root_initialized
    if root is not None and root.winfo_exists():
        root.lift()  # 如果窗口已经存在，将其提升到最前面
        return
        
    root = tk.Tk()
    root.title("补发表格处理")

    global table_path
    global table_name

    # 参数设置
    # 窗口大小
    window_width = 350 if mode == 0 else 450
    window_height = 220 if mode == 0 else 270
    # label字体设置
    label_font_fam = '楷体'
    label_font_size = 15
    label_color = '#000000'
    # button字体设置
    button_font_fam = '楷体'
    button_font_size = 8
    button_color = '#000000'
    button_x_offset = 26
    button_x_interval = 100  if mode == 0 else 125  # button x轴位置间隔设置
    button_y_offset = 82
    # entry设置
    entry_font_fam = '楷体'
    entry_font_size = 11
    entry_color = '#000000'
    entry_width = 40 if mode == 0 else 38
    # create_button 第一排设置
    create_button_up_x_offset = 26
    create_button_up_y_offset = 134
    create_button_up_y_interval = 70 if mode == 0 else 80  # button y轴位置间隔设置
    # create_button 第二排设置
    create_button_down_x_offset = 26
    create_button_down_y_offset = 155
    create_button_down_y_interval = 58 if mode == 0 else 68  # button y轴位置间隔设置
    # 快捷复制表名按钮设置
    label_copy_font_fam = "Arial"  # 设置字体类型
    label_copy_font_size = 13       # 设置字体大小
    label_copy_color = "black"      # 设置字体颜色
    label_copy_x_offset = 10        # 设置x坐标
    label_copy_x_interval = 40        # 设置y坐标
    label_copy_y_offset = 190 if mode == 0 else 220   # 设置y坐标
    # form label设置
    form_label_x_offset = 300 if mode == 0 else 385
    form_label_y_offset = 190 if mode == 0 else 220



    # 线程化的函数，避免界面冻结
    # def run_in_thread(func, *args):
    #     threading.Thread(target=func, args=args).start()

    # 线程化的函数
    def run_in_thread(func, *args):
        def wrapper():
            result = func(*args)
            root.after(0, lambda: handle_result(result))
        thread = threading.Thread(target=wrapper)
        thread.start()

    # 处理结果的回调函数
    def handle_result(result):
        output_filename, all_order_numbers, shop_order_numbers = result
        create_buttons(output_filename, all_order_numbers, shop_order_numbers)

    # 动态按钮列表
    dynamic_buttons = []

    # 定义清空动态按钮的函数
    def clear_dynamic_buttons():
        for widget in dynamic_buttons:
            widget.destroy()
        dynamic_buttons.clear()
        or_entry.delete(0, tk.END)

    # 创建动态按钮的函数
    def create_buttons(output_filename, all_order_numbers, shop_order_numbers):
        # 删除之前的动态按钮
        for widget in dynamic_buttons:
            widget.destroy()
        dynamic_buttons.clear()

        # 订单编号label
        order_numbers_label = tk.Label(root, text="处理后数据: ", font=f'{label_font_fam} {label_font_size-5} bold', fg=label_color)
        order_numbers_label.place(x=create_button_up_x_offset, y=create_button_up_y_offset-20)
        dynamic_buttons.append(order_numbers_label)

        # 创建复制文件名按钮
        copy_filename_button = tk.Label(root, text="文件名", font=f'{button_font_fam} {button_font_size} bold', fg=button_color, cursor='hand2', compound='center')
        copy_filename_button.bind("<Button-1>", lambda event: pyperclip.copy(output_filename))
        copy_filename_button.place(x=create_button_up_x_offset, y=create_button_up_y_offset)
        dynamic_buttons.append(copy_filename_button)

        # 创建复制全部订单编号按钮
        copy_all_orders_button = tk.Label(root, text="全部订单编号", font=f'{button_font_fam} {button_font_size} bold', fg=button_color, cursor='hand2', compound='center')
        copy_all_orders_button.bind("<Button-1>", lambda event: pyperclip.copy('\n'.join(all_order_numbers)))
        copy_all_orders_button.place(x=create_button_up_x_offset+create_button_up_y_interval, y=create_button_up_y_offset)
        dynamic_buttons.append(copy_all_orders_button)

        # 创建每个店铺的按钮
        button_x_offset = create_button_down_x_offset
        button_y_offset = create_button_down_y_offset
        button_y_interval = create_button_down_y_interval
        for shop, orders in shop_order_numbers.items():
            shop_name = simplify_shop_name(shop)
            shop_button = tk.Label(root, text=shop_name, font=f'{button_font_fam} {button_font_size-1} bold', fg=button_color, cursor='hand2', compound='center')
            shop_button.bind("<Button-1>", lambda event, orders=orders: pyperclip.copy('\n'.join(orders)))
            # command=lambda o='\n'.join(orders): pyperclip.copy(o)
            shop_button.place(x=button_x_offset, y=button_y_offset)
            dynamic_buttons.append(shop_button)
            button_x_offset += button_y_interval

    # 设置窗口位于屏幕的右上部分中间
    scnwidth, scnheight = root.maxsize()
    # 计算窗口的x坐标为屏幕宽度减去窗口宽度，y坐标为0（屏幕最上端）
    x_offset = scnwidth - window_width - 100
    y_offset = (scnheight - window_height) / 2 - 100
    tmp = '%dx%d+%d+%d' % (window_width, window_height, x_offset, y_offset)
    root.geometry(tmp)

    # 禁止更改窗口大小
    root.resizable(False, False)  # 禁止更改大小
    root.update()  # 强制刷新

    # 原始表格处理部分
    # 标题
    or_label = tk.Label(root, text="原始表格处理", compound='center', font=f'{label_font_fam} {label_font_size} bold', fg=label_color)
    or_label.place(x=10, y=10)

    # 输入框
    or_entry_var = tk.StringVar()
    or_entry = tk.Entry(root, textvariable=or_entry_var, width=entry_width,font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    or_entry.place(x=10, y=40)

    # 选择文件按钮 桌面目录
    or_button_choose = tk.Button(root, cursor="hand2", text="选择桌面文件", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda:choose_table(or_entry))
    or_button_choose.place(x=button_x_offset+0*button_x_interval, y=button_y_offset)
    # 选择文件按钮 当前项目目录
    or_button_choose = tk.Button(root, cursor="hand2", text="选择项目文件", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda:choose_table(or_entry, False))
    or_button_choose.place(x=button_x_offset+1*button_x_interval, y=button_y_offset)
    # 开始处理按钮
    or_button_sure = tk.Button(root, state="disabled", text="开始处理", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda:run_in_thread(process_table, table_name))
    or_button_sure.place(x=button_x_offset+2*button_x_interval, y=button_y_offset)
    # 清空输入框按钮
    or_button_clear = tk.Button(root, cursor="hand2", text="清空", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda: clear_dynamic_buttons())
    or_button_clear.place(x=button_x_offset+2.7*button_x_interval, y=button_y_offset)
    # 监控输入框的变化，并更新按钮状态
    or_entry_var.trace("w", lambda *args: update_button_state(or_entry, or_button_sure, *args))


    # 店铺名称字典 key为简称 value为完整店铺名 '潮洁居家日用旗舰店-天猫', '余猫旗舰店-天猫', '团洁3504猫宁-天猫', '团洁旗舰店-天猫', '潮洁873猫宁-天猫'
    shop_names = {
        "潮洁": "潮洁居家日用旗舰店-天猫",
        "余猫": "余猫旗舰店-天猫",
        "团洁": "团洁3504猫宁-天猫",
        "潮洁873": "潮洁873猫宁-天猫"
    }

    #使用pyperclip.copy复制shopa_names字典的value值
    for index, (short_name, full_name) in enumerate(shop_names.items()):
        label = tk.Label(root, text=short_name, cursor='hand2', compound='center',
                        font=f'{label_copy_font_fam} {label_copy_font_size-5} bold', fg=label_copy_color)
        label.place(x=label_copy_x_offset + index * label_copy_x_interval, y=label_copy_y_offset)  # 根据索引设置x坐标
        label.bind("<Button-1>", lambda event, i=index: pyperclip.copy(full_name))

    def open_form_folder():
        os.startfile(os.path.join(os.getcwd(), "form"))

    # 右下角打开项目根目录下form文件夹按钮
    open_form_folder_button = tk.Label(root, text="form", cursor='hand2', compound='center', font=f'{button_font_fam} {button_font_size} bold', fg=button_color)
    open_form_folder_button.bind("<Button-1>", lambda event: open_form_folder())
    open_form_folder_button.place(x=form_label_x_offset, y=form_label_y_offset)

    root.protocol("WM_DELETE_WINDOW", on_close)  # 设置关闭窗口时的回调函数
    root.mainloop()

def on_close():
    global root, root_initialized
    root.destroy()
    root = None
    root_initialized = False  # 窗口关闭后重置标志

# 定义一个回调函数，在主线程中调用create_window
def call_create_window():
    global root
    if root is not None and root.winfo_exists():
        root.lift()  # 如果窗口已经存在，将其提升到最前面
        root.focus_force()  # 强制窗口获得焦点
    else:
        create_window(mode=1)
        # root.after(0, create_window)  # 如果窗口不存在，创建新窗口

if __name__ == "__main__":
    # 测试
    create_window() # 创建窗口