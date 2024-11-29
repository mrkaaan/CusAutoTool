import os
import json
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk, messagebox
import threading
from datetime import datetime
import pyperclip
import shutil
import re
import openpyxl

from organize_table import process_table

table_path = ''
table_name = ''
root = None
root_initialized = False  # 标志 root 是否已经初始化

DEFAULT_VALUES = {
    "defaults": {
        "window_name": "千牛工作台",
        "table_name": "xxx_处理结果.xlsx",
        "table_path": "",
        "notic_shop_name": "",
        "notic_mode": 2,
        "show_logistics": True,
        "logistics_mode": 1,
        "use_today": "",
        "test_mode": "0",
        "is_write": True
    },
    "last_used": {}
}

CONFIG_FILE = 'notic_config.json'

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        # 如果文件不存在，则创建一个默认的配置
        save_config(DEFAULT_VALUES)  # 保存默认配置
        return DEFAULT_VALUES

def get_defaults(config):
    return config.get("defaults", {})

def get_last_used(config):
    return config.get("last_used", {})

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

# 桌面的表格
def choose_table(entry_var, open_desktop=True):
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
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
    )

    # 如果文件存在
    if file_path_temp:
        table_path = file_path_temp
        table_name = os.path.basename(file_path_temp)

        # 更新 StringVar 的内容
        entry_var.set(table_name)

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
    global window, window_initialized
    if window is not None and window.winfo_exists():
        window.lift()  # 如果窗口已经存在，将其提升到最前面
        return
        
    window = tk.Tk()
    window.title("通知补发信息填写")

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
            window.after(0, lambda: handle_result(result))
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
        order_numbers_label = tk.Label(window, text="处理后数据: ", font=f'{label_font_fam} {label_font_size-5} bold', fg=label_color)
        order_numbers_label.place(x=create_button_up_x_offset, y=create_button_up_y_offset-20)
        dynamic_buttons.append(order_numbers_label)

        # 创建复制文件名按钮
        copy_filename_button = tk.Label(window, text="文件名", font=f'{button_font_fam} {button_font_size} bold', fg=button_color, cursor='hand2', compound='center')
        copy_filename_button.bind("<Button-1>", lambda event: pyperclip.copy(output_filename))
        copy_filename_button.place(x=create_button_up_x_offset, y=create_button_up_y_offset)
        dynamic_buttons.append(copy_filename_button)

        # 创建复制全部订单编号按钮
        copy_all_orders_button = tk.Label(window, text="全部订单编号", font=f'{button_font_fam} {button_font_size} bold', fg=button_color, cursor='hand2', compound='center')
        copy_all_orders_button.bind("<Button-1>", lambda event: pyperclip.copy('\n'.join(all_order_numbers)))
        copy_all_orders_button.place(x=create_button_up_x_offset+create_button_up_y_interval, y=create_button_up_y_offset)
        dynamic_buttons.append(copy_all_orders_button)

        # 创建每个店铺的按钮
        button_x_offset = create_button_down_x_offset
        button_y_offset = create_button_down_y_offset
        button_y_interval = create_button_down_y_interval
        for shop, orders in shop_order_numbers.items():
            shop_name = simplify_shop_name(shop)
            shop_button = tk.Label(window, text=shop_name, font=f'{button_font_fam} {button_font_size-1} bold', fg=button_color, cursor='hand2', compound='center')
            shop_button.bind("<Button-1>", lambda event, orders=orders: pyperclip.copy('\n'.join(orders)))
            # command=lambda o='\n'.join(orders): pyperclip.copy(o)
            shop_button.place(x=button_x_offset, y=button_y_offset)
            dynamic_buttons.append(shop_button)
            button_x_offset += button_y_interval

    # 设置窗口居中
    scnwidth, scnheight = window.maxsize()
    tmp = '%dx%d+%d+%d' % (
        window_width, window_height, (scnwidth - window_width) / 2, (scnheight - window_height) / 2)  # 参数为（‘窗口宽 x 窗口高 + 窗口位于x轴位置 + 窗口位于y轴位置）
    window.geometry(tmp)

    # 禁止更改窗口大小
    window.resizable(False, False)  # 禁止更改大小
    window.update()  # 强制刷新

    # 原始表格处理部分
    # 标题
    or_label = tk.Label(window, text="原始表格处理", compound='center', font=f'{label_font_fam} {label_font_size} bold', fg=label_color)
    or_label.place(x=10, y=10)

    # 输入框
    or_entry_var = tk.StringVar()
    or_entry = tk.Entry(window, textvariable=or_entry_var, width=entry_width,font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    or_entry.place(x=10, y=40)

    # 选择文件按钮 桌面目录
    or_button_choose = tk.Button(window, cursor="hand2", text="选择桌面文件", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda:choose_table(or_entry))
    or_button_choose.place(x=button_x_offset+0*button_x_interval, y=button_y_offset)
    # 选择文件按钮 当前项目目录
    or_button_choose = tk.Button(window, cursor="hand2", text="选择项目文件", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda:choose_table(or_entry, False))
    or_button_choose.place(x=button_x_offset+1*button_x_interval, y=button_y_offset)
    # 开始处理按钮
    or_button_sure = tk.Button(window, state="disabled", text="开始处理", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda:run_in_thread(process_table, table_name))
    or_button_sure.place(x=button_x_offset+2*button_x_interval, y=button_y_offset)
    # 清空输入框按钮
    or_button_clear = tk.Button(window, cursor="hand2", text="清空", font=f'{button_font_fam} {button_font_size} bold',fg=button_color, command=lambda: clear_dynamic_buttons())
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
        label = tk.Label(window, text=short_name, cursor='hand2', compound='center',
                        font=f'{label_copy_font_fam} {label_copy_font_size-5} bold', fg=label_copy_color)
        label.place(x=label_copy_x_offset + index * label_copy_x_interval, y=label_copy_y_offset)  # 根据索引设置x坐标
        label.bind("<Button-1>", lambda event, i=index: pyperclip.copy(full_name))

    def open_form_folder():
        os.startfile(os.path.join(os.getcwd(), "form"))

    # 右下角打开项目根目录下form文件夹按钮
    open_form_folder_button = tk.Label(window, text="form", cursor='hand2', compound='center', font=f'{button_font_fam} {button_font_size} bold', fg=button_color)
    open_form_folder_button.bind("<Button-1>", lambda event: open_form_folder())
    open_form_folder_button.place(x=form_label_x_offset, y=form_label_y_offset)

    window.protocol("WM_DELETE_WINDOW", on_close)  # 设置关闭窗口时的回调函数
    window.mainloop()


# 假设这是你要调用的函数 a
def function_a(**kwargs):
    print("Function a called with parameters:")
    for key, value in kwargs.items():
        print(f"{key}: {value}")
    # 在这里添加你的实际处理逻辑

# 创建窗口并收集参数
def create_parameter_window(mode=0):
    '''
        :param mode: 0 当前文件启动 1 被其他程序调用启动 由于在utils中调用是非主线程，会导致窗口大小不同，设置了不同的窗口大小和位置参数
    '''
    print(f"create_window mode: {mode} {'主程序模式' if mode == 0 else '被调用模式'}")
    global root, root_initialized
    if root is not None and root.winfo_exists():
        root.lift()  # 如果窗口已经存在，将其提升到最前面
        return
    # 读取配置文件
    notic_config = load_config()

    # 获取全局表格变量
    global table_path
    global table_name

    # 修改 on_ok 函数以保存上次使用的值
    def on_ok():
        # 收集所有参数
        notification_reissue_parameter = {
            'window_name': window_name_var.get(),
            'table_name': table_name,
            'table_path': table_path,
            'notic_shop_name': notic_shop_name_var.get(),
            'notic_mode': notic_mode_var.get(),
            'show_logistics': show_logistics_var.get(),
            'logistics_mode': logistics_mode_var.get(),
            'use_today': use_today_var.get() or datetime.now().strftime('%Y-%m-%d'),
            'test_mode': int(test_mode_var.get()) if test_mode_var.get().isdigit() else 0,
            'is_write': is_write_var.get(),
        }

        # 检查必填项是否为空
        # 检查表格是否为默认名称 xxx_处理结果.xlsx
        if notification_reissue_parameter['table_name'] == 'xxx_处理结果.xlsx':
            messagebox.showwarning("警告", "请选择处理结果表格")
            return
        if not notification_reissue_parameter['window_name']:
            messagebox.showwarning("警告", "窗口名称不能为空")
            return
        if not notification_reissue_parameter['table_name']:
            messagebox.showwarning("警告", "表格名称不能为空")
            return
        if not notification_reissue_parameter['notic_shop_name']:
            messagebox.showwarning("警告", "通知店铺名称不能为空")
            return
        # 判断日期格式是否为 yyyy-mm-dd
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', notification_reissue_parameter['use_today']):
            messagebox.showwarning("警告", "日期格式错误，请使用 yyyy-mm-dd 格式")
            return
        # 判断文件路径table_path是否存在
        if not os.path.exists(table_path):
            messagebox.showwarning("警告", "表格文件不存在")
            return

        # 保存上次使用的值
        save_config({"defaults":DEFAULT_VALUES.get("defaults"),"last_used": notification_reissue_parameter})

        # 关闭窗口
        root.destroy()

        # 调用函数 a
        function_a(**notification_reissue_parameter)


    def set_today_date():
        use_today_var.set(datetime.now().strftime('%Y-%m-%d'))
    
    # 使用默认值
    def use_defaults():
        global table_path
        global table_name
        # 获取全局表格变量
        defaults = get_defaults(notic_config)
        window_name_var.set(defaults['window_name'])
        table_name_var.set(defaults['table_name'])
        table_name = ''
        table_path = ''
        notic_shop_name_var.set(defaults['notic_shop_name'])
        notic_mode_var.set(defaults['notic_mode'])
        show_logistics_var.set(defaults['show_logistics'])
        logistics_mode_var.set(defaults['logistics_mode'])
        use_today_var.set(defaults['use_today'] or datetime.now().strftime('%Y-%m-%d'))
        test_mode_var.set(defaults['test_mode'])
        is_write_var.set(defaults['is_write'])

    # 使用上次内容
    def use_last_used():
        global table_path
        global table_path
        last_used = get_last_used(notic_config)
        if last_used:
            window_name_var.set(last_used['window_name'])
            table_name = last_used['table_name'].split('/')[-1]
            table_name_var.set(table_name)
            table_path = last_used['table_name']
            notic_shop_name_var.set(last_used['notic_shop_name'])
            notic_mode_var.set(last_used['notic_mode'])
            show_logistics_var.set(last_used['show_logistics'])
            logistics_mode_var.set(last_used['logistics_mode'])
            use_today_var.set(last_used['use_today'] or datetime.now().strftime('%Y-%m-%d'))
            test_mode_var.set(last_used['test_mode'])
            is_write_var.set(last_used['is_write'])

    # 识别店铺名称
    def identify_shop_name(notic_shop_name_var, notic_shop_name_combobox):
        global table_path
        global table_name
        print(f"identify_shop_name notic_shop_name_var: {table_path}")
        if not os.path.exists(table_path):
            messagebox.showwarning("警告", "表格文件不存在")
            return

        try:
            workbook = openpyxl.load_workbook(table_path)
            sheet_names = workbook.sheetnames

            if len(sheet_names) == 1 and sheet_names[0].lower() == 'sheet1':
                messagebox.showwarning("警告", "表格只有一个 sheet 且名称为 sheet1，请处理表格后再来通知")
                return

            shop_names = []
            is_correct_table = False
            for sheet_name in sheet_names:
                simplified_name = simplify_shop_name(sheet_name)
                if simplified_name not in shop_names:
                    shop_names.append(simplified_name)
                # 通过双重判断 是否含有 '潮洁' '余猫' 两个关键词 判断是否打开正确表格
                if '潮洁' in sheet_name.lower() or '余猫' in sheet_name.lower():
                    is_correct_table = True

            if not is_correct_table:
                messagebox.showwarning("警告", "表格不正确，请选择正确的表格")
                return
            
            if shop_names:
                # 去除shop_names中的‘全部店铺’
                if '全部店铺' in shop_names:
                    shop_names.remove('全部店铺')
                notic_shop_name_var.set(shop_names[0])  # 默认选择第一个识别到的店铺名称
                notic_shop_name_combobox['values'] = shop_names  # 更新下拉菜单选项
            else:
                notic_shop_name_var.set('自行输入或识别表格获取')
                notic_shop_name_combobox['values'] = ['潮洁居家', '余猫日用', '自行输入或识别表格获取']
        except Exception as e:
            messagebox.showerror("错误", f"读取表格时发生错误: {e}")

    # 创建主窗口
    root = tk.Tk()
    root.title("通知补发参数设置")

    # 设置窗口大小
    window_width = 400 if mode == 0 else 500
    window_height = 350 if mode == 0 else 400

    # label字体设置
    label_font_fam = '楷体'
    label_font_size = 11
    label_color = '#000000'
    label_weight = 'bold'
    label_x_offset = 10
    label_y_offset = 10
    label_y_interval = 30
    # entry设置
    entry_font_fam = '楷体'
    entry_font_size = 11
    entry_color = '#000000'
    entry_width = 22
    entry_x_offset = 100
    entry_x_interval = 100  if mode == 0 else 125  # entry x轴位置间隔设置
    entry_y_offset = 10
    entry_y_interval = 30
    # label提示文字设置 灰色提示文字设置
    label_notice_font_fam = '楷体'
    label_notice_font_size = 9
    label_notice_color = '#808080'
    label_notic_weight = 'normal'
    # button字体设置
    button_font_fam = '楷体'
    button_font_size = 10
    button_color = '#000000'
    button_x_offset = 140
    button_y_offset = 300
    # 功能按钮设置
    set_button_x_offset = window_width - 150
    set_button_y_offset = window_height - 40
    set_button_x_interval = 30  if mode == 0 else 125  # button x轴位置间隔设置
    
    # 设置窗口居中
    scnwidth, scnheight = root.maxsize()
    tmp = '%dx%d+%d+%d' % (
        window_width, window_height, (scnwidth - window_width) / 2, (scnheight - window_height) / 2)  # 参数为（‘窗口宽 x 窗口高 + 窗口位于x轴位置 + 窗口位于y轴位置）
    root.geometry(tmp)

    # 禁止更改窗口大小
    root.resizable(False, False)  # 禁止更改大小
    root.update()  #

    # 窗口名
    window_name_label = tk.Label(root, text="窗口名称:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    window_name_label.place(x=label_x_offset, y=label_y_offset)
    window_name_var = tk.StringVar(value='千牛工作台')
    window_name_entry = tk.Entry(root, textvariable=window_name_var, width=entry_width, font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    window_name_entry.place(x=entry_x_offset, y=entry_y_offset)

    # 表格名
    table_name_label = tk.Label(root, text="表格名称:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    table_name_label.place(x=label_x_offset, y=label_y_offset+1*label_y_interval)
    table_name_var = tk.StringVar(value='xxx_处理结果.xlsx')
    table_name_entry = tk.Entry(root, textvariable=table_name_var, width=entry_width, font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    table_name_entry.place(x=entry_x_offset, y=entry_y_offset+1*entry_y_interval)
    # 选择文件按钮 桌面目录
    or_button_choose_desktop = tk.Label(root, cursor="hand2", text="选择桌面", font=f'{button_font_fam} {button_font_size-2} bold', fg=button_color)
    or_button_choose_desktop.bind("<Button-1>", lambda event:choose_table(table_name_var))
    or_button_choose_desktop.place(x=entry_x_offset+int(1.8*entry_x_interval), y=entry_y_offset+1*entry_y_interval)
    # 选择文件按钮 当前项目目录
    or_button_choose_project = tk.Label(root, cursor="hand2", text="选择项目", font=f'{button_font_fam} {button_font_size-2} bold', fg=button_color)
    or_button_choose_project.bind("<Button-1>", lambda event:choose_table(table_name_var,False))
    or_button_choose_project.place(x=entry_x_offset+int(2.4*entry_x_interval), y=entry_y_offset+1*entry_y_interval)

    # 通知店铺名
    notic_shop_name_label = tk.Label(root, text="店铺名称:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    notic_shop_name_label.place(x=label_x_offset, y=label_y_offset+2*label_y_interval)
    notic_shop_name_var = tk.StringVar()
    notic_shop_name_combobox = ttk.Combobox(root, textvariable=notic_shop_name_var, values=['潮洁居家', '余猫日用', '自行输入或识别表格获取'], width=entry_width)
    notic_shop_name_combobox.place(x=entry_x_offset, y=entry_y_offset+2*entry_y_interval)
    identify_shop_name_button = tk.Label(root, text="识别店铺名称", cursor='hand2', compound='center', font=f'{entry_font_fam} {entry_font_size-2} bold', fg=entry_color)
    identify_shop_name_button.bind("<Button-1>", lambda event: identify_shop_name(notic_shop_name_var, notic_shop_name_combobox))
    identify_shop_name_button.place(x=entry_x_offset+int(1.8*entry_x_interval), y=entry_y_offset+2*entry_y_interval)

    # 通知模式
    notic_mode_label = tk.Label(root, text="通知模式:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    notic_mode_label.place(x=label_x_offset, y=label_y_offset+3*label_y_interval)
    notic_mode_var = tk.IntVar(value=2)
    notic_mode_radio1 = tk.Radiobutton(root, text="输入通知", variable=notic_mode_var, value=1)
    notic_mode_radio1.place(x=entry_x_offset, y=entry_y_offset+3*entry_y_interval)
    notic_mode_radio2 = tk.Radiobutton(root, text="窗口通知", variable=notic_mode_var, value=2)
    notic_mode_radio2.place(x=entry_x_offset+entry_x_interval, y=entry_y_offset+3*entry_y_interval)

    # 是否显示物流公司
    show_logistics_label = tk.Label(root, text="显示物流:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    show_logistics_label.place(x=label_x_offset, y=label_y_offset+4*label_y_interval)
    show_logistics_var = tk.BooleanVar(value=True)
    show_logistics_checkbox = tk.Checkbutton(root, variable=show_logistics_var)
    show_logistics_checkbox.place(x=entry_x_offset, y=entry_y_offset+4*entry_y_interval)
    show_logistics_note_label = tk.Label(root, text="仅输入通知模式生效", font=f'{label_notice_font_fam} {label_notice_font_size} {label_notic_weight}', fg=label_notice_color)
    show_logistics_note_label.place(x=entry_x_offset+60, y=entry_y_offset+4*entry_y_interval)

    # 物流模式
    logistics_mode_label = tk.Label(root, text="物流模式:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    logistics_mode_label.place(x=label_x_offset, y=label_y_offset+5*label_y_interval)
    logistics_mode_var = tk.IntVar(value=1)
    logistics_mode_radio1 = tk.Radiobutton(root, text="自动识别", variable=logistics_mode_var, value=1)
    logistics_mode_radio1.place(x=entry_x_offset, y=entry_y_offset+5*entry_y_interval)
    logistics_mode_radio2 = tk.Radiobutton(root, text="手动输入", variable=logistics_mode_var, value=2)
    logistics_mode_radio2.place(x=entry_x_offset+entry_x_interval, y=entry_y_offset+5*entry_y_interval)

    # 目录日期 指的是如果所选文件在项目目录且非今天日期文件夹下，需要修改这个参数，如果是在桌面目录下，则不需要保持默认值或为空或今天日期，程序会自动复制一份到项目目录的今日日期文件夹下
    use_today_label = tk.Label(root, text="目录日期:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    use_today_label.place(x=label_x_offset, y=label_y_offset+6*label_y_interval)
    use_today_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
    use_today_entry = tk.Entry(root, textvariable=use_today_var, width=int(0.7*entry_width), font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    use_today_entry.place(x=entry_x_offset, y=entry_y_offset+6*entry_y_interval)
    use_today_today_button = tk.Label(root, text="当天日期", cursor='hand2', compound='center', font=f'{entry_font_fam} {entry_font_size-2} bold', fg=entry_color)
    use_today_today_button.bind("<Button-1>", lambda event: set_today_date())
    use_today_today_button.place(x=entry_x_offset+int(1.3*entry_x_interval), y=entry_y_offset+6*entry_y_interval)
    use_today_note_label = tk.Label(root, text="为空默认当天", font=f'{label_notice_font_fam} {label_notice_font_size} {label_notic_weight}', fg=label_notice_color)
    use_today_note_label.place(x=entry_x_offset+int(1.6*entry_x_interval), y=entry_y_offset+6*entry_y_interval)

    # 测试模式
    test_mode_label = tk.Label(root, text="测试模式:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    test_mode_label.place(x=label_x_offset, y=label_y_offset+7*label_y_interval)
    test_mode_var = tk.StringVar(value='0')
    test_mode_entry = tk.Entry(root, textvariable=test_mode_var, width=int(0.7*entry_width), font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    test_mode_entry.place(x=entry_x_offset, y=label_y_offset+7*label_y_interval)
    test_mode_note_label = tk.Label(root, text="0为不测试", font=f'{label_notice_font_fam} {label_notice_font_size} {label_notic_weight}', fg=label_notice_color)
    test_mode_note_label.place(x=entry_x_offset+int(1.3*entry_x_interval), y=label_y_offset+7*label_y_interval)

    # 是否回写表格
    is_write_label = tk.Label(root, text="是否回写:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    is_write_label.place(x=label_x_offset, y=label_y_offset+8*label_y_interval)
    is_write_var = tk.BooleanVar(value=True)
    is_write_checkbox = tk.Checkbutton(root, variable=is_write_var)
    is_write_checkbox.place(x=entry_x_offset, y=label_y_offset+8*label_y_interval)

    # OK 按钮
    ok_button = tk.Button(root, text="通知补发", command=on_ok)
    ok_button.place(x=button_x_offset, y=button_y_offset)

    # 使用默认值按钮
    use_defaults_button = tk.Label(root, text="默认值", cursor='hand2', compound='center', font=f'{button_font_fam} {button_font_size-1} bold', fg=button_color)
    use_defaults_button.bind("<Button-1>", lambda event: use_defaults())
    use_defaults_button.place(x=set_button_x_offset, y=set_button_y_offset)

    # 使用上次内容按钮
    use_last_used_button = tk.Label(root, text="上次值", cursor='hand2', compound='center', font=f'{button_font_fam} {button_font_size-1} bold', fg=button_color)
    use_last_used_button.bind("<Button-1>", lambda event: use_last_used())
    use_last_used_button.place(x=set_button_x_offset+2*set_button_x_interval, y=set_button_y_offset)

    # 运行主循环
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

# 测试
if __name__ == "__main__":
    create_parameter_window()