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
from plyer import notification
import configparser

import automation as au


config = configparser.ConfigParser()
config.read('config.ini')
# 访问环境变量
window_open_mode = config['default'].get('WINDOW_OPEN_MODE')
# 尝试将 TRY_NUMBER 转换为整数
window_open_mode = int(window_open_mode)

table_path = ''
table_name = ''
window = None
window_initialized = False  # 标志 window 是否已经初始化

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

def show_toast(title, message, timeout=0.2):
    notification.notify(
        title=title,
        message=message,
        app_name="提醒",
        timeout=timeout,
        toast=True
    )

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


# 创建窗口并收集参数
def create_window(mode=0):
    '''
        :param mode: 0 当前文件启动 1 被其他程序调用启动 由于在utils中调用是非主线程，会导致窗口大小不同，设置了不同的窗口大小和位置参数
    '''
    print(f"create_window mode: {mode} {'主程序模式' if mode == 0 else '被调用模式'}")
    global window, window_initialized
    if window is not None and window.winfo_exists():
        window.lift()  # 如果窗口已经存在，将其提升到最前面
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
            print(f"文件不存在：{table_path}")
            messagebox.showwarning("警告", "表格文件不存在")
            return

        # 保存上次使用的值
        save_config({"defaults":DEFAULT_VALUES.get("defaults"),"last_used": notification_reissue_parameter})

        # 关闭窗口
        window.destroy()
        show_toast("提醒", "开始补发通知...")
        # 调用函数 a
        au.notification_reissue(**notification_reissue_parameter)


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
        global table_name
        global table_path
        last_used = get_last_used(notic_config)
        if last_used:
            window_name_var.set(last_used['window_name'])
            table_name = last_used['table_name'].split('/')[-1]
            table_name_var.set(table_name)
            table_path = last_used['table_path']
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
            print(f"文件不存在：{table_path}")
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
    window = tk.Tk()
    window.title("通知补发参数设置")

    # 设置窗口大小
    window_width = 400 if mode == 0 else 520
    window_height = 350 if mode == 0 else 350

    # label字体设置
    label_font_fam = '楷体'
    label_font_size = 11
    label_color = '#000000'
    label_weight = 'bold'
    label_x_offset = 10
    label_y_offset = 10
    label_y_interval = 30  if mode == 0 else 32  # label y轴位置间隔设置
    # entry设置
    entry_font_fam = '楷体'
    entry_font_size = 11
    entry_color = '#000000'
    entry_width = 22 if mode == 0 else 20  # entry 宽度设置
    entry_x_offset = 100 if mode == 0 else 125  # entry x轴位置间隔设置
    entry_x_interval = 100  if mode == 0 else 130  # entry x轴位置间隔设置
    entry_y_offset = 10
    entry_y_interval = 30  if mode == 0 else 30  # entry y轴位置间隔设置
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
    set_button_x_offset = window_width - 150 if mode == 0 else window_width - 146  # button x轴位置间隔设置
    set_button_y_offset = window_height - 40 if mode == 0 else window_height - 43  # button y轴位置间隔设置
    set_button_x_interval = 30  if mode == 0 else 32  # button x轴位置间隔设置
    
    # 设置窗口位于屏幕的右上部分中间
    scnwidth, scnheight = window.maxsize()
    # 计算窗口的x坐标为屏幕宽度减去窗口宽度，y坐标为0（屏幕最上端）
    x_offset = scnwidth - window_width - 100
    y_offset = (scnheight - window_height) / 2 - 100
    tmp = '%dx%d+%d+%d' % (window_width, window_height, x_offset, y_offset)
    window.geometry(tmp)

    # 禁止更改窗口大小
    window.resizable(False, False)  # 禁止更改大小
    window.update()  #

    # 窗口名
    window_name_label = tk.Label(window, text="窗口名称:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    window_name_label.place(x=label_x_offset, y=label_y_offset)
    window_name_var = tk.StringVar(value='千牛工作台')
    window_name_entry = tk.Entry(window, textvariable=window_name_var, width=entry_width, font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    window_name_entry.place(x=entry_x_offset, y=entry_y_offset)

    # 表格名
    table_name_label = tk.Label(window, text="表格名称:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    table_name_label.place(x=label_x_offset, y=label_y_offset+1*label_y_interval)
    table_name_var = tk.StringVar(value='xxx_处理结果.xlsx')
    table_name_entry = tk.Entry(window, textvariable=table_name_var, width=entry_width, font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    table_name_entry.place(x=entry_x_offset, y=entry_y_offset+1*entry_y_interval)
    # 选择文件按钮 桌面目录
    or_button_choose_desktop = tk.Label(window, cursor="hand2", text="选择桌面", font=f'{button_font_fam} {button_font_size-2} bold', fg=button_color)
    or_button_choose_desktop.bind("<Button-1>", lambda event:choose_table(table_name_var))
    or_button_choose_desktop.place(x=entry_x_offset+int(1.8*entry_x_interval), y=entry_y_offset+1*entry_y_interval)
    # 选择文件按钮 当前项目目录
    or_button_choose_project = tk.Label(window, cursor="hand2", text="选择项目", font=f'{button_font_fam} {button_font_size-2} bold', fg=button_color)
    or_button_choose_project.bind("<Button-1>", lambda event:choose_table(table_name_var,False))
    or_button_choose_project.place(x=entry_x_offset+int(2.4*entry_x_interval), y=entry_y_offset+1*entry_y_interval)

    # 通知店铺名
    notic_shop_name_label = tk.Label(window, text="店铺名称:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    notic_shop_name_label.place(x=label_x_offset, y=label_y_offset+2*label_y_interval)
    notic_shop_name_var = tk.StringVar()
    notic_shop_name_combobox = ttk.Combobox(window, textvariable=notic_shop_name_var, values=['潮洁居家', '余猫日用', '自行输入或识别表格获取'], width=(entry_width if mode == 0 else entry_width-2))
    notic_shop_name_combobox.place(x=entry_x_offset, y=entry_y_offset+2*entry_y_interval)
    identify_shop_name_button = tk.Label(window, text="识别店铺", cursor='hand2', compound='center', font=f'{entry_font_fam} {entry_font_size-2} bold', fg=entry_color)
    identify_shop_name_button.bind("<Button-1>", lambda event: identify_shop_name(notic_shop_name_var, notic_shop_name_combobox))
    identify_shop_name_button.place(x=entry_x_offset+int(1.8*entry_x_interval), y=entry_y_offset+2*entry_y_interval)

    # 通知模式
    notic_mode_label = tk.Label(window, text="通知模式:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    notic_mode_label.place(x=label_x_offset, y=label_y_offset+3*label_y_interval)
    notic_mode_var = tk.IntVar(value=2)
    notic_mode_radio1 = tk.Radiobutton(window, text="输入通知", variable=notic_mode_var, value=1)
    notic_mode_radio1.place(x=entry_x_offset, y=entry_y_offset+3*entry_y_interval)
    notic_mode_radio2 = tk.Radiobutton(window, text="窗口通知", variable=notic_mode_var, value=2)
    notic_mode_radio2.place(x=entry_x_offset+entry_x_interval, y=entry_y_offset+3*entry_y_interval)

    # 是否显示物流公司
    show_logistics_label = tk.Label(window, text="显示物流:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    show_logistics_label.place(x=label_x_offset, y=label_y_offset+4*label_y_interval)
    show_logistics_var = tk.BooleanVar(value=True)
    show_logistics_checkbox = tk.Checkbutton(window, variable=show_logistics_var)
    show_logistics_checkbox.place(x=entry_x_offset, y=entry_y_offset+4*entry_y_interval)
    show_logistics_note_label = tk.Label(window, text="仅输入通知模式生效", font=f'{label_notice_font_fam} {label_notice_font_size} {label_notic_weight}', fg=label_notice_color)
    show_logistics_note_label.place(x=entry_x_offset+60, y=entry_y_offset+4*entry_y_interval)

    # 物流模式
    logistics_mode_label = tk.Label(window, text="物流模式:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    logistics_mode_label.place(x=label_x_offset, y=label_y_offset+5*label_y_interval)
    logistics_mode_var = tk.IntVar(value=1)
    logistics_mode_radio1 = tk.Radiobutton(window, text="自动识别", variable=logistics_mode_var, value=1)
    logistics_mode_radio1.place(x=entry_x_offset, y=entry_y_offset+5*entry_y_interval)
    logistics_mode_radio2 = tk.Radiobutton(window, text="手动输入", variable=logistics_mode_var, value=2)
    logistics_mode_radio2.place(x=entry_x_offset+entry_x_interval, y=entry_y_offset+5*entry_y_interval)

    # 目录日期 指的是如果所选文件在项目目录且非今天日期文件夹下，需要修改这个参数，如果是在桌面目录下，则不需要保持默认值或为空或今天日期，程序会自动复制一份到项目目录的今日日期文件夹下
    use_today_label = tk.Label(window, text="目录日期:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    use_today_label.place(x=label_x_offset, y=label_y_offset+6*label_y_interval)
    use_today_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
    use_today_entry = tk.Entry(window, textvariable=use_today_var, width=int(0.7*entry_width), font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    use_today_entry.place(x=entry_x_offset, y=entry_y_offset+6*entry_y_interval)
    use_today_today_button = tk.Label(window, text="当天日期", cursor='hand2', compound='center', font=f'{entry_font_fam} {entry_font_size-2} bold', fg=entry_color)
    use_today_today_button.bind("<Button-1>", lambda event: set_today_date())
    use_today_today_button.place(x=entry_x_offset+int(1.3*entry_x_interval), y=entry_y_offset+6*entry_y_interval)
    use_today_note_label = tk.Label(window, text="为空默认当天", font=f'{label_notice_font_fam} {label_notice_font_size} {label_notic_weight}', fg=label_notice_color)
    use_today_note_label.place(x=entry_x_offset+int(1.6*entry_x_interval), y=entry_y_offset+6*entry_y_interval)

    # 测试模式
    test_mode_label = tk.Label(window, text="测试模式:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    test_mode_label.place(x=label_x_offset, y=label_y_offset+7*label_y_interval)
    test_mode_var = tk.StringVar(value='0')
    test_mode_entry = tk.Entry(window, textvariable=test_mode_var, width=int(0.7*entry_width), font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    test_mode_entry.place(x=entry_x_offset, y=label_y_offset+7*label_y_interval)
    test_mode_note_label = tk.Label(window, text="0为不测试", font=f'{label_notice_font_fam} {label_notice_font_size} {label_notic_weight}', fg=label_notice_color)
    test_mode_note_label.place(x=entry_x_offset+int(1.3*entry_x_interval), y=label_y_offset+7*label_y_interval)

    # 是否回写表格
    is_write_label = tk.Label(window, text="是否回写:", font=f'{label_font_fam} {label_font_size} {label_weight}', fg=label_color)
    is_write_label.place(x=label_x_offset, y=label_y_offset+8*label_y_interval)
    is_write_var = tk.BooleanVar(value=True)
    is_write_checkbox = tk.Checkbutton(window, variable=is_write_var)
    is_write_checkbox.place(x=entry_x_offset, y=label_y_offset+8*label_y_interval)

    # OK 按钮
    ok_button = tk.Button(window, text="通知补发", command=on_ok)
    ok_button.place(x=button_x_offset, y=button_y_offset)

    # 使用默认值按钮
    use_defaults_button = tk.Label(window, text="默认值", cursor='hand2', compound='center', font=f'{button_font_fam} {button_font_size-1} bold', fg=button_color)
    use_defaults_button.bind("<Button-1>", lambda event: use_defaults())
    use_defaults_button.place(x=set_button_x_offset, y=set_button_y_offset)

    # 使用上次内容按钮
    use_last_used_button = tk.Label(window, text="上次值", cursor='hand2', compound='center', font=f'{button_font_fam} {button_font_size-1} bold', fg=button_color)
    use_last_used_button.bind("<Button-1>", lambda event: use_last_used())
    use_last_used_button.place(x=set_button_x_offset+2*set_button_x_interval, y=set_button_y_offset)

    # 运行主循环
    window.protocol("WM_DELETE_WINDOW", on_close)  # 设置关闭窗口时的回调函数
    window.mainloop()


def on_close():
    global window, window_initialized
    window.destroy()
    window = None
    window_initialized = False  # 窗口关闭后重置标志

# 定义一个回调函数，在主线程中调用create_window
def call_create_window():
    global window
    if window is not None and window.winfo_exists():
        window.lift()  # 如果窗口已经存在，将其提升到最前面
        window.focus_force()  # 强制窗口获得焦点
    else:
        create_window(mode=window_open_mode)
        # window.after(0, create_window)  # 如果窗口不存在，创建新窗口


# 使用上一次数据直接通知
def notic_last_data():
    # 读取配置文件
    notic_config = load_config()
    last_data = get_last_used(notic_config)
    if last_data is not None:
        au.notification_reissue(**last_data)
        print("使用上次数据通知补发")
        show_toast("提醒", "使用上次数据通知补发...")
    else:
        print("没有上次数据")
        show_toast("提醒", "没有上次数据")

# 测试
if __name__ == "__main__":
    create_window()