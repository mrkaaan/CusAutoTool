import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import threading
from datetime import datetime
import pyperclip
import shutil

from organize_table import process_original_table, process_original_number

file_path_or = ''
file_name_or = ''
file_path_erp = ''
file_name_erp = ''

# 线程化的函数，避免界面冻结
def run_in_thread(func, *args):
    threading.Thread(target=func, args=args).start()

# 选择原始表格
def choose_original_table(enter):
    global file_path_or
    global file_name_or

    # 设置默认打开目录为桌面
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

    file_path_temp = filedialog.askopenfilename(
        title="选择文件",
        initialdir=desktop_path,  # 默认打开桌面
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    if file_path_temp:
        # 追加新的文本
        # current_text = or_entry.get()  # 获取当前的文本
        enter.delete(0, tk.END)  # 清除输入框中的所有文本
        file_path_or = file_path_temp
        file_name_or = os.path.basename(file_path_temp)
        enter.insert(0, file_path_or)  # 插入新的组合文本

        # 获取当前日期的年月日
        today = datetime.now().strftime('%Y-%m-%d')
        target_dir = f"./form/{today}"

        # 如果目标目录不存在，则创建
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)

         # 获取目标文件路径
        file_name = os.path.basename(file_path_temp)
        target_file_path = os.path.join(target_dir, file_name)

        # 复制文件到目标目录
        try:
            shutil.copy(file_path_temp, target_file_path)
            print(f"文件已成功复制到 {target_file_path}")
        except Exception as e:
            print(f"文件复制失败：{e}")

# 选择ERP导出表格
def choose_erp_export_table(enter):
    global file_path_erp
    global file_name_erp

    # 获取当前日期的年月日
    today = datetime.now().strftime('%Y-%m-%d')
    default_dir = f"./form/{today}"
    
    # 如果目录不存在，则创建
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)
    
    # 打开文件对话框，设置初始目录为默认目录
    file_path_temp = filedialog.askopenfilename(
        title="选择ERP导出表格文件",
        initialdir=default_dir,  # 默认打开目录
        filetypes=[("CSV files", "*.csv")]
    )

    if file_path_temp:
        # 追加新的文本
        # current_text = or_entry.get()  # 获取当前的文本
        enter.delete(0, tk.END)  # 清除输入框中的所有文本
        file_path_erp = file_path_temp
        file_name_erp = os.path.basename(file_path_temp)
        enter.insert(0, file_path_erp)  # 插入新的组合文本

# 复制ERP导出表格表明
def copy_erp_table_name(shop_number):
    # 获取今天的日期字符串
    today_date = datetime.now().strftime('%Y-%m-%d')
    table_name = ''
    if shop_number == '1':
        table_name = f'{today_date}_余猫_ERP二次导出表格'
    elif shop_number == '2':
        table_name = f'{today_date}_潮洁_ERP二次导出表格'
    else:
        pass
    # 将今天的日期复制到剪贴板
    pyperclip.copy(table_name)
    # 可以在这里打印一些信息，确认日期已复制
    # print("今天的日期已复制到剪贴板：", today_date)

# 检测Entry内容变化
def update_button_state(entry_var, button, *args):
    if entry_var.get() == "":
        button.config(state='disabled', cursor="")
    else:
        button.config(state='normal', cursor="hand2")

def main():    # 主窗口
    root = tk.Tk()
    root.title("补发表格处理")

    global file_path_or
    global file_name_or
    global file_path_erp
    global file_name_erp

    # 参数设置
    # 窗口大小
    window_width = 520
    window_height = 250
    # label字体设置
    label_font_fam = '楷体'
    label_font_size = 15
    label_color = '#000000'
    # entry设置
    entry_font_fam = '楷体'
    entry_font_size = 15
    entry_color = '#000000'
    entry_width = 47

    # 设置窗口居中
    scnwidth, scnheight = root.maxsize()
    tmp = '%dx%d+%d+%d' % (
        window_width, window_height, (scnwidth - window_width) / 2, (scnheight - window_height) / 2)  # 参数为（‘窗口宽 x 窗口高 + 窗口位于x轴位置 + 窗口位于y轴位置）
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
    # 清空输入框按钮
    or_button_clear = tk.Button(root, cursor="hand2", text="清空", command=lambda:or_entry.delete(0, tk.END))
    or_button_clear.place(x=485, y=36)
    # 选择文件按钮
    or_button_choose = tk.Button(root, cursor="hand2", text="选择文件", command=lambda:choose_original_table(or_entry))
    or_button_choose.place(x=180, y=70)
    # 开始处理按钮
    or_button_sure = tk.Button(root, state="disabled", text="开始处理", command=lambda:run_in_thread(process_original_table, file_name_or))
    or_button_sure.place(x=270, y=70)
    # 监控输入框的变化，并更新按钮状态
    or_entry_var.trace("w", lambda *args: update_button_state(or_entry, or_button_sure, *args))

    # ERP导出表格处理部分
    # 标题
    erp_label = tk.Label(root, text="ERP导出表格处理", compound='center', font=f'{label_font_fam} {label_font_size} bold', fg=label_color)
    erp_label.place(x=10, y=130)
    # 快捷复制表明
    erp_table_name_1 = tk.Label(root, text="余猫", cursor='hand2',compound='center', font=f'{label_font_fam} {label_font_size-5} bold', fg=label_color)
    erp_table_name_1.place(x=10, y=190)
    erp_table_name_1.bind("<Button-1>", lambda event: copy_erp_table_name('1'))
    erp_table_name_2 = tk.Label(root, text="潮洁", cursor='hand2',compound='center', font=f'{label_font_fam} {label_font_size-5} bold', fg=label_color)
    erp_table_name_2.place(x=50, y=190)
    erp_table_name_2.bind("<Button-1>", lambda event: copy_erp_table_name('2'))
    # 输入框
    erp_entry_var = tk.StringVar()
    erp_entry = tk.Entry(root, textvariable=erp_entry_var,width=entry_width,font=f'{entry_font_fam} {entry_font_size}', fg=entry_color)
    erp_entry.place(x=10, y=160)
    # 清空输入框按钮
    erp_button_clear = tk.Button(root, cursor="hand2", text="清空", command=lambda:erp_entry.delete(0, tk.END))
    erp_button_clear.place(x=485, y=156)
    # 选择文件按钮
    erp_button_choose = tk.Button(root, cursor="hand2", text="选择文件", command=lambda:choose_erp_export_table(erp_entry))
    erp_button_choose.place(x=180, y=190)
    # 处理按钮
    erp_button_sure = tk.Button(root, state="disabled", text="开始处理", command=lambda:run_in_thread(process_original_number, file_name_erp))
    erp_button_sure.place(x=270, y=190)
    # 监控输入框的变化，并更新按钮状态
    erp_entry_var.trace("w", lambda *args: update_button_state(or_entry, erp_button_sure, *args))

    root.mainloop()

if __name__ == "__main__":
    main()