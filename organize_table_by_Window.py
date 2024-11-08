import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import threading

# 您的处理函数
def process_original_table(input_filename, form_folder):
    # 您的处理逻辑
    pass

def process_original_number(input_filename, shop_name, form_folder):
    # 您的处理逻辑
    pass

# 线程化的函数，避免界面冻结
def run_in_thread(func, *args):
    threading.Thread(target=func, args=args).start()

# 处理原始表格
def handle_original_table():
    file_path = filedialog.askopenfilename(title="选择原始表格文件", filetypes=[("CSV files", "*.csv")])
    if file_path:
        run_in_thread(process_original_table, file_path, './form')

# 处理ERP导出表格
def handle_erp_export_table():
    file_path = filedialog.askopenfilename(title="选择ERP导出表格文件", filetypes=[("CSV files", "*.csv")])
    if file_path:
        run_in_thread(process_original_number, file_path, '余猫', './form')

# 主窗口
root = tk.Tk()
root.title("补发单号表格处理")

# 原始表格处理部分
tk.Label(root, text="原始表格处理").pack()
original_table_entry = tk.Entry(root, width=50)
original_table_entry.pack()
original_table_button = tk.Button(root, text="选择文件", command=handle_original_table)
original_table_button.pack()

# ERP导出表格处理部分
tk.Label(root, text="ERP导出表格处理").pack()
erp_export_entry = tk.Entry(root, width=50)
erp_export_entry.pack()
erp_export_button = tk.Button(root, text="选择文件", command=handle_erp_export_table)
erp_export_button.pack()

root.mainloop()