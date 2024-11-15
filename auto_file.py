import subprocess

bat_file_path = r'D:\project\auto_customer\auto_file.bat'
# file_path = r'D:\test.png'
file_path = r'C:\Users\Kan\Documents\Captura\2024-11-15-11-50-10.mp4'

# 使用 subprocess.run 执行批处理文件
try:
    result = subprocess.run([bat_file_path, file_path], check=True)
    print("批处理文件执行成功")
except subprocess.CalledProcessError as e:
    print(f"批处理文件执行失败: {e}")


# pip install pynput