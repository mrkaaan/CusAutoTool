import pychrome
import keyboard

# 假设 Chrome 已经启动并启用了远程调试端口，默认是9222
browser = pychrome.Browser(url="http://127.0.0.1:9222")

def find_tab_by_url(tabs, url):
    for tab in tabs:
        # 确保我们获取的是字符串形式的 URL
        if hasattr(tab, 'url') and isinstance(tab.url, str) and url in tab.url:
            return tab
    return None

def on_hotkey():
    target_url = "https://www.kdocs.cn/l/cv9FaonW5wT8"
    
    # 获取所有标签页
    tabs = browser.list_tab()
    
    # 查找目标 URL 的标签页
    target_tab = find_tab_by_url(tabs, target_url)
    
    if target_tab:
        try:
            # 连接到找到的标签页
            target_tab.start()

            # 执行 JavaScript 来聚焦该标签页
            target_tab.call_method("Page.bringToFront")

            # 可选：执行更多 JavaScript 操作
            result = target_tab.call_method("Runtime.evaluate", 
                                           expression=f"window.location.href='{target_url}';",
                                           await_response=True)
            print(result)
        finally:
            # 断开连接
            target_tab.stop()
    else:
        print(f"No tab found with URL {target_url}")

# 修改快捷键为更简单的组合，比如 ctrl+alt+t
keyboard.add_hotkey('ctrl+alt+t', on_hotkey)

print("Listening for hotkey... Press Ctrl+C to exit.")
try:
    keyboard.wait('esc')  # 持续监听直到按下 ESC 键退出
except KeyboardInterrupt:
    pass