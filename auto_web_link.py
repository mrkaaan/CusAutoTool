# import webbrowser
# import time
# from pywinauto import Desktop
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.service import Service as ChromeService
# from selenium.webdriver.edge.service import Service as EdgeService
# from selenium.webdriver.chrome.options import Options as ChromeOptions
# from selenium.webdriver.edge.options import Options as EdgeOptions
# from subprocess import CREATE_NO_WINDOW
# import os

# def get_driver(browser, headless=False):
#     if browser.lower() == 'chrome':
#         options = ChromeOptions()
#         if headless:
#             options.add_argument('--headless')
#         service = ChromeService(executable_path=r"path\to\chromedriver.exe", creationflags=CREATE_NO_WINDOW)
#         driver = webdriver.Chrome(service=service, options=options)
#     elif browser.lower() == 'edge':
#         options = EdgeOptions()
#         if headless:
#             options.add_argument('--headless')
#         service = EdgeService(executable_path=r"path\to\msedgedriver.exe", creationflags=CREATE_NO_WINDOW)
#         driver = webdriver.Edge(service=service, options=options)
#     else:
#         raise ValueError("Unsupported browser. Please choose 'chrome' or 'edge'.")
#     return driver

# def open_or_focus_url(url, title, browser='chrome'):
#     # 获取所有浏览器窗口
#     desktop = Desktop(backend="uia")
#     browser_class_name = {
#         'chrome': 'Chrome_WidgetWin_1',
#         'edge': 'ApplicationFrameWindow'
#     }.get(browser.lower(), None)
    
#     if not browser_class_name:
#         raise ValueError("Unsupported browser. Please choose 'chrome' or 'edge'.")
    
#     browser_windows = desktop.windows(class_name=browser_class_name)
    
#     for window in browser_windows:
#         try:
#             # 激活窗口并获取其内容
#             window.set_focus()
#             time.sleep(0.5)  # 等待窗口加载
            
#             # 使用 Selenium 查找特定标签页
#             driver = get_driver(browser)
#             driver.get('about:blank')  # 打开一个空白页面以初始化驱动
            
#             # 切换到每个标签页并检查 URL
#             handles = driver.window_handles
#             for handle in handles:
#                 driver.switch_to.window(handle)
#                 if url in driver.current_url:
#                     print(f"Found tab with URL: {url}")
#                     driver.quit()
#                     return
            
#             driver.quit()
#         except Exception as e:
#             print(f"Error checking window: {e}")
    
#     # 如果没有找到匹配的标签页，则打开新标签页
#     print(f"Opening new tab with URL: {url}")
#     webbrowser.register(browser, None, webbrowser.BackgroundBrowser(webbrowser.Path(r"path\to\browser.exe")))
#     webbrowser.get(browser).open_new_tab(url)
