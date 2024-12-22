自动打开网页
AutoHotkey 主要通过窗口标题和进程名来管理窗口，无法实现判断浏览器是否打开了特定链接
使用 Python 的 pywinauto 和 selenium 库提供了更强大的功能来处理这类任务

pip install pywinauto selenium

---

selenium 的效率太低了

之前用的是autohotkey，重复打开相同网页
换用python的pychrome，如果已经有相同网页，则切换到该网页，不会重复打开

pip install pychrome


https://www.cnblogs.com/--kisaragi--/p/15241080.html

需要打开远程调试功能
先找到chrome.exe的文件位置
cd C:\Program Files\Google\Chrome\Application
开启远程控制命令
chrome.exe --remote-debugging-port=9222


127.0.0.1:9222/json/list

chorme://inspect/