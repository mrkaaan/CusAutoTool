# CusAutoTool - 客服自动化工具使用说明

> 个人开发项目备忘录，记录各模块功能和关键实现细节

---

## 项目背景
CusAutoTools基于Python开发，通过自动化操作和快捷键集成解决电商客服日常工作中的重复性操作。**核心思路**：通过快捷键触发自动化流程，减少人工点击。

基于 [Joey Gambler的python-WinGUI](https://github.com/JoeyGambler/python-WinGUI) 项目二次开发，继承MIT开源协议。

---

## 功能列表
1. **文件替换功能**
   - 在输入框中输入指定文字，按下快捷键后，将剪切板中的文字替换为需要的图片或视频文件，便于快速发送。
2. **应用快捷键调用**
   - 为没有快捷键的应用设置快捷键，通过快捷键直接调用指定应用。
3. **旺店通ERP操作**
   - 快速操作旺店通ERP系统，按下快捷键后，按照预设的坐标点击并输入指定内容，实现产品的补发操作。
4. **补发通知自动化**
   - 处理ERP导出的补发单号表格，将其转换为程序需要的格式，并根据处理后的表格自动点击千牛接待台，给不同用户发送补发单号通知。

---

## 环境准备

### 基础环境
- Miniconda 3.7+
- Python 3.8+

### 依赖安装
1. **必须使用conda环境**（已有environment.yml）
   ```bash
   conda env create -f environment.yml
   conda activate gui
   # 需手动安装的库（不修改原环境配置）
   conda install pandas openpyxl
   ```
> environment.yml 文件为WinGUI项目中继承而来，避免修改出错，额外的两个库选择手动安装

2. **必须手动创建**：
   - 在`config/`目录新建`window_config.ini`（内容如下）：
     ```ini
     [defaults]
     WINDOW_OPEN_MODE = 0  # 0=小窗口模式，1=大窗口模式
     ```
> 这个文件是所有使用tkinter工具所需要的配置，**必须手动创建**，否则会报错，WINDOW_OPEN_MODE 用于所有 Tk 窗口，解决因不同调用方式导致的显示比例差异（如直接调用窗口文件、通过 auto_key 调用或在不同电脑上运行）。每窗口设有两套尺寸，通过 WINDOW_OPEN_MODE 选择：0 为小尺寸，1 为大尺寸。

---
## 项目结构
```
CusAutoTool/
├── config/               # 配置文件目录
│   ├── handles.json      # 软件窗口句柄缓存
│   ├── hotstrings_cn.json    # 快捷键映射配置
│   ├── coordinates.json      # 屏幕坐标配置
│   ├── notic_config.json     # 通知参数配置
│   └── window_config.ini     # 窗口显示配置（需手动创建）
│
├── src/                  # 核心代码
│   ├── main.py           # 程序入口
│   ├── WinGUI.py         # 窗口操作基类（继承自python-WinGUI）
│   ├── auto_operation.py # ERP/千牛自动化操作
│   ├── utils.py          # 快捷键管理核心
│   └── [其他功能模块].py   
│
├── scripts/              # 辅助脚本
│   └── copy_clipboard.bat # 文件剪切板操作脚本
│
├── form/                 # 数据表格目录
├── image/                # 界面元素截图 不可删除
└── temp/                 # 运行时临时文件
```
---

## 重要配置文件

| 文件路径                  | 作用说明                                                                 | 注意事项                      |
|--------------------------|--------------------------------------------------------------------------|-------------------------------|
| config/hotstrings.json | 定义文字→文件映射（旧版）                    | key只有拼音无中文，已不更新          |
| config/hotstrings_cn.json | 定义文字→文件映射（如输入"差价"自动粘贴差价链接图片）                    | 新增关键词需手动维护          |
| config/coordinates.json  | 存储按钮坐标（通过WinGUI自动生成）                                      | 更换显示器后需要重新校准      |
| config/handles.json      | 缓存软件窗口句柄（加速打开常用软件）                                    | open_sof方法创建，调用窗口的句柄，自动维护勿修改      |
| config/window_config.ini | 控制TKinter弹窗尺寸（0=小尺寸模式，1=大尺寸模式）                        | 必须手动创建！                |
| config/notic_config.ini | `notification_reissue` 配置参数：默认值和上次使用值                        | 默认值不可删除                |


**config/config.ini**  
```ini
[paths]
COPY_CLIPBOARD_PATH = scripts/copy_clipboard.bat
[names]
HOT_FILE_NAME = hotstrings_cn.json
[numbers]
TRY_NUMBER = 1 
WINDOW_OPEN_MODE = 1 
```
> `[paths]` 记录工具和映射文件的目录。
> `[names]` 包含热字符串文件名。
> `[numbers]` 中的 `TRY_NUMBER` 指定在WinGUI中查找特定图像时的重试次数。
> `WINDOW_OPEN_MODE` 已弃用，详情见下文。

**window_config.ini**  
由于不同电脑上Tk窗口显示比例可能存在差异，单独设立此配置文件以适配不同的显示需求。该文件不被Git管理，因此在新拉项目时需自行创建并正确配置。
```ini
[defaults]
WINDOW_OPEN_MODE = 0  # 选择窗口尺寸（0为较小尺寸，1为较大尺寸）
```

---

## 代码结构说明
### █ 核心入口：src/main.py
```python
# 启动流程（所有配置都从config.py导入）
if __name__ == "__main__":
    setup_pyautogui()            # 关闭pyautogui的安全保护（防止鼠标飞出屏幕暂停程序）
    setup_logging()              # 日志记录到config/dev.log（自动分割10MB文件）
    window_name = r"千牛接待台"    # 要操作的窗口名称（必须与软件标题完全一致！）
    window_name_erp = r"旺店通ERP" 
    load_coordinates_from_json(r'../config/coordinates.json')  # 加载按钮坐标配置
    
    # 命令行参数处理（--test进入测试模式）
    input_args = setup_arguments()  
    if input_args.test:
        test()  # 运行测试函数
    else:
        main()  # 正式运行
```

### █ 配置中心：src/config.py
```python
# 关键配置函数说明（这些函数会被main.py调用）
def setup_arguments():
    '''处理命令行参数，目前只支持--test进入测试模式'''
    parser = argparse.ArgumentParser(description="Main script and testing.")
    parser.add_argument('--test', action='store_true', help='Run tests')
    return parser.parse_args()

def setup_pyautogui():
    '''配置自动化参数：关闭安全保护+设置操作间隔0.1秒'''
    pyautogui.FAILSAFE = False  # 重要！否则鼠标移到左上角会触发异常
    pyautogui.PAUSE = 0.1       # 每个操作后暂停0.1秒（太快会出错）

def setup_logging():
    '''日志配置：记录到config/dev.log，文件超过10MB自动分割'''
    logger.add("../config/dev.log", rotation="10 MB")

def setup_bat_path():
    '''获取scripts/copy_clipboard.bat的绝对路径（文件操作依赖这个）'''
    return ut.get_bat_path()  # 具体实现在utils.py
```

### █ 快捷键绑定（main函数内部）
```python
def main():
    # 所有功能都通过这个数组绑定快捷键（格式：快捷键 → 执行函数）
    hotkey_actions = [
        # 示例1：按alt+shift+e打开ERP软件
        {'key': 'alt+shift+e', 'func': ut.open_sof, 'args': ['旺店通ERP']},
        
        # 示例2：按F2触发文件粘贴功能
        {'key': 'f2', 'func': ac.on_press_clipboard},
        
        # 示例3：ERP补发操作（带复杂参数）
        {'key': 'ctrl+b+num 7', 
         'func': ao.erp_common_action_1, 
         'args': ['旺店通ERP']},
         
        # 更多快捷键见实际代码...
    ]
    ut.auto_key(hotkey_actions)  # 启动快捷键监听
```

---

### █ 核心工具包


#### █ utils.auto_key() - 快捷键引擎
**核心作用**：将快捷键绑定到具体功能，并处理多线程执行

##### 工作流程伪代码
```python
def auto_key(hotkeys):
    # 阶段1：过滤重复快捷键
    创建 seen_keys 集合
    for 每个快捷键配置 in hotkeys:
        if 快捷键已存在:
            提示重复并跳过
        else:
            添加到 filtered_hotkeys 列表

    # 阶段2：定义执行逻辑
    def 执行函数(func, args):
        if 需要开线程（use_thread参数）:
            创建守护线程执行目标函数
            if 需要重复执行（redo参数）:
                同时启动两个线程
        else:
            直接执行函数
            if 需要重复执行:
                连续调用两次

    # 阶段3：绑定监听
    for 每个过滤后的快捷键 in filtered_hotkeys:
        绑定到 keyboard 库的 add_hotkey 方法
        通过lambda传递参数（特别注意闭包问题！）

    # 阶段4：保持监听
    设置退出快捷键（Shift+Ctrl+E）
    进入 keyboard.wait() 持续监听状态
```

##### 关键参数说明
| 参数名      | 作用                                                                 |
|------------|----------------------------------------------------------------------|
| use_thread | 是否使用新线程执行（防止阻塞主线程，适用于长时间操作如自动回复）     |
| redo       | 是否重复执行（某些场景需要连续触发两次才能生效）                     |
| key        | 快捷键组合字符串（格式示例："ctrl+shift+q"）                         |

##### 线程管理
- 使用`Thread(target=函数).start()`创建守护线程
- 示例场景：自动回复功能需要持续监听消息，必须独立线程运行
- 注意：线程中不能直接操作UI组件（如TKinter元素）


##### 典型调用示例
```python
# 在main.py中的实际使用
ut.auto_key([
    {'key': 'f2', 'func': 文件粘贴函数}, 
    {'key': 'ctrl+b+num 7', 
     'func': ERP操作函数, 
     'args': ['旺店通ERP'], 
     'use_thread': True,  # 需要长时间运行故启用线程
     'redo': False}
])
```

#### █ utils.py - 基础工具
依赖keyboard、win32con、win32gu、pypercli等库，主要操作剪切板以及对键盘和鼠标
- **open_sof()**：通过窗口句柄快速打开软件（缓存句柄到handles.json）
- **move_mouse()**：移动鼠标到预设坐标（调试用）

#### █ auto_operation.py - ERP/千牛自动化操作核心
依赖于WinGUI二次开发
- **截图存放**：`image/`目录下的png文件（如erp_confirm_button.png）
```python
# 基于 WinGUI 的自动化操作
class WinGUI:  # 来自 src/WinGUI.py
    def get_app_screenshot(self): ...  # 获取窗口截图
    def move_and_click(self, x, y): ...  # 移动鼠标并点击
    def locate_icon(self, img_name): ...  # 通过图像匹配定位按钮

# 典型功能实现
def erp_common_action_1(window_name):
    '''ERP标准补发流程'''
    app = WinGUI(window_name)  # 初始化窗口控制器
    app.move_and_click(*load_coordinates('清空按钮'))  # 点击清空按钮
    app.move_and_click(*load_coordinates('今天日期'))  # 选择当天日期
    # 更多步骤...

def wait_a_moment_by_qianniu(window_name):
    '''千牛自动回复'''
    app = WinGUI(window_name)
    while 运行标志:
        if 发现新消息弹窗():
            app.move_and_click(输入框坐标)  # 定位到聊天窗口
            pyautogui.typewrite("请稍等正在查询...")  # 发送预设话术
        time.sleep(2)  # 每2秒检测一次
def handle_auto_send_price_link():
   '''自动发送差价链接'''

def win_key():
   '''窗口切换快捷键（绑定鼠标侧键）'''
```



#### █ notification_reissue_window.py - 补发通知
- **call_create_window()**：创建TKinter配置窗口
- **notic_last_data()**：使用上次配置自动发送通知
- 依赖`config/notic_config.json`存储配置


#### █ auto_copy_clipboard_latest.py - 文件替换工具
核心功能：根据输入框文字自动粘贴对应文件
```python
def on_press_clipboard(keyword='hp'):
    '''
    工作流程：
    1. 监听当前输入框内容（如用户输入"好评截图"）
    2. 匹配`hotstrings_cn.json`中的预设关键词
    3. 调用`scripts/copy_clipboard.bat`将对应文件复制到剪切板
    4. 自动按下Ctrl+V完成粘贴
    '''
    # 示例配置文件内容
    # hotstrings_cn.json:
    # {
    #   "好评截图": "D:/images/好评截图.jpg",
    #   "差价链接": "D:/links/差价.html"
    # }
```

**文件映射配置**：
- 配置文件：`config/hotstrings_cn.json`
- 格式：`{"关键词": "文件路径"}`

**bat脚本作用**：
- 功能：将指定文件复制到剪切板
- 调用方式：`scripts/copy_clipboard.bat 文件路径`


**维护建议**
1. **关键词设计**：
   - 尽量简短易记（如"hp"对应好评截图）
   - 避免重复映射

2. **文件管理**：
   - 统一存放路径（如D:/images/）
   - 定期更新映射文件

---

#### █ notification_reissue_window.py - 补发通知工具
```python
# 核心功能：自动发送补发通知
def call_create_window():
    '''创建配置窗口'''
    if 窗口已存在:
        激活现有窗口
    else:
        创建新窗口(mode=window_open_mode)  # 支持两种窗口尺寸

def notic_last_data():
    '''使用上次配置发送通知'''
    读取config/notic_config.json中的参数
    调用auto_operation.notification_reissue()
```

##### 工作流程
1. **数据准备**：
   - 从ERP导出补发订单表格（CSV格式）
   - 使用`organize_table`整理数据

2. **通知发送**：
   - 读取整理后的Excel文件
   - 逐个客户发送补发单号
   - 标记已通知客户

3. **参数配置**：
   - 配置文件：`config/notic_config.json`
   - 包含：
     - 默认参数（窗口显示用）
     - 上次使用的参数

##### 特殊机制
1. **多店铺支持**：
   - 不同店铺使用不同通知模板
   - 支持自定义点击步骤

2. **测试模式**：
   - 不实际发送消息
   - 不修改Excel文件

3. **历史记录**：
   - 每次通知结果记录到日志
   - 支持断点续传

##### 典型调用
```python
# 快捷键绑定示例
{'key': 'ctrl+t+num 2', 'func': call_create_window}  # 打开配置窗口
{'key': 'ctrl+t+num 3', 'func': notic_last_data}    # 使用上次配置直接发送
```

---

#### █ organize_table_window.py - 表格整理工具
```python
# 核心功能：ERP表格整理
def call_create_window():
    '''创建表格整理窗口'''
    if 窗口已存在:
        激活现有窗口
    else:
        创建新窗口(mode=window_open_mode)
```

##### 主要作用
1. **数据拆分**：
   - 按店铺分割订单数据
   - 生成多个工作表

2. **格式转换**：
   - CSV → Excel
   - 自动添加表头

3. **文件管理**：
   - 按日期归档（form/YYYY-MM-DD/）
   - 保留原始数据备份

##### 使用场景
- ERP导出数据后一键整理
- 快速生成分店铺报表
---

## 特殊机制说明
### ▶ WinGUI 窗口控制器


#### 核心功能
```python
# 初始化
app = WinGUI("窗口标题")  # 必须与目标窗口标题完全一致

# 常用方法
app.get_app_screenshot()  # 获取窗口截图（保存到temp/app.png）
app.move_and_click(x, y)  # 移动鼠标到(x,y)并点击
app.locate_icon("按钮截图.png")  # 通过图像匹配定位按钮
```

#### 工作原理
1. **窗口定位**：
   - 使用`win32gui`库根据窗口标题查找句柄
   - 自动将目标窗口置顶并最大化

2. **图像匹配**：
   - 依赖OpenCV的模板匹配算法
   - 匹配精度可配置（默认相似度>90%）

3. **坐标系统**：
   - 所有坐标基于屏幕左上角(0,0)
   - 支持相对坐标移动（rel_remove_and_click）

#### 维护建议
1. **截图规范**：
   - 截图存放于`image/`目录
   - 命名规则：`功能_按钮名.png`
   - 建议尺寸：50x50像素

2. **性能优化**：
   - 首次运行会生成窗口截图（temp/app.png）
   - 后续操作直接使用缓存截图

3. **调试技巧**：
   - 使用`get_workscreen_screenshot()`获取全屏截图
   - 通过`check_icon()`验证按钮是否存在

### ▶ 坐标定位系统
#### 实现方式
```python
# 坐标来源（二选一）
1. 手动配置：直接修改config/coordinates.json
   - 格式：{"清空按钮": [125, 330]}

2. 自动获取：通过WinGUI的locate_icon()
   - 示例：app.locate_icon("清空按钮.png")
```


**常见问题**：
| 现象                 | 解决方案                          |
|----------------------|-----------------------------------|
| 点击位置偏移         | 检查截图是否匹配，重新生成坐标     |
| 找不到按钮           | 确认窗口名称正确，调整截图区域     |  
| 部分分辨率失效       | 在对应分辨率电脑重新采集坐标       |


### ▶ 多窗口模式适配
- 通过`window_config.ini`中的WINDOW_OPEN_MODE参数切换：
  - Mode 0：适合笔记本小屏幕（窗口尺寸较小）
  - Mode 1：适合外接显示器（调整控件位置）

### ▶ 异常处理机制
- 所有自动化操作都有**重试机制**（config.ini的TRY_NUMBER）
- 按下`Shift+Ctrl+E`可强制退出程序

---

## 注意事项
1. **必须存在但不上传的文件**：
   - config/window_config.ini
   - temp/app.png（运行时自动生成）
   - config/handles.json （运行时自动生成）
   - config/dev.log 将 WinGUI 的 log 日志移至 config.py，并更新部分内容。后续开发内容使用 print 而非 log。


2. **调试技巧**：
   - 使用`python src/main.py --test`进入测试模式

---

## 致谢声明
本项目基于 [Joey Gambler的python-WinGUI](https://github.com/JoeyGambler/python-WinGUI) 开发，核心窗口操作类`WinGUI.py`直接继承自原项目，遵循MIT协议（完整协议见LICENSE文件）。

## 开发中功能

结合以上功能，选择一个合适的技术方法，将项目封装为一个操作友好的应用，并实现以下功能：
- 桌面的悬浮球（类似utools的悬浮球），可以拖动，可以关闭，可以最小化，可以右键呼出菜单
- 悬浮球带有右下角托盘图标 右键点击可以呼出菜单
- 悬浮球无论使用什么技术实现都要可以与python原本的功能进行通信，便于后续功能的开发
- 悬浮球的菜单以及右下角托盘的菜单要有的功能
   - 结束程序
   - 重启程序   
（为了便于你的理解，我这里解释一下我的程序的启动流程是，保证进入项目下的src目录下，使用python main.py运行，使用ctrl+c进行结束）

我是个人开发的工具，使用起来比较要求界面的美观以及响应速度快，以及开发相对简单
1 选用Electron（结合 Python 和前端技术），因为我本身有一定的前端基础，electron也已经入门，而且也美观，这个暂时不用提供代码，而是讲解一下python这边要用什么技术处理呢，比如Flask 或 FastAPI ，这两个是什么，讲解一下

2 解答我关于打包的疑惑，因为我的最终目标是悬浮窗与原本的python为一体，我是否需要先将我的项目打包再创建悬浮窗，还是先构建悬浮窗再打包。或者换句话说悬浮窗与python原本是同级别还是包含关系

3 有关项目依赖还有一个疑惑，我的项目依赖于conda的虚拟环境，虚拟环境配置好后我个人开发的时候需要 conda activate gui，
虽然可以将 Conda 环境导出为 requirements.txt 文件：
conda list --export > requirements.txt
然后在打包时指定该文件，确保所有依赖被包含。
但是如果在conda activate gui激活的情况下直接不指定requirements.txt文件打包会是这么样子呢？默认把conda环境下的所有包都打包进去吗？

4 还有就是打包时是否需要指定什么参数吗，因为我的项目比较大涉及到比如键盘操作，循环监听，异步，tk窗口等，直接打包可以吗