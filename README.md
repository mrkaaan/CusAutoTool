# CusAutoTool - 客服自动化工具使用说明

> 个人开发项目备忘录，记录各模块功能和关键实现细节

---

## 项目背景
CusAutoTools基于Python开发，通过自动化操作和快捷键集成解决电商客服日常工作中的重复性操作。**核心思路**：通过快捷键触发自动化流程，减少人工点击。

基于 [Joey Gambler的python-WinGUI](https://github.com/JoeyGambler/python-WinGUI) 项目二次开发，继承MIT开源协议。

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


#### █ 快捷键引擎：utils.auto_key()
**核心作用**：将快捷键绑定到具体功能，并处理多线程执行

#### 工作流程伪代码
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

#### 1. utils.py - 基础工具
依赖keyboard、win32con、win32gu、pypercli等库，主要操作剪切板以及对键盘和鼠标
- **open_sof()**：通过窗口句柄快速打开软件（缓存句柄到handles.json）
- **move_mouse()**：移动鼠标到预设坐标（调试用）

#### 2. auto_operation.py - ERP/千牛自动化
依赖于WinGUI二次开发，核心功能包括：
- **erp_common_action_1()**：ERP通用操作（选择日期+清空商品+输入备注）
- **handle_auto_send_price_link()**：自动发送差价链接
- **win_key()**：窗口切换快捷键（绑定鼠标侧键）

#### 3. auto_copy_clipboard_latest.py - 文件替换工具
核心功能：根据输入框文字自动粘贴对应文件
**on_press_clipboard()** 工作流程：
1. 监听当前输入框文字
2. 匹配`hotstrings_cn.json`中的预设关键词
3. 调用`scripts/copy_clipboard.bat`将对应文件复制到剪切板
4. 自动执行Ctrl+V粘贴

#### 4. notification_reissue_window.py - 补发通知
- **call_create_window()**：创建TKinter配置窗口
- **notic_last_data()**：使用上次配置自动发送通知
- 依赖`config/notic_config.json`存储配置

---

## 特殊机制说明
### ▶ 坐标定位
- **原理**：通过`WinGUI.py`的find_image()方法匹配界面截图
- **截图存放**：`image/`目录下的png文件（如erp_confirm_button.png）

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

