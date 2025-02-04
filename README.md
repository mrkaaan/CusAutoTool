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
> 这个文件是所有使用tkinter工具所需要的配置，**必须手动创建**，否则会报错

---

## 代码结构说明
### █ 核心入口：src/main.py
```python
# 主要逻辑流程
if __name__ == "__main__":
    setup_pyautogui()          # 禁用pyautogui的安全措施
    setup_logging()            # 日志配置到config/dev.log
    load_coordinates_from_json()  # 加载坐标配置文件
    main()                     # 正式运行
```

**hotkey_actions数组**定义了所有快捷键绑定：
- `alt+shift+e` → 打开ERP软件（utils.open_sof）
- `ctrl+shift+q` → 千牛自动备注（auto_operation.run_once_remarks_by_qianniu）
- `f2` → 粘贴预设文件（auto_copy_clipboard_latest.on_press_clipboard）

---

### █ 关键模块说明
#### 1. utils.py - 基础工具
- **setup_logging()**：配置日志到`config/dev.log`
- **open_sof()**：通过窗口句柄快速打开软件（缓存句柄到handles.json）
- **move_mouse()**：移动鼠标到预设坐标（调试用）

#### 2. auto_operation.py - ERP/千牛自动化
- **erp_common_action_1()**：ERP通用操作（选择日期+清空商品+输入备注）
- **handle_auto_send_price_link()**：自动发送差价链接
- **win_key()**：窗口切换快捷键（绑定鼠标侧键）

#### 3. auto_copy_clipboard_latest.py - 智能粘贴
- **on_press_clipboard()**核心逻辑：
  1. 监听当前输入框文字
  2. 匹配`hotstrings_cn.json`中的预设关键词
  3. 调用`scripts/copy_clipboard.bat`将对应文件复制到剪切板
  4. 自动执行Ctrl+V粘贴

#### 4. notification_reissue_window.py - 补发通知
- **call_create_window()**：创建TKinter配置窗口
- **notic_last_data()**：使用上次配置自动发送通知
- 依赖`config/notic_config.json`存储配置

---

## 重要配置文件
| 文件路径                   | 作用说明                                                                 |
|---------------------------|--------------------------------------------------------------------------|
| config/hotstrings_cn.json | 关键词映射文件，格式：`{"hp": "D:/好评截图.jpg"}`                        |
| config/coordinates.json   | 屏幕坐标配置，记录ERP按钮位置等                                         |
| config/config.ini         | 全局路径配置：[paths]COPY_CLIPBOARD_PATH = scripts/copy_clipboard.bat  |

---

## 特殊机制说明
### ▶ 坐标定位系统
- **原理**：通过`WinGUI.py`的find_image()方法匹配界面截图
- **截图存放**：`image/`目录下的png文件（如erp_confirm_button.png）
- **坐标缓存**：首次找到按钮后坐标会存入coordinates.json

### ▶ 多窗口模式适配
- 通过`window_config.ini`中的WINDOW_OPEN_MODE参数切换：
  - Mode 0：适合笔记本小屏幕（窗口尺寸较小）
  - Mode 1：适合外接显示器（调整控件位置）

### ▶ 异常处理机制
- 所有自动化操作都有**重试机制**（默认重试3次）
- 按下`Shift+Ctrl+E`可强制退出程序
- 崩溃时会自动清理临时文件（temp/目录）

---

## 注意事项
1. **必须存在但不上传的文件**：
   - config/window_config.ini
   - temp/app.png（运行时自动生成）

2. **功能依赖**：
   - ERP操作需要保持旺店通窗口名称一致
   - 发送文件功能依赖Everything搜索工具

3. **调试技巧**：
   - 使用`python src/main.py --test`进入测试模式
   - 按`Ctrl+Shift+Alt+M`快速移动鼠标到调试位置

---

## 致谢声明
本项目基于 [Joey Gambler的python-WinGUI](https://github.com/JoeyGambler/python-WinGUI) 开发，核心窗口操作类`WinGUI.py`直接继承自原项目，遵循MIT协议（完整协议见LICENSE文件）。

