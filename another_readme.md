以下是重新整理后的规范README文档：

# CusAutoTool - 客服自动化工具

## 目录
- [项目简介](#项目简介)
- [功能特性](#功能特性)
- [环境要求](#环境要求)
- [快速开始](#快速开始)
  - [安装依赖](#安装依赖)
  - [配置文件](#配置文件)
  - [运行程序](#运行程序)
- [项目结构](#项目结构)
- [功能详解](#功能详解)
- [注意事项](#注意事项)
- [致谢与许可](#致谢与许可)

## 项目简介
CusAutoTool是一款基于Python开发的客服效率提升工具，通过自动化操作和快捷键集成，显著优化以下工作场景：
- 电商客服高频操作自动化
- ERP系统批量操作
- 聊天界面快速响应
- 数据处理与通知自动化

基于 [Joey Gambler 的 python-WinGUI](https://github.com/JoeyGambler/python-WinGUI) 项目二次开发，继承MIT开源协议。

## 功能特性
### 核心功能模块
1. **智能内容替换**
   - 文本转多媒体：将预设关键词自动替换为图片/视频文件
   - 地址信息自动补全：智能拼接物流信息
   - 模板消息快速输入

2. **ERP自动化**
   - 一键补发操作
   - 智能表单填写
   - 仓库管理快捷操作
   - 批量订单处理

3. **聊天界面增强**
   - 客户信息快速复制
   - 自动应答机制
   - 差价链接快速发送
   - 消息模板管理

4. **数据处理工具**
   - Excel表格智能整理
   - 补发通知批量生成
   - 数据格式自动转换

## 环境要求
### 基础环境
- Miniconda 3.7+
- Python 3.8+

### 依赖安装
```bash
# 创建虚拟环境
conda env create -f environment.yml
conda activate gui

# 额外依赖（需手动安装）
conda install pandas openpyxl
```

## 快速开始
### 配置文件准备
1. 在`config/`目录下创建`window_config.ini`文件：
   ```ini
   [defaults]
   WINDOW_OPEN_MODE = 0  # 0-小窗口模式 1-大窗口模式
   ```
2. 根据实际需求配置`hotstrings_cn.json`文件

### 启动程序
```bash
python src/main.py
```

## 项目结构
```
CusAutoTool/
├── config/               # 配置文件目录
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
├── image/                # 界面元素截图
└── temp/                 # 运行时临时文件
```

## 功能详解
### 核心模块说明
1. **WinGUI 集成**
   - 继承自`python-WinGUI`的窗口操作基类
   - 实现窗口定位、坐标点击、图像识别等基础功能
   - 应用示例：ERP界面自动化操作

2. **快捷键引擎**
   - 使用`keyboard`库管理全局热键
   - 支持多线程操作防止界面阻塞
   - 配置文件：`hotstrings_cn.json`

3. **ERP自动化流程**
   ```python
   # 示例：ERP补发操作
   def erp_common_action_1(window_name, add_product=False, warehouse='cz', remark='补发'):
       app = WinGUI(window_name)
       app.click(coordinates['today_button'])
       app.input_text(coordinates['remark_field'], remark)
       ...
   ```

4. **智能剪切板**
   - 基于`pyperclip`和bat脚本实现
   - 文件路径自动转换机制
   - 支持多媒体文件快速插入

## 注意事项
⚠️ **重要配置**
1. 必须手动创建`config/window_config.ini`文件
2. `handles.json`和`temp/`目录会被程序自动生成
3. 图像识别依赖`image/`目录中的截图文件

🔧 **常见问题**
- 出现窗口定位错误时：
  1. 检查`coordinates.json`中的坐标配置
  2. 确认目标窗口名称与配置一致
  3. 调整`WINDOW_OPEN_MODE`参数

- 快捷键失效处理：
  1. 确认无其他程序占用热键
  2. 检查`hotstrings_cn.json`文件格式
  3. 重启程序重新加载配置

## 致谢与许可
### 项目继承
本项目基于以下开源项目开发：
- **[python-WinGUI](https://github.com/JoeyGambler/python-WinGUI)** 
  - 作者：Joey Gambler
  - 协议：MIT License
  - 主要贡献：提供窗口操作基础框架

### 开源协议
本项目遵循MIT开源协议，详见[LICENSE](LICENSE)文件。