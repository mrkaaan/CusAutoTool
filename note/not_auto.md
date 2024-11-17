
auto_file.py
根据条件修改剪切板的内容的函数。启动函数条件是判断是否按下空格，函数里面每次被执行都会读取剪切板，
为了优化加入了两次空格间隔判断、set、两次剪切板内容是否相同判断、判断按键从空格改为Ctrl+空格
造成这种延迟的原因 1监听键盘 2 判断是否输入Ctrl+空格 3读取剪切板
主要还是1和3所以 目前这个函数造成的电脑延迟还是会导致打字有卡顿
改用C语言尝试
安装环境 mingw-64
https://sourceforge.net/projects/mingw-w64/
在files这个界面，你打开Tollchains targetting Win64，再打开Personal Builds，点入mingw builds,点 8.1.0，threads posix，再选seh

解压 复制bin目录路径
添加到用户环境变量Path

安装vscode插件 c/c++

ctrl+shift+p 输入c/c++
设置编译器路径为刚才的bin路径的gcc
选择IntelliSense模式为 gcc-x64(legacy)

保持界面为.c文件
终端-配置任务 选择gcc 会生成一个task.json文件
继续选择.c文件
终端-运行生成任务 会生成一个exe文件
