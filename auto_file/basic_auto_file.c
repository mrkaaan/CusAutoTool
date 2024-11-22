#include <stdio.h>
#include <stdlib.h>
#include <windows.h>

int main() {
    SetConsoleOutputCP(CP_UTF8);  // 设置输出代码页为UTF-8
    SetConsoleCP(CP_UTF8);        // 设置输入代码页为UTF-8

    printf("程序将在后台静默运行5秒，记录这期间按下的键盘内容...\n");

    Sleep(5000);  // 等待5秒

    printf("在这5秒内，你按下的按键如下:\n");

    system("pause");  // 等待用户输入，防止窗口关闭

    return 0;
}