#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>

#define CHECK_INTERVAL 500 // 检查间隔时间（毫秒）
#define MAX_CLIPBOARD_SIZE 1024

char bat_file_path[MAX_PATH];
char hotstrings_json[MAX_PATH];
char previous_clipboard_content[MAX_CLIPBOARD_SIZE];

typedef struct Hotstring {
    char *hotstring;
    char *file_path;
} Hotstring;

Hotstring hotstrings[100]; // 假设最多有100个热字符串
int hotstring_count = 0;

BOOL LoadConfig(const char *filename) {
    FILE *fp = fopen(filename, "r");
    if (!fp) return FALSE;

    fscanf(fp, "%s", bat_file_path);
    fclose(fp);

    return TRUE;
}

BOOL LoadHotstrings(const char *filename) {
    FILE *fp = fopen(filename, "r");
    if (!fp) return FALSE;

    char line[256];
    while (fgets(line, sizeof(line), fp)) {
        char *hotstring = strtok(line, ":");
        char *file_path = strtok(NULL, "\n");

        if (hotstring && file_path) {
            hotstrings[hotstring_count].hotstring = strdup(hotstring);
            hotstrings[hotstring_count].file_path = strdup(file_path);
            hotstring_count++;
        }
    }

    fclose(fp);
    return TRUE;
}

void ExecuteBat(const char *bat_file_path, const char *file_path) {
    char command[MAX_PATH + MAX_PATH];
    snprintf(command, sizeof(command), "\"%s\" \"%s\"", bat_file_path, file_path);
    system(command);
    printf("Executed %s\n", file_path);
}

LRESULT CALLBACK KeyboardHookProc(int nCode, WPARAM wParam, LPARAM lParam) {
    static BOOL ctrl_pressed = FALSE;
    static DWORD last_checked_time = 0;

    if (nCode >= 0) {
        KBDLLHOOKSTRUCT *pKeyBoard = (KBDLLHOOKSTRUCT *)lParam;

        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
            if (pKeyBoard->vkCode == VK_CONTROL) {
                ctrl_pressed = TRUE;
            } else if (pKeyBoard->vkCode == VK_SPACE && ctrl_pressed) {
                DWORD current_time = GetTickCount();
                if (current_time - last_checked_time > CHECK_INTERVAL) {
                    last_checked_time = current_time;

                    OpenClipboard(NULL);
                    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
                    if (hData) {
                        wchar_t *text = GlobalLock(hData);
                        if (text) {
                            char clipboard_text[MAX_CLIPBOARD_SIZE];
                            WideCharToMultiByte(CP_UTF8, 0, text, -1, clipboard_text, MAX_CLIPBOARD_SIZE, NULL, NULL);
                            GlobalUnlock(hData);

                            if (strcmp(clipboard_text, previous_clipboard_content) != 0) {
                                strcpy(previous_clipboard_content, clipboard_text);

                                for (int i = 0; i < hotstring_count; i++) {
                                    int len = strlen(hotstrings[i].hotstring);
                                    if (strlen(clipboard_text) >= len &&
                                        strcmp(clipboard_text + strlen(clipboard_text) - len, hotstrings[i].hotstring) == 0) {
                                        ExecuteBat(bat_file_path, hotstrings[i].file_path);
                                    }
                                }
                            }
                        }
                    }
                    CloseClipboard();
                }
            }
        } else if (wParam == WM_KEYUP || wParam == WM_SYSKEYUP) {
            if (pKeyBoard->vkCode == VK_CONTROL) {
                ctrl_pressed = FALSE;
            }
        }
    }

    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

int main() {
    if (!LoadConfig("config.ini")) {
        printf("Failed to load config.ini\n");
        return 1;
    }

    if (!LoadHotstrings("hotstrings.json")) {
        printf("Failed to load hotstrings.json\n");
        return 1;
    }

    HHOOK hook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProc, NULL, 0);
    if (!hook) {
        printf("SetWindowsHookEx failed: %lu\n", GetLastError());
        return 1;
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnhookWindowsHookEx(hook);

    for (int i = 0; i < hotstring_count; i++) {
        free(hotstrings[i].hotstring);
        free(hotstrings[i].file_path);
    }

    return 0;
}



