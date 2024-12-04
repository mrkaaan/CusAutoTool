import pyperclip
import re
import time

def is_full_address(text):
    # 判断条件：
    # 1. 长度足够长（假设最小长度为20字符）
    # 2. 包含连续的数字（手机号码）
    # 3. 包含中文逗号用于分隔姓名、电话和地址
    # 4. 可选：包含"-"符号（虚拟号码）

    # 条件1: 长度足够长
    if len(text) < 20:
        return False, False
    
    # 条件2: 包含连续的数字（手机号码）
    has_consecutive_digits = re.search(r'\d{11}', text) is not None
    if not has_consecutive_digits:
        return False, False
    
    # 条件3: 包含中文逗号
    has_chinese_comma = '，' in text
    if not has_chinese_comma:
        return False, False
    
    # 条件4: 包含"-"符号（可选）
    has_dash = '-' in text
    
    return True, has_dash

def split_addr_info(addr_info):
    # 使用中文逗号分割地址信息为三部分
    parts = addr_info.split('，')
    if len(parts) == 3:
        return parts[0], parts[1], parts[2]  # 姓名, 电话, 地址
    else:
        return None, None, None

def replace_phone_in_address(address, phone):
    # 使用中文逗号分割并替换电话部分
    name, old_phone, address_part = split_addr_info(address)
    if name and old_phone and address_part:
        updated_addr_info = f"{name}，{phone}，{address_part}"
        return updated_addr_info
    return address

def listen_clipboard_changes():
    print("Waiting for two clipboard changes...")
    previous_content = pyperclip.paste()
    changed_contents = []
    
    while len(changed_contents) < 2:
        current_content = pyperclip.paste()
        
        if current_content != previous_content:
            print("Clipboard content has changed.")
            changed_contents.append(current_content)
            previous_content = current_content
            
            if len(changed_contents) == 2:  # 已经收集了两次变化的内容
                break
        
        time.sleep(1)  # 等待一段时间后再次检查
    
    # 判断并处理收集到的两个内容
    if len(changed_contents) == 2:
        addr_info = None
        phone_num = None
        
        for content in changed_contents:
            is_addr, _ = is_full_address(content)
            if is_addr:
                addr_info = content
            elif re.match(r'^1[3-9]\d{9}$', content):  # 检查是否是单独的电话号码
                phone_num = content
        
        if addr_info and phone_num:
            print("Processing the address with the new phone number...")
            updated_addr_info = replace_phone_in_address(addr_info, phone_num)
            pyperclip.copy(updated_addr_info)
            print(f"Updated Address Info: {updated_addr_info}")
        else:
            print("Failed to identify a full address and a separate phone number.")
    else:
        print("Failed to collect two clipboard changes.")


if __name__ == '__main__':
    # 调用函数开始监听
    listen_clipboard_changes()