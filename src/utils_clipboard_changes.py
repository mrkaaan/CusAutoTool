import pyperclip
import re
import time
import utils as ut

def is_full_address(text):
    # 判断条件：
    # 1. 长度足够长（假设最小长度为20字符）
    # 2. 包含连续的数字（手机号码）
    # 3. 包含中文逗号用于分隔姓名、电话和地址
    # 4. 可选：包含"-"符号（虚拟号码）

    # 条件1: 长度足够长
    if len(text) < 20:
        return False
    
    # 条件2: 包含连续的数字（手机号码）
    has_consecutive_digits = re.search(r'\d{11}', text) is not None
    if not has_consecutive_digits:
        return False
    
    # 条件3: 包含中文逗号
    has_chinese_comma = '，' in text
    if not has_chinese_comma:
        return False
    
    # 条件4: 包含"-"符号（可选）
    # has_dash = '-' in text
    
    return True

def is_phone_number(text):
    # 判断条件：
    # 1. 长度为11位
    # 2. 有且仅为数字

    # 条件1: 长度为11位
    if len(text)!= 11:
        return False
    
    # 条件2: 有且仅为数字
    if not text.isdigit():
        return False
    
    return True

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
    if name and old_phone and address_part and phone:
        updated_addr_info = f"{name}，{phone}，{address_part}"
        return updated_addr_info
    return None

def listen_clipboard_changes():
    print("Waiting for two clipboard changes...")
    ut.show_toast('提醒', '请复制完整的地址信息到剪贴板，然后再复制电话号码到剪贴板...')
    previous_content = pyperclip.paste()
    changed_contents = {}
    
    while True:
        current_content = pyperclip.paste()
        
        if current_content != previous_content:
            is_address = is_full_address(current_content)    
            is_phone = is_phone_number(current_content)
            
            if is_address:
                changed_contents['address'] = current_content
                print(f"Address: {current_content}")
                ut.show_toast('提醒', '已复制完整的地址信息到剪贴板')
            elif is_phone:
                changed_contents['phone'] = current_content
                print(f"Phone: {current_content}")
                ut.show_toast('提醒', '已复制电话号码到剪贴板')

            previous_content = current_content

        # 判断是否收集到了两次变化的内容
        if len(changed_contents) == 2:
            break
        
        time.sleep(0.2)
        
    
    # 判断并处理收集到的两个内容
    if len(changed_contents) == 2:
        try:
            print("Processing the address with the new phone number...")
            updated_addr_info = replace_phone_in_address(changed_contents['address'], changed_contents['phone'])
            if not updated_addr_info:
                print("Error: Failed to update the address info.")
                ut.show_toast('提醒', '更新地址信息失败 请重试')
                return
            pyperclip.copy(updated_addr_info)
            print(f"Updated Address Info: {updated_addr_info}")
            ut.show_toast('提醒', '已更新地址信息到剪贴板')
        except Exception as e:
            print(f"Error: {e}")


if __name__ == '__main__':
    # 调用函数开始监听
    listen_clipboard_changes()