import auto_operation as ao     # 自动操作模块
import utils as ut      # 工具模块
import utils_clipboard_changes as uc   # 剪切板工具模块
import auto_copy_clipboard_latest as ac   # 自动识别剪切板并替换为图片或视频模块
import organize_table_window as tb_win
import notification_reissue_window as nr_win
from config import setup_arguments, setup_pyautogui, setup_logging

def main():
    # ---------- 快捷键启动 -------------
    hotkey_actions = [
        # {'key': 'alt+shift+e', 'func': ut.open_sof, 'args': ['旺店通ERP',265632,1,'CoolWindow']},
        {'key': 'alt+shift+e', 'func': ut.open_sof, 'args': ['旺店通ERP'], 'use_thread':False}, # 打开制定软件
        
        # {'key': 'ctrl+shift+q', 'func': ao.run_once_remarks_by_qianniu, 'args': [window_name, True, True, 2]}, # 添加备注 并 取消标记 不需点击备注按钮
        {'key': 'ctrl+shift+w', 'func': ao.run_once_remarks_by_qianniu, 'args': [window_name, False, True, 2]}, # 添加备注 并 取消标记 需点击备注按钮

        {'key': 'f2', 'func': ac.on_press_clipboard}, # 自动识别输入框并替换剪切板内容为图片或视频文件
        {'key': 'ctrl+space', 'func': ac.on_press_clipboard}, # 自动识别输入框并替换剪切板内容为图片或视频文件
        {'key': 'ctrl+shift+space', 'func': ac.clear_clipboard, 'use_thread':False}, # 清空剪切板

        {'key': 'ctrl+shift+x', 'func': ut.update_clipboard_express_company, 'use_thread':False}, # 默认剪切板为物流单号 拼凑为指定格式
        {'key': 'f3', 'func': ac.clear_clipboard, 'use_thread':False}, # 清空剪切板
        
        {'key': 'ctrl+t+num 1', 'func': tb_win.call_create_window, 'use_thread':False}, # 打开窗口，用于整理表格
        {'key': 'ctrl+t+num 2', 'func': nr_win.call_create_window, 'use_thread':False}, # 打开窗口，用于通知补发
        {'key': 'ctrl+t+num 3', 'func': nr_win.notic_last_data, 'use_thread':False}, # 使用上次数据直接通知补发

        {'key': 'ctrl+shift+l', 'func': ao.handle_auto_send_price_link, 'args': [window_name, 1]}, # 发送差价链接
        {'key': 'ctrl+shift+g', 'func': ac.on_press_clipboard, 'args':['hp']}, # 自动识别输入框并替换剪切板内容为图片或视频文件 发送好评截图
        {'key': 'ctrl+shift+d', 'func': uc.listen_clipboard_changes, 'use_thread':False}, # 拼接完整地址 默认复制一次收货人信息一次电话
        {'key': 'ctrl+shift+a', 'func': ao.run_once_copy_username_by_qianniu, 'args': [window_name]}, # 复制用户名

        {'key': 'ctrl+b+num 1', 'func': ao.erp_select_today, 'args': ['旺店通ERP']}, # erp 选择今天
        {'key': 'ctrl+b+num 2', 'func': ao.erp_clear_product, 'args': ['旺店通ERP']}, # erp 清空产品
        {'key': 'ctrl+b+num 3', 'func': ao.erp_input_remarks, 'args': ['旺店通ERP', '补发']}, # erp 输入备注 "补发"
        {'key': 'ctrl+b+num 4', 'func': ao.erp_input_remarks, 'args': ['旺店通ERP', '补发金属转接头']}, # erp 输入备注 "补发金属转接头"
        {'key': 'ctrl+b+num 5', 'func': ao.erp_add_product_notes, 'args': ['旺店通ERP', 'cz']}, # erp 添加备注产品 潮州
        {'key': 'ctrl+b+num 6', 'func': ao.erp_add_product_notes, 'args': ['旺店通ERP', 'sz']}, # erp 添加备注产品 深圳

        {'key': 'ctrl+b+num 7', 'func': ao.erp_common_action_1, 'args': ['旺店通ERP']}, # erp 常用操作1 不选择仓库 不添加产品 选择今天日期 清空商品

        {'key': 'ctrl+b+num 8', 'func': ao.erp_common_action_1, 'args': ['旺店通ERP', False, 'cz', '补发']}, # erp 常用操作1 潮州仓 选择今天日期 清空商品 选择仓库 输入备注
        {'key': 'ctrl+b+num 9', 'func': ao.erp_common_action_1, 'args': ['旺店通ERP', False, 'sz', '补发']}, # erp 常用操作1 深圳仓 选择今天日期 清空商品 选择仓库 输入备注
        {'key': 'ctrl+b+num /', 'func': ao.erp_common_action_1, 'args': ['旺店通ERP', True, 'cz', '补发']}, # erp 常用操作1 潮州仓 选择今天日期 清空商品 选择仓库 添加商品备注 输入备注
        {'key': 'ctrl+b+num *', 'func': ao.erp_common_action_1, 'args': ['旺店通ERP', True, 'sz', '补发']}, # erp 常用操作1 深圳仓 选择今天日期 清空商品 选择仓库 添加商品备注 输入备注

        {'key': 'ctrl+r+num 1', 'func': ao.erp_common_action_2, 'args': ['旺店通ERP', 'sz', ['内23'], '补发金属转接头']}, # erp 常用操作2 转接头23
        {'key': 'ctrl+r+num 2', 'func': ao.erp_common_action_2, 'args': ['旺店通ERP', 'sz', ['内24'], '补发金属转接头']}, # erp 常用操作2 转接头23
        {'key': 'ctrl+r+num 6', 'func': ao.erp_common_action_2, 'args': ['旺店通ERP', 'sz', ['内23', '内24'], '补发金属转接头']}, # erp 常用操作2 转接头23 24
        {'key': 'ctrl+r+num 4', 'func': ao.erp_common_action_2, 'args': ['旺店通ERP', 'sz', ['万能接头【细'], '补发铜嘴万能接头']}, # erp 常用操作2 转接头23 24
        {'key': 'ctrl+r+num 5', 'func': ao.erp_common_action_3, 'args': ['旺店通ERP', 'sz', '补发金属转接头']}, # erp 常用操作2 其他转接头
        {'key': 'ctrl+r+num 7', 'func': ao.erp_common_action_2, 'args': ['旺店通ERP', 'sz', ['内23', '万能接头【细'], '补发金属转接头']}, # erp 常用操作2 其他转接头

    ]
    ut.auto_key(hotkey_actions)

    

def test():
    print("Running tests...")
    # ---------- 快捷键启动 -------------
    hotkey_actions = [
        {'key': 'ctrl+t+num 1', 'func': tb_win.call_create_window, 'use_thread':False}, # 打开窗口，用于整理表格
        {'key': 'ctrl+t+num 2', 'func': nr_win.call_create_window, 'use_thread':False} # 打开窗口，用于通知补发

        # {'key': 'alt+shift+e', 'func': ut.open_sof, 'args': ['旺店通ERP',265632,1,'CoolWindow']},
        # {'key': 'alt+shift+e', 'func': ut.open_sof, 'args': ['旺店通ERP'], 'use_thread':False}, # 打开制定软件
        
        # {'key': 'ctrl+shift+q', 'func': ao.run_once_remarks_by_qianniu, 'args': [window_name, True, 2]}, # 添加备注 并 取消标记
        # {'key': 'ctrl+shift+w', 'func': ao.run_once_unmark_by_qianniu, 'args': [window_name, 2]}, # 取消标记

        # {'key': 'f2', 'func': ac.on_press_clipboard}, # 自动识别输入框并替换剪切板内容为图片或视频文件
        # {'key': 'ctrl+space', 'func': ac.on_press_clipboard}, # 自动识别输入框并替换剪切板内容为图片或视频文件
        # {'key': 'ctrl+shift+space', 'func': ac.clear_clipboard, 'use_thread':False}, # 清空剪切板

        # {'key': 'ctrl+shift+x', 'func': ut.update_clipboard_express_company, 'use_thread':False}, # 默认剪切板为物流单号 拼凑为指定格式
        # {'key': 'f3', 'func': ac.clear_clipboard, 'use_thread':False}, # 清空剪切板
        
        # {'key': 'ctrl+shift+alt+z', 'func': tb_win.call_create_window, 'use_thread':False}, # 打开窗口，用于整理表格
        # {'key': 'ctrl+shift+alt+a', 'func': nr_win.call_create_window, 'use_thread':False}, # 打开窗口，用于通知补发
        # {'key': 'ctrl+shift+alt+q', 'func': nr_win.notic_last_data, 'use_thread':False}, # 使用上次数据直接通知补发

        # {'key': 'ctrl+shift+l', 'func': ao.handle_auto_send_price_link, 'args': [window_name, 1]}, # 发送差价链接
        # {'key': 'ctrl+shift+g', 'func': ac.on_press_clipboard, 'args':['hp']}, # 自动识别输入框并替换剪切板内容为图片或视频文件 发送好评截图
        # {'key': 'ctrl+shift+d', 'func': uc.listen_clipboard_changes, 'use_thread':False}, # 拼接完整地址 默认复制一次收货人信息一次电话
        # {'key': 'ctrl+shift+a', 'func': ao.run_once_copy_username_by_qianniu, 'args': [window_name]} # 复制用户名


    ]
    ut.auto_key(hotkey_actions)

    # def ceshi(keya):
    #     # 79 means Num 1
    #     if keyboard.is_pressed(79):
    #         print(keya)
    # # keyboard.add_hotkey("ctrl + Num 1",ceshi,args=('aaa',))
    # keyboard.add_hotkey("ctrl + num 1", lambda: print("Num 1 pressed!"))
    # keyboard.wait("ctrl + esc")


    # ---------- 快捷键启动 -------------

    # ---------- 通知补发单号 -------------
    '''
        表格必须经过格式化，有整理过后的原始单号以及物流单号
        不必担心循环被手动终止导致下次启动有问题，做了相应的防误触处理
        停止运行组合键为 ctrl+shift+e
    '''
    notification_reissue_parameter = {
        'window_name':window_name, # 窗口名称
        'table_name':'2024-11-27_230453_处理结果.xlsx',  # 表格名称
        'table_path': '', # 预留位置 当前逻辑比较畸形避免处可以用于出错后续优化
        'notic_shop_name':'潮洁居家', # 通知店铺名称
        'notic_mode':2,       # 通知模式  1：输入框通知 2：补发窗口按钮通知
        'show_logistics':False, # 是否显示物流公司 输入框通知模式下生效
        'logistics_mode':1,    # 物流模式 1自动识别物流公司 2手动输入物流公司
        'use_today':'2024-11-27', # 使用日期 如果为空则使用当天日期
        'test_mode':1, # 测试模式 0：不测试 若测试则输入测试数量
        'is_write':False, # 是否回写表格
    }
    # ao.notification_reissue(**notification_reissue_parameter)

    # ---------- 表格处理 -------------
    # tb.process_table('11月27日天猫补发单号.csv')
    # ao.erp_aciton_box()
    # tb_win.create_window(0)
if __name__ == "__main__":
    # 配置参数
    setup_pyautogui()
    setup_logging()
    window_name = r"千牛接待台"  # 窗口名称
    window_name_erp = r"旺店通ERP"
    input_args = setup_arguments()

    
    if input_args.test:
        test()
    else:
        main()