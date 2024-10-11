import win32gui
import win32api
import win32con
import ctypes
from ctypes import wintypes
from pynput import mouse, keyboard
from concurrent.futures import ThreadPoolExecutor
import pywintypes
import time
import win32clipboard


# 定义 ctypes 结构体用于发送鼠标事件
class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        class _MOUSEINPUT(ctypes.Structure):
            _fields_ = [("dx", ctypes.c_long),
                        ("dy", ctypes.c_long),
                        ("mouseData", ctypes.c_ulong),
                        ("dwFlags", ctypes.c_ulong),
                        ("time", ctypes.c_ulong),
                        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

        _fields_ = [("mi", _MOUSEINPUT)]

    _fields_ = [("type", ctypes.c_ulong), ("input_union", _INPUT)]


# 设置窗口大小为 639x504
def set_window_size(hwnd):
    width, height = 639, 504
    if win32gui.IsWindow(hwnd):
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, width, height, win32con.SWP_NOMOVE)
        except pywintypes.error as e:
            print(f"Error in SetWindowPos for hwnd {hwnd}: {e}")
    else:
        print(f"Invalid window handle: {hwnd}")


# 获取所有 Chrome 窗口句柄
def enumerate_chrome_windows():
    windows = []

    def window_callback(hwnd, windows):
        if 'Chrome' in win32gui.GetWindowText(hwnd):
            set_window_size(hwnd)
            windows.append(hwnd)

    win32gui.EnumWindows(window_callback, windows)
    return windows


# 提取公共代码以减少重复
def send_click_message(hwnd, message, x, y):
    lparam = win32api.MAKELONG(x, y)
    try:
        win32gui.SendMessage(hwnd, message, 0, lparam)
    except pywtypes.error as e:
        print(f"Error sending message to hwnd {hwnd}: {e}")


def send_mouse_click(hwnd, x, y, button='left', double_click=False):
    try:
        # 获取窗口客户区坐标
        rect = win32gui.GetClientRect(hwnd)
        client_x = x + rect[0]  # 相对于客户区的 x
        client_y = y + rect[1]  # 相对于客户区的 y

        if button == 'left':
            send_click_message(hwnd, win32con.WM_LBUTTONDOWN, client_x, client_y)
            time.sleep(0.001)  # 减少响应时间
            send_click_message(hwnd, win32con.WM_LBUTTONUP, client_x, client_y)
            if double_click:
                time.sleep(0.001)
                send_click_message(hwnd, win32con.WM_LBUTTONDBLCLK, client_x, client_y)
                send_click_message(hwnd, win32con.WM_LBUTTONUP, client_x, client_y)
        elif button == 'right':
            send_click_message(hwnd, win32con.WM_RBUTTONDOWN, client_x, client_y)
            time.sleep(0.001)  # 减少响应时间
            send_click_message(hwnd, win32con.WM_RBUTTONUP, client_x, client_y)
            if double_click:
                time.sleep(0.001)
                send_click_message(hwnd, win32con.WM_RBUTTONDBLCLK, client_x, client_y)
                send_click_message(hwnd, win32con.WM_RBUTTONUP, client_x, client_y)
    except pywintypes.error as e:
        print(f"Error sending mouse click to hwnd {hwnd}: {e}")


# 发送键盘按键消息，支持组合键
def send_key_press(hwnd, keys):
    try:
        if isinstance(keys, list):
            for key in keys:
                win32gui.SendMessageTimeout(hwnd, win32con.WM_KEYDOWN, key, 0, win32con.SMTO_ABORTIFHUNG, 100)
            for key in reversed(keys):
                win32gui.SendMessageTimeout(hwnd, win32con.WM_KEYUP, key, 0, win32con.SMTO_ABORTIFHUNG, 100)
        else:
            win32gui.SendMessageTimeout(hwnd, win32con.WM_KEYDOWN, keys, 0, win32con.SMTO_ABORTIFHUNG, 100)
            win32gui.SendMessageTimeout(hwnd, win32con.WM_KEYUP, keys, 0, win32con.SMTO_ABORTIFHUNG, 100)
    except pywintypes.error as e:
        print(f"Error sending key press to hwnd {hwnd}: {e}")


# 发送中文字符通过剪贴板
def send_chinese_chars_via_clipboard(hwnds, chinese_chars):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(chinese_chars)
    win32clipboard.CloseClipboard()

    for hwnd in hwnds:
        send_key_press(hwnd, [win32con.VK_CONTROL, ord('V')])  # 发送 Ctrl + V


# 键盘按键映射函数
def map_key_to_vk(key):
    try:
        if hasattr(key, 'char') and key.char is not None:
            char = key.char
            if 'a' <= char <= 'z':  # 小写字母
                return ord(char.upper())  # 返回大写字母的虚拟键码
            elif 'A' <= char <= 'Z':  # 大写字母
                return ord(char)
            elif '0' <= char <= '9':  # 数字
                return ord(char)
        else:
            return {
                keyboard.Key.enter: win32con.VK_RETURN,
                keyboard.Key.space: win32con.VK_SPACE,
                keyboard.Key.backspace: win32con.VK_BACK,
                keyboard.Key.shift: win32con.VK_SHIFT,
                keyboard.Key.ctrl_l: win32con.VK_CONTROL,
                keyboard.Key.alt_l: win32con.VK_MENU,
                keyboard.Key.tab: win32con.VK_TAB,
                keyboard.Key.esc: win32con.VK_ESCAPE,
            }.get(key, 0)
    except Exception as e:
        print(f"Error mapping key {key}: {e}")
        return 0


# 键盘按键事件监听器优化
def on_press(key):
    try:
        if hasattr(key, 'char') and key.char:  # 处理可打印字符
            if key.char.isascii():  # 只处理ASCII字符
                vk_code = map_key_to_vk(key)
                if vk_code:
                    print(f"Key {key} pressed, vk_code: {vk_code}")
                    sync_perform_action('keypress', keys=vk_code)
            else:
                # 处理中文输入
                chinese_input = key.char  # 假设输入了中文字符
                hwnds = enumerate_chrome_windows()
                if hwnds:
                    send_chinese_chars_via_clipboard(hwnds, chinese_input)  # 发送中文字符到所有窗口
                else:
                    print("No Chrome windows found for Chinese input.")
    except Exception as e:
        print(f"Error processing key press: {e}")


# 鼠标点击事件监听器
def on_click(x, y, button, pressed):
    if pressed:
        print(f"Mouse clicked at ({x}, {y}) with {button}")
        sync_perform_action('click', x, y, button=button.name)


# 启动监听器
def start_listener():
    with ThreadPoolExecutor() as executor:
        mouse_listener = mouse.Listener(on_click=on_click)
        keyboard_listener = keyboard.Listener(on_press=on_press)
        mouse_listener.start()
        keyboard_listener.start()

        mouse_listener.join()
        keyboard_listener.join()


# 同步执行鼠标和键盘操作
def sync_perform_action(action, x=None, y=None, hwnds=None, button='left', keys=None):
    if hwnds is None:
        windows = enumerate_chrome_windows()
    else:
        windows = hwnds

    if not windows:
        print("No Chrome windows found.")
        return

    with ThreadPoolExecutor() as executor:
        futures = []
        for hwnd in windows:
            if action == 'click':
                futures.append(executor.submit(send_mouse_click, hwnd, x, y, button))
            elif action == 'keypress':
                futures.append(executor.submit(send_key_press, hwnd, keys))

        for future in futures:
            future.result()


if __name__ == "__main__":
    start_listener()


