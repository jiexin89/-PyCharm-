import ctypes
import win32gui
import win32api
import win32con
import pywintypes
import logging
from pynput import mouse, keyboard
from concurrent.futures import ThreadPoolExecutor
import time

# 日志设置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 定义 ctypes 结构体用于发送鼠标事件
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

# 全局变量控制同步功能
sync_enabled = True

# 获取所有 Chrome 窗口句柄
def enumerate_chrome_windows():
    windows = []

    def window_callback(hwnd, _):
        class_name = win32gui.GetClassName(hwnd)
        window_title = win32gui.GetWindowText(hwnd)

        if class_name == 'Chrome_WidgetWin_1' and 'Google Chrome' in window_title:
            windows.append(hwnd)

    win32gui.EnumWindows(window_callback, None)
    logging.info(f"Found Chrome windows: {windows}")
    return windows

# 设置窗口大小为 639x520
def set_window_size(hwnd, width=639, height=520):
    if win32gui.IsWindow(hwnd):
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, width, height, win32con.SWP_NOMOVE)
            logging.info(f"Set window size for hwnd {hwnd} to {width}x{height}")
        except pywintypes.error as e:
            logging.error(f"Error in SetWindowPos for hwnd {hwnd}: {e}")

# 激活窗口，减少延迟
def activate_window(hwnd):
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.01)  # 将延迟缩短到 0.01 秒
        logging.info(f"Activated hwnd {hwnd}")
    except Exception as e:
        logging.error(f"Error activating hwnd {hwnd}: {e}")

# 同步鼠标点击到多个窗口
def sync_mouse_click(x, y, button='left', hwnds=None, skip_activate_main_hwnd=None):
    if not sync_enabled or not hwnds:
        return

    button_down_msg = win32con.WM_LBUTTONDOWN if button == 'left' else win32con.WM_RBUTTONDOWN
    button_up_msg = win32con.WM_LBUTTONUP if button == 'left' else win32con.WM_RBUTTONUP

    with ThreadPoolExecutor(max_workers=len(hwnds)) as executor:  # 增加并发线程数量
        for hwnd in hwnds:
            if hwnd == skip_activate_main_hwnd:
                continue
            executor.submit(send_mouse_event, hwnd, x, y, button_down_msg, button_up_msg)

def send_mouse_event(hwnd, x, y, button_down, button_up):
    try:
        activate_window(hwnd)

        lparam = win32api.MAKELONG(int(x), int(y))

        win32gui.PostMessage(hwnd, button_down, 0, lparam)
        win32gui.PostMessage(hwnd, button_up, 0, lparam)
        logging.info(f"Mouse click sent to hwnd {hwnd} at ({x}, {y})")
    except Exception as e:
        logging.error(f"Error sending mouse click to hwnd {hwnd}: {e}")

# 键盘按键映射函数
def map_key_to_vk(key):
    try:
        if isinstance(key, keyboard.Key):
            if key == keyboard.Key.space:
                return win32con.VK_SPACE
            elif key == keyboard.Key.enter:
                return win32con.VK_RETURN
            elif key == keyboard.Key.backspace:
                return win32con.VK_BACK
            elif key == keyboard.Key.tab:
                return win32con.VK_TAB
        return key.vk if isinstance(key, keyboard.KeyCode) else None
    except AttributeError:
        return None

# 同步键盘按键到多个窗口
def sync_key_press(keys, hwnds=None, skip_activate_main_hwnd=None):
    if not sync_enabled or not hwnds:
        return

    with ThreadPoolExecutor(max_workers=len(hwnds)) as executor:  # 增加并发线程数量
        for hwnd in hwnds:
            if hwnd == skip_activate_main_hwnd:
                continue
            executor.submit(send_key_event, hwnd, keys)

def send_key_event(hwnd, keys):
    try:
        activate_window(hwnd)

        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, keys, 0)
        logging.info(f"Key press '{keys}' sent to hwnd {hwnd}")
    except Exception as e:
        logging.error(f"Error sending key press to hwnd {hwnd}: {e}")

# 鼠标点击事件处理函数
def on_click(x, y, button, pressed):
    if pressed:
        hwnds = enumerate_chrome_windows()
        main_hwnd = win32gui.GetForegroundWindow()
        sync_mouse_click(x, y, button=button.name, hwnds=hwnds, skip_activate_main_hwnd=main_hwnd)

# 键盘按键事件处理函数
def on_press(key):
    global sync_enabled

    if key == keyboard.Key.f8:  # 切换群控同步
        sync_enabled = not sync_enabled
        logging.info(f"Sync {'enabled' if sync_enabled else 'disabled'}")
    elif key == keyboard.Key.f9:  # 重新排列窗口
        hwnds = enumerate_chrome_windows()
        rearrange_windows(hwnds)
    else:
        vk_code = map_key_to_vk(key)
        if vk_code:
            hwnds = enumerate_chrome_windows()
            main_hwnd = win32gui.GetForegroundWindow()
            sync_key_press(vk_code, hwnds=hwnds, skip_activate_main_hwnd=main_hwnd)

# 重新排列窗口
def rearrange_windows(hwnds, width=639, height=504):
    screen_width = win32api.GetSystemMetrics(0)
    screen_height = win32api.GetSystemMetrics(1)

    cols = screen_width // width
    rows = screen_height // height

    for i, hwnd in enumerate(hwnds):
        row = i // cols
        col = i % cols
        x = col * width
        y = row * height

        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, width, height, 0)
            logging.info(f"Rearranged hwnd {hwnd} to ({x}, {y}, {width}, {height})")
        except pywintypes.error as e:
            logging.error(f"Error rearranging hwnd {hwnd}: {e}")

# 启动鼠标和键盘监听器
def start_listener():
    with ThreadPoolExecutor() as executor:
        mouse_listener = mouse.Listener(on_click=on_click)
        keyboard_listener = keyboard.Listener(on_press=on_press)
        mouse_listener.start()
        keyboard_listener.start()

        mouse_listener.join()
        keyboard_listener.join()

if __name__ == "__main__":
    hwnds = enumerate_chrome_windows()
    for hwnd in hwnds:
        set_window_size(hwnd)
    start_listener()
