import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

import time
import random
import os
import ctypes
from ctypes import wintypes
import pyautogui
import win32gui
import win32con
import keyboard

# 定义常量
KEYEVENTF_KEYUP = 0x0002
VK_TAB = 0x09
VK_RETURN = 0x0D
VK_LMENU = 0xA0  # 左Alt键的虚拟键码


# 按键操作函数
def send_key_event(key_code: int, event_up: bool = False):
    """发送按键事件到系统"""
    extra_info = ctypes.c_ulong(0)
    flags = 0
    if event_up:
        flags = KEYEVENTF_KEYUP
    input_struct = InputI(ki=KeyBdInput(0, key_code, flags, 0, ctypes.pointer(extra_info)))
    input_event = Input(ctypes.c_ulong(1), input_struct)
    User32.SendInput(1, ctypes.byref(input_event), ctypes.sizeof(input_event))

def perform_alt_tab():
    """执行Alt+Tab切换窗口操作"""
    send_key_event(VK_LMENU, False)  # 按下左Alt
    send_key_event(VK_TAB, False)  # 按下Tab
    send_key_event(VK_TAB, True)  # 释放Tab
    send_key_event(VK_LMENU, True)  # 释放左Alt

def bring_window_to_foreground(window_title: str) -> bool:
    """尝试将指定标题的窗口置于前台"""
    hwnd = win32gui.FindWindow(None, window_title)
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        return True
    print(f"未找到窗口：'{window_title}'")
    return False

def find_image_center(image_path, similarity_threshold=0.8) -> tuple[int, int] | None:
    """在屏幕上查找图像中心位置"""
    center = pyautogui.locateCenterOnScreen(image_path, confidence=similarity_threshold)
    if center:
        return center
    print(f"未能定位图像，路径：'{image_path}'")
    return None

def simulate_transaction(barcode_file_path: str, min_delay: float = 1.0, max_delay: float = 600.0):
    """模拟交易流程（新增无限循环版本）"""

    def stop_simulation(event):
        nonlocal run_simulation
        run_simulation = False

    run_simulation = True  # 控制循环运行状态的变量

    keyboard.on_press_key('esc', stop_simulation)  # 监听Esc键

    while run_simulation:
        if not os.path.exists(barcode_file_path):
            print(f"文件不存在：'{barcode_file_path}'")
            break

        with open(barcode_file_path, 'r', encoding='utf-8') as file:
            barcodes = file.read().splitlines()
        if not barcodes:
            print("文件中无有效条形码数据。")
            break
        selected_barcode = random.choice(barcodes)
        print(f"选择随机条形码：{selected_barcode}")

        if not bring_window_to_foreground("陇之情·新商通"):
            continue

        input_field_center = find_image_center("C:/work/bi.png", 0.1)
        if input_field_center is None:
            continue

        pyautogui.click(*input_field_center)
        time.sleep(0.5)

        pyautogui.typewrite(selected_barcode)
        pyautogui.press('enter')
        time.sleep(0.5)  # 输入后适当延迟
        pyautogui.press('+')
        time.sleep(0.5)  # 按键间隔
        pyautogui.press('enter')
        time.sleep(0.5)  # 按键间隔
        pyautogui.press('delete')  # 每次模拟交易后清除因无法负库存交易屏幕余留的商品，防止影响下一次模拟交易

        delay_seconds = random.uniform(min_delay, max_delay)
        time.sleep(delay_seconds)

    keyboard.unhook_all()  # 当循环结束时取消所有键盘监听

if __name__ == "__main__":
    barcode_file_path = "C:/work/bc.txt"
    simulate_transaction(barcode_file_path)

    # 根据run_simulation的状态给出相应的结束提示
    print("模拟交易已因用户按下Esc键而停止。")

    keyboard.unhook_all()  # 结束时取消所有键盘监听

def select_barcode_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        barcode_file_path_entry.delete(0, tk.END)
        barcode_file_path_entry.insert(tk.END, file_path)

def start_simulation():
    """
    从GUI中启动模拟交易流程。
    """
    global barcode_file_path_entry
    barcode_file_path = barcode_file_path_entry.get()
    if not barcode_file_path:
        messagebox.showerror("错误", "未选择条形码文件")
        return

    simulate_transaction(barcode_file_path)

def create_gui():
    """
    创建图形用户界面。
    """
    global root, barcode_file_path_entry

    root = tk.Tk()
    root.title("陇之情·新商通伴侣")

    # 条形码文件路径输入框及选择按钮
    barcode_file_path_frame = tk.Frame(root)
    barcode_file_path_label = tk.Label(barcode_file_path_frame, text="条形码文件路径:")
    barcode_file_path_entry = tk.Entry(barcode_file_path_frame, width=50)
    select_file_button = tk.Button(barcode_file_path_frame, text="选择文件", command=select_barcode_file)
    barcode_file_path_label.pack(side=tk.LEFT)
    barcode_file_path_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
    select_file_button.pack(side=tk.LEFT)

    # 启动模拟交易按钮
    start_simulation_button = tk.Button(root, text="开始模拟交易", command=start_simulation)

    # 布局
    barcode_file_path_frame.pack(padx=10, pady=10)
    start_simulation_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()