from time import sleep
import win32gui
import win32con
import pyautogui
from pyautogui import locateCenterOnScreen


# pyautogui.PAUSE = 2 默认等待时间
# print(pyautogui.size())
# 获取窗口句柄
hwnd = win32gui.FindWindow('OpusApp',None)
if hwnd == 0:
    pyautogui.alert(text='你尚未打开wps文件', title='错误提示')
else:
    # 窗口置于最前
    win32gui.SetForegroundWindow(hwnd)
    # 显示为最大化
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    sleep(0.5)
    # 通过图片寻找编辑受限按钮并点击
    p1=locateCenterOnScreen('./1.png')
    pyautogui.click(p1.x,p1.y)
    sleep(0.5)
    # 通过图片寻找停止保护按钮并点击
    p2=locateCenterOnScreen('./2.png')
    pyautogui.click(p2.x,p2.y)
    sleep(0.5)
    pyautogui.write('123')
    pyautogui.press('enter')
