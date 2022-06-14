from time import sleep
import win32gui
import win32con
import pyautogui

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
    # 找到编辑受限按钮并点击
    pyautogui.click(632,1027)
    # 找到停止保护按钮坐标1675,228点击
    sleep(1)
    pyautogui.click(1675,326)
    sleep(0.5)
    pyautogui.write('123')
    pyautogui.press('enter')











