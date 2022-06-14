import win32gui
import pyautogui 


# 获取窗口句柄
hwnd = win32gui.FindWindow('OpusApp',None)
print(hwnd)

# 显示窗口状态
# win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED) 
# 窗口置于最前
win32gui.SetForegroundWindow(hwnd)

# tid = win32gui.FindWindowEx(hwnd,None,'NetUIRibbonTab','审阅')
# print(tid)

# 打开界面，我工作中没有这个界面，不知道是什么原因
pyautogui.press('alt')
pyautogui.press('W')
pyautogui.press('E')

# 打开审阅
pyautogui.press('alt')
pyautogui.press('R')
# 打开限制编辑
pyautogui.press('P')
pyautogui.press('E')


# 实在找不到句柄，干脆用最古老的方法
pyautogui.press('tab',3)
pyautogui.press('enter')
pyautogui.write('123')
pyautogui.press('enter')
















