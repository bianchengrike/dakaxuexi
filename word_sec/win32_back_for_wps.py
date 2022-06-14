from time import sleep
import win32gui
import win32con
import win32api
import os
import sys


def m_click(hwnd,x,y):           #后台可用 
    win32api.PostMessage(hwnd,win32con.WM_LBUTTONDOWN,0,win32api.MAKELONG(x,y))
    win32api.PostMessage(hwnd,win32con.WM_LBUTTONUP,0,win32api.MAKELONG(x,y))  

def kb_enter(hwnd):
       win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, 13, 0)
       win32api.SendMessage(hwnd, win32con.WM_KEYUP, 13, 0)

def kb_input(hwnd,text):
    for item in text:
        win32api.PostMessage(hwnd, win32con.WM_CHAR, item, 0)

def restart_self():
    python = sys.executable
    os.execl(python,python,*sys.argv)

if __name__ == '__main__':
    fhwnd = win32gui.FindWindow('OpusApp',None)
    print(fhwnd)
    win32gui.SetWindowPos(fhwnd,win32con.HWND_TOPMOST, 0,0,1280,1024, win32con.SWP_SHOWWINDOW)
    m_click(fhwnd,636,1011)
    sleep(1)
    m_click(fhwnd,1033,227)
    sleep(1)
    mbhwnd=win32gui.FindWindow('bosa_sdm_Microsoft Office Word 11.0','取消保护文档')
    sleep(1)
    kb_input(mbhwnd,b'VaBQbVFtVAikIV')
    sleep(1)
    kb_enter(mbhwnd)
