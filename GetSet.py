# coding=utf-8
__author__ = 'Ayaneru'

import sys
import os.path
import win32api
import win32clipboard as w
import win32con
import time
import pythoncom
import pyHook
import win32gui
import codecs

''' window_hwnd = FindWindow(None, "Form1")
button_hwnd = FindWindowEx(window_hwnd, None, 'ThunderRT6CommandButton', None)
text_hwnd = FindWindowEx(window_hwnd, None, 'ThunderRT6TextBox', None)
print window_hwnd
#print hex(window_hwnd)
print button_hwnd
#获取文本
#print GetWindowText(button_hwnd)
print text_hwnd

#设置文本
SendMessage(text_hwnd, win32con.WM_SETTEXT, None, '320924')
#回车
PostMessage(text_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
#PostMessage(text_hwnd, win32con.WM_KEYUP, None, win32con.VK_RETURN)
#按钮点击
#SendMessage(button_hwnd, win32con.BM_CLICK, 0, -1)'''

f = codecs.open('D:\\write.txt', 'w+', 'utf-8')
f.write('111\n')
f.write('干干\n')
f.close()