# -*- coding:utf-8 -*-

from typeclip import *
import win32con
import time
import pythoncom
import pyHook
import win32gui


def onKeyboardEvent(event):
    if event.Key == 'Rcontrol':
        while True:
            time.sleep(1)
            mouse_move(500,500)
            time.sleep(0.5)
            mouse_click(525,298)
            time.sleep(1.5)
            mouse_click(525,298)
            time.sleep(3)
            mouse_click(1030,292)
            time.sleep(1)
            mouse_click(1030,292)
            time.sleep(1)
            mouse_click(988,479)
            time.sleep(0.5)
            mouse_click(988,479)
            time.sleep(13)
            mouse_click(946,383)
            time.sleep(39.5)
            mouse_click(815,281)
            time.sleep(5)
            mouse_click(815,281)
            time.sleep(5.1)
            mouse_click(815,281)
            time.sleep(2)
            mouse_click(815,281)
            time.sleep(2)
            mouse_click(815,281)
            time.sleep(4)
            mouse_click(407,259)
            time.sleep(0.9)
            mouse_click(443,159)
            time.sleep(4)
            mouse_click(380,101)
            time.sleep(2.5)
    return True


print('''
try kancolle 1-5A''')



global a
a = -1

# 创建一个“钩子”管理对象
hm = pyHook.HookManager()

# 监听所有键盘事件
hm.KeyDown = onKeyboardEvent

hm.HookKeyboard()
# 一直监听，直到手动退出程序
pythoncom.PumpMessages(1000)






### aa.split('\r\n') \t
### resource@ https://gist.github.com/chriskiehl/2906125
