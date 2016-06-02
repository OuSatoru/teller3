# -*- coding:utf-8 -*-

from typeclip import *
import win32con
import time
import pythoncom
import pyHook
import win32gui


def onKeyboardEvent(event):
    HWND = win32gui.GetForegroundWindow()
    tt = win32gui.GetWindowText(HWND)

    if tt.find(FORE_TITLE) > -1 and event.Key == 'Rcontrol':
        global a
        a += 1
        print(a)
        if a < len(csvs):
            splitted = csvs[a].strip().split(',')
            idnum = splitted[2]
            card = splitted[0]
            time.sleep(0.8)
            typer('71216')
            time.sleep(0.02)
            typer(idnum)
            time.sleep(0.02)
            typer('0')
            press('enter', 'enter')
            time.sleep(1.4)
            mouse_absolute(205, 250, 455, 250)
            time.sleep(0.02)
            mouse_rclick(420, 250)
            time.sleep(0.08)
            mouse_click(480, 268)
            time.sleep(0.01)
            guest = get_text()
            time.sleep(0.01)
            press_over_time('num_lock', 3)
            typer('71210')
            typer(guest)
            press('enter')
            typer('0')
            press('enter')
            typer('1')
            typer(card)
            press_over_time('enter', 4)
            typer('320919011351258147')

    return True


print('''
改关系
按右CTRL开始''')

with open('C:\\sb.csv') as csvfile:
    csvs = csvfile.readlines()

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
