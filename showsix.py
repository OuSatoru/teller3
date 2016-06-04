# -*- coding:utf-8 -*-

import typeclip
import win32con
import time
import pythoncom
import pyHook
import win32gui
from ast import literal_eval
import re


def onKeyboardEvent(event):
    HWND = win32gui.GetForegroundWindow()
    tt = win32gui.GetWindowText(HWND)
    if '考试试卷'.encode('gbk') in tt and event.Key == 'Rcontrol':
        workerw = win32gui.FindWindowEx(HWND, None, 'WorkerW', None)
        while win32gui.FindWindowEx(workerw, None, 'ReBarWindow32', None) == 0:
            workerw = win32gui.FindWindowEx(HWND, workerw, 'WorkerW', None)
        rebar = win32gui.FindWindowEx(workerw, None, 'ReBarWindow32', None)
        combo = win32gui.FindWindowEx(rebar, None, 'ComboBoxEx32', None)
        button = win32gui.FindWindowEx(combo, None, 'ToolbarWindow32', None)
        combo2 = win32gui.FindWindowEx(combo, None, 'ComboBox', None)
        link = win32gui.FindWindowEx(combo2, None, 'Edit', None)
        time.sleep(0.5)
        win32gui.SendMessage(link, win32con.WM_LBUTTONDOWN, 0, 0)
        time.sleep(0.01)
        win32gui.SendMessage(link, win32con.WM_LBUTTONUP, 0, 0)
        time.sleep(0.01)
        win32gui.SendMessage(link, win32con.WM_SETTEXT, None,
                             'javascript:if(clipboardData.setData("text",document.documentElement.outerHTML))')
        time.sleep(0.01)
        win32gui.SendMessage(button, win32con.WM_LBUTTONDOWN, 0, 0)
        time.sleep(0.01)
        win32gui.SendMessage(button, win32con.WM_LBUTTONUP, 0, 0)
        time.sleep(0.01)
        win32gui.PostMessage(button, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
        time.sleep(0.01)
        win32gui.PostMessage(button, win32con.WM_KEYUP, win32con.VK_RETURN, None)
        time.sleep(0.2)
        present = presentqestion(str(typeclip.get_text()))
        win32gui.SendMessage(link, win32con.WM_LBUTTONDOWN, 0, 0)
        time.sleep(0.01)
        win32gui.SendMessage(link, win32con.WM_LBUTTONUP, 0, 0)
        time.sleep(0.01)
        win32gui.SendMessage(link, win32con.WM_SETTEXT, None,
                             'javascript:alert("题目: %s\n选项: %s\n答案: %s)' %
                             (ANS[present-1][0], ANS[present-1][1], ANS[present-1][2]))
        time.sleep(0.01)
        win32gui.SendMessage(button, win32con.WM_LBUTTONDOWN, 0, 0)
        time.sleep(0.01)
        win32gui.SendMessage(button, win32con.WM_LBUTTONUP, 0, 0)
        time.sleep(0.01)
        win32gui.PostMessage(button, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
        time.sleep(0.01)
        win32gui.PostMessage(button, win32con.WM_KEYUP, win32con.VK_RETURN, None)


def parseanswer(s):
    if s.startswith('[{') or s.startswith('b\'[{'):
        ans = []
        s1 = s.replace('null', 'None')
        l = literal_eval(s1)
        for each in l:
            ans.append([each['topic']['content'], each['topic']['topicOption'], each['topic']['answer']])
        return ans

def presentqestion(html):
    f = re.findall(r'\[.*\d.*\].*CHECKED', html)
    return len(f)


ANS = parseanswer(typeclip.get_text())
hm = pyHook.HookManager()
hm.KeyDown = onKeyboardEvent
hm.HookKeyboard()

pythoncom.PumpMessages(1000)
