# -*- coding:utf-8 -*-
__author__ = 'Ayaneru'

import win32con
from win32gui import *
import time
import csv
import pythoncom
import pyHook

LEI = "WindowsForms10.Window.8.app.0.378734a"
CSVLOCAL = ''

window_hwnd = FindWindow(LEI, None)
#print hex(window_hwnd)
window1_hwnd = FindWindowEx(window_hwnd, None, LEI, None)
#print hex(window1_hwnd)
window2_hwnd = FindWindowEx(window_hwnd, window1_hwnd, LEI, None)
#print hex(window2_hwnd)
for i in range(4):
    window1_hwnd = FindWindowEx(window1_hwnd, None, LEI, None)
    #print hex(window1_hwnd)

windowtab_hwnd = FindWindowEx(window1_hwnd, None, LEI, None)
windowcha_hwnd = FindWindowEx(window1_hwnd, windowtab_hwnd, LEI, None)

window5_hwnd =FindWindowEx(windowtab_hwnd,None,LEI,None)
window5_hwnd =FindWindowEx(window5_hwnd,None,LEI,None)

window11_hwnd = FindWindowEx(window5_hwnd, None, LEI, None)
#print hex(window11_hwnd)
window12_hwnd = FindWindowEx(window5_hwnd, window11_hwnd, LEI, None)
#print hex(window12_hwnd)
print_hwnd = FindWindowEx(window12_hwnd, None, LEI, None)
#print hex(print_hwnd)

window21_hwnd = FindWindowEx(window2_hwnd, None, LEI, None)
#print hex(window21_hwnd)
window22_hwnd = FindWindowEx(window2_hwnd, window21_hwnd, LEI, None)
#print hex(window22_hwnd)
window211_hwnd = FindWindowEx(window21_hwnd, None, LEI, None)
#print hex(window211_hwnd)
window212_hwnd = FindWindowEx(window21_hwnd, window211_hwnd, LEI, None)
#print hex(window212_hwnd)
window212_hwnd = FindWindowEx(window212_hwnd, None, LEI, None)
#print hex(window212_hwnd)
location_hwnd = FindWindowEx(window212_hwnd, None, LEI, None)
#print hex(location_hwnd)
for j in range(5):
    location_hwnd = FindWindowEx(window212_hwnd, location_hwnd, LEI, None)
#    print hex(location_hwnd)

window221_hwnd = FindWindowEx(window22_hwnd, None, LEI, None)
#print hex(window221_hwnd)
ID_hwnd = FindWindowEx(window221_hwnd, None, 'WindowsForms10.EDIT.app.0.378734a', None)
#print hex(ID_hwnd)

tID_hwnd = FindWindowEx(windowcha_hwnd, None, 'WindowsForms10.EDIT.app.0.378734a', None)
tName_hwnd = FindWindowEx(windowcha_hwnd, tID_hwnd, 'WindowsForms10.EDIT.app.0.378734a', None)
tAdd_hwnd = FindWindowEx(windowcha_hwnd, None, LEI, None)
for k in range(3):
    tAdd_hwnd = FindWindowEx(windowcha_hwnd, tAdd_hwnd, LEI, None)

print('ID '+ hex(ID_hwnd))
print('print '+hex(print_hwnd))
print('local '+hex(location_hwnd))
print('tName '+hex(tName_hwnd))
print('tID '+hex(tID_hwnd))
print('tAdd '+hex(tAdd_hwnd))

'''ID = '320924199107260272'
NAME = u'王'
SendMessage(ID_hwnd, win32con.WM_SETTEXT, None, ID)
time.sleep(0.05)
PostMessage(ID_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
time.sleep(0.2)

print GetWindowText(location_hwnd)
time.sleep(0.05)
SendMessage(print_hwnd, win32con.BM_CLICK, 0, -1)'''


SendMessage(print_hwnd, win32con.BM_CLICK, 0, -1)







pythoncom.PumpMessages()
