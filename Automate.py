# -*- coding:utf-8 -*-
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


def onKeyboardEvent(event):

    if event.Key == 'Rcontrol':
        global a
        a += 1
        print(a)

        if a < len(csvs):
            local = ''
            a_person = people(csvs[a].rstrip())
            print(a_person.name + a_person.idnum + a_person.phone)
            time.sleep(0.5)
            win32gui.SendMessage(ID_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
            time.sleep(0.01)
            win32gui.SendMessage(ID_hwnd, win32con.WM_LBUTTONUP, 0, 0)
            win32gui.SendMessage(ID_hwnd, win32con.WM_SETTEXT, None, a_person.idnum)
            time.sleep(0.35)
            win32gui.PostMessage(ID_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
            time.sleep(0.03)






            if len(a_person.phone) == 11:


                if win32gui.FindWindow(None, '输入身份证号码')>0:
                    cant = win32gui.FindWindow(None, '输入身份证号码')
                    print('cant' + hex(cant))
                    time.sleep(0.01)
                    qd = win32gui.FindWindowEx(cant, None, 'Button', None)
                    time.sleep(0.03)
                    win32gui.SendMessage(qd, win32con.BM_CLICK, 0, -1)
                    time.sleep(0.02)
                    win32gui.SendMessage(tID_hwnd, win32con.WM_SETTEXT, None, a_person.idnum)
                    time.sleep(0.01)
                    win32gui.SendMessage(tName_hwnd, win32con.WM_SETTEXT, None, a_person.name)
                    time.sleep(0.02)

                    win32gui.SendMessage(tAdd_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                    time.sleep(0.01)
                    win32gui.SendMessage(tAdd_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                    time.sleep(0.02)
                    win32gui.PostMessage(tAdd_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
                    time.sleep(0.01)
                    win32gui.PostMessage(tAdd_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, None)

                    #win32gui.SendMessage(tAdd_hwnd, win32con.BM_CLICK, 0, -1)
                    time.sleep(0.01)
                    local = '东台市安丰镇'



                print('first is ' + local)
                if local != '东台市安丰镇':
                    print('no')
                    local = win32gui.GetWindowText(location_hwnd)
                    local = local.strip()
                    print('2nd is ' + local)

                    print('here')
                    print(local)

                else:
                    local = '江苏省东台市安丰镇'
                    local = local.strip().encode('gb2312')









                win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                time.sleep(0.01)
                win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                time.sleep(0.02)
                win32gui.PostMessage(print_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
                time.sleep(0.01)
                win32gui.PostMessage(print_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, None)
                #time.sleep(0.01)
                #win32gui.SendMessage(print_hwnd, win32con.BM_CLICK, 0, -1)

            #press('enter','enter')
        #print a_person.location


        '''elif event.Key == 'Rmenu':
            print '91041'
            print a
            if a < len(csvs):
                a_person = people(csvs[a])
                time.sleep(0.8)
                typer('91041A')
                #press('enter','caps_lock','a','0')
                typer(a_person.idnum)
                #press('0','enter','enter')
                #typer('21000101')
                if a_person.gender == 1:
                    press('1')
                    #typer(a_person.birthdate)
                    #press('m','r','enter')
                else:
                    press('2')
                    #typer(a_person.birthdate)
                    #press('m','s','enter')
                #press('2','enter','enter','1')
                typer(a_person.phone)
                press('enter','enter')
                typer(a_person.card)
                press_over_time('enter',4)
                typer('03504')

                press('enter','enter')

        elif event.Key == 'Rshift':
            print a
            print '91043'
            if a < len(csvs):
                a_person = people(csvs[a])
                time.sleep(0.8)
                typer('91043a')
                #press('enter','caps_lock','a','0')
                typer(a_person.idnum)
                #press('0','enter','enter')
                #typer('21000101')


                press_over_time('enter',11)
                typer(a_person.card)
                press_over_time('enter',5)

                typer('ynyn')'''
    return True


class people:
    idnum = ""
    birthdate = ""
    gender = 0
    #    married = 0
    phone = ""
    location = ""
    card = ""
    name = ""
    #timethru = ""

    def __init__(self, comma_splitted_string=""):
        splitted_list = comma_splitted_string.split(',')
        self.idnum = splitted_list[4]
        self.phone = splitted_list[5]
        #self.location = splitted_list[6]
        #self.card = splitted_list[4].rstrip()
        self.birthdate = self.idnum[6:14]
        self.gender = eval(self.idnum[16]) % 2
        self.name = splitted_list[2].strip()


print('''
右Ctrl-10101 右Alt-91041 右Shift-91043
有重名的手动输入，再下一个
使用英文输入法
 ---All rights reserved by Wang Cong---''')


LEI = "WindowsForms10.Window.8.app.0.378734a"


window_hwnd = win32gui.FindWindow(LEI, None)
#print hex(window_hwnd)
window1_hwnd = win32gui.FindWindowEx(window_hwnd, None, LEI, None)
#print hex(window1_hwnd)
window2_hwnd = win32gui.FindWindowEx(window_hwnd, window1_hwnd, LEI, None)
#print hex(window2_hwnd)
for i in range(4):
    window1_hwnd = win32gui.FindWindowEx(window1_hwnd, None, LEI, None)
    #print hex(window1_hwnd)

windowtab_hwnd = win32gui.FindWindowEx(window1_hwnd, None, LEI, None)
windowcha_hwnd = win32gui.FindWindowEx(window1_hwnd, windowtab_hwnd, LEI, None)

window5_hwnd =win32gui.FindWindowEx(windowtab_hwnd,None,LEI,None)
window5_hwnd =win32gui.FindWindowEx(window5_hwnd,None,LEI,None)

window11_hwnd = win32gui.FindWindowEx(window5_hwnd, None, LEI, None)
#print hex(window11_hwnd)
window12_hwnd = win32gui.FindWindowEx(window5_hwnd, window11_hwnd, LEI, None)
#print hex(window12_hwnd)
print_hwnd = win32gui.FindWindowEx(window12_hwnd, None, LEI, None)
#print hex(print_hwnd)

window21_hwnd = win32gui.FindWindowEx(window2_hwnd, None, LEI, None)
#print hex(window21_hwnd)
window22_hwnd = win32gui.FindWindowEx(window2_hwnd, window21_hwnd, LEI, None)
#print hex(window22_hwnd)
window211_hwnd = win32gui.FindWindowEx(window21_hwnd, None, LEI, None)
#print hex(window211_hwnd)
window212_hwnd = win32gui.FindWindowEx(window21_hwnd, window211_hwnd, LEI, None)
#print hex(window212_hwnd)
window212_hwnd = win32gui.FindWindowEx(window212_hwnd, None, LEI, None)
#print hex(window212_hwnd)
location_hwnd = win32gui.FindWindowEx(window212_hwnd, None, LEI, None)
#print hex(location_hwnd)
for j in range(5):
    location_hwnd = win32gui.FindWindowEx(window212_hwnd, location_hwnd, LEI, None)
#    print hex(location_hwnd)

thru_hwnd = win32gui.FindWindowEx(window212_hwnd, location_hwnd, LEI, None)
thru_hwnd = win32gui.FindWindowEx(window212_hwnd, thru_hwnd, LEI, None)
thru_hwnd = win32gui.FindWindowEx(window212_hwnd, thru_hwnd, LEI, None)


window221_hwnd = win32gui.FindWindowEx(window22_hwnd, None, LEI, None)
#print hex(window221_hwnd)
ID_hwnd = win32gui.FindWindowEx(window221_hwnd, None, 'WindowsForms10.EDIT.app.0.378734a', None)
#print hex(ID_hwnd)

tID_hwnd = win32gui.FindWindowEx(windowcha_hwnd, None, 'WindowsForms10.EDIT.app.0.378734a', None)
tName_hwnd = win32gui.FindWindowEx(windowcha_hwnd, tID_hwnd, 'WindowsForms10.EDIT.app.0.378734a', None)
tAdd_hwnd = win32gui.FindWindowEx(windowcha_hwnd, None, LEI, None)
for k in range(3):
    tAdd_hwnd = win32gui.FindWindowEx(windowcha_hwnd, tAdd_hwnd, LEI, None)

print('ID '+ hex(ID_hwnd))
print('print '+hex(print_hwnd))
print('local '+hex(location_hwnd))
print('tName '+hex(tName_hwnd))
print('tID '+hex(tID_hwnd))
print('tAdd '+hex(tAdd_hwnd))
print('thru '+hex(thru_hwnd))

csvfile = open('C:\\from.csv')
csvs = csvfile.readlines()
destfile = codecs.open('C:\\dest.txt', 'w+', 'utf-8')

for each_line in csvs:
    splitted = each_line.rstrip().split(',')
    ID = splitted[4]
    name = splitted[2].strip()
    phone = splitted[5]
    sex = eval(ID[16]) % 2

    time.sleep(0.5)
    win32gui.SendMessage(ID_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
    time.sleep(0.01)
    win32gui.SendMessage(ID_hwnd, win32con.WM_LBUTTONUP, 0, 0)
    win32gui.SendMessage(ID_hwnd, win32con.WM_SETTEXT, None, ID)
    time.sleep(0.35)
    win32gui.PostMessage(ID_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
    time.sleep(0.03)
    if win32gui.FindWindow(None, '输入身份证号码')>0:
        location = '东台市安丰镇'
        #thru = '21000101'
        cant = win32gui.FindWindow(None, '输入身份证号码')
        print('cant' + hex(cant))
        time.sleep(0.01)
        qd = win32gui.FindWindowEx(cant, None, 'Button', None)
        time.sleep(0.03)
        win32gui.SendMessage(qd, win32con.BM_CLICK, 0, -1)
        time.sleep(0.02)
    else:
        time.sleep(0.5)
        location = win32gui.GetWindowText(location_hwnd)
        #thrutemp = win32gui.GetWindowText(thru_hwnd)
        #thrutemp = '-'.split(thrutemp)
        #thru = ''.join(thrutemp)

    whole = (','.join(['A', ID, name, str(sex), '0156', phone, location, '224221']) + '\n').decode('gb2312')
    destfile.write(whole)

destfile.close()
csvfile.close()