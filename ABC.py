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
    
    if tt.find(FORE_TITLE) > -1:
        if event.Key == 'Rcontrol':
            global a
            a += 1
            print(a)
            if a < len(csvs):
                a_person = people(csvs[a])
                time.sleep(0.8)
                if len(a_person.phone) == 11:
                    type_mobile(a_person, HWND)
                elif len(a_person.phone) == 8:
                    type_phone(a_person, HWND)
                else:
                    type_none(a_person, HWND)
                #press('enter','enter')
            #print a_person.location
                

        elif event.Key == 'Rmenu':
            print('91041')
            print(a)
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
            print(a)
            print('91043')
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
                
                typer('ynyn')
    return True


class people:
    idnum = ""
    birthdate = ""
    gender = 0
    #    married = 0
    phone = ""
    location = ""
    card = ""

    def __init__(self, comma_splitted_string=""):
        splitted_list = comma_splitted_string.split(',')
        self.idnum = splitted_list[0]
        self.phone = splitted_list[1]
        self.location = splitted_list[2]
        self.card = splitted_list[4].rstrip()
        self.birthdate = self.idnum[6:14]
        self.gender = eval(self.idnum[16]) % 2

def type_mobile(a_person, HWND):
    typer('101013')
    press('enter','caps_lock','a','0')
    typer(a_person.idnum)
    press('0','enter','enter')
    typer('21000101')
    if a_person.gender == 1:
        press('1')
        typer(a_person.birthdate)
        press('m','r','enter')
    else:
        press('2')
        typer(a_person.birthdate)
        press('m','s','enter')
    typer('2015601561')
    typer(a_person.phone)
    typer('000644')
    press('enter','enter')
    typer('224200')
    press('caps_lock')
    ###typer地址
    ###typer(locationsWB[a_person.locationA])
    ###typer(locationsWB[a_person.locationB])
    #print a_person.locationC
    ###typer(a_person.locationC.rstrip())
    print(a_person.location)
    for one in a_person.location:
        
        win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
    press_over_time('enter',4)
    typer('320919011351258147')

def type_phone(a_person, HWND):
    typer('101013')
    press('enter','caps_lock','a','0')
    typer(a_person.idnum)
    press('0','enter','enter')
    typer('21000101')
    if a_person.gender == 1:
        press('1')
        typer(a_person.birthdate)
        press('m','r','enter')
    else:
        press('2')
        typer(a_person.birthdate)
        press('m','s','enter')
    typer('2015601561')
    typer(a_person.phone)
    press('enter')
    typer('000644')
    press('enter','enter')
    typer('224200')
    press('caps_lock')
    ###typer地址
    ###typer(locationsWB[a_person.locationA])
    ###typer(locationsWB[a_person.locationB])
    #print a_person.locationC
    ###typer(a_person.locationC.rstrip())
    print(a_person.location)
    for one in a_person.location:

        win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
    press_over_time('enter',4)
    typer('320919011351258147')

    '''typer('101013')
    press('enter','caps_lock','a','0')
    typer(a_person.idnum)
    press('0','enter','enter')
    typer('21000101')
    if a_person.gender == 1:
        press('1')
        typer(a_person.birthdate)
        press('m','r','enter')
    else:
        press('2')
        typer(a_person.birthdate)
        press('m','s','enter')
    press('2','enter','enter','1','enter')
    #typer(a_person.phone)
    typer('000644')
    press('enter')
    typer('0515'+a_person.phone)
    press('enter','enter')
    typer('224200')
    press('caps_lock')
    ###typer地址
    ###typer(locationsWB[a_person.locationA])
    ###typer(locationsWB[a_person.locationB])
    #print a_person.locationC
    ###typer(a_person.locationC.rstrip())
    print a_person.location
    for one in a_person.location:
        
        win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
    press_over_time('enter',4)
    typer('320919011351258147')'''

def type_none(a_person, HWND):
    typer('101013')
    press('enter','caps_lock','a','0')
    typer(a_person.idnum)
    press('0','enter','enter')
    typer('21000101')
    if a_person.gender == 1:
        press('1')
        typer(a_person.birthdate)
        press('m','r','enter')
    else:
        press('2')
        typer(a_person.birthdate)
        press('m','s','enter')
    press('2','enter','enter','1')
    press('enter')
    typer('000644')
    press('enter','enter')
    typer('224200')
    press('caps_lock')
    ###typer地址
    ###typer(locationsWB[a_person.locationA])
    ###typer(locationsWB[a_person.locationB])
    #print a_person.locationC
    ###typer(a_person.locationC.rstrip())
    print(a_person.location)
    for one in a_person.location:
        
        win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
    press_over_time('enter',4)
    typer('320919011351258147')


print('''
右Ctrl-10101 右Alt-91041 右Shift-91043
有重名的手动输入，再下一个
使用英文输入法
 ---All rights reserved by Wang Cong---''')

csvfile = open('C:\\test.csv')
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
