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
                    typer('101013')
                    time.sleep(0.05)
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

                    print('start enter')
                    press('enter','enter')
                    typer('224200')
                    press('caps_lock')

                    

                    print('first is ' + local)
                    if local != '东台市安丰镇':
                        print('no')
                        local = win32gui.GetWindowText(location_hwnd)
                        local = local.strip()
                        print('2nd is ' + local)

                        print('here')
                        print(local)
                        for one in local:
                            win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
                    else:
                        local = '江苏省东台市安丰镇'
                        local = local.strip().encode('gb2312')
                        for one in local:
                            win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
                    
                    


                    press_over_time('enter',4)
                    typer('320919011351258147')
                    '''time.sleep(0.1)



                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                    time.sleep(0.01)
                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                    time.sleep(0.02)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
                    time.sleep(0.01)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, None)
                    #time.sleep(0.01)
                    #win32gui.SendMessage(print_hwnd, win32con.BM_CLICK, 0, -1)'''
                elif len(a_person.phone) == 8:
                    typer('101013')
                    time.sleep(0.05)
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

                    print('start enter')
                    press('enter','enter')
                    typer('224200')
                    press('caps_lock')



                    print('first is ' + local)
                    if local != '东台市安丰镇':
                        print('no')
                        local = win32gui.GetWindowText(location_hwnd)
                        local = local.strip()
                        print('2nd is ' + local)

                        print('here')
                        print(local)
                        for one in local:
                            win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
                    else:
                        local = '江苏省东台市安丰镇'
                        local = local.strip().encode('gb2312')
                        for one in local:
                            win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)




                    press_over_time('enter',4)
                    typer('320919011351258147')
                    '''time.sleep(0.1)



                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                    time.sleep(0.01)
                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                    time.sleep(0.02)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
                    time.sleep(0.01)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, None)'''
                else:
                    typer('101013')
                    time.sleep(0.05)
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
                    typer('86033393')
                    press('enter')
                    typer('000644')

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

                    print('start enter')
                    press('enter','enter')
                    typer('224200')
                    press('caps_lock')



                    print('first is ' + local)
                    if local != '东台市安丰镇':
                        print('no')
                        local = win32gui.GetWindowText(location_hwnd)
                        local = local.strip()
                        print('2nd is ' + local)

                        print('here')
                        print(local)
                        for one in local:
                            win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
                    else:
                        local = '江苏省东台市安丰镇'
                        local = local.strip().encode('gb2312')
                        for one in local:
                            win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)




                    press_over_time('enter',4)
                    typer('320919011351258147')
                    '''time.sleep(0.1)



                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                    time.sleep(0.01)
                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                    time.sleep(0.02)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
                    time.sleep(0.01)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, None)'''
                #press('enter','enter')
            #print a_person.location
                

        elif event.Key == 'Rshift':
            pass

        '''elif event.Key == 'Rshift':
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

    def __init__(self, comma_splitted_string=""):
        splitted_list = comma_splitted_string.split(',')
        self.idnum = splitted_list[0]
        self.phone = splitted_list[1].strip()
        #self.location = splitted_list[6]
        #self.card = splitted_list[4].rstrip()
        self.birthdate = self.idnum[6:14]
        self.gender = eval(self.idnum[16]) % 2
        #self.name = splitted_list[2].strip()

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
    press('2','enter','enter','1')
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
    print('end1')
    time.sleep(0.1)
    print('end2')
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
        win32gui.SendMessage(tAdd_hwnd, win32con.BM_CLICK, 0, -1)
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
    for one in local:
        win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)

    print('start enter')

    press_over_time('enter',4)
    typer('320919011351258147')
    win32gui.SendMessage(print_hwnd, win32con.BM_CLICK, 0, -1)

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
    print(a_person.location)
    for one in a_person.location:

        
        win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
    press_over_time('enter',4)
    typer('320919011351258147')

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
按右CTRL开始
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

csvfile = open('C:\\fe.csv')
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
