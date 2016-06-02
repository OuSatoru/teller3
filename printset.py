# -*- coding:utf-8 -*-

import sys
import os.path
import win32api
import win32clipboard as w
import win32con
import time
import pythoncom
import pyHook
import win32gui

# A const for time.sleep
SLEEP_TIME = .05
#A string to check foreground window
FORE_TITLE = 'moni'

#Giant dictonary to hold key name and VK value
VK_CODE = {'backspace': 0x08,
           'tab': 0x09,
           'clear': 0x0C,
           'enter': 0x0D,
           'shift': 0x10,
           'ctrl': 0x11,
           'alt': 0x12,
           'pause': 0x13,
           'caps_lock': 0x14,
           'esc': 0x1B,
           'spacebar': 0x20,
           'page_up': 0x21,
           'page_down': 0x22,
           'end': 0x23,
           'home': 0x24,
           'left_arrow': 0x25,
           'up_arrow': 0x26,
           'right_arrow': 0x27,
           'down_arrow': 0x28,
           'select': 0x29,
           'print': 0x2A,
           'execute': 0x2B,
           'print_screen': 0x2C,
           'ins': 0x2D,
           'del': 0x2E,
           'help': 0x2F,
           '0': 0x30,
           '1': 0x31,
           '2': 0x32,
           '3': 0x33,
           '4': 0x34,
           '5': 0x35,
           '6': 0x36,
           '7': 0x37,
           '8': 0x38,
           '9': 0x39,
           'a': 0x41,
           'b': 0x42,
           'c': 0x43,
           'd': 0x44,
           'e': 0x45,
           'f': 0x46,
           'g': 0x47,
           'h': 0x48,
           'i': 0x49,
           'j': 0x4A,
           'k': 0x4B,
           'l': 0x4C,
           'm': 0x4D,
           'n': 0x4E,
           'o': 0x4F,
           'p': 0x50,
           'q': 0x51,
           'r': 0x52,
           's': 0x53,
           't': 0x54,
           'u': 0x55,
           'v': 0x56,
           'w': 0x57,
           'x': 0x58,
           'y': 0x59,
           'z': 0x5A,
           'numpad_0': 0x60,
           'numpad_1': 0x61,
           'numpad_2': 0x62,
           'numpad_3': 0x63,
           'numpad_4': 0x64,
           'numpad_5': 0x65,
           'numpad_6': 0x66,
           'numpad_7': 0x67,
           'numpad_8': 0x68,
           'numpad_9': 0x69,
           'multiply_key': 0x6A,
           'add_key': 0x6B,
           'separator_key': 0x6C,
           'subtract_key': 0x6D,
           'decimal_key': 0x6E,
           'divide_key': 0x6F,
           'F1': 0x70,
           'F2': 0x71,
           'F3': 0x72,
           'F4': 0x73,
           'F5': 0x74,
           'F6': 0x75,
           'F7': 0x76,
           'F8': 0x77,
           'F9': 0x78,
           'F10': 0x79,
           'F11': 0x7A,
           'F12': 0x7B,
           'F13': 0x7C,
           'F14': 0x7D,
           'F15': 0x7E,
           'F16': 0x7F,
           'F17': 0x80,
           'F18': 0x81,
           'F19': 0x82,
           'F20': 0x83,
           'F21': 0x84,
           'F22': 0x85,
           'F23': 0x86,
           'F24': 0x87,
           'num_lock': 0x90,
           'scroll_lock': 0x91,
           'left_shift': 0xA0,
           'right_shift ': 0xA1,
           'left_control': 0xA2,
           'right_control': 0xA3,
           'left_menu': 0xA4,
           'right_menu': 0xA5,
           'browser_back': 0xA6,
           'browser_forward': 0xA7,
           'browser_refresh': 0xA8,
           'browser_stop': 0xA9,
           'browser_search': 0xAA,
           'browser_favorites': 0xAB,
           'browser_start_and_home': 0xAC,
           'volume_mute': 0xAD,
           'volume_Down': 0xAE,
           'volume_up': 0xAF,
           'next_track': 0xB0,
           'previous_track': 0xB1,
           'stop_media': 0xB2,
           'play/pause_media': 0xB3,
           'start_mail': 0xB4,
           'select_media': 0xB5,
           'start_application_1': 0xB6,
           'start_application_2': 0xB7,
           'attn_key': 0xF6,
           'crsel_key': 0xF7,
           'exsel_key': 0xF8,
           'play_key': 0xFA,
           'zoom_key': 0xFB,
           'clear_key': 0xFE,
           '+': 0xBB,
           ',': 0xBC,
           '-': 0xBD,
           '.': 0xBE,
           '/': 0xBF,
           '`': 0xC0,
           ';': 0xBA,
           '[': 0xDB,
           '\\': 0xDC,
           ']': 0xDD,
           "'": 0xDE}

'''no longer needed
locationsWB = {'一': 'g ',
               '二': 'fg ',
               '三': 'dg ',
               '四': 'lh ',
               '五': 'gg ',
               '六': 'uy ',
               '七': 'ag ',
               '八': 'wty ',
               '九': 'vt ',
               '十': 'fgh ',
               '红安': 'xa pv ',
               '新榆': 'usr swgj ',
               '联合': 'buwg '}'''


def press(*args):
    #'''
    #one press, one release.
    #accepts as many arguments as you want. e.g. press('left_arrow', 'a','b').
    #'''
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, 0, 0)
        time.sleep(SLEEP_TIME)
        win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)


def pressAndHold(*args):
    #'''
    #press and hold. Do NOT release.
    #accepts as many arguments as you want.
    #e.g. pressAndHold('left_arrow', 'a','b').
    #'''
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, 0, 0)
        time.sleep(SLEEP_TIME)


def pressHoldRelease(*args):
    #'''
    #press and hold passed in strings. Once held, release
    #accepts as many arguments as you want.
    #e.g. pressAndHold('left_arrow', 'a','b').

    #this is useful for issuing shortcut command or shift commands.
    #e.g. pressHoldRelease('ctrl', 'alt', 'del'), pressHoldRelease('shift','a')
    #'''
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, 0, 0)
        time.sleep(SLEEP_TIME)
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(.1)


def release(*args):
    #'''
    #release depressed keys
    #accepts as many arguments as you want.
    #e.g. release('left_arrow', 'a','b').
    #'''
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)


def typer(string=None, *args):
    ## time.sleep(4)
    for i in string:
        if i == ' ':
            win32api.keybd_event(VK_CODE['spacebar'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['spacebar'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '!':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['1'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['1'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '@':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['2'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['2'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '{':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['['], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['['], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '?':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['/'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['/'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == ':':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE[';'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE[';'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '"':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['\''], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['\''], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '}':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE[']'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE[']'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '#':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['3'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['3'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '$':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['4'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['4'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '%':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['5'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['5'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '^':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['6'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['6'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '&':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['7'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['7'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '*':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['8'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['8'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '(':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['9'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['9'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == ')':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['0'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['0'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '_':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['-'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['-'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '=':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['+'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['+'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '~':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['`'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['`'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '<':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE[','], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE[','], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == '>':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['.'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['.'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'A':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['a'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['a'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'B':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['b'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['b'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'C':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['c'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['c'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'D':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['d'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['d'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'E':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['e'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['e'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'F':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['f'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['f'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'G':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['g'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['g'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'H':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['h'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['h'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'I':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['i'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['i'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'J':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['j'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['j'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'K':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['k'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['k'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'L':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['l'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['l'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'M':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['m'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['m'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'N':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['n'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['n'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'O':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['o'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['o'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'P':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['p'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['p'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'Q':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['q'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['q'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'R':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['r'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['r'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'S':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['s'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['s'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'T':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['t'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['t'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'U':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['u'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['u'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'V':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['v'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['v'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'W':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['w'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['w'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'X':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['x'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['x'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'Y':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['y'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['y'], 0, win32con.KEYEVENTF_KEYUP, 0)

        elif i == 'Z':
            win32api.keybd_event(VK_CODE['left_shift'], 0, 0, 0)
            win32api.keybd_event(VK_CODE['z'], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE['left_shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(VK_CODE['z'], 0, win32con.KEYEVENTF_KEYUP, 0)

        else:
            win32api.keybd_event(VK_CODE[i], 0, 0, 0)
            time.sleep(SLEEP_TIME)
            win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)


def press_over_time(args, n):
    for i in range(n):
        press(args)
        time.sleep(SLEEP_TIME)


def type_whole(cString):
    pp = cString.split('\r\n')
    m = 0
    for i in pp:
        m += 1
        ss = i.split('\t')
        n = 0
        for j in ss:
            typer(j)
            time.sleep(SLEEP_TIME)
            n += 1
            if n != len(ss):
                press('right_arrow')
                time.sleep(SLEEP_TIME)
        if m != len(pp):
            press('down_arrow')
            time.sleep(SLEEP_TIME)
        press_over_time('left_arrow', len(ss) - 1)





def getText():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_TEXT)
    w.CloseClipboard()
    return d


def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_TEXT, aString)
    w.CloseClipboard()


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
                time.sleep(0.5)
                win32gui.PostMessage(ID_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
                time.sleep(0.03)
                

                
                

                
                if len(a_person.phone) == 11:
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
                        local = win32gui.GetWindowText(location_hwnd)
                        local = local.strip()
                        for one in local[0:18]:
                            win32gui.PostMessage(HWND, win32con.WM_CHAR, ord(one), 0)
                    
                    


                    press_over_time('enter',4)
                    typer('320919011351258147')
                    time.sleep(0.1)



                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                    time.sleep(0.01)
                    win32gui.SendMessage(print_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                    time.sleep(0.02)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
                    time.sleep(0.01)
                    win32gui.PostMessage(print_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, None)
                    #time.sleep(0.01)
                    #win32gui.SendMessage(print_hwnd, win32con.BM_CLICK, 0, -1)
                elif len(a_person.phone) == 8:
                    type_phone(a_person, HWND)
                else:
                    type_none(a_person, HWND)
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

    def __init__(self, comma_splitted_string=""):
        splitted_list = comma_splitted_string.split(',')
        self.idnum = splitted_list[4]
        self.phone = splitted_list[5]
        #self.location = splitted_list[6]
        #self.card = splitted_list[4].rstrip()
        self.birthdate = self.idnum[6:14]
        self.gender = eval(self.idnum[16]) % 2
        self.name = splitted_list[2].strip()

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
