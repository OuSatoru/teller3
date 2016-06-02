# coding=utf-8
from typeclip import *
import time
from bs4 import BeautifulSoup
import win32gui

print('博客'.encode('gbk') in win32gui.GetWindowText(int('0x00020B28', 16)))

