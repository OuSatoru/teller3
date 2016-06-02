# coding=utf-8
from typeclip import *
import time
from bs4 import BeautifulSoup
import win32gui
import re
import xlrd

data = xlrd.open_workbook('whole.xlsx')
table = data.sheets()[0]
row = table.nrows

s = table.row_values(4170)

print(s)
print(repr(s))