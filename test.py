# coding=utf-8
import typeclip
import time
from bs4 import BeautifulSoup
import win32gui
import re
import json
from ast import literal_eval

s = typeclip.get_text()
soup = BeautifulSoup(s, 'lxml')
l = soup.body.find(CHECKED value = )
print(l.get_text())