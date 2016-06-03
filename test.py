# coding=utf-8
from typeclip import *
import time
from bs4 import BeautifulSoup
import win32gui
import re
import json
from ast import literal_eval

with open('json.txt') as f:
    s = f.read().replace('null', 'None')
    l = json.loads(s[0])
    print(l["topic"])
