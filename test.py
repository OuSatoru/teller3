# coding=utf-8
import typeclip
import time
from bs4 import BeautifulSoup
import win32gui
import re
import json
from ast import literal_eval


def parseanswer(s):
    if s.startswith('[{'):
        ans = []
        s1 = s.replace('null', 'None').lstrip('b')
        l = literal_eval(s1)
        print(l)
        print(type(l))
        print(type(l[0]))
        for each in l:
            ans.append([each['topic']['content'], each['topic']['topicOption'], each['topic']['answer']])
        return ans

print(typeclip.get_text())
ANS = parseanswer(typeclip.get_text().decode('gbk'))

print(ANS[2][0], ANS[2][1], ANS[2][2])

'''s = str(typeclip.get_text())
print(s)
f = re.findall(r'\[.*\d.*\].*CHECKED', s)
print(f)'''