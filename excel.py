# -*- coding:utf-8 -*-

import xlrd

data = xlrd.open_workbook('whole.xlsx')
table = data.sheets()[0]
row = table.nrows

with open('whole.csv', 'w+') as f:
    for i in range(row):
        s1 = table.row_values(i)[0]
        for c in ['(', ')', '（', '）', ' ', '\ue004', '\xa0', '\u2022', '\n', '\r']:
            s1 = s1.replace(c, '')
        s2 = table.row_values(i)[1]
        l2 = s2.split('|')
        del l2[0]
        s2 = ','.join(l2)
        for c in ['(', ')', '（', '）', ' ', '\ue004', '\xa0', '\u2022', '\n', '\r']:
            s2 = s2.replace(c, '')
        s2 = s2.rstrip(',C')
        print(i)
        if s1+ ',' + s2 != '':
            f.write(s1 + ',' + s2 + '\n')

