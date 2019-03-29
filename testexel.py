# -*- coding: utf-8 -*-
import xlrd
from ipwhois import IPWhois
import ipwhois
import time

start_time = time.time()
rb = xlrd.open_workbook('C:/Users/Пользователь/PycharmProjects/exwork/exel.xls',
                        formatting_info=True)
sheet = rb.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

r_count = sheet.nrows
print('Row count: ', r_count)

x = 0
y = 11
x1 = 1
y1 = 0

while int((vals[y])[2]) >= 1.0:

    try:
        w = IPWhois((vals[y])[x])
        res = w.lookup_rdap(retry_count=0)
        resp = res.get('objects')
        resp1 = list(resp.keys())
        resp2 = resp1[0]
        resp3 = resp.get(resp2)
        resp4 = resp3.get('contact')
        resp5 = resp4.get('name')

        print('Owner: ', resp5)

        y = y + 1

    except ipwhois.exceptions.IPDefinedError as e:

        print('Error: ', '%s' % e)

        y = y + 1

print("--- %s seconds ---" % (time.time() - start_time))

# TODO Прпробовать перепелить код на перл под питон
# TODO Разобраться с получением инфы о посещениях по ip
# TODO Разобраться с openpyxl