# -*- coding: utf-8 -*-
import xlrd
from ipwhois import IPWhois
import ipwhois
import time

p = 0
x = 0
y = 11
x1 = 1
y1 = 0

start_time = time.time()

rb = xlrd.open_workbook('C:/Users/Пользователь/PycharmProjects/exwork/exew.xls',
                        formatting_info=True)
maxsheets = rb.nsheets
print('Sheets count: ', maxsheets)



while int(maxsheets) >= p:
    sheet = rb.sheet_by_index(p)
    vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    r_count = sheet.nrows

    try:

        w = IPWhois((vals[y])[x])
        res = w.lookup_rdap(retry_count=0)
        resp = res.get('objects')
        resp1 = list(resp.keys())
        resp2 = resp1[0]
        resp3 = resp.get(resp2)
        resp4 = resp3.get('contact')
        resp5 = resp4.get('name')

        print(resp5)

        y = y + 1
        if int((vals[y])[2]) <= 1.0:

            p = p + 1
            y = 11
            x = 0
            print("--- %s seconds ---" % (time.time() - start_time), p)
            continue
        if p == int(maxsheets):
            break

    except ipwhois.exceptions.IPDefinedError as e:

        print('Error: ', '%s' % e)

        y = y + 1
        if int((vals[y])[2]) >= 1.0:
            p = p + 1
            print("--- %s seconds ---" % (time.time() - start_time), "Sheet: ", p)
            continue
        if p == int(maxsheets):
            break

print("--- %s seconds ---" % (time.time() - start_time), "Sheet: ", p)

# TODO Прпробовать перепелить код на перл под питон
# TODO Разобраться с получением инфы о посещениях по ip
# TODO Разобраться с openpyxl