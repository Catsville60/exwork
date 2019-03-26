# -*- coding: utf-8 -*-
import socket
import xlrd
from ipwhois import IPWhois
import ipwhois
import signal

import time
from pprint import pprint


rb = xlrd.open_workbook('C:/Users/Пользователь/PycharmProjects/untitled1/exel.xls',
                        formatting_info=True)
sheet = rb.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

r_count = sheet.nrows
print('Row count: ', r_count)

x = 0
y = 11
x1 = 1
y1 = 0
s = socket.socket()

while y <= r_count:
        if int((vals[y])[2]) > 1.0:
            try:
                try:
                    ais = socket.gethostbyaddr((vals[y])[x])

                    w = IPWhois((vals[y])[x])
                    res = w.lookup_rdap(retry_count=0)
                    resp = res.get('objects')
                    resp1 = list(resp.keys())
                    resp2 = resp1[0]
                    resp3 = resp.get(resp2)
                    resp4 = resp3.get('contact')
                    resp5 = resp4.get('name')

                    print('Host: ', ais[0], ' Owner: ', resp5)

                    y = y + 1

                except socket.herror:

                    w = IPWhois((vals[y])[x])
                    ress = w.lookup_rdap(retry_count=0)
                    respp = ress.get('objects')
                    respp1 = list(respp.keys())
                    respp2 = respp1[0]
                    respp3 = respp.get(respp2)
                    respp4 = respp3.get('contact')
                    respp5 = respp4.get('name')

                    print('Host: Not found! Owner: ', respp5)

                    y = y + 1

            except ipwhois.exceptions.IPDefinedError as e:

                print('%s' % e)

                y = y + 1
        else:
            break






