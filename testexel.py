# -*- coding: utf-8 -*-
import sys
import socket
import xlrd
from ipwhois import IPWhois
from PyQt5.QtWidgets import QWidget, QPushButton, QApplication
from PyQt5.QtCore import QCoreApplication

rb = xlrd.open_workbook('c:/Users/Пользователь/PycharmProjects/untitled1/tet.xls', formatting_info=True)
sheet = rb.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

r_count = sheet.nrows
print('Row count: ', r_count)

x = 0
y = 11
x1 = 1
y1 = 0
class test(IPWhois):
    def whoisid(self):
        while y <= r_count:
            try:
                ais = socket.gethostbyaddr((vals[y])[x])

                w = IPWhois((vals[y])[x])
                res = w.lookup_whois()
                resp = res.get('nets')
                resp1 = resp[0]
                resp2 = resp1.get('description')

                print('Host: ', ais[0], ' Owner: ', resp2)

                y = y + 1
            except socket.herror:
                w = IPWhois((vals[y])[x])
                res = w.lookup_whois()
                resp = res.get('nets')
                resp1 = resp[0]
                resp2 = resp1.get('description')

                print('Host: Not found! Owner: ', resp2)

                y = y + 1



