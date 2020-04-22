# -*- coding: utf-8 -*-

import json
import requests
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

rb = xlrd.open_workbook('C:/Users/Пользователь/Desktop/Статистика.xls',
                        formatting_info=True)
maxsheets = rb.nsheets

print('Script is start working')
print('Sheets count: ', maxsheets)

while int(maxsheets) != p:
    sheet = rb.sheet_by_index(p)
    vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    r_count = sheet.nrows

    try:
        if int((vals[y])[2]) <= 50.01:
            p = p + 1
            y = 11
            x = 0

            print('*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*Sheet: ', p)
            continue

        w = IPWhois((vals[y])[x])

        res = w.lookup_rdap(retry_count=0)
        resp = res.get('objects')
        resp1 = res.get('asn_description')
        #print(resp1)

        ip = ((vals[y])[x])
        response = requests.get("https://freeapi.robtex.com/ipquery/" + ip)
        respo = json.loads(response.text)
        respo1 = respo['pas']
        respo2 = respo1[0]
        respo3 = respo2['o']
        print(respo3)

        y = y + 1

    except ipwhois.exceptions.IPDefinedError as e:
        print('Ошибка: IP не определен', '%s' % e)
        y = y + 1

    except ipwhois.exceptions.HTTPLookupError as er:
        print('Ошибка: поиск RDAP не удался', '%s' % er)
        y = y + 1

    except ipwhois.exceptions.ASNRegistryError as err:
        print('Ошибка: поиск ASN не удался', '%s' % err)
        y = y + 1

    except ValueError:
        print('Файл пустой')
        break
    except IndexError as erro:
        print('Название сайта получить не удалось')
        y = y + 1
    except KeyError as error:
        print('Название сайта получить не удалось')
        y = y + 1

print('Script is finish --- %s seconds ---' % (time.time() - start_time),)