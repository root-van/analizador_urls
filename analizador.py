import json 
import requests
import time
from openpyxl import Workbook
from openpyxl.styles import Font

x = 1
f = open("urls_sospechosas.txt", "r")
book = Workbook()
apis = book.active
apis['B2'] = 'Urls'
apis['C2'] = 'Fecha de Analisis'
apis['D2'] = 'Total de Analisis'
apis['E2'] = 'Analisis positivos'
apis['F2'] = 'Clasificacion'
apis['B2'].font = Font(color='000000', bold=True)
apis['C2'].font = Font(color='000000', bold=True)
apis['D2'].font = Font(color='000000', bold=True)
apis['E2'].font = Font(color='000000', bold=True)
apis['F2'].font = Font(color='000000', bold=True)
i = 3
for linea in f:
    print(x)
    if x == 5:
        time.sleep(62.00)
        api_url = 'https://www.virustotal.com/vtapi/v2/url/report'
        params = dict(apikey='f3f87547a514d73af003c74f408a46e85ddf94abba6c9f695c8ce5aba66b7348',
        resource = linea, scan=0)
        response = requests.get(api_url, params=params)
        if response.status_code == 200:
          result=response.json()
          print(json.dumps(result, sort_keys=False, indent=4))
          apis['B'+str(i)] = str(linea)
          apis['C'+str(i)] = str(result['scan_date'])
          apis['D'+str(i)] = str(result['total'])
          apis['E'+str(i)] = str(result['positives'])
          if result['positives'] <= 3:
            apis['F'+str(i)] = 'Baja'
          elif result['positives'] >3 or result['positives'] <=10:
            apis['F'+str(i)] = 'Media'
          else:
            apis['F'+str(i)] = 'Alta'
        
    else:
        api_url = 'https://www.virustotal.com/vtapi/v2/url/report'
        params = dict(apikey='713089e0e43eba82067282e69b9158c8fc85cd67c40f5abc289b98d8fbac68a8',
        resource = linea, scan=0)
        response = requests.get(api_url, params=params)
        if response.status_code == 200:
          result=response.json()
          print(json.dumps(result, sort_keys=False, indent=4))
          apis['B'+str(i)] = str(linea)
          apis['C'+str(i)] = str(result['scan_date'])
          apis['D'+str(i)] = str(result['total'])
          apis['E'+str(i)] = str(result['positives'])
          if result['positives'] <= 3:
            apis['F'+str(i)] = 'Baja'
          elif result['positives'] >3 or result['positives'] <=10:
            apis['F'+str(i)] = 'Media'
          else:
            apis['F'+str(i)] = 'Alta'
    i = i + 1
    x = x + 1
book.save('reporte_analizador_urls.xlsx')
f.close()
