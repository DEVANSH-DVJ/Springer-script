import sys
import wget
import xlrd

import requests
from bs4 import BeautifulSoup

loc = ("Springer Ebooks.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
max = len(sheet.col(4))

list = []

try:
    if len(sys.argv) == 1:
        x, y = 2, max
    elif sys.argv[1] == '-l':
        if len(sys.argv) == 2:
            print('Parameter missing')
            exit(0)
        for l in range(2, len(sys.argv)):
            val = int(sys.argv[l])
            if val > max-2 or val < 1:
                print('Parameter out of bounds', val)
                exit(0)
            list.append(val+1)
    elif len(sys.argv) == 2:
        x, y = int(sys.argv[1])+1, int(sys.argv[1])+2
    elif len(sys.argv) == 3:
        x, y = int(sys.argv[1])+1, int(sys.argv[2])+2
    else:
        print("'-l' is required as the first parameter for list")
        exit(0)

except:
    print('Enter only numbers as parameter')
    exit(0)

if list == []:
    if y > max or x < 2:
        print('Parameter out of bounds')
        exit(0)
    list = range(x, y)

for i in list:
    page = requests.get(sheet.cell_value(i,4))
    site = BeautifulSoup(page.content, 'html.parser')
    down_button = site.find_all('a', attrs={'title':'Download this book in PDF format'})
    if len(down_button) == 0:
        continue
    url = 'https://link.springer.com' + down_button[0].get_attribute_list('href')[0]
    print(str(i-1) + '. ' + url)

    ed = str(sheet.cell_value(i,3));
    if ed[-2:] == '.0':
        ed = ed[:-2]
    name = sheet.cell_value(i,1) + ' - ' + sheet.cell_value(i,2) + ' ~~ ' + ed + '.pdf'
    for j in range(len(name)):
        if name[j] == '\n':
            name = name[:j] + ' ' + name[j+1:]
    print(str(i-1) + '. ' + name)

    wget.download(url, '../' + name)
    print('\n\n')
