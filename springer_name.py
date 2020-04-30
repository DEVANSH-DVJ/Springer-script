import xlrd
import requests
from bs4 import BeautifulSoup
import wget

loc = ("Springer Ebooks.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

for i in range(2, len(sheet.col(4))):
    x = str(sheet.cell_value(i,3));

    if x[-2:] == '.0':
        x = x[:-2]

    name = sheet.cell_value(i,1) + ' - ' + sheet.cell_value(i,2) + ' ~~ ' + x + '.pdf'
    for j in range(len(name)):
        if name[j] == '\n':
            name = name[:j] + ' ' + name[j+1:]
    print(str(i) + '. ' + name)
