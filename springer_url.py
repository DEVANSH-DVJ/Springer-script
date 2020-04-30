import xlrd
import requests
from bs4 import BeautifulSoup
import wget

loc = ("Springer Ebooks.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

for i in range(115, len(sheet.col(4))):
    # print(sheet.cell_value(i,4))
    page = requests.get(sheet.cell_value(i,4))
    s = BeautifulSoup(page.content, 'html.parser')
    x = s.find_all('a', attrs={'title':'Download this book in PDF format'})
    if len(x) == 0:
        continue
    y = x[0].get_attribute_list('href')
    z = 'https://link.springer.com' + y[0]
    print(str(i) + '. ' + z)
