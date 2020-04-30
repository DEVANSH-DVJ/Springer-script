import xlrd
import requests
from bs4 import BeautifulSoup

loc = ("Springer Ebooks.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

for i in range(2, len(sheet.col(4))):
    page = requests.get(sheet.cell_value(i,4))
    site = BeautifulSoup(page.content, 'html.parser')
    down_button = site.find_all('a', attrs={'title':'Download this book in PDF format'})
    if len(down_button) == 0:
        continue
    url = 'https://link.springer.com' + down_button[0].get_attribute_list('href')[0]
    print(str(i-1) + '. ' + url)
