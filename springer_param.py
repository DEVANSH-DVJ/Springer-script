import sys
import xlrd

loc = ("Springer Ebooks.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
max = len(sheet.col(4))

if len(sys.argv) == 1:
    x, y = 2, max
elif len(sys.argv) == 2:
    try:
        x, y = int(sys.argv[1])+1, int(sys.argv[1])+2
    except:
        print('Enter only numbers as parameter')
        exit(0)
else:
    try:
        x, y = int(sys.argv[1])+1, int(sys.argv[2])+2
    except:
        print('Enter only numbers as parameter')
        exit(0)

if y > max or x < 2:
    print('Parameter out of bounds')
    exit(0)

for i in range(x, y):
    print(str(i-1))
