import sys
import xlrd

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
    print(str(i-1))
