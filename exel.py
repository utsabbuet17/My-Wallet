import xlrd

loc = (r"‪decmo.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

print(sheet.cell_value(1,0))

# For row 0 and column 0
sheet.cell_value(0, 0)

for i in range(sheet.ncols):
    print(sheet.cell_value(0, i))