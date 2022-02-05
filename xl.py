import xlwings as xw

wb = xw.Book("111.xlsx")

sht = wb.sheets[0]

print(sht.range('a1').value)
sht.range('a1').api.IndentLevel = 0