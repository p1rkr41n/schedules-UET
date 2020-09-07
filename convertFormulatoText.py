import os
import xlwings as xw
wbxl=xw.Book('out.xlsx')
for row in  range (2, 11):
    cellOf1 = "K" + str(row)
    cellOf2 = "L"+ str(row)
    cellOf3 = "M"+ str(row)
    wbxl.sheets['raw_copy'].range(cellOf1).value = wbxl.sheets['raw_copy'].range(cellOf1).value
    wbxl.sheets['raw_copy'].range(cellOf2).value = wbxl.sheets['raw_copy'].range(cellOf2).value
    wbxl.sheets['raw_copy'].range(cellOf3).value = wbxl.sheets['raw_copy'].range(cellOf3).value
wbxl.save()
os.system("C:\Windows\System32\\taskkill.exe /IM EXCEL.EXE /F")