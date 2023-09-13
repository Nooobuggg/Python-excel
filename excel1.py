import openpyxl

wb =  openpyxl.load_workbook(r'C:/Users/Admin/Desktop/Python/自动化文件处理/excel/exp1.xlsx')

sheet = wb['Sheet1']

for i in range(1,9):
    print(i, sheet.cell(row=i,column=3).value) #cell()方法:传入数值row & column 如:(row = n，column = m) 可获取excel表格的第mn单元格