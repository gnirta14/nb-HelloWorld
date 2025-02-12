import openpyxl as xl
wb = xl.load_workbook('openpyxl_transactions.xlsx')
print('**** wsn', wb.sheetnames)
sheet = wb['Sheet1']

c1 = sheet['a1']
c2 = sheet.cell(1,2)
smr = sheet.max_row

# update the price cell and write it to a new excel file.
for row in range(2, smr +1 ):
    cell = sheet.cell(row, 3)
    corrected_price = cell.value * 0.9
    corrected_price_cell = sheet.cell(row,4)
    corrected_price_cell.value = corrected_price


wb.save('openpyxl_transactions2.xlsx')
