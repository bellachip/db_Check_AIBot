import openpyxl

ex = openpyxl.load_workbook('db_check.xlsx')
sheet = ex["Sheet1"]
first_name_arr = []
last_name_arr = []

for i in range(3):
    if i >= 2:
        first_name = sheet['A' + str(i)].value
        first_name_arr.append(first_name)
        last_name = sheet['B' + str(i)].value
        last_name_arr.append(last_name)

c1 = sheet['C1']
c1.value = 'no'

ex.save('db_check.xlsx')
