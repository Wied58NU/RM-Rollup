#!/anaconda3/bin/python

import xlrd

workbook = xlrd.open_workbook('/Users/jeffreywiedemann/Desktop/Resource_Planning/jeff_for_python.xlsx')

worksheet = workbook.sheet_by_name('Summary')

emp = worksheet.cell(1, 0).value
emp_group = worksheet.cell(0, 0).value
emp_work_weeks = worksheet.cell(16, 2).value
emp_project_hours = worksheet.cell(39, 2).value

print (emp)
print (emp_group)
print (emp_work_weeks)
print (emp_project_hours)


