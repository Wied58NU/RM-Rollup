import xlrd

//Open work book
workbook = xlrd.open_workbook('/Users/jeffreywiedemann/Desktop/Resource_Planningjeff_for_python.xlsx')

worksheet = workbook.sheet_by_name('Summary')

emp = worksheet.cell(0, 1).value
emp_group = worksheet.cell(0, 0).value
emp_work_weeks = worksheet.cell(16, 2).value
emp_project_houes = worksheet.cell(39, 2).value
