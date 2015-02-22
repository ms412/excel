import xlwt

workbook = xlwt.Workbook()

sheet = workbook.add_sheet('Summary')
sheet.write(0,0, 'foobar')

workbook.save('foobar.xls')