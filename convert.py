from readfile import convertXlsSheet
from xlrd import open_workbook
import xlwt

if __name__ == "__main__":



    wb = open_workbook('Book3.xls')
    saveDict = {}

  #  print wb

    sheet_names = wb.sheet_names()


    for sheet in sheet_names:
        print 'Current Sheet', sheet
        xlsSheet = wb.sheet_by_name(sheet)
        xls =convertXlsSheet(xlsSheet)
        saveDict[sheet]=xls.ReadSheet()
        print 'SaveDict', saveDict

    sumWb = xlwt.Workbook()
    sheet = sumWb.add_sheet('Summary')

    row = 0
    column = 1

    for key, value in saveDict.iteritems():
        print 'Dict1',key,value
        row = row + 1
        column = 1
        sheet.write(row,0,key)
        for key1,value1 in value.iteritems():
            sheet.write(row,column,value1)
            column = column + 1





    sumWb.save('foobar.xls')