from xlrd import open_workbook


class convertXlsSheet():

    def __init__(self, ws):

        self.WS = ws

        self.PageInfo = {}

        self.QrowStart = 1
        self.QrowStop = 22
        self.GenderRow = 24
        self.AgeRow = 27
        self.SchoolRow = 30
        self.LangRow = 33
        self.NationalRow = 36
        self.HabitationRow = 39
        self.NeighborRow = 42

        self.Col_A = 0
        self.Col_B = 1
        self.Col_C = 2
        self.Col_D = 3
        self.Col_E = 4
        self.Col_F = 5
        self.Col_G = 6
        self.Col_H = 7

        self.Question_B = 6
        self.Question_C = 5
        self.Question_D = 4
        self.Question_E = 3
        self.Question_F = 2
        self.Question_G = 1
        self.Question_H = 0

        self.Gender_B = 1
        self.Gender_C = 2

        self.Age_B = 1
        self.Age_C = 1
        self.Age_D = 2
        self.Age_E = 3
        self.Age_F = 4

        self.School_B = 1
        self.School_C = 2
        self.School_D = 3
        self.School_E = 4
        self.School_F = 5

        self.Lang_B = 1
        self.Lang_C = 2
        self.Lang_D = 3
        self.Lang_E = 4

        self.National_B = 1
        self.National_C = 2

        self.Habitation_B = 1
        self.Habitation_C = 2
        self.Habitation_D = 3
        self.Habitation_E = 4

        self.Neighbor_B = 1
        self.Neighbor_C= 2


    def ReadSheet(self):
       # print self.WS.nrows
        result = []
        for nrow in range (self.WS.nrows):
           # print nrow

            if nrow in range(self.QrowStart,self.QrowStop):
               # self.Question(self.WS.row(nrow))
             #   result.append(self.Question(nrow))
                self.PageInfo[nrow]=self.Question(nrow)
            #self.PageInfo['Question']= result
          #  print 'nrow',nrow

            if nrow == self.GenderRow:
                #self.PageInfo['Gender']=self.Gender(nrow)
                self.PageInfo[nrow]=self.Gender(nrow)

            if nrow == self.AgeRow:
                #self.PageInfo['Age']=self.Age(nrow)
                self.PageInfo[nrow]=self.Age(nrow)

            if nrow == self.SchoolRow:
                #self.PageInfo['School']=self.School(nrow)
                self.PageInfo[nrow]=self.School(nrow)

            if nrow == self.LangRow:
                #self.PageInfo['Lang']=self.Lang(nrow)
                self.PageInfo[nrow]=self.Lang(nrow)

            if nrow == self.NationalRow:
               # self.PageInfo['National']=self.National(nrow)
                self.PageInfo[nrow]=self.National(nrow)

        #    if nrow == self.NationalRow:
               # self.PageInfo['National']=self.National(nrow)
         #       self.PageInfo[nrow]=self.National(nrow)

            if nrow == self.HabitationRow:
                #self.PageInfo['Habitation']=self.Habitation(nrow)
                self.PageInfo[nrow]=self.Habitation(nrow)

            if nrow == self.NeighborRow:
                #self.PageInfo['Neighbor']=self.Neighbor(nrow)
                self.PageInfo[nrow]=self.Neighbor(nrow)


      #  print 'Page', self.PageInfo

        return self.PageInfo

    def Question(self,Row):
        print 'Question'
        valueResult = 'na'
        for Column in range(1,8):
        #    print 'Q',(CurrentRow,Column),self.WS.cell(CurrentRow,Column).value
         #   print 'Column',Column
            valueCell = self.WS.cell(Row,Column).value
          #  print 'CellValue',valueCell
            if valueCell == 1.0:
                #    print 'SET'
                if Column == self.Col_B:
                    valueResult = self.Question_B
                elif Column == self.Col_C:
                    valueResult = self.Question_C
                elif Column == self.Col_D:
                    valueResult = self.Question_D
                elif Column == self.Col_E:
                    valueResult = self.Question_E
                elif Column == self.Col_F:
                    valueResult = self.Question_F
                elif Column == self.Col_G:
                    valueResult = self.Question_G
                elif Column == self.Col_H:
                    valueResult = self.Question_H
                else:
                    valueResult = False

        return valueResult
    
    def Gender(self,Row):
        print 'Gender'
        valueResult = 'na'
        for Column in range(1,3):
        #    print 'Q',(CurrentRow,Column),self.WS.cell(CurrentRow,Column).value
          #  print 'Column',Column
            cellValue = self.WS.cell(Row,Column).value
           # print 'ceellvalue', cellValue,Column
            if cellValue == 1.0:
               # print '#'
                if Column == self.Col_B:
                    valueResult = self.Gender_B
                elif Column == self.Col_C:
                    valueResult = self.Gender_C
                else:
                    valueResult = False
          #  else:
           #     print'not found'

        return valueResult

    def Age(self,Row):
        print 'Age'
        valueResult = 'na'
        for Column in range(1,6):
            cellValue = self.WS.cell(Row,Column).value
       #     print 'ceellvalue', cellValue,Column
            if cellValue == 1.0:
               # print '#'
                if Column == self.Col_C:
                    valueResult = self.Age_C
                elif Column == self.Col_D:
                    valueResult = self.Age_D
                elif Column == self.Col_E:
                    valueResult = self.Age_E
                elif Column == self.Col_F:
                    valueResult = self.Age_F
                else:
                    valueResult = False

        return valueResult

    def School(self,Row):
        print 'School'
        valueResult = 'na'
        for Column in range(1,6):
            cellValue = self.WS.cell(Row,Column).value
            if cellValue == 1.0:
                if Column == self.Col_B:
                    valueResult = self.School_B
                elif Column == self.Col_C:
                    valueResult = self.School_C
                elif Column == self.Col_D:
                    valueResult = self.School_D
                elif Column == self.Col_E:
                    valueResult = self.School_E
                elif Column == self.Col_F:
                    valueResult = self.School_F
                else:
                    valueResult = False

        return valueResult

    def Lang(self,Row):
        print 'Language'
        valueResult = 'na'
        for Column in range(1,5):
            cellValue = self.WS.cell(Row,Column).value
            if cellValue == 1.0:
                if Column == self.Col_B:
                    valueResult = self.Lang_B
                elif Column == self.Col_C:
                    valueResult = self.Lang_C
                elif Column == self.Col_D:
                    valueResult = self.Lang_D
                elif Column == self.Col_E:
                    valueResult =self.Lang_E
                else:
                    valueResult = False

        return valueResult

    def National(self,Row):
        print 'Nationality'
        valueResult = 'na'
        for Column in range(1,3):
            cellValue = self.WS.cell(Row,Column).value
            if cellValue == 1.0:
                if Column == self.Col_B:
                    valueResult = self.National_B
                elif Column == self.Col_C:
                    valueResult = self.National_C
                else:
                    valueResult = False

        return valueResult

    def Habitation(self,Row):
        print 'Habitation'
        valueResult = 'na'
        for Column in range(1,5):
            cellValue = self.WS.cell(Row,Column).value
            if cellValue == 1.0:
                if Column == self.Col_B:
                    valueResult = self.Habitation_B
                elif Column == self.Col_C:
                    valueResult = self.Habitation_C
                elif Column == self.Col_D:
                    valueResult = self.Habitation_D
                elif Column == self.Col_E:
                    valueResult =self.Habitation_E
                else:
                    valueResult = False

        return valueResult

    def Neighbor(self,Row):
        print 'Neighbor'
        valueResult = 'na'
        for Column in range(1,3):
            cellValue = self.WS.cell(Row,Column).value
            if cellValue == 1.0:
                if Column == self.Col_B:
                    valueResult = self.Neighbor_B
                elif Column == self.Col_C:
                    valueResult = self.Neighbor_C
                else:
                    valueResult = False

        return valueResult



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



