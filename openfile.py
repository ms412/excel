from xlrd import open_workbook

wb = open_workbook('Book2.xls')

question = []
results = {}

for s in wb.sheets():
    print 'Sheets',s.name
    for cur_row in range(s.nrows):
        if cur_row in range(1,22):
            print 'Current Row', cur_row
            print 'Row', s.row(cur_row)
            for cur_col in range(1,8):
                print 'Q',(cur_row,cur_col),s.cell(cur_row,cur_col).value
             #   print s.cell(cur_row,cur_col).value
                cell_val = s.cell(cur_row,cur_col).value
                if cell_val >= 1.0:
                #    print 'SET'
                    if cur_col == 1:
                        result_val = 6
                    elif cur_col == 2:
                        result_val = 5
                    elif cur_col == 3:
                        result_val = 4
                    elif cur_col == 4:
                        result_val = 3
                    elif cur_col == 5:
                        result_val = 2
                    elif cur_col == 6:
                        result_val = 1
                    elif cur_col == 7:
                        result_val = 0
                    else:
                        result_val = False
                    print result_val
            question.append(result_val)
        results['Q'] = question
          #  question.append(result_val)
       # print 'Question', question
        if cur_row == 24:
            for cur_col in range (1,2):
                print 'Gender',(cur_row,cur_col),s.cell(cur_row,cur_col).value
                cell_val = s.cell(cur_row,cur_col).value

                if cell_val >= 1.0:
                    if cur_col == 1:
                        result_val = 'm'
                    elif cur_col == 2:
                        result_val = 'w'
                    else:
                        result_val = False
            results[cur_row] = result_val
            print 'Gender', result_val
            question.append(result_val)

        if cur_row == 27:
            for cur_row in range (1,5):
                print 'Age',(cur_row,cur_col),s.cell(cur_row,cur_col).value
                cell_val = s.cell(cur_row,cur_col).value

                if cell_val >= 1.0:
                    if cur_col == 2:
                        result_val = 6
                    elif cur_col == 3:
                        result_val = 5
                    elif cur_col == 4:
                        result_val = 3
                    elif cur_col == 5:
                        result_val = 2
                    else:
                        result_val = False

            print 'Age', result_val
            results[cur_row] = result_val

        if cur_row == 30:
            for cur_row in range (1,5):
                print 'School',(cur_row,cur_col),s.cell(cur_row,cur_col).value
                cell_val = s.cell(cur_row,cur_col).value

                if cell_val >= 1.0:
                    if cur_col == 1:
                        result_val = 6
                    elif cur_col == 2:
                        result_val = 5
                    elif cur_col == 3:
                        result_val = 3
                    elif cur_col == 4:
                        result_val = 2
                    elif cur_col == 5:
                        result_val = 1
                    else:
                        result_val = False

            print 'School', result_val
            results['School'] = result_val

        if cur_row == 33:
            for cur_row in range (1,5):
                print 'Lang',(cur_row,cur_col),s.cell(cur_row,cur_col).value
                cell_val = s.cell(cur_row,cur_col).value

                if cell_val >= 1.0:
                    if cur_col == 1:
                        result_val = 6
                    elif cur_col == 2:
                        result_val = 5
                    elif cur_col == 3:
                        result_val = 3
                    elif cur_col == 4:
                        result_val = 2
                    else:
                        result_val = False

            print 'Lang', result_val
            results['Language'] = result_val
              #  else:
               #     print 'ZERO'
print 'Results',results

       # for col in range(s.ncols):
        #    print 'Cellname',(row,col),'-'
         #   print 'Cell',s.cell(row,col).value

