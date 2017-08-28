import xlrd
import xlsxwriter

#set path
input_workbook = 'C:/Users/agao/Desktop/ComputerNames.xlsx'
output_workbook ='C:/Users/agao/Desktop/practice.xlsx'

#read file
read_workbook = xlrd.open_workbook(input_workbook)

#list of sheets
#print book.sheet_names()

read_sheet = read_workbook.sheet_by_index(0)
print read_sheet.name

nrow = read_sheet.nrows
ncol = read_sheet.ncols

print "Row: " + str(nrow)
print "Column: " + str(ncol)

#print read_sheet.cell(rowx=0,colx=0)
#print read_sheet.row_values(0)
#print read_sheet.col_values(0)

#write file
write_workbook = xlsxwriter.Workbook(output_workbook)
write_sheet = write_workbook.add_worksheet("AG Practice")

for i in range (1,nrow):
    value = read_sheet.cell(i,0).value
    print value
    write_sheet.write(i,0,value+100)
write_workbook.close()
print "Done"

'''
for i in range(nrow):
    for j in range (ncol):
        value = read_sheet.cell(i,j).value
        print value
        write_sheet.write(j,i,value)
write_workbook.close()
print "Done"
'''

