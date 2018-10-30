# -*- coding: utf-8 -*-

import xlrd
import xlwt

originData = xlrd.open_workbook('origin.xlsx')
sheets = originData.sheets()
firstSheet = sheets[0]
secondSheet = sheets[1]

firstStudents = []
for i in range (2, firstSheet.nrows):
	firstStudents.append(firstSheet.row_values(i))

firstStudents.sort(key = lambda student : student[3], reverse = True)

output = xlwt.Workbook(encoding = 'utf-8')
outSheet = output.add_sheet('result')
length = len(firstSheet.row_values(1))
for i in range (length):
	outSheet.write(0, i, firstSheet.row_values(1)[i])

outSheet.write(0, 9, "总分")

for i in range (len(firstStudents)):
	for j in range (len(firstStudents[i])):
		outSheet.write(i+1, j, firstStudents[i][j])
	sum = 0
	for j in range (3, 9):
		sum = sum + firstStudents[i][j]
	outSheet.write(i+1, 9, sum)
		
output.save('result.xls')
print "done"