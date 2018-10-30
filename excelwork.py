# -*- coding: utf-8 -*-

import xlrd
import xlwt

originData = xlrd.open_workbook('origin.xlsx')
sheets = originData.sheets()
firstSheet = sheets[0]
secondSheet = sheets[1]

titles = firstSheet.row_values(1)
firstStudents = []
for i in range (2, firstSheet.nrows):
	firstStudents.append(firstSheet.row_values(i))

def sortSubject(col):
	title = titles[col]
	titles.insert(col+1, title + "rank")
	firstStudents.sort(key = lambda student : student[col], reverse = True)
	preScore = 0;
	preRank = 0;
	for i in range (len(firstStudents)):
		if firstStudents[i][col] == preScore:
			firstStudents[i].insert(col+1, preRank)
		else:
			preRank = i+1
			firstStudents[i].insert(col+1, i+1)
		preScore = firstStudents[i][col];
	return True
	
sortSubject(3)
sortSubject(5)
sortSubject(7)
sortSubject(9)
sortSubject(11)
sortSubject(13)
firstStudents.sort(key = lambda student : student[1])

output = xlwt.Workbook(encoding = 'utf-8')
outSheet = output.add_sheet('result')

for i in range (len(titles)):
	outSheet.write(0, i, titles[i])

#outSheet.write(0, 10, "总分")

for i in range (len(firstStudents)):
	for j in range (len(firstStudents[i])):
		outSheet.write(i+1, j, firstStudents[i][j])
#	sum = 0
#	for j in range (3, 9):
#		sum = sum + firstStudents[i][j]
#	outSheet.write(i+1, 9, sum)
		
output.save('result.xls')
print "done"