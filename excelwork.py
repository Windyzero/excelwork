# -*- coding: utf-8 -*-

import xlrd
import xlwt
import Tkinter
import tkMessageBox
import ConfigParser
 
 
cf = ConfigParser.ConfigParser()
cf.read('config.ini')

levelStr = cf.get("config", "improveLevels")
improveLevels = levelStr.split(",")
improveLevels = map(int, improveLevels)
pointPerLevel = cf.getfloat("config", "pointPerLevel")
subjectsCount = cf.getint("config", "subjectsCount") 

originData = xlrd.open_workbook('origin.xlsx')
sheets = originData.sheets()
firstSheet = sheets[0]
secondSheet = sheets[1]

titles = firstSheet.row_values(1)
titles.append("总分")
firstStudents = []
secondStudents = []

for i in range (2, firstSheet.nrows):
	firstStudents.append(firstSheet.row_values(i))

for i in range (2, secondSheet.nrows):
	secondStudents.append(secondSheet.row_values(i))	
	
def sortSubject(list, col):
	list.sort(key = lambda student : student[col], reverse = True)
	preScore = 0;
	preRank = 0;
	for i in range (len(list)):
		if list[i][col] == 0:
			list[i].insert(col+1, 0)
		elif list[i][col] == preScore:
			list[i].insert(col+1, preRank)
		else:
			preRank = i+1
			list[i].insert(col+1, i+1)
		preScore = list[i][col];
	return True

def sortAllSubjects(list, startCol, count, step):
	for i in range(count):
		sortSubject(list, startCol + i * step)

def calculateSum(list, startCol, count, step):
	for i in range (len(list)):
		sum = 0
		for j in range (count):
			sum = sum + list[i][startCol + j * step]
		list[i].append(sum)
	
def getLevelByRank(rank):
	for i in range(len(improveLevels)):
		if rank <= improveLevels[i]:
			return i
	return len(improveLevels)
	
def calPointByTwoRank(small, big):
	maxRank = improveLevels[len(improveLevels) - 1]
	if big > maxRank:
		big = maxRank
	if small > maxRank:
		small = maxRank
	bigLevel = getLevelByRank(big)
	point = 0
	step = bigLevel
	while (step >= 0):
		if step > 0 and improveLevels[step - 1] > small:
			pointPerRank = pointPerLevel / (improveLevels[step] - improveLevels[step - 1])
#			print "%d, %d, pointPerRank: %f"%(improveLevels[step], improveLevels[step-1], pointPerRank)
			point = point + (big - improveLevels[step - 1]) * pointPerRank
			big = improveLevels[step - 1]
			step = step - 1
		elif step > 0:
			pointPerRank = pointPerLevel / (improveLevels[step] - improveLevels[step - 1])
			point = point + (big - small) * pointPerRank
			return point
		else:
			pointPerRank = pointPerLevel / (improveLevels[step])
			point = point + (big - small) * pointPerRank
			return point
	return 0
	
def calculateImprovePoint(student, startCol):
	curRank = student[startCol]
	preRank = student[startCol+1]
#	print "curRank: %d"%curRank
#	print "preRank: %d"%preRank
	point = 0
	if curRank == 0 or preRank == 0 or curRank == preRank:
		point = 0
	elif curRank < preRank:
		point = calPointByTwoRank(curRank, preRank)
	else:
		point = -calPointByTwoRank(preRank , curRank)
#	print "point:%f"%point
	student.insert(startCol+2, point)
		

calculateSum(firstStudents, 3, subjectsCount, 1)
calculateSum(secondStudents, 3, subjectsCount, 1)
sortAllSubjects(firstStudents, 3, subjectsCount + 1, 2)
sortAllSubjects(secondStudents, 3, subjectsCount + 1, 2)

firstStudents.sort(key = lambda student : student[1])
secondStudents.sort(key = lambda student : student[1])

firstDict = {}
for i in range(len(firstStudents)):
	firstDict[firstStudents[i][1]] = firstStudents[i]

resultStudents = []
for i in range(len(secondStudents)):
	resultStudents.append(secondStudents[i])
	fStudent = firstDict.get(resultStudents[i][1])
	if fStudent != None:
		for j in range(subjectsCount + 1):
			resultStudents[i].insert(5 + j * 3, fStudent[4 + j * 2])
		del firstDict[resultStudents[i][1]]
	else:
		for j in range(subjectsCount + 1):
			resultStudents[i].insert(5 + j * 3, 0)

leftStudents = firstDict.values()
for i in range(len(leftStudents)):
	for j in range(subjectsCount + 1):
		leftStudents[i][3 + j * 3] = 0
		leftStudents[i].insert(4 + j * 3, 0)
	resultStudents.append(leftStudents[i])

resultStudents.sort(key = lambda student : student[1])

for i in range(len(resultStudents)):
	for j in range (subjectsCount + 1):
		calculateImprovePoint(resultStudents[i], 4 + j * 4)
			
for i in range(subjectsCount + 1):
	titles.insert(4 + i * 4, "排名")
	titles.insert(5 + i * 4, "上次排名")
	titles.insert(6 + i * 4, "进步分")

output = xlwt.Workbook(encoding = 'utf-8')
outSheet = output.add_sheet('result')

for i in range (len(titles)):
	outSheet.write(0, i, titles[i])

for i in range (len(resultStudents)):
	for j in range (len(resultStudents[i])):
		outSheet.write(i+1, j, resultStudents[i][j])
		
output.save('result.xls')
tkMessageBox.showinfo(title="说明", message="计算完成！")
