# -*- coding: utf-8 -*-
"""
Created on Mon Jan 22 15:04:59 2018

@author: jzp
"""

"""
运行环境：Python 3
形码码表文件名：phrase.xls
码表格式：
	A列：汉字
	C列：第一形码
	D列：第二形码
	读取的总行数可以在代码第31行修改（0~31014）
词组文件：phrase.txt
词组文件格式：
	每行一个词组
完成后，会在当前目录下生成Shape.txt，手动将其改成Unicode，之后才能被输入法程序使用
"""

import xlrd
import xlwt
import sys
import os

firstShape = {}
secondShape = {}

readWorkBook = xlrd.open_workbook(u'phrase.xls')
sheet_name = readWorkBook.sheet_names()[0]
readSheet = readWorkBook.sheet_by_name(sheet_name)

for row in range(0,31014):
	ch = readSheet.cell(row,0).value
	firstShape[ch] = readSheet.cell(row,2).value
	secondShape[ch] = readSheet.cell(row,3).value

fp = open("phrase.txt",'rb')

phraseDict = {}

for line in fp.readlines():
	word = line.decode('gbk').strip('\n').strip('\r')
	length = len(word)
	encode = ""
	try:
		if length==2:
			if word[0] in secondShape:
				encode = firstShape[word[0]]+secondShape[word[0]]
			else:
				encode = firstShape[word[0]]+firstShape[word[0]]
			if word[1] in secondShape:
				encode += firstShape[word[1]]+secondShape[word[1]]
			else:
				encode += firstShape[word[1]]+firstShape[word[1]]
		elif length==3:
			if word[0] in secondShape:
				encode = firstShape[word[0]]+secondShape[word[0]]+firstShape[word[1]]+firstShape[word[2]]
			else:
				encode = firstShape[word[0]]+firstShape[word[0]]+firstShape[word[1]]+firstShape[word[2]]
		else:
			encode = firstShape[word[0]]+firstShape[word[1]]+firstShape[word[2]]+firstShape[word[3]]
		if encode in phraseDict:
			phraseDict[encode].append(word)
		else:
			phraseDict[encode] = [word]
	except KeyError:
		print(word)
fp.close()

sortedEncodeList = sorted(phraseDict.keys(), reverse=False)

os.remove("Shape.txt")
shape = open("Shape.txt",'a',encoding="utf-8")

for encode in sortedEncodeList:
	for phrase in phraseDict[encode]:
		shape.write(phrase+'\t'+encode+'\n')



