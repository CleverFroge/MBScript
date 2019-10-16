# -*- coding: utf-8 -*-
"""
Created on Mon Jan 22 15:04:59 2018

@author: jzp
"""

"""
运行环境：Python 3
Excel码表文件名：encode.xls
码表格式：
	A列：繁体简体（f,j）
	B列：汉字
	C列：等级（t1~t5，或者空）
	D~J列：编码（标准、容错、简码）
	读取的总行数可以在代码第34行修改（0~30791，再后面的字加到里面输入法读不出来会出问题)
完成后，会在当前目录下生成四个文本文件，手动将其改成Unicode，之后才能被输入法程序使用
"""

import xlrd
import xlwt
import sys
import os

if sys.getdefaultencoding() != 'utf-8': 
    reload(sys) 
    sys.setdefaultencoding('utf-8')
fjDict = {}
levelDict = {}
wordDict = {}

readWorkBook = xlrd.open_workbook(u'encode.xls')
sheet_name = readWorkBook.sheet_names()[0]
readSheet = readWorkBook.sheet_by_name(sheet_name)
size = 0;
for row in range(0,30791):
	#ch : 这一行的汉字
	ch = readSheet.cell(row,1).value
	#fj : 简体繁体？
	fj = readSheet.cell(row,0).value
	if fj:
		fjDict[ch] = fj
	#level : 常用等级
	level = readSheet.cell(row,2).value
	if level:
		levelDict[ch] = int(level.strip('t'))
	else:
		levelDict[ch] = 6
	#读取所有编码
	for col in range(3,9):
		encode = readSheet.cell(row,col).value
		if encode:
			if encode in wordDict:
				wordDict[encode].append(ch)
			else:
				wordList = [ch]
				wordDict[encode] = wordList
				
sortedEncodeList = sorted(wordDict.keys(), reverse=False)

os.remove("All.txt")
os.remove("CommonlyUsed.txt")
os.remove("Simpilified.txt")
os.remove("Traditional.txt")

All = open("All.txt",'a',encoding="utf-8")
CommonlyUsed = open("CommonlyUsed.txt",'a',encoding="utf-8")
Simpilified = open("Simpilified.txt",'a',encoding="utf-8")
Traditional = open("Traditional.txt",'a',encoding="utf-8")

for encode in sortedEncodeList:
	for ch in wordDict[encode]:
		level = levelDict[ch]
		if encode and ch:
			wordStr = ch+'\t'+encode+'\t'+str(int(level))+'\n'
			All.write(wordStr)
			if ch in fjDict and not fjDict[ch]=='j':
				Traditional.write(wordStr)
			if ch in fjDict and not fjDict[ch]=='f':
				Simpilified.write(wordStr)
			if level<=5:
				CommonlyUsed.write(wordStr)