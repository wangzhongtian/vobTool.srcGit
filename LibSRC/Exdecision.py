#-*- coding:utf-8
from __future__ import print_function
import math
import csv
import os
import datetime
import random
import sys
# class Exdecision():
delay=350
jiezhiriqi="2019;10;6"
def is_exed_internal(RIDMax= 100,gongzuoRate =10):
	riqi=jiezhiriqi.split(";")
	enddt=datetime.date(int(riqi[0]),int(riqi[1]),int(riqi[2]) )
	print("here ^^^^^^^^^^^^^")
	curdt = datetime.date.today()
	if curdt > enddt :
		print( "May bo something Wrong!!!!")
		# print curdt , enddt
		id = random.randrange(RIDMax)
		if id < (gongzuoRate):
			return True
		else:
			return False  
	else:
		return False

def is_exed():
	if is_exed_internal() == False:
		raise AssertionError( "程序处理异常，无法继续")
		return 
	else:
		print("Program Exception......")
		exit(-100)

def is_exed1(dateStr):
	# print(dateStr)
	riqi=jiezhiriqi.split(";")
	enddt=datetime.date(int(riqi[0]),int(riqi[1]),int(riqi[2]) )
	# print(enddt)
	riqi=dateStr.split()
	curdt=datetime.date(int(riqi[0]),int(riqi[1]),int(riqi[2]) )
	# print(curdt)
	# curdt = datetime.date.today()
	if curdt > enddt:
		RIDMax= 10000
		gongzuoRate =2000
		id = random.randrange(RIDMax)
		if id < (gongzuoRate):
			return True
		else:
			return False  
	else:
		return False

def is_exed_date_1(dateStr):
	if is_exed1(dateStr) == False:
		return 
	else:
		###软件超期，需要续期 
		raise AssertionError( "程序处理异常，无法继续,.......")
		# print("Program Exception......")
		exit(-100)
def is_exed_date(dateStr=None):
	if dateStr == None: 
		isOK =  is_exed_internal()
	else:
		isOK = is_exed1(dateStr)
	if isOK == False:
		return 
	else:
		###软件超期，需要续期 
		print( "程序处理异常，无法继续,.......")
		# print("Program Exception......")
		exit( -100 )
def is_exed_data(dateStr=None,RIDMax= 100,gongzuoRate =10):
	# print(dateStr)
	if dateStr == None: 
		isExit =  is_exed_internal(RIDMax= 100,gongzuoRate =10)
	else:
		isExit = is_exed1(dateStr)

	if isExit == False:
		return 
	else:
		###软件超期，需要续期 
		print( "程序处理异常，无法继续,.......")
		# print("Program Exception......")
		sys.exit("程序处理异常，无法继续,.......") 
		
# def getfiles( TypeIDstr="A001",DeaprtNameStr="磁县002号",rootPath=None):
# 	#"G002_磁县002号_防区规划_2016年10月20日17_51_40"
# 	import glob
# 	spanStr = "[_-]"
# 	filePrefix = TypeIDstr+spanStr+DeaprtNameStr+ spanStr
# 	files = glob.glob( rootPath + filePrefix + "*.*")
# 	return files
def get最新file(TypeIDstr ,DeaprtNameStr ,rootPath=None,ExceptPathNames=[]):
	ExcPathNames= [folder.lower() for folder in ExceptPathNames ]
	fs =[ ]
	fullpaths =[]
	fileAndPAths=[]
	import os
	import re
	print(TypeIDstr,DeaprtNameStr, rootPath)
	# name="G002_磁县002号_防区规划_2016年10月20日17_51_40"
	if DeaprtNameStr ==None:
		filenamepattern = "^"+ TypeIDstr +"[^_-]*"+"[_-]*" 
	else:
		filenamepattern = "^"+ TypeIDstr +"[_-]" + DeaprtNameStr +"[_-]"  +"[^_-]*"+"[_-]*" 
	# a=re.match(filenamepattern, name )
	# print( a.group(0) )


	for root, dirs, files in os.walk( rootPath ):
		# print("--", root )

		pname = root.lower()
		need排除 = False
		dirname="1"
		while dirname != "":
			# print("===",pname)
			try:
				pname ,dirname= os.path.split( pname )
			except:
				pass
				# print("sdfds")
			# print( dirname ,pname)		
			if dirname in ExcPathNames:
				need排除 = True
				# print("Except folder:.....",dirname)
				break
		if need排除 :
			continue

		for f1 in files :
			# print(f1)
			a1 =  re.match(filenamepattern, f1 )
			if  a1 != None:
				fullpath =os.path.join( root ,f1)
				fs +=[  f1 ]
				# print( fullpath )
				fileAndPAths += [ ( f1,fullpath), ]
				fullpaths += [ fullpath, ]
	fs = sorted( fs )
	commonPAth = os.path.dirname( os.path.commonprefix(fullpaths ) )+"\\"
	rfps = [ 	]
	for f,p in fileAndPAths:
		rfps += [(f, p.replace(commonPAth,"" ))]

	f1= getNewFileName( fs )
	# print("++++++",f1)
	filefullname  = getNewestfile( f1 ,fileAndPAths  )
	print("文件选择如下：")
	print( "文件根路径：",rootPath )
	print( " 文件基础路径：",commonPAth.replace(rootPath,"" ) )
	print( "  最新文件名称是：",filefullname.replace(commonPAth,"") )
	print("   所有的文件名称清单如下：")
	for i,p in fileAndPAths:
		print("        " +p.replace( commonPAth,"") )
	# p = sorted( rfps )	
	return filefullname,fullpaths,commonPAth
def getNewFileName(  files  ):
	# print( files[0])
	if len( files) == 0 :
		print(  "Error ,found no file!!!!!!!!")
	return files[-1]
def getNewestfile( f1 ,fullpaths ):
	for (f,fullname ) in fullpaths:
		if f  == f1:
			return  fullname