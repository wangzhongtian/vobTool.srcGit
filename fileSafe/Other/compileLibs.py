#-*- coding:utf-8
#### compile the  py files to EXE （standalone EXE ，include the  ordinary stdlibs ）,for ironpyton 。

#ipy compileipy2exe.py
import pyc
# import clr
# clr.AddReference( "stdlibALLs.dll")
import os, glob
import fnmatch

def srcfileProc(files,Folder):
	newfiles=""
	for f1 in files.split():
		newfiles += " " +os.path.join(Folder,f1 )
	return newfiles.strip()

def getLibPys(ipLib_Rootpath= r"C:\\z\\IronPython\\Lib" ):
	gb1 = glob.glob(ipLib_Rootpath+r"\*.py")
	gb2 =  glob.glob(ipLib_Rootpath+r"\encodings\*.py")
	gb=gb1+gb2 
	return gb

def compileAppExe(files ,EntryFilename ,LibPathName=None):
	fs= files.split()
	gb1 = []
	gb2=[]
	# print(EntryFilename )
	gb  =  ["/target:exe ",]   
	gb += ["/standalone " ]
	gb += ["/main:" + EntryFilename  ]
	# print LibPathName
	if LibPathName != None :
		# print LibPathName,getLibPys( LibPathName )
		gb2 += getLibPys( LibPathName ) 
	gb = gb + gb1 + gb2 + fs
	# for g in gb :
	# 	print (g)
	pyc.Main(  gb)
	
def movefile2Tgr( src,tgr):
	import shutil
	import os
	if ( os.path.exists( tgr) ) :
		# print tgr 
		os.remove( tgr )
	shutil.move(src, tgr)

def copyfile2Tgr( src,tgr):
	import shutil
	import os
	if ( os.path.exists( tgr) ) :
		# print tgr 
		os.remove( tgr )
	shutil.copy(src, tgr)

def Compile2TgrExe(EntryFilename,files,desFolder ):
	compileAppExe(files ,EntryFilename )
