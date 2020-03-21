#-*- coding:utf-8
#### compile the  py files to EXE （standalone EXE ，include the  ordinary stdlibs ）,for ironpyton 。

#ipy compileipy2exe.py 

# import sys
# sys.path.append(r'C:\Program Files\IronPython 2.7\Lib')
# sys.path.append(r'C:\Program Files\IronPython 2.7')

import clr
clr.AddReference( "StdLibAlls.dll")
# clr.AddReference( "stdlib.dll")
import os, glob
import fnmatch
import pyc
import compileLibs

# rootDisk="c:\\"
# SrcRoot="Cur\\2016-10-25"
# LibPathName= r"C:\\z\\IronPython\\Lib" #  maybe None
root文件夹 = os.environ["root文件夹"].strip() 
rootDisk, SrcRoot = os.path.splitdrive(root文件夹) 
libpath = os.path.join(SrcRoot,"libs")
LibPathNames=[ libpath  ,]
print root文件夹,rootDisk, SrcRoot , libpath 
desRootFolder=os.path.join( SrcRoot ,r"Compiled")

def getLibFiles( paths=[]):
	gb1 =[]
	for p1 in paths:
		p=  p1+r"\*.py"  
		# print p
		gb1 += glob.glob( p )
	return gb1
def Compile2ToolsExe( _tobeCopiedfiles1,_srcfiles,_EntryFilename,_SubFolder):
	libfiles = getLibFiles( LibPathNames )
	for  lf in libfiles:
		_srcfiles += lf +"\r\n"
	SrcFolder = os.path.join( rootDisk,SrcRoot,_SubFolder )
	desFolder= os.path.join( rootDisk,desRootFolder,_SubFolder )
	desFolder= os.path.join( rootDisk,desRootFolder )
	_srcfiles = compileLibs.srcfileProc(_srcfiles,SrcFolder)
	_EntryFilename=_EntryFilename.strip( )

	if not os.path.exists(desFolder ):
		os.makedirs(desFolder )
	EntryFilename1 = os.path.join( SrcFolder,_EntryFilename  )
	print(EntryFilename1)
	compileLibs.compileAppExe(_srcfiles ,EntryFilename1 ,None) 

	src=_EntryFilename.lower().strip().replace(".py",".exe")
	tgrexe =  os.path.join(desFolder,src)
	compileLibs.movefile2Tgr(src, tgrexe)
	files2=_tobeCopiedfiles1.strip().split()
	for f1 in files2:
		srcfilename =os.path.join(SrcFolder,f1 )
		tgrFilename= os.path.join( desFolder,f1 )
		compileLibs.copyfile2Tgr(srcfilename,tgrFilename)

LibPathName= None #  maybe None
SubFolder= "src"
EntryFilename = """00-backup.py
"""
srcfiles="""
filesafe.py
glCls.py
"""
tobeCopiedfiles1=  "__future__.py"
Compile2ToolsExe( tobeCopiedfiles1,srcfiles,EntryFilename,SubFolder)



EntryFilename = """10-restore.py
"""
srcfiles="""
filesafe.py
glCls.py
"""
tobeCopiedfiles1=  "__future__.py"
Compile2ToolsExe( tobeCopiedfiles1,srcfiles,EntryFilename,SubFolder)


