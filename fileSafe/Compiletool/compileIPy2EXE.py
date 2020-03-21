#-*- coding:utf-8
#### compile the  py files to EXE （standalone EXE ，include the  ordinary stdlibs ）,for ironpyton 。
#ipy64 compileipy2exe.py 

import clr
import System
import System.IO

# def getFullNames(dllname):
#     def getEnvs(envnmame):
#         tgt1=System.EnvironmentVariableTarget.Machine
#         tgt2=System.EnvironmentVariableTarget.User
#         tgt3=System.EnvironmentVariableTarget.Process
#         Libpa =[]
#         for tgt in (tgt3,tgt1,tgt2):
#             Libpaths = System.Environment.GetEnvironmentVariable(envnmame,tgt)
#             #print(Libpaths,tgt );print()
#             if Libpaths  == None:
#                 pass
#             else:
#                 Libpa += Libpaths.split(";")
#         print()
#         return Libpa
#     libpaths = [".\\"]+getEnvs("libpath") + getEnvs("path")
#     #print( "---------------",libpaths);print()
#     for p in libpaths:#.split(";"):
#         filename = p+"\\"+dllname
#         if ( System.IO.File.Exists(filename) ):
#             print("---Load Dlls:" ,filename)
#             return filename
#     print( "Can not find Lib:" ,dllname)
#     return None

# dllname=r"""STDLIBAlls.DLL""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )
# clr.AddReference( "stdlib.dll")
import os, glob
import fnmatch
import pyc
import compileLibs

# rootDisk="c:\\"
# SrcRoot="Cur\\2016-10-25"
# LibPathName= r"C:\\z\\IronPython\\Lib" #  maybe None


def getLibFiles( paths=[]):
	gb1 =[]
	for p1 in paths:
		p=  p1+r"\*.py"  
		# print p
		gb1 += glob.glob( p )
	return gb1
def Compile2ToolsExePy( _tobeCopiedfiles1,_srcfiles,_EntryFilename,_SubFolder):
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

	compileLibs.compileAppExe(_srcfiles ,EntryFilename1 ,None) 

	src=_EntryFilename.lower().strip().replace(".py",".exe")
	tgrexe =  os.path.join(desFolder,src)
	compileLibs.movefile2Tgr(src, tgrexe)
	print  "编译目标文件：",tgrexe  
	files2=_tobeCopiedfiles1.strip().split()
	for f1 in files2:
		srcfilename =os.path.join(SrcFolder,f1 )
		tgrFilename= os.path.join( desFolder,f1 )

		compileLibs.copyfile2Tgr(srcfilename,tgrFilename)
def Compile2ToolsExeBat( _tobeCopiedfiles1,_srcfiles,_EntryFilename,_SubFolder):
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

	# compileLibs.compileAppExe(_srcfiles ,EntryFilename1 ,None) 
	cmdTpl ="ipyc /target:exe  /standalone /platform:x64 /main:{:}  {}" 
	cmdStr =  cmdTpl.format( EntryFilename1 ,_srcfiles+ " "+EntryFilename1 )
	print( cmdStr )
	os.system(cmdStr )
	# gb += ["" ]
	# gb += ["/main:" + EntryFilename  ]
	src=_EntryFilename.lower().strip().replace(".py",".exe")
	tgrexe =  os.path.join(desFolder,src)
	compileLibs.movefile2Tgr(src, tgrexe)
	print  "编译目标文件：",tgrexe  
	files2=_tobeCopiedfiles1.strip().split()
	for f1 in files2:
		srcfilename =os.path.join(SrcFolder,f1 )
		tgrFilename= os.path.join( desFolder,f1 )

		compileLibs.copyfile2Tgr(srcfilename,tgrFilename)
def Compile2ToolsExe( _tobeCopiedfiles1,_srcfiles,_EntryFilename,_SubFolder):
	Compile2ToolsExeBat( _tobeCopiedfiles1,_srcfiles,_EntryFilename,_SubFolder)
# root文件夹 = """  E:\ipy\电子围栏Tools  """.strip()
# rootDisk, SrcRoot = os.path.splitdrive(root文件夹)
# libpath = os.path.join(SrcRoot,"libs")
# LibPathNames=[ libpath  ,]
# print( root文件夹,rootDisk, SrcRoot , libpath )
# desRootFolder=os.path.join( SrcRoot , r"Tools\\Compiled"  )
root文件夹 = """  F:\\ipy\\fileSafe """.strip()
rootDisk, SrcRoot = os.path.splitdrive(root文件夹)
libpath = os.path.join( "f:\\ipy\\libs ".strip() )
# libpath=os.path.abspath( libpath )
LibPathNames=[ libpath  ,]
desRootFolder=os.path.join( SrcRoot , r"Compiled"  )
SubFolder= "src"

EntryFilename ="""00-Backup.py
"""
srcfiles="""
00-Backup.py
filesafe.py
glCls.py
10-restore.py
"""
tobeCopiedfiles1=  "__future__.py "
Compile2ToolsExe( tobeCopiedfiles1,srcfiles,EntryFilename,SubFolder)

EntryFilename ="""10-restore.py
"""
srcfiles="""
00-Backup.py
filesafe.py
glCls.py
10-restore.py
"""
tobeCopiedfiles1=  "__future__.py "
Compile2ToolsExe( tobeCopiedfiles1,srcfiles,EntryFilename,SubFolder)



