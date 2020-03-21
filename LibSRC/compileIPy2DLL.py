#-*- coding:utf-8
#### compile the  py files to EXE （standalone EXE ，include the  ordinary stdlibs ）,for ironpyton 。
from __future__ import print_function
import clr
import System
import System.IO

def getFullNames(dllname):
	libpaths= ".\\;"+System.Environment.GetEnvironmentVariable("path") +";"+System.Environment.GetEnvironmentVariable("libpath")
	for p in libpaths.split(";"):
		filename = p+"\\"+dllname
		if ( System.IO.File.Exists(filename) ):
			print("---Load Dlls:" ,filename)
			return filename
	return None

# dllname=r"""stdlib.DLL""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )

# clr.AddReferenceToFileAndPath( "stdlib.dll")
import os, glob
import fnmatch
import pyc
import os
import sys
def compileStd(ipLib_Rootpath):
	gb1 = glob.glob(ipLib_Rootpath+r"\*.py")
	gb2 =  glob.glob(ipLib_Rootpath+r"\encodings\*.py")
	gb=gb1+gb2 
	gb.append("/out:StdLibALLs")  

	# print ["/target:dll",]  + gb
	pyc.Main(["/target:dll"]+gb)
def compileAppDLL( files,outname ="AppAssembly"):
	fs= files.split()
	gb= fs
	gb.append("/out:"+outname)  
	# gb.append("/out:StdLib") 
	# print ["/target:dll",]  + gb
	pyc.Main(["/target:dll"]+gb)

def compileAppExe(fies ,EntryFilename ):
	fs= files.split()
	gb= fs
	gb =  ["/target:exe ",]  + gb
	gb += ["/standalone" ]
	gb += ["/main:"+EntryFilename  ] 
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
	src=EntryFilename.lower().strip().replace(".py",".exe")
	tgrexe =  os.path.join(desFolder,src)
	print( src,tgrexe )
	movefile2Tgr(src, tgrexe)

def compileDirs2AppDLL( paths=[],outname ="AppAssembly"):
	gb1 =[]
	for p1 in paths:
		# print("---" ,p1)
		p=  p1+r"\*.py"  #os.path.join( p1,r"\*.py")
		# print( p)
		gb1 += glob.glob( p )
		# for i in gb1 :print(i)
	with open("resp.txt","w")  as f1:
		for i in gb1:
			print(i ,file=f1)
	_srcfiles= " ".join( gb1)
	libpath=os.path.dirname( __file__ )
	cmdTpl ="ipy pyc.py  @resp.txt  /out:{}" 
	cmdStr =  cmdTpl.format(  os.path.join(libpath,outname)  )
	print( cmdStr )
	os.system( cmdStr )
def mainLibComp():
	curdir =os.getcwd()
	# libpath=os.path.dirname( __file__ )
	# print( "---------",libpath ,__file__)
	# print( "---------",os.path.abspath( curdir ) )
	# compileDirs2AppDLL( [os.path.abspath( curdir ) ],"AppAssembly");
	# print()
	# import shutil
	# shutil.copyfile(os.path.abspath("./AppAssembly.dll"),os.path.abspath("../libs/AppAssembly.dll") )
	# return
	# sys.exit()
	ipLib_Rootpath1 = r"C:/Program Files/IronPython 2.7/Lib"
	# ipLib_Rootpath1 = r"C:\\关键\\IronPython\\Lib"
	# ipLib_Rootpath1 = r"C:\\IronPython\\Lib"
	ipLib_Rootpath2  =ipLib_Rootpath1+r"\encodings"
	compileDirs2AppDLL([ ipLib_Rootpath1,ipLib_Rootpath2] ,"StdLibALLs")	
def mainAppComp():
	curdir =os.getcwd()
	# libpath=os.path.dirname( __file__ )
	# print( "---------",libpath ,__file__)
	# print( "---------",os.path.abspath( curdir ) )
	compileDirs2AppDLL( [os.path.abspath( curdir ) ],"AppAssembly");
	print()
	import shutil
	shutil.copyfile(os.path.abspath("./AppAssembly.dll"),os.path.abspath("../libs/AppAssembly.dll") )
	return
	
if __name__ == "__main__":
	mainAppComp()
	mainLibComp()
# mainComp()