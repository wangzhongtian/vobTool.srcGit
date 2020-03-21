#-*- coding: UTF-8
# from commonFuncs import *
#import win32com.client
from __future__ import print_function
import clr
# dllname= r"C:\z\_a\nsbd\调试报告\src\sft.ipy\STDLIB.DLL"
dllname=r"""StdLibALLs.DLL"""
# clr.LoadAssemblyFromFileWithPath(dllname)
clr.AddReferenceToFile( dllname)

import math
import csv
import os
import datetime
import random
import codecs
import clr
from System.Runtime.InteropServices import Marshal
clr.AddReferenceByPartialName("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass
import XLSCombine
import datetime

def CombineProcXls( path,Prefix,filenameTpl="Tem_" ):##filenameTpl+2016年10月16日14_42_12.xlsx
	tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xls"
	curDt = datetime.datetime.now()
	filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
	filename1 = os.path.join(path,filename )
	xlsCombinerocObj =XLSCombine.xlsCombinerocCls( )
	print()
	print(filename1 )
	print()	
	xlsCombinerocObj.合并Xls( filename1)
	return filename1
def Combine2Xls( path,Prefix,filenameTpl="Tem_",srcFilenames=[] ):##filenameTpl+2016年10月16日14_42_12.xlsx
	tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xls"
	curDt = datetime.datetime.now()
	filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
	filename1 = os.path.join(path,filename )
	xlsCombinerocObj =XLSCombine.xlsCombinerocCls( )
	print()
	print(filename1 )
	print()	
	xlsCombinerocObj.合并2Xls( filename1,srcFilenames)	
	return filename1
import csv, codecs, cStringIO
def movefile2Tgr( src,tgr):
	import shutil
	import os
	if ( os.path.exists( tgr) ) :
		# print tgr 
		os.remove( tgr )
	shutil.move(src, tgr)
def convertFile2utf8(srcfile,Spanlines=0):
		temfile="TEM_____aaaaaaa.csv"
		line=0;
		with codecs.open(temfile, 'w',encoding="utf_8_sig") as csvfile:
			with codecs.open(srcfile,"r" ) as cfile:
				for row in cfile:
						d2= unicode(row,"mbcs",'ignore' ).strip()  
						if Spanlines <= line:
							csvfile.write( d2 + "\r\n")
						else:
							# print( "------",d2 )
							pass
						line += 1
		csvfile.close()
		cfile.close()
		# print("---+----",srcfile)
		movefile2Tgr(temfile,srcfile )
class CSVProcCls:
	# 返回一个dict的数组
	def readin_1_CSVFile(self,CsvFilename,SpanLines =0		):
		rows =[]
		# print("++++++++++++++++++++", SpanLines)
		try:
			convertFile2utf8( CsvFilename,SpanLines)
			# print("++++++++++++++++++++", SpanLines)
			with codecs.open(CsvFilename, 'r',encoding="utf_8_sig") as csvfile:
				spamreader = csv.DictReader(csvfile, delimiter=',', quotechar='"')
				for row in spamreader:
					rows += [row]
		except:
			print("--+--Error in Read CSV FIle:", CsvFilename)
			pass
		# print(rows)
		return rows
	def readCsvFiles(self,CsvFilenames ,SpanLines =0	):
		rows =[]
		for Csvfile in CsvFilenames:
			rows += self.readin_1_CSVFile( Csvfile ,SpanLines)
		return rows

	def 导出到CSV文件(self,srcfilename,tgrfilePrefix,srcsheetnames,rootpath=""):
		xlCSV = 6
		xlCSVWindows = 23	
		tgrFiles =[]	


		import System
		# import INFITF
		t = System.Type.GetTypeFromProgID("Excel.Application")
		self.curApp = System.Activator.CreateInstance(t)
		print( "转换文件和Sheet：",end=" ")
		print( srcfilename ,srcsheetnames[0])
		# os.startfile(srcfilename)
		# self.curApp = ApplicationClass()
		# self.curApp = Marshal.GetActiveObject("Excel.Application")
		self.curApp.Visible =True
		self.curApp.Application.DisplayAlerts = False
		self.curApp.Application.Workbooks.Open(srcfilename)
		self.curbook=self.curApp.Application.ActiveWorkbook
		if rootpath == "":
			dataPath = self.curbook.path 
		else :
			dataPath = rootpath
		for sheetname in srcsheetnames:
			try:
				fileName = os.path.join(dataPath,tgrfilePrefix+sheetname+".csv")
				self.curSheet = self.curbook.Sheets(sheetname).Activate()
				self.curbook.SaveAs(fileName, xlCSV ,False)
				print("Save to file Ok:",fileName)
			except:
				print("-- --Save to CVS Error.",sheetname,fileName)
				# self.curbook.Close()
				continue
				pass
			print("导出表页,success.:",end="")
			print( sheetname ," ,in file ",srcfilename)
			tgrFiles += [fileName ]
		self.curbook.Close()
		self.curApp.Application.Quit()
		self.temFiles += tgrFiles
		return tgrFiles
		
	def 清理CSV文件(self):
		xlCSV = 6
		self.curApp = ApplicationClass()
		self.curApp = Marshal.GetActiveObject("Excel.Application")
		# self.curApp = win32com.client.GetActiveObject("excel.application")
		self.curApp.Application.DisplayAlerts = False
		self.books=self.curApp.Application.Workbooks
		for book in self.books:
			filename =book.Name
			extension = filename [-4:]
			extension= extension.upper() 
			if extension ==".CSV":
				pass
				book.Close()
	def outDict2Xls(self,dict_of_Dict_Data,fieldnamesStr,filenameStr,rootpath,Prefix):
		self.outDict2CSV(dict_of_Dict_Data,fieldnamesStr,filenameStr )
		CombineProcXls(rootpath,Prefix, filenameStr )

	def outDict2CSV(self,dict_of_Dict_Data,fieldnamesStr,filenameStr ):
		fieldnames  = fieldnamesStr.split()
		titleRow = dict()
# 		dataRow  = dict()
		for data in fieldnames:
			titleRow[data] = data
		filename= os.path.join( filenameStr+".csv")
		with codecs.open(filename, 'w',encoding="utf_8_sig") as csvfile: 
			spamwriter = csv.DictWriter(csvfile, fieldnames, restval='', extrasaction='ignore', dialect='excel')
			spamwriter.writerow( titleRow )
			ks = sorted( dict_of_Dict_Data.keys() ) 
			for k in  ks :
				dataRow = dict_of_Dict_Data[k]
				spamwriter.writerow(dataRow)
		os.startfile(filename)	
		return 	[ filename ]
	def __init__(self):
		self.temFiles=[]
		pass

class CSVProcArraryCls:
	def outDict2CSV(self,rootpath,filename1,mibiaodict,fieldnamesStr):
		fieldnames  = fieldnamesStr.split()
		# for fd in fieldnames: print( fd)
		row0 = dict()
		row1 = dict()
		for data in fieldnames:
			row0[data] =data
		filename= os.path.join( rootpath,filename1)
		with codecs.open(filename, 'w',encoding="utf_8_sig") as csvfile:#,encoding="utf_8_sig"
			spamwriter = csv.DictWriter(csvfile, fieldnames, restval='', extrasaction='ignore', dialect='excel')
			spamwriter.writerow(row0)

			ks = list(mibiaodict.keys()) 
			for k in sorted( ks ):
				v= mibiaodict[k]
				v1 = v
				idx = 0
				# print(v)
				for ii in v1 :
					keyname = fieldnames[idx]
					row1[ keyname ] = ii
					idx += 1
				# print(row1)
				spamwriter.writerow(row1)
		os.startfile(filename)

	def ReadinCSV2Arrary_1(self,rootpath,tgrfilename,srcsheetnames,tgrArray):
		for sheetname in srcsheetnames:
				fileName = os.path.join(rootpath, tgrfilename+sheetname+".csv")
				try:
					print("处理表页：",end="")
					print( sheetname)
					with codecs.open(fileName,"r",encoding="utf_8") as csvfile:
					    spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
					    for row in spamreader:
					    	rowArray =[]
					    	# print(row)
					    	for d1 in row:
					    		# print(d1)
					    		rowArray += [d1,]
					    	tgrArray += [rowArray]
				except:
					print("--Error in  Read CSV FIle:", fileName)
					pass
	
	def ReadinCSV2Arrary(self,rootpath,tgrfilename,srcsheetnames,tgrArray):
		for sheetname in srcsheetnames:
				fileName = os.path.join(rootpath, tgrfilename+sheetname+".csv")
				try:
					print("处理表页：",end="")
					print( sheetname)
					with codecs.open(fileName,"r" ) as csvfile:
					    spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
					    for row in spamreader:
					    	rowArray =[]
					    	# print(row)
					    	for d1 in row:
					    		d2= unicode(d1,"mbcs",'ignore' )  
					    		rowArray += [d2,]
					    	tgrArray += [rowArray]
				except:
					print("----Error in  Read CSV FIle:", fileName)
					pass

	
	def 导出到CSV文件(self,srcfilename,tgrfilename,srcsheetnames,rootpath=""):
		xlCSV = 6
		xlCSVWindows = 23		
		#self.curApp = win32com.client.GetActiveObject("excel.application")
		self.curApp = ApplicationClass()

		self.curApp = Marshal.GetActiveObject("Excel.Application")

		self.curApp.Application.DisplayAlerts = False
		self.curApp.Application.Workbooks.Open(srcfilename)
		self.curbook=self.curApp.Application.ActiveWorkbook
		# dataPath = self.curbook.path 
		# print( dataPath)
		if rootpath == "":
			dataPath = self.curbook.path 
		else :
			dataPath = rootpath
		for sheetname in srcsheetnames:
			try:
				fileName = os.path.join(dataPath,tgrfilename+sheetname+".csv")
				self.curSheet = self.curbook.Sheets(sheetname).Activate()

				self.curbook.SaveAs(fileName, xlCSV ,False)
				print("Save to file Ok:",fileName)
			except:
				print("----Save to CVS Error.",sheetname,fileName)
				# self.curbook.Close()
				pass
			print("导出表页,success.:",end="")
			print( sheetname ," ,in file ",srcfilename)
		self.curbook.Close()
					
	def __init__(self):
		pass
	def ReadinCSV2Arrary_2(self,rootpath,tgrfilename,srcsheetnames,tgrArray):
		for sheetname in srcsheetnames:
				fileName = os.path.join(rootpath, tgrfilename+sheetname+".csv")
				try:
					print("处理表页：",end="")
					print( sheetname)
					with codecs.open(fileName,"r" ) as csvfile:
					    spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
					    for row in spamreader:
					    	rowArray =[]
					    	# print(row)
					    	for d1 in row:
					    		d2= unicode(d1,"mbcs",'ignore' )  
					    		rowArray += [d2,]
					    	tgrArray += [rowArray]
				except:
					print("----Error in  Read CSV FIle:", fileName)
					pass
	def outDict2Xls(self,dict_of_Dict_Data,fieldnamesStr,filenameStr,rootpath,Prefix):
		self.outDict2CSV(dict_of_Dict_Data,fieldnamesStr,filenameStr )
		CombineProcXls(rootpath,Prefix, filenameStr )