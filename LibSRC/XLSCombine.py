#-*- coding: UTF-8
from __future__ import print_function
# from commonFuncs import *
# import win32com.client
import math
import csv
import os
import datetime
import random
import clr
from System.Runtime.InteropServices import Marshal
clr.AddReferenceByPartialName("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass
import csv, codecs, cStringIO

class xlsCombinerocCls:
	def 合并2Xls(self,xlsfileName,srcFilenames=[]):
		xlCSV = 6
		xlOpenXMLWorkbook = 51
		xlExcel12 = 50
		self.curApp = ApplicationClass()
		self.curApp = Marshal.GetActiveObject("Excel.Application")
		self.curApp.Application.DisplayAlerts = False
		self.books=self.curApp.Application.Workbooks
		newbooks = self.books.Add()
		newbooks.SaveAs( xlsfileName, xlExcel12 )
		AfterSheet = newbooks.Sheets(1)

		for srcfile in srcFilenames:
			try:
				os.startfile( srcfile )
			except:
				pass
			book=self.curApp.Application.ActiveWorkbook
			rootpath,filename = os.path.split( srcfile )
			if  filename == book.Name:
				book.Sheets(1).Move( AfterSheet )
		newbooks.Save(  )

	def 合并Xls(self,fileName):
		xlCSV = 6
		xlOpenXMLWorkbook = 51
		xlExcel12 = 50
		self.curApp = ApplicationClass()
		self.curApp = Marshal.GetActiveObject("Excel.Application")
		self.curApp.Application.DisplayAlerts = False
		self.books=self.curApp.Application.Workbooks
		newbooks = self.books.Add()
		newbooks.SaveAs( fileName, xlExcel12 )
		AfterSheet = newbooks.Sheets(1)
		for book in self.books:
			filename =book.Name
			extension = filename [-4:]
			extension= extension.upper() 
			if extension ==".CSV":
				pass
				book.Sheets(1).Move( AfterSheet )
		newbooks.Save(  )
	def 关闭CSV文件(self):
		xlCSV = 6
		self.curApp = win32com.client.GetActiveObject("excel.application")
		self.curApp.Application.DisplayAlerts = False
		self.books=self.curApp.Application.Workbooks
		for book in self.books:
			filename =book.Name
			extension = filename [-4:]
			extension= extension.upper() 
			if extension ==".CSV":
				pass
				book.Close()


		self.curbook=self.curApp.Application.ActiveWorkbook
		orgFile=self.curbook.FullName
		self.dataPath= os.path.join(orgFile,"..")
		print("File outputpath：",self.dataPath)
		self.removeTemFiles()
		
	def removeTemFiles(self):
		import os
		from os.path import join, getsize
		for root, dirs, files in os.walk(self.dataPath):
			for file1 in files:
				if file1[-4:].lower() in [".1kml",".csv"]:
					os.remove(os.path.join(root,file1)  ) # don't visit CVS directories
			
	def __init__(self ):
# 		self.fileName = filename
		pass