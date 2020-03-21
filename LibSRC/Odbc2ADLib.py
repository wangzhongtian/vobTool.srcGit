#-*- coding:utf-8 -*-
from __future__ import print_function
#######################  ######################################
import sys
import clr
import System
import System.IO

def getFullNames(dllname):
    def getEnvs(envnmame):
        tgt1=System.EnvironmentVariableTarget.Machine
        tgt2=System.EnvironmentVariableTarget.User
        tgt3=System.EnvironmentVariableTarget.Process
        Libpa =[]
        for tgt in (tgt3,tgt1,tgt2):
            Libpaths = System.Environment.GetEnvironmentVariable(envnmame,tgt)
            #print(Libpaths,tgt );print()
            if Libpaths  == None:
                pass
            else:
                Libpa += Libpaths.split(";")
        print()
        return Libpa
    libpaths = [".\\"]+getEnvs("libpath") + getEnvs("path")
    #print( "---------------",libpaths);print()
    for p in libpaths:#.split(";"):
        filename = p+"\\"+dllname
        if ( System.IO.File.Exists(filename) ):
            print("---Load Dlls:" ,filename)
            return filename
    print( "Can not find Lib:" ,dllname)
    return None

dllname=r"""STDLIBalls.DLL""" ;
clr.AddReferenceToFileAndPath( getFullNames(dllname) )

import datetime
import os
#import xlsProLib
import datetime
import re
import csv
import codecs
import Exdecision
clr.AddReference("System.Data")
import System.Data.Odbc


class gCfgDataCls():
	字段文件名="_file"
	字段sheetName="_shName"

class titlesCls():
	titles=[]
	NumberTitles="数量 百分比 单价 金额 价格 认量".split()
	filedmappings=[]
	@classmethod
	def getfieldTypeMappings(cls):
		ftMapping=""
		for field in cls.titles:
			type1="text"
			for numbertile in cls.NumberTitles:
				# print( numbertile,cls.NumberTitles)
				if numbertile in field:
					type1 = "numeric"
					print( field ,"is numeric ")
			ftMapping +='{:} {:} '.format(field,type1)
			cls.filedmappings += [(field,type1 )]
		# print( ftMapping)
		return ftMapping

	@classmethod
	def gentitles(cls, row ):
		# cls.titles= []
		for key in row:
			# print( key )
			key1 = key.strip()
			val  = row[ key ]
			if key1 != key:
				del row[ key]
			try:
				row[ key1 ] = val.strip()
			except:
				pass
			if key1 not in cls.titles:
				# print( key1)
				cls.titles += [ key1]
			# for i in cls.titles: print(i,end=";")
	@classmethod
	def getSelectFieldStr(cls):
		return " sum(归一金额) as Value  "


import re
class xlstbl2AD():
	# curApp = ApplicationClass()
	def __init__(self,xlsfilename):
		self.xlsfilename= xlsfilename
		self.RowBegin=0
		pattern2='''Driver={:};Provider=Microsoft.ACE.OLEDB.12.0;DBQ={:};IMEX=0;IgnoreCalcError=true;AllowFormula=false;Extended Properties="Excel 12.0 Xml;HDR=YES;EmptyTextMode=NullAsEmpty;IMEX=0";'''
		DriverName = '{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}'
		connectionString = pattern2.format(DriverName,self.xlsfilename )
		# connectionString= pattern12
		print(self.xlsfilename  )
		# print( connectionString )
		self.connection = System.Data.Odbc.OdbcConnection(connectionString)
		# print(connection.Database, connection.DataSource )
		try:
			self.connection.Open()
		except Exception as e:
			print( e )
			print(connectionString)

			sys.exit()
		tblinfo = self.connection.GetSchema("Tables") 
		self.Tbls=[]
		for row in tblinfo.Rows:
			for col in tblinfo.Columns:
				# print("----")
				if  col.ColumnName == "TABLE_NAME":
					#print(col.ColumnName )
				#print("===")
					# print(row[col])
					tblname = row[col]
					if tblname[-1:] == "$" or tblname[-2:] == "$'":
						# if tblname
						# print("-----------",tblname)
						self.Tbls +=[tblname]
						# print( tblname)
		# print("  --- ")
	def xls2AofD(self ,sheetsNamehead=[]):
		# sheets= self.book.sheets
		if len( sheetsNamehead ) >0:
			pattern = "|^".join( sheetsNamehead )
		else:
			return
		a_dDatas=[]
		for sn  in  self.Tbls :
			# print(sn)# = sheet.Name
			if re.match( pattern,sn)  or re.match( pattern,sn.replace("'","")) :
				a_dDatas += self.procSheet(  sn )
			# print("",sn,self.book.Name)
		return a_dDatas
	def procSheet(self, sheetName):
		command = System.Data.Odbc.OdbcCommand();
		command.Connection = self.connection;
		# connection.Open();
		if self.RowBegin ==1:
			sqlstr="select * from [{}]".format(sheetName)
		else:
			sqlstr="select * from [{}A{}:Z]".format(sheetName.replace("'",""),self.RowBegin)
		# print( sqlstr)
		command.CommandText = sqlstr
		reader1 = command.ExecuteReader();
		adDatas =[]
		idx =0
		curid =""
		while( reader1.Read()):
			idx += 1
			d1= dict()
			curid=""
			fn =""

			for i in range(0, reader1.FieldCount):
				try:
					if reader1.IsDBNull(i):
						continue
					fn =reader1.GetName(i);#rint( "-"+fn+"-")
					val =reader1.GetProviderSpecificValue(i)
					val1=""
					t2 = reader1.GetProviderSpecificFieldType(i)
					if t2.FullName =="System.String":
						pass
						val1 = val
						# print(val, t2.FullName)
					if t2.FullName == "System.Double":
						pass
						val1 = "{}".format( val )
						# print( type(val), t2.FullName)
					if t2.FullName == "System.DateTime":
						# print( val,type(val), t2.FullName)
						val1= "{}年{}月{}日  {:02d}:{:02d}:{:02d}".format( val.Year,val.Month,val.Day,val.Hour,val.Minute,val.Second)
						# print( val1 )
					# continue
					if fn.lower() =="id":
						curid  = val1
					d1[fn.strip()] =val1.strip();#print(fn)
				except Exception  as e :
					print(idx,"_",i,curid,fn, sheetName ,e)
					pass
			if len(d1 ) >0:
				d1[gCfgDataCls.字段文件名]= self.xlsfilename
				d1[gCfgDataCls.字段sheetName]=sheetName[:-1]
				# print(d1[cfgfile.字段文件名],d1[cfgfile.字段sheetName]   )
				adDatas +=[ d1 ] 
			if self.MaxRecords >0 and self.MaxRecords  < idx:
				break

		reader1.Close()
		return adDatas
	@classmethod
	def xls2AD(cls,filenames=[], sheetnameHead=[],SpanLines=0,MaxRecords=-1):
		adDatas =[]
		for filename in filenames:
			xlsObj =xlstbl2AD( filename )
			xlsObj.RowBegin = SpanLines+1
			xlsObj.MaxRecords = MaxRecords
			a_dDatas = xlsObj.xls2AofD(sheetnameHead) #'A','B',
			adDatas += a_dDatas
			xlsObj.connection.Close()
		return adDatas


class xlstblWrite():
	@classmethod
	def xls2AD(cls,xlsfilename, sheetname,adData):
		titlesCls.titles=[]
	 	# map( titlesCls.gentitles,a_d_datas)
	 	# print()
		map(titlesCls.gentitles, adData)
		writeObj= cls(xlsfilename)
		writeObj.DataFields= titlesCls.titles
		tblname = sheetname
		writeObj.createTbl( tblname )
		# writeObj.ClearData( tblname )
		for row in adData:
			writeObj.Adddata( row, tblname )

		writeObj.connection.Close()
	@classmethod
	def AD2Xls(cls,xlsfilename, sheetname,adData,titleField=[]):
		writeObj= cls(xlsfilename)
		if titleField ==[]:
			# print( "Title Field is Empty ")
			titlesCls.titles=[]
			map(titlesCls.gentitles, adData)
			writeObj.DataFields= titlesCls.titles
			# print( "Title Field is : ",titlesCls.titles)
			# print( "DataFields is : ",writeObj.DataFields)
		else:
			writeObj.DataFields= titleField
			titlesCls.titles = titleField	
		tblname = sheetname
		# writeObj.ClearData( tblname )		
		writeObj.createTbl( tblname )
		cnt=1
		try:
			writeObj.connection.Close()
			writeObj.connection.Open()	
		except:
			pass
		for row in adData:
			# print(row[1] )
			# cnt=cnt+1
			# print( tblname,cnt )
			if cnt % 10000 == 0 :
				writeObj.connection.Close()				
				writeObj.connection.Open()
				writeObj.Adddata( row, tblname )
			else:
				writeObj.Adddata( row, tblname )
			# if cnt >40000 :
			# 	import sys
			# 	sys.exit()

		writeObj.connection.Close()
	# curApp = ApplicationClass()

	@classmethod
	def saveDD2Xls(cls,xlsfilename, sheetname,ddData,titleField=[]):
		
	 	# map( titlesCls.gentitles,a_d_datas)
	 	# print()
		writeObj= cls(xlsfilename)
		if titleField ==[]:
			map(titlesCls.gentitles, ddData)
			writeObj.DataFields= titlesCls.titles
		else:
			writeObj.DataFields= titleField
		titlesCls.titles=writeObj.DataFields
		tblname = sheetname
		# writeObj.ClearData( tblname)
		writeObj.createTbl( tblname )

		for row in ddData.values():
			writeObj.Adddata( row, tblname )

		writeObj.connection.Close()
	# curApp = ApplicationClass()
	def ClearData(self,TblName ):
		SqlHead ='''CREATE TABLE {}  ( '''.format(TblName)
		SqlTail=''' ) '''
		SqlBody=""
		dBFieldINfoStr= titlesCls.getfieldTypeMappings()
		self.fieldinfos = dBFieldINfoStr.split()
		cnt= len( self.fieldinfos) 
		for idx in range( 0,cnt,2):
			SqlBody += self.fieldinfos[idx] +" " +self.fieldinfos[idx+1] +","
		SqlBody =SqlBody[:-1]
		CreateTblSQL_Main =SqlHead  + SqlBody + SqlTail
		

		# truncate table 
		return ###################################
		CreateTblSQL_Main ="DELETE FROM   {:}  where  序号 = 'A00012321301' ".format(tblname)

		command = System.Data.Odbc.OdbcCommand();
		command.Connection = self.connection;
		print(CreateTblSQL_Main);print()
		command.CommandText = CreateTblSQL_Main
		try:
			reader1 = command.ExecuteNonQuery();
			self.connection.Close()
		except Exception as e:
			print(CreateTblSQL_Main);print()
			print(e)
	def __init__(self,xlsfilename):
		self.xlsfilename= xlsfilename
		pattern2='''Driver={:};Provider=Microsoft.ACE.OLEDB.12.0;DBQ={:};Readonly=false;IgnoreCalcError=true;AllowFormula=false;Extended Properties="Excel 12.0 Xml;HDR=YES;EmptyTextMode=NullAsEmpty;IMEX=1";'''
		connectionString = pattern2.format('{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}',xlsfilename )
		# print( connectionString )
		print(self.xlsfilename  )		
		self.connection = System.Data.Odbc.OdbcConnection(connectionString)
		self.connection.Open()

	def createTbl(self,TblName):
		SqlHead ='''CREATE TABLE {}  ( '''.format(TblName)
		SqlTail=''' ) '''
		SqlBody=""
		dBFieldINfoStr= titlesCls.getfieldTypeMappings()
		self.fieldinfos = dBFieldINfoStr.split()
		cnt= len( self.fieldinfos) 
		for idx in range( 0,cnt,2):
			SqlBody += self.fieldinfos[idx] +" " +self.fieldinfos[idx+1] +","
		SqlBody =SqlBody[:-1]
		CreateTblSQL_Main =SqlHead  + SqlBody + SqlTail
		
		command = System.Data.Odbc.OdbcCommand();
		command.Connection = self.connection;

		print(CreateTblSQL_Main);print()
		command.CommandText = CreateTblSQL_Main
		try:
			reader1 = command.ExecuteNonQuery();
			self.connection.Close()
		except Exception as e:
			print(CreateTblSQL_Main);print()
			print(e)
			


	def getInsertSQL(self,nvPairDict,TblName):
		sqlhead = "insert into {} ( ".format(TblName)
		sqlFileds =" "
		sqlValus=" ) values ( "
		sqlTail=" )"
		
		for key in nvPairDict.keys():
				if key.strip() == "":
					continue
				if key.strip()  not in self.fieldinfos:
					continue 
				v = nvPairDict[key]

				if v==None :
					continue
					v="无效"
				else:
					v = "{}".format(v) 
				if v=="":
					continue
					v="无效"
				if key in self.DataFields:
					if  v == "无效":
						v="0.0"
					v=v.replace(",","").replace("¥","").replace("￥","").replace("?","")
					if v[:1] == "(" and v[-1:] == ")":
						v="-"+v[1:-2]
					# print( key ,v,end=";" )
				sqlFileds += key +","
				if (key,"numeric") in titlesCls.filedmappings:
						# print(key,"numeric",v.strip())
						sqlValus  +=" "+ v.strip()  +","
				else:
						sqlValus  +="'"+ v.strip()  +"',"
		sqlFileds = sqlFileds[:-1]
		sqlValus  = sqlValus[:-1]
		# print()

		return sqlhead + sqlFileds + sqlValus +  sqlTail
	def Adddata(self,nvPairDict,TblName):
		try:
			SqlStr  = self.getInsertSQL(nvPairDict,TblName)
			# print(SqlStr)
			command = System.Data.Odbc.OdbcCommand();
			# self.connection.Open()
			command.Connection = self.connection; ##############
			command.CommandText = SqlStr
			try:
				reader1 = command.ExecuteNonQuery();
           
			except Exception as e:
				print(SqlStr);print()
				print(e)
			# self.connection.Close()
			command.Connection = None
			command.Dispose()
			# self.connection.Close()
		except  ValueError as e:
			print("增加数据发生错误，数据库已经回退，具体数据如下:\n",e)
			# self.conn.rollback()
			return 
		return 


class xlstblWriteExistTbl(xlstblWrite):
	def createTbl(self,TblName):
		SqlHead ='''CREATE TABLE {}  ( '''.format(TblName)
		SqlTail=''' ) '''
		SqlBody=""
		dBFieldINfoStr= titlesCls.getfieldTypeMappings()
		self.fieldinfos = dBFieldINfoStr.split()
		cnt= len( self.fieldinfos) 
		for idx in range( 0,cnt,2):
			SqlBody += self.fieldinfos[idx] +" " +self.fieldinfos[idx+1] +","
		SqlBody =SqlBody[:-1]
		CreateTblSQL_Main =SqlHead  + SqlBody + SqlTail
		
		# command = System.Data.Odbc.OdbcCommand();
		# command.Connection = self.connection;

		# print(CreateTblSQL_Main);print()
		# command.CommandText = CreateTblSQL_Main
		try:
			# reader1 = command.ExecuteNonQuery();
			self.connection.Close()
		except Exception as e:
			# print(CreateTblSQL_Main);print()
			print(e)
			

class mainCls():
	@classmethod
	def main(cls):
		adDatas = xlstbl2AD.xls2AD( cfgfile.updateFiles,cfgfile.sheetNameHead )
		print( len(adDatas) )
		return
		for row in adDatas:
				for k,v in row.items():
					pass
					print(k,v,end=";",sep="=")
				print()
				print()

# mainCls.main()





