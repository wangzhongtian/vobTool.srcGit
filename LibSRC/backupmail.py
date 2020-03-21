# -*- coding: UTF-8 -*-
from __future__ import print_function

import clr
import System
import System.IO
import sys
def getFullNames(dllname):
    Libpaths = System.Environment.GetEnvironmentVariable("libpath")
    if Libpaths == None:
        Libpaths =""

    path1 = System.Environment.GetEnvironmentVariable("path")
    if path1 == None:
        path1 =""
    libpaths= ".\\;"+path1+";"+Libpaths
    # print(libpaths)
    for p in libpaths.split(";"):
        filename = p+"\\"+dllname
        if ( System.IO.File.Exists(filename) ):
            print("---Load Dlls:" ,filename)
            return filename
    print( "Can not find Lib:" ,dllname)
    return None

dllname=r"""AppAssembly.dll""" ;
clr.AddReferenceToFileAndPath( getFullNames(dllname) )

dllname=r"""STDLIBalls.DLL""" ;
clr.AddReferenceToFileAndPath( getFullNames(dllname) )

from System.Runtime.InteropServices import Marshal

import re
import string
import os
import sys
# import time
import re
# sys.exit()
from os.path import join, getsize
import io
from collections import  *
import System

import Odbc2ADLib

class glCls:    
    g_邮箱名称=""
    gFolderBoxLst=()
    gExpireDays =1145
    rubbishFolderObject = None  
    ExceptFolder ="tt;RSS 源;草稿;垃圾邮件;已删除邮件;建议的联系人;对话操作设置;新闻源;快速步骤设置".split(";")
  

constAliasName= "Alias.MailBack.xls.lnk" 
#MailBackPath = "\\current\\System\\maillog\\"
ForReading= "r" #以只读方式打开文件。不能写这个文件。 
ForWriting ="w" 
ForAppending="a"

strCodingUnicode="utf-8"

TristateUseDefault =-2 #使用系统默认值打开文件。 
TristateTrue =-1 #以 Unicode 方式打开文件。 
TristateFalse = 0 #以 ASCII 方式打开文件。 

olMSG=3  #shouled be 3
olMSGUnicode = 9
constUserProperty_isSaved="LastBackVersion"
olYesNo = 6
olDateTime = 5
olText = 1
#
olMail= 43
olNote = 44
olMeetingCancellation =54
olMeetingRequest = 53 
olMeetingResponseNegative =55
olMeetingResponsePositive =56
olMeetingResponseTentative =57

olAppointment = 26
olJournal = 42
olTask =48
olTaskRequest =49
olTaskRequestAccept =51
olTaskRequestDecline =52
olTaskRequestUpdate =50
olAddressEntries =21
olAddressEntry =8
olAddressList =7
olAddressLists =20
olContact =40


olNotSave=1
olSave=2


from datetime  import *
class fileProc():
    @classmethod
    def getDTFileName(cls,Prefix,filenameTpl):
        tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xlsx"
        curDt = datetime.now()
        filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
        return filename
    @classmethod
    def tplProc(cls,rootPath,PreFix,A2Prefix,A3Prefix,tplXlsFileName="基础数据\\reportTpl.xlsx"):
        # tplXlsFileName ="基础数据\\reportTpl.xlsx"
        # rootPath = cfgfile.root文件夹
        xlsTplName = os.path.join(rootPath,tplXlsFileName)
        tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
        xlstgrName =os.path.join( rootPath,tgrFile)
        srcfileObj =System.IO.FileInfo(xlsTplName )
        srcfileObj.CopyTo( xlstgrName )
        return xlstgrName
def getComAppObject(ComAppName):
    # AppObj = win32com.client.Dispatch(ComAppName)
    AppObj=  Marshal.GetActiveObject(ComAppName)
    AppObj =AppObj.Application
    #AppObj.Visible = True
   # AppObj=Outlook.ApplicationClass()
    return AppObj

def  ItemNew(ItemInfo): 
    # print("N;;;;;;;;;;;;;;;;;;;;;;;", ItemInfo )
    if (ItemInfo.srcItem.Class == olMeetingResponseNegative) or \
        (ItemInfo.srcItem.Class ==  olMeetingResponsePositive  )  or  \
        (ItemInfo.srcItem.Class ==  olMeetingResponseTentative ):
       print( "New type of file met,Title is",ItemInfo.srcItem.Subject , "Item Class type: ",ItemInfo.srcItem.Class )
       return False
    # print("----",ItemInfo.FileFullName )
    a= False
    for f1 in glCls.folders:
   		 a = os.path.exists( os.path.join(f1,ItemInfo.leibieFilename) );
   		 if a == True :
   		 	break;
    return not a

def  time2Str( d_t ):
    timestr =str(d_t.Year)+"-"+str(d_t.Month)+"-"+str(d_t.Day) + "_"+str(d_t.Hour) \
      +  "-"+ str(d_t.Minute) + "-"+ str(d_t.Second) +"-"+str(d_t.Millisecond  )
    return timestr

def OADT2PythonDT(d_t):
    return d_t

# 1970 年 1 月 1 日 间的 毫秒数
def getTimeLength(temTime):
    ot= System.DateTime(1970,1,1,0,0,0)
    delta= temTime-ot
    lenTime=  delta.TotalMilliseconds 
    return lenTime
    
class temDataStruct:
    a=0
class clsItemInfo:
    adMailDictObj = []
    ItemCnt=0
def getAllSubfoldes(rootpath ):
	a = os.listdir(rootpath)
	folders=[]
	for b in a:
		b1= os.path.join( rootpath,b)
		if os.path.isdir( b1 ):
			# print(b1 )
			folders += [b1]
	folders = sorted( folders,reverse= True)	
	for i in folders:
		print(i)	
	return folders
def main_entry(MailBackPath):
    try:
        olApplication = getComAppObject("Outlook.Application")
        myNameSpace = olApplication.GetNamespace("MAPI")
    except:
        print("--Please open Outlook. Error,exit out")	
        import sys
        sys.exit()
        return
    glCls.folders = getAllSubfoldes(MailBackPath)
    # return 
    t= datetime.now()
    t_str= "{:04d}_{:02d}_{:02d}_{:02d}.{:02d}.{:02d}".format(t.year,t.month,t.day,t.hour,t.minute,t.second)
    try:
        for rootFolder in  myNameSpace.Folders:
            # print("processing Folder: ", rootFolder.Name)
            ItemInfo=clsItemInfo()
            ItemInfo.ItemCnt =0
            ItemInfo.DumedCnt=0
            ItemInfo.MailBackPath = MailBackPath

            ItemInfo.StartDtimeStr = t_str
            print(ItemInfo.StartDtimeStr)
            ItemInfo.rootFolder = rootFolder 
            rubbishFolder = "ttttt"

            try:
                glCls.rubbishFolderObject = ItemInfo.rootFolder.Folders(rubbishFolder)
            except:
                glCls.rubbishFolderObject = ItemInfo.rootFolder.Folders.Add(rubbishFolder)
                # print("New Folder Created." ,rubbishFolder)
            for folderObj in ItemInfo.rootFolder.Folders:
                folderName = folderObj.Name
                if folderName not in glCls.ExceptFolder +[rubbishFolder]:
                    # PrintInfo( "folder found,Procesing Folder："+ folderName)
                    dealSubFolder(ItemInfo,folderName,folderObj)


            # break
    except:
        print("------root Foleder Processing  Error",rootFolder.Name)
        sys.exit()
    #olApplication.Quit() 
    gFieldName="发件人 修改时间 标题 文件名 全名"
    xlstgrName = fileProc.tplProc(glCls.rootpath,"G118" ,"邮件备份","王中天","reportTpl.xlsx")
    Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "汇总表",clsItemInfo.adMailDictObj ,gFieldName.split())            
            

def  dealSubFolder( ItemInfo,subFolderName,srcFolderObj  ) :
    # print("343425")
    try:
        ItemInfo.srcFolder= subFolderName
        print( "Processing belowBox，begin" , ItemInfo.srcFolder, " in ", ItemInfo.rootFolder.Name )
        DEAL_FolderBox( ItemInfo ,srcFolderObj )
    except:
        print( " -----Processing belowBox ，Unexception Error Happened !!，total Dumped Count is " , ItemInfo.DumedCnt,",end"  )
        sys.exit()        
    PrintErrinfo("Processing belowBox，completed!   ..." + ItemInfo.srcFolder )

def AppendLog2File(ItemInfo):
	maildictObj = dict()
	maildictObj ["发件人" ] = ItemInfo.SenderName;#print( "发件人" ,maildictObj ["发件人" ])
	maildictObj ["修改时间" ] = time2Str( ItemInfo.MidificationTime);#print( "修改时间" ,maildictObj ["修改时间" ])
	maildictObj ["标题" ] = ItemInfo.title;#print( "标题",maildictObj ["标题" ])
	maildictObj ["文件名" ] = ItemInfo.FileName;#print("文件名", maildictObj ["文件名" ])
	maildictObj ["全名" ] = ItemInfo.Relativepath;
	print("全名", maildictObj ["全名" ])
	clsItemInfo.adMailDictObj += [maildictObj]

def  mailSaveProc( srcItem, filenameStr  ):
    try :
         t = srcItem.SaveAs(filenameStr , olMSGUnicode )#olMSG  olMSGUnicode)
    except:
        print("mail Save Error",filenameStr )
        # return 
        try:
            srcItem.Move( glCls.rubbishFolderObject )
        except:
            print("mail ---Move-- Error",filenameStr )
        #srcItem.Delete() ##???????????
def postProcMail( srcItem ):
    try:
        d= System.TimeSpan(glCls.gExpireDays,0,0,0,0)
        if srcItem.Class not in( olMail,olNote ,olAppointment,olJournal,olTask, ) :
             return 
        d_t = System.DateTime.Now
        d_t =d_t - d
        mailcTime = OADT2PythonDT(srcItem.CreationTime) 
        if  mailcTime < d_t :
            # print( "Mail can be deleted !")
            srcItem.Delete()
    except:
            print("Error in cal datetime!!!")

def  DEAL_Contact(ItemInfo,srcFolderObj):
    if ItemInfo.srcFolder != "联系人":
        return ;
    mysrcFolder = srcFolderObj 
    cnt =  mysrcFolder.Items.count
    print( "Total Msg count is:",cnt )
    if( cnt<= 0):
        PrintInfo( "No Item found")
        return;
    i=0;
    ItemInfo.RowOffset= 0
    try:
        for ItemInfo.srcItem in mysrcFolder.Items:
            i = i+1
            try:
                print(ItemInfo.srcItem.FullName,end= ";" )
                print(ItemInfo.srcItem.Email1Address,end= ";" )
                print(ItemInfo.srcItem.Email2Address,end= ";" )
                print(ItemInfo.srcItem.CompanyName,end= ";" )
                print(ItemInfo.srcItem.MobileTelephoneNumber,end= ";" )                
                print( i )
            except:
                pass
    except:
        print(  "-???????????????---- Dump to : fail," +ItemInfo.srcFolder )
        
def  DEAL_FolderBox(ItemInfo,srcFolderObj):
    # DEAL_Contact(ItemInfo,srcFolderObj);
    # return;
    mysrcFolder = srcFolderObj 
    # ItemInfo.itembackpath = ItemInfo.MailBackPath + ItemInfo.srcFolder + "\\"
    ItemInfo.itembackpath = ItemInfo.MailBackPath  +"\\"
    cnt =  mysrcFolder.Items.count
    print( "Total Msg count is:",cnt )
    if( cnt<= 0):
        PrintInfo( "No Item found")
        return;
    i=0;
    ItemInfo.RowOffset= 0
    ItemInfo.srcItem= mysrcFolder.Items(cnt)
    curtime =System.DateTime.Today
    while( ItemInfo.srcItem!=None ):
        ItemInfo.title =""
        ItemInfo.ItemCnt=ItemInfo.ItemCnt+1
        make_filename(ItemInfo,i )
        # print( ".make name ok...................." )
        if  ItemNew(ItemInfo) == True :
            filenameStr = os.path.abspath(ItemInfo.FileFullName)
            try:
                mailSaveProc(ItemInfo.srcItem ,filenameStr )
                # ItemInfo.FileFullName = os.path.join( "RootPATHZHANWEI",ItemInfo.Relativepath)
                AppendLog2File(ItemInfo)
                ItemInfo.DumedCnt=ItemInfo.DumedCnt+1
                i=i+1
            except:
                print(  "----- Dump to : fail," +filenameStr )
            ItemInfo.DumedCnt=ItemInfo.DumedCnt+1
            i=i+1

        else:
            #print( "file already introduced out...",cnt )
            pass
        postProcMail(ItemInfo.srcItem )
        cnt=cnt-1
        if( cnt< 1):
            break;
        ItemInfo.srcItem= mysrcFolder.Items(cnt)
def getsrcMsgModifyTime(MsgItemObject):
    try:
      Time_1 =OADT2PythonDT( MsgItemObject.ReceivedTime )
    except:
      Time_1=None
    try:
        if(  Time_1  == None ) :
            Time_1  =OADT2PythonDT( MsgItemObject.LastModificationTime)
    except:
      # Time_1=datetime.datetime.today()
      Time_1 =System.DateTime.Today
      #print( "New receiv time---++-:", "None" )
    return Time_1
def makdirs(fullpath):
	try :
		os.makedirs(fullpath)
	except:
		pass

def makdirs_1(fullpath):
    fullpath=os.path.abspath(fullpath)
    if os.path.exists(fullpath) == True :
        return True
    drive=os.path.splitdrive(fullpath) [0]+"\\"
    #print ("drive is:",drive)
    # print(  "Will Make file path ,if it is not exist:",fullpath )
    try:
        FileRootPath=os.path.split(fullpath)[0]
        # print( "Upper folder is:",FileRootPath )
        if os.path.exists(FileRootPath) == True :
           os.mkdir(fullpath)
        elif drive != FileRootPath:
           makdirs(FileRootPath) 
        else:
            return False
        return True
    except:
        # print( "mk msg folder Error!" )
        return False
   
def make_filename(ItemInfo,i ):
    ItemInfo.row= ItemInfo.RowOffset +i
    rePattern= "[\\\\/:\*\?\"\<\>\|\t\r\n']" #,"gim"   # \ / : * ? " < > | \r\n
   # print( "==========Here",ItemInfo.RowOffset,i )
    dtTime= OADT2PythonDT(ItemInfo.srcItem.LastModificationTime )
    # print( "==========Here")    
    tl=  getTimeLength(dtTime)

    Idstr="{:.1f}".format(tl)
    #print( "-==========Here")    
    ItemInfo.SenderName="" 
    ItemInfo.MidificationTime =getsrcMsgModifyTime(ItemInfo.srcItem) #new Date(Time_1)
    # print ("--------Time length is :",tl)
    # ItemInfo.title=""
    ItemInfo.isAction =""   #+ ItemInfo.srcItem.FlagIcon
    #print ItemInfo.srcItem.Class 
    if ItemInfo.srcItem.Class== olMeetingRequest  or \
         ItemInfo.srcItem.Class== olMail :
           ItemInfo.SenderName=  ItemInfo.srcItem.SenderName 
    elif  ItemInfo.srcItem.Class== olAppointment :
            ItemInfo.SenderName=  ItemInfo.srcItem.Organizer
    else : 
        pass
   # print ItemInfo.srcItem.Class 
    title=ItemInfo.srcItem.Subject
    # title=title.replace("'","''").replace('"','""')
    ItemInfo.title=re.sub(rePattern,"_",title) #+";"
    #print( "subject is:",ItemInfo.srcItem.Subject )
    strTitle= ItemInfo.title
    if( len(strTitle) >200):
      strTitle=strTitle[0:200]
      print( "too long ,Truncated." )
      pass

    ItemInfo.FileName=""+strTitle +"_"+ Idstr+ ".msg" 
    fileMianName=str(ItemInfo.MidificationTime.Year) +"-"+str(ItemInfo.MidificationTime.Month)
    # print( "-----------too long ,Truncated." )
    fullpath=os.path.join( ItemInfo.itembackpath ,ItemInfo.StartDtimeStr ,fileMianName,ItemInfo.srcFolder)
    makdirs(fullpath)
    ItemInfo.datePath=fileMianName
    ItemInfo.FileFullName = os.path.join(fullpath,ItemInfo.FileName)
    ItemInfo.Relativepath= os.path.join( "RootPATHZHANWEI",ItemInfo.StartDtimeStr ,fileMianName,ItemInfo.srcFolder, ItemInfo.FileName)
    ItemInfo.leibieFilename= os.path.join( fileMianName,ItemInfo.srcFolder, ItemInfo.FileName)
def  PrintErrinfo(info):
	print( info+"\r\n" )
def  PrintInfo( info ):
    print( info +"\r\n"  )


def RestoreXlsFromMsgFile( filename,rootPath ):
    os.startfile(filename)
    #WshShell.Run("E:\\Current\\System\\a\\TDProject\\WScript\\mailbackup\\182.msg",2, false)
    isexit= False
   # filetime=datetime.datetime.today()
    loopCnt=1000
    while( isexit == False ):
        loopCnt=loopCnt-1
        if(loopCnt < 1):
           break
        print( " waiting..",loopCnt )
        time.sleep( 0.01)
        olApplication =None
        try:
            olApplication= getComAppObject("Outlook.Application")
        except:
            olApplication=None
            pass
        if(olApplication != None ):
            Insp=olApplication.ActiveInspector()
            if( Insp != None ) :
                isexit= True
                #===============================================================
                #===============================================================
                filetime=getsrcMsgModifyTime(Insp.CurrentItem)
                Log1MsgFile(Insp.CurrentItem,filename,rootPath)
                olApplication=None 
                Insp=None 
                fileMianName=str(filetime.year) +"-" +str(filetime.month)        
                oldName=  filename
                newPath=rootPath+"\\"+fileMianName+"\\"
                if os.path.exists(newPath) == False :
                    os.mkdir(newPath)
                FilePureName=os.path.split(filename)[1]
                print( oldName,newPath+ FilePureName )
                #time.sleep(0.2)
                os.rename(oldName,newPath+ temDataStruct.FileName)
                #win32api.MoveFile(oldName,newPath+ FilePureName)

def Log1MsgFile(srcItem,filename,rootPath): 
    rePattern= "[\\\\/:\*\?\"\<\>\|\t\r\n]" #,"gim"   # \ / : * ? " < > | \r\n
    #print filename 
    dtTime= OADT2PythonDT( srcItem.LastModificationTime)  
    #print "dt timeis ",dtTime 
    Idstr=""+str(getTimeLength(dtTime))  
    #print "receiv time----:"+ Idstr 

    #print("Modify time----:"+ Idstr)
    temDataStruct.SenderName="" 
    temDataStruct.MidificationTime =getsrcMsgModifyTime(srcItem) #new Date(Time_1)
    temDataStruct.title=""
    temDataStruct.isAction =""   #+ ItemInfo.srcItem.FlagIcon

    #print(srcItem.Class)
    if srcItem.Class== olMeetingRequest  or \
         srcItem.Class== olMail :
           temDataStruct.SenderName=  srcItem.SenderName 
    elif  srcItem.Class== olAppointment :
            temDataStruct.SenderName=  srcItem.Organizer
    else : 
        pass
    #print(srcItem.Class)
    temDataStruct.title=srcItem.Subject #+";"
    #print("subject is:",srcItem.Subject)
    strTitle= temDataStruct.title
    #print("hello :----:"+ strTitle)
    if( len(strTitle) >200):
      #strTitle= ( datetime.datetime.now())
      strTitle=strTitle[0:200]
      print("too long ,Truncated." )
      pass 
    temDataStruct.FileName=os.path.split(filename)[1]
    temDataStruct.FileFullName = rootPath+"\\"+filename   
    filetime=temDataStruct.MidificationTime
    
    fileMianName=str(filetime.year) +"-" +str(filetime.month)
    #print "xc cvc", temDataStruct.MidificationTime , filetime 
    XmlFileName=rootPath+fileMianName +".txt"
    dataFileFs =System.IO.File.AppendText( XmlFileName   ) #io.open(XmlFileName, ForAppending,1,strCodingUnicode)  
    xlsProLib.MakeXmlSheetData_2_Fs(dataFileFs ,temDataStruct)
    dataFileFs.close()
    isClosed=False
    while(isClosed ==False):
        try:
            srcItem.Close( olNotSave)
        except:
            pass
        else:
            isClosed=True
        time.sleep(0.1)

def  DEAL_FileFolder(fileFolderName,rootPath):
    #NaList=("邮件撤回失败", "撤回","未传递","未送达")
    for root, dirs, files in os.walk(fileFolderName):
        for name in files:
            if(name[-4:]==".msg"):
                if "撤回"==name[0:2] or "未传递"==name[0:3] or "未送达"==name[0:3]  \
                or "邮件撤回"==name[0:4] :
                    continue
                filename=join(root,name)
                #print ("Proc File:"+filename)
                RestoreXlsFromMsgFile( filename,rootPath )
            #
def FileRestore_Main():
    msgFolder=r"G:\maillog\2009\tem\\"
    # all the xlc file and tgrfilepath will in this direcotry
    rootPath=r"G:\maillog\2009\收件箱\\"
    DEAL_FileFolder(msgFolder,rootPath)




if __name__ == "__main__":

    # glCls.ExceptFolder ="tt;RSS 源;草稿;垃圾邮件;已删除邮件;建议的联系人;对话操作设置;新闻源;快速步骤设置".split(";")

    # for fol in glCls.ExceptFolder:
    #     print( fol ) 
    glCls.gExpireDays = 160
    glCls.rootpath = r"D:\wangzht\maillog\\"  
    main_entry(glCls.rootpath)


    # maildictObj = dict()
    # maildictObj ["发件人" ] = "Test";#print( "发件人" ,maildictObj ["发件人" ])
    # maildictObj ["修改时间" ] ="191223123-12321";#print( "修改时间" ,maildictObj ["修改时间" ])
    # maildictObj ["标题" ] = "''ewqreqwr''";#print( "标题",maildictObj ["标题" ])
    # maildictObj ["文件名" ] = "3213";#print("文件名", maildictObj ["文件名" ])
    # maildictObj ["全名" ] = "12321321";#print("全名", maildictObj ["全名" ])
    # clsItemInfo.adMailDictObj += [maildictObj]

    # for d in clsItemInfo.adMailDictObj:
    #     for k,v in d.items():
    #         print( k,v)
    #     print()
