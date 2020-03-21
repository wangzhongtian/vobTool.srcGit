#-*- coding:utf-8 -*-
#-*- coding: UTF-8

from __future__ import print_function
import clr
# dllname= r"C:\z\_a\nsbd\调试报告\src\sft.ipy\STDLIB.DLL"
# dllname=r"""STDLIBAlls.DLL""" ;clr.AddReferenceToFileAndPath( dllname)
# dllname=r"""AppAssembly.dll""" ;clr.AddReferenceToFileAndPath( dllname)
import System
import System.IO
import System.DateTime
import sys
import LoadDlls

dllname=r"StdLibALLs.dll";clr.AddReferenceToFileAndPath( LoadDlls.getFullNames(dllname) )

import os
import datetime
import shutil
import gzip
import sys
from glCls import * 
import codecs

glbUNIT = 10000.0*1000.0
glb_rongrenshu= 1.0  # in seconds
class dtDelta:
    def __init__(self,basictimeStr="1970年1月1日 0:0:0",timeZoneoffset=8.0*3600):
        self.jizhunDTstr = basictimeStr
        self.tzOffset = timeZoneoffset
        self.jizhunDT= System.DateTime.Parse(self.jizhunDTstr)
        self.jizhunoffset = float(self.jizhunDT.Ticks)/glbUNIT+ self.tzOffset
        # print("------------jizhun offsetis：---------------" ,self.jizhunoffset,self.jizhunDT.Ticks )# FileSystemInfo.LastWriteTime)
    def getSecondsElapseUTC(self,SysDatetimeObj):
        a=float(SysDatetimeObj.Ticks)/glbUNIT - self.jizhunoffset 
        i= int(a)
        # print("getSecondsElapseUTC--:",i,SysDatetimeObj.Ticks)
        # if i == 1447204343:
        	# print("====",a,i,SysDatetimeObj.Ticks-self.jizhunDT.Ticks- 8*3600*10000*1000)
        return a#int(a)
    def getfileElapsedTick(self, filename):
        fi = System.IO.FileInfo( filename)
        #local
        fNet1 = fi.LastWriteTime #Utc
        return self.getSecondsElapseUTC( fNet1)
    def getSystime(self,mtimeOffset):
        # print( systime.Ticks  )
        Ticks = (mtimeOffset +  self.jizhunoffset )* glbUNIT 
        systime = System.DateTime(int(Ticks) )
        # a=float(SysDatetimeObj.Ticks)/glbUNIT - self.jizhunoffset 
        # i= int(a)
        # print("getSystime：：：", systime )
        # if i == 1447204343:
        	# print("====",a,i,SysDatetimeObj.Ticks-self.jizhunDT.Ticks- 8*3600*10000*1000)
        return systime #int(a)  
g_dtDeltaObj = dtDelta()  
   
def gettimeinSeconds( timestr):
	f= float( timestr)
	i = int(f)
	# if "1434426500.00" == timestr:
	# 	print("====" ,f, timestr,i)
	return f

def getRelativepath(pathname,rootpath):
        rp=os.path.relpath(pathname, start=rootpath) 
        return rp.replace("\\","/")
def setFileAttrCleared_old( name ):
       NewFlag  = 0x00
       # win32file.SetFileAttributesW( name,NewFlag) 
def setFileAttrCleared( name ):
       NewFlag  = 0x00
       fi = System.IO.FileInfo( name)
       # win32file.SetFileAttributesW( name,NewFlag) 
       fi.Attributes = fi.Attributes & ~System.IO.FileAttributes.Hidden & ~System.IO.FileAttributes.ReadOnly
def isLink( name ):
        #System.IO.FileAttributes
        attributes1 = System.IO.File.GetAttributes( name )
        if ((attributes1 & System.IO.FileAttributes.ReparsePoint) == System.IO.FileAttributes.ReparsePoint):
            return True
        else:
            return False
                 
def removePath( path ):
        for root, dirs, files in os.walk(path, topdown=False):
            for name in files:
                f1 =os.path.join(root, name)
                setFileAttrCleared(f1 )
               # if True == os.path.exists( f1):
                os.remove( f1 )
            for name in dirs:
                os.rmdir( os.path.join(root, name) )
        os.rmdir( path )

class  fileinfoCls:
    def __init__(self,path="",isDir=None ,FileID = -1):
        self.name=path.replace("\\","/")
        if isDir != None:
             self.Isdir = isDir
        else:
            if( path != ""):
                self.Isdir = os.path.isdir(path)
            else:
                self.Isdir =  None   
        self.mtime = 0.0
        self.FileID = FileID
        self.LastLableTime =glCls.strcurtime+"_"+glCls.CurReleaseStr
       
        if( path != ""):
            try:
                self.mtime = g_dtDeltaObj.getfileElapsedTick(path)
            except:
                sys.exit()
                # print("get File mtime Error...........",path)
                self.mtime =0.0
        self.statu ="删除"
##        在新的库中可能的状态为:
##        删除，新库中不存在；
##        新建，新库中新建的文件；
##        不变：新库中文件与上一版本中相同
##        修改：新库中文件已经被修改
    def setStatu(self,statuStr):

        self.statu =statuStr
    def setNewFileID(self):
        self.FileID = getFileNumber()
    def isSame(self,Obj):
        # print( " in isSame:")
        libFilename= os.path.abspath(self.name).replace("\\","/")
        sysfile = os.path.abspath( Obj.name ).replace("\\","/")
        if libFilename != sysfile :
               print(self.name,"--??????--",Obj.name ,"\r\n",libFilename , sysfile)
               sys.exit()
               return False
        #print(sysfile, libFilename )
        # print(self.Isdir,  Obj.Isdir  )
        if (self.Isdir !=  Obj.Isdir ):
            # print()
            # print( " dir Property is not SAME")
            # print(self.name,"--??????--",Obj.name)
            # sys.exit()
            return False
        if  self.Isdir == True:
               return True
     	# rongrenshu= 2.0/100.0 #2.0/10000.0
        rongrenshu= glb_rongrenshu 
        # print("in IsSame:",Obj.mtime,Obj.mtime )
        b = Obj.mtime > self.mtime -rongrenshu and Obj.mtime < self.mtime + rongrenshu
        if b == False: 
            pass
            # print("--??--","{:.10f}".format(Obj.mtime),"{:.10f}".format(self.mtime -rongrenshu),"{:.10f}".format(self.mtime + rongrenshu)  )
            # sys.exit()
        # print("out  isSame")
        return b
        
    def fileArchived(self):
       return 32

    def setFileArchiveCleared(self):
       archived = win32con.FILE_ATTRIBUTE_ARCHIVE
       not_archived  = ~archived
       file_flag = win32file.GetFileAttributesW( self.name )
       NewFlag  = file_flag & not_archived
       win32file.SetFileAttributesW(self.name,NewFlag)           
    def fileinfoStr(self):
        if int(self.mtime) == 0:
            stra= str(a.tm_year)+"-"+str(a.tm_mon )+"-"+str(a.tm_mday )+"-"+str(a.tm_hour )+"-"+str(a.tm_min)+"-"+str(a.tm_sec)
            #stra += "::"+ str(self.mtime) #"{:.9f}".format(f)
            stra += "::"+ "{:.9f}".format(self.mtime)
            return stra
        else:
            return None
    def save1Field2FS(self, fs1,dataStr,Sep=","):
        fs1.write( dataStr )
        fs1.write( Sep)

    def save2FS(self, fs1):
        self.save1Field2FS(fs1,self.statu,Sep=",")
        self.save1Field2FS(fs1,self.LastLableTime ,Sep=",")
        self.save1Field2FS(fs1, str(self.FileID ) ,Sep=",")
        # self.save1Field2FS(fs1, str(self.mtime) ,Sep=",")
        self.save1Field2FS(fs1, "{:.9f}".format(self.mtime),Sep=",")
 
        isDirStr = "F"
        if self.Isdir == True:
            isDirStr = "D"
        
        self.save1Field2FS(fs1, isDirStr ,Sep=",")
        relName = getRelativepath(self.name,glCls.fullsourceFolder  )
        self.save1Field2FS(fs1,relName,Sep="")

    def ldfromStr(self,lineStr,defaultStatu):
        if( lineStr.strip( ) == "" ):
             return
        # print( "ldfromStr 1 ")
        i = lineStr.find( ",")
        st = lineStr[:i]
        self.laststatu =  st
        self.statu=st
        if( defaultStatu !=None ):
            self.statu  = defaultStatu
            
        # print( "ldfromStr 2 ")
        j=lineStr.find( ",",i+1)
        st = lineStr[i+1:j]
        self.LastLableTime  =  st
        # print( "ldfromStr 3 ")        
        i=j+1
        j=lineStr.find( ",",i+1)
        st = lineStr[i:j]
        self.FileID = int( st ,10)

        # print( "ldfromStr 4 ")        
        i=j+1
        j=lineStr.find( ",",i+1)
        st = lineStr[i:j]
        # print("===",st )
        self.mtime = gettimeinSeconds( st)
        # print("===",self.mtime )
        # print( "ldfromStr 5 ")        
        i=j+1
        j=lineStr.find( ",",i)
        st = lineStr[i:j]
        if st == "D" :
            self.Isdir  = True
        else:
            self.Isdir  = False
        # print( "ldfromStr 6 ")            
        name1 = lineStr[j+1:].strip( )
        if  name1[:1] == "\\" :
            name1 =name1[1:]
        # print("----------------",glCls.fullsourceFolder)
        self.name =os.path.join( glCls.fullsourceFolder, name1 ).replace("\\","/")
        # self.name =self.name
        # print( "ldfromStr 7 ")        
        if self.FileID  > glCls.FileNumer:
            glCls.FileNumer =  self.FileID
        # print( "ldfromStr 8 ")   
    def genlayerFolder(self):
        i = self.FileID
       # tgrpath = str(0) ;#str( i/1000 )
        key = 0xFFF
        bits = 12
        L1 = ( i & (key <<(12*3) )  )>>(12*3)
        L2 = ( i & (key <<(12*2) )  ) >>(12*2)
        L3 = ( i & (key <<(12*1) )  )>>(12*1)
        return os.path.join( str( L1),str(L2),str(L3) )
            
    def genFileName(self):
        tgrpath = self.genlayerFolder()
        tgrfolder = os.path.join( glCls.tgrVobroot,tgrpath)
        print("--->: ",tgrfolder )        
        if os.path.exists( tgrfolder  ) != True:
            print("--->: ",tgrfolder )
            os.makedirs(tgrfolder )
        import shutil
        #tgrname = str(self.FileID)
        tgrname = os.path.join( tgrfolder, str(self.FileID)) 
        #shutil.copyfile( self.name, tgrname)
        CompressFile( self.name, tgrname)
    def genRestoreSorcFolder(self):
        #glCls.最新VobFile
        pass###
    def restorefile(self):
        tgrpath = self.genlayerFolder()
        vobroot =   glCls.VobFolderRoot
        tgrfolder = os.path.join( vobroot,self.LastLableTime,tgrpath).replace("\\","/")
        import shutil
        tgrname = os.path.join( tgrfolder, str(self.FileID))
        parentPath = os.path.dirname( self.name)#.replace("\\","/") 
        if os.path.exists( parentPath  ) != True:
            # print("----: ",parentPath )
            os.makedirs( parentPath )
        #print(".....",tgrname,self.name  )
        #shutil.copyfile(  tgrname,self.name)
        deCompressFile(  tgrname ,self.name ) #.replace("\\","/")
        # mt = int( self.mtime)
        # mt =  self.mtime
        # print( self.mtime ,"---" ,self.name )
        syst1= g_dtDeltaObj.getSystime( self.mtime )
        filename = os.path.abspath(self.name)
        # print( "----------------------: change time to :",filename ,";",syst1 )
        fi = System.IO.FileInfo( filename)
        #local
        # fNet1 = fi.LastWriteTime #Utc
        fi.LastWriteTime = syst1
        # os.utime(self.name,(mt , mt ) )
    def  SaveFile2Vob(self):
        if self.Isdir == True:
            pass
            #print( "Proc file :",self.name )
        else:
			print( "copy file :",self.name )
			self.genFileName()
			pass
def regFilename( orgName ):
    # print("regFilename 1:" ,orgName)
    name = os.path.abspath( orgName )
    # print("regFilename 2:",name)
    return  name.replace("\\","/")
class fileDictCls:
    def __init__(self):
        self.fileinfo_dict= dict()
    def AddFileinfo( self, fileinfoCls_Obj):
        # self.fileinfo_dict[fileinfoCls_Obj.name.lower()] = fileinfoCls_Obj
        # print("::: in AddFileinfo",fileinfoCls_Obj.name )
        # sys.exit()
        self.fileinfo_dict[fileinfoCls_Obj.name ] = fileinfoCls_Obj
    def delFileinfo(self,  name ):
        try:
            del self.fileinfo_dict[ name ]  #  lower()
        except:
            print("---","Key not find" , name )
        #print( "----",self.fileinfo_dict[ name])
    def  getFileinfo(self,path):
    #    print(path,"in getFileinfo:",len(self.fileinfo_dict ) )

       cnt=0;
       try:
            # print(path )
            for u in self.fileinfo_dict.keys():
                pass
                cnt+=1
                # if u == path :
                #     print( "found ",u )
            v1 = self.fileinfo_dict[path]
            # print(v1," found")
            return  v1 #.lower()
       except:
            print("Exception in getFileinfo ")
            # sys.exit()
            return None
    def printFileInfos(self):
        a = len(self.fileinfo_dict)
        print(a)

    def SaveFileInfos(self,SSCNAme):
        # f1 =gzip.open( SSCNAme,"wt",encoding="utf8")
        f1 =codecs.open( SSCNAme,"wt",encoding="utf8")
        for key in self.fileinfo_dict.keys():
            # for key ,value  in  g工段表_All :
            value = self.fileinfo_dict[key ]
            value.save2FS(f1)
            f1.write("\n")
        f1.close()
    def ldFileInfos(self,SSCNAme,defaultStatu=None):
        if SSCNAme == None:
            print( "空白版本，将建立新版本文件")
            return
        # f1 =gzip.open( SSCNAme,"rt",encoding="utf8") 
        f1 =codecs.open( SSCNAme,"rt",encoding="utf8") 
        # print( f1, SSCNAme,"in ldFileInfos")
        for line in f1: # for key ,value  in  g工段表_All :
            # print(line)
           # value = self.fileinfo_dict[key]
            # print(line,"in ldFileInfos:")
            value = fileinfoCls()
            value.ldfromStr( line,defaultStatu )
            # print( "ldfromStr:",value.name)
            #exit(-1)
            if( value.name == "" ):
                # print(value.name,";" )
                pass
            else:
                # print("name:",value.name)
                name1 = regFilename(value.name )
                # print("=====,归一化文件名",name1 )
                self.fileinfo_dict[ name1 ] = value #.lower()
        f1.close()


class dirSafe:
    def __init__(self):
        self.fileinfoObj =None
        self.fileDictObj = fileDictCls()
        self.exceptFileExt=".pyc ".lower().split()
        self.exceptFolders="__pycache__ builddir".lower().split()
        self.Updated =False
        pass
    
    def dirmain(self,rootdir=None):
        if rootdir == None :
            rootdir = glCls.fullsourceFolder  
        for path in os.listdir( rootdir ):
            filename =os.path.join( rootdir,  path)
            if os.path.isdir( filename ):
                # print("Here!!! " ,rootdir,path)
                if path.lower()  in self.exceptFolders:
                    print( "Skipped:" ,path  )
                    continue
                elif isLink( filename ) :
                    print( "Skipped Link:" , filename )
                    continue
                else:
                    self.dirmain( filename )
                    self.FucntionDir( os.path.join( rootdir,path) ,rootdir,path )

            if os.path.isfile( filename ):
                mainname,ext = os.path.splitext(path)
                if ext.lower() in self.exceptFileExt:
                    pass
                else:
                    self.FucntionFile( os.path.join( rootdir,path),rootdir,path )                    
    def procFile(self ):
        # return 
        self.fileinfoObj.SaveFile2Vob()
    def functionProc(self, fullfilename1,root,subname,isDir):
        fullfilename=fullfilename1.replace("\\","/")
        # print("Cur file/dir  info :",fullfilename,root,subname,"isDir=",isDir)
        self.fileinfoObj = fileinfoCls(fullfilename,isDir,-1)
        v = self.fileDictObj.getFileinfo(fullfilename)
        # sys.exit()
        while True:
            # print("''''",v,fullfilename)
            if( v == None):## folder or file is new created
                print("新建")
                self.fileinfoObj.setNewFileID()
                self.fileinfoObj.setStatu("新建")
                self.fileDictObj.AddFileinfo(self.fileinfoObj)
                self.procFile()
                self.Updated =True
                break
            bo = v.isSame( self.fileinfoObj )
            if bo == True :
            	# print("不变")
                v.setStatu("不变")
                self.fileDictObj.AddFileinfo( v )
            else:
            	print("--修改")
                self.fileinfoObj.setStatu("修改")
                self.Updated =True
                self.fileinfoObj.setNewFileID()
                self.fileDictObj.AddFileinfo( self.fileinfoObj )
                self.procFile()
            break

    def FucntionFile(self, fullfilename,root,subname ):
        glCls.fileCnt +=1
        #print( glCls.fileCnt )
        self.functionProc( fullfilename,root,subname, False)

    def FucntionDir(self, fullfilename,root,subname ):
         self.functionProc( fullfilename,root,subname, True)

    def printFileinfos(self):
        self.fileDictObj.printFileInfos()
    def SaveFileinfos(self):
        print( self.SSCNAme )
        self.fileDictObj.SaveFileInfos(self.SSCNAme )
    def ldFileinfos(self,defaultStatu="删除"):
        try:
            self.fileDictObj.ldFileInfos(glCls.最新VobFile,defaultStatu)
        except:
            print("ldFileinfos,Execption!!,exit out")
            
            return
class dirRestore( dirSafe):
    def __init__(self):
           self.delPaths=[]
           dirSafe.__init__(self )
    def doDelPath(self):
        #print( self.delPaths)
        for pathroot in self.delPaths:
              #print( self.delPaths ) 
              doesExist = os.path.exists(pathroot)
              if doesExist == True:
                     try:
                            v= self.fileDictObj.fileinfo_dict[pathroot]
                     except:
                            v =None
                     if v != None:
                            if v.statu != "C":
                                   print( "Cleared::::",v.name)
                                   del self.fileDictObj.fileinfo_dict[pathroot]
                     Isdir = os.path.isdir( pathroot)
                     if Isdir == True:
                            print( "删除过时或者无效的旧文件（夹）：",pathroot)
                            setFileAttrCleared(pathroot)
                            removePath(pathroot)
                     else:
                            print( "删除过时或者无效的旧文件：",pathroot)
                            setFileAttrCleared(pathroot)
                            os.remove( pathroot)
        self.delPaths=[]
        
    def restore(self,NewValue):
        if NewValue.Isdir  == True:
          if os.path.exists(NewValue.name) == False:
            print( "创建新文件(夹)：",NewValue.name)
            os.makedirs(NewValue.name )
          return
        print( "：",NewValue.name)
        NewValue.restorefile()
       
    def restoreAll(self):
       for key,v in self.fileDictObj.fileinfo_dict.items():
            if v.statu == "C":
                    #print( " C:::::----", v.name,v.Isdir )
                    self.restore( v )

    def delPath(self):
        temName = self.fileinfoObj.name.lower() 
        if  temName in self.exceptFolders:
            print( temName )
            return 
        mainname,ext = os.path.splitext(temName)
        if ext.lower() in self.exceptFileExt: 
            # print( temName )
            return

        self.delPaths.append( self.fileinfoObj.name )

    def functionProc(self, fullfilename1,root,subname,isDir):
        # print("functionProc",  fullfilename1 )
        fullfilename=fullfilename1.replace("\\","/")        
        self.fileinfoObj = fileinfoCls(fullfilename,isDir,-1)
        v = self.fileDictObj.getFileinfo(fullfilename)
        while True:
            if( v == None):##
                print( "库中无此文件：",fullfilename)   
                self.delPath()
                break
            if( v.statu =="删除"):##
                # print( "库中标记为删除的文件：",fullfilename)  
                self.delPath()
                break
            bo = v.isSame( self.fileinfoObj )
            if bo == True :
                #print( "库中版本相同的文件：",fullfilename) 
                pass
            else:
                print( "版本不同：",fullfilename)
                v.setStatu("C")
                self.delPath()
                break
            break
    def dictProc(self):
        print("in dictProc")
        for key ,v in self.fileDictObj.fileinfo_dict.items():
            self.fileinfoObj = fileinfoCls(key,None,-1)
            # if "中线渠道建筑物桩号推算" in  key:
            # 	print( [key])
            # print( key,v)
            # print(key,"--------------:")
            while True:
                if v.statu == "删除":
                     self.fileinfoObj =v  
                     self.delPath()
                    #  print( "库中标记为删除的文件：",key)  
                     break
                bo = v.isSame( self.fileinfoObj )
                if bo == True :
                    # print( "版本 相同 的文件：",key) 
                    pass
                else:
                    v.statu ="C"
                    print( "版本 不同 的文件：",key)
                    print(v.name,v.Isdir ,v.statu,v.mtime)
                    print(self.fileinfoObj.name,self.fileinfoObj.Isdir ,self.fileinfoObj.statu,self.fileinfoObj.mtime)
                    self.delPath()
                break
            # print(key,"::::::--------------:")

        self.doDelPath()
        self.restoreAll(   )
      
class dirSafeBackup(  dirSafe):
    def Save2Vobs( self,vobFolder,sourceFolder,sourceroot):

        init_internal(vobFolder ,sourceFolder ,sourceroot  )  
        self.SSCNAme=glCls.VobFileName
        
        self.ldFileinfos(defaultStatu="删除")
        rfolder = os.path.join(sourceroot,sourceFolder)
        # print("====",rfolder,sourceFolder,sourceroot)
        self.dirmain()#(rfolder ) 
        # print( "codecs")
        if self.haveUpdate() :
            self.printFileinfos()
            self.SaveFileinfos();
        else:
            print(":: Have no update folder or file in DOC folder")
    def saveUpdate(self,*arg):
        try:
            logfile = os.path.join(glCls.tgrVobroot,"update.log" )
            # f1 =codecs.open( SSCNAme,"rt",encoding="utf8") 
            fobj =codecs.open(logfile ,"at",encoding="utf8")
            for i in arg:
                fobj.write( i )
                fobj.write(",")
            fobj.write("\n")
            fobj.close()
        except:
            print("Exception in saveUpdate ")
    def getRelativeName(self,vobj):
        try:
            name =vobj.name
            isDirStr = "F" if vobj.Isdir == True else "D"
            relName = getRelativepath(name,glCls.fullsourceFolder  )
            return isDirStr +","+ relName
        except:
            print("getRelativeName Exception")
    def haveUpdate(self):
        print( "变更汇总：")
        cnt=1
        ret= False
        obj = self.fileDictObj
        curLabel= glCls.strcurtime+"_"+glCls.CurReleaseStr
        for key in obj.fileinfo_dict.keys():
            value = obj.fileinfo_dict[key ]
            if "不变" == value.statu:
                continue

            if "新建" == value.statu or "修改" == value.statu  :
                if  curLabel != value.LastLableTime  :
                    continue
                else:                
                    self.saveUpdate(str(cnt),value.statu,self.getRelativeName(value) )
                    cnt +=1
                    ret = True;
                    continue
            laststatu = "删除"
            try :
                laststatu = value.laststatu
                print("------------------", laststatu ,statu)
            except:
                pass

            if "删除" == value.statu :
                if  "删除" == laststatu :
                    continue
                else:
                    self.saveUpdate(str(cnt),value.statu,laststatu,self.getRelativeName(value) )
                    cnt  += 1
                    ret = True; 
                    continue
            print("未处理的情况：",laststatu,value.statu ,value.LastLableTime )                                  

        print()
        return ret 


class dirSafeRestore(dirRestore):
    def RestoreFromVobs( self,vobFolder,sourceFolder,sourceroot,releaseID=None):
        # glObj= glCls()
        init_internal(vobFolder ,sourceFolder ,sourceroot ,doCreateVobPath =False,releaseID=releaseID)  
        self.SSCNAme=glCls.VobFileName
        self.ldFileinfos(defaultStatu = None)
        # return
        print( "--------------处理目标文件夹中的文件：")
        self.dirmain( )
        self.printFileinfos()
        print( "----------dirmain completed!!----：")
        self.doDelPath()
        print( "----------doDelPath completed!!----：")
        self.printFileinfos()
        self.dictProc()
        print( "----------dictProc completed!!----：")

#-*- coding:utf-8 -*-
import os
import datetime

def  getdisks():
    w=ord( 'A' )
    # import win32file
    disklst=[]
    # i = win32file.GetLogicalDrives()
    for ii in range ( 0,27):
        i=2**ii
        dis= chr(w +ii) +":"
        # print( dis)
        if os.path.exists(dis):
            disklst += chr(w +ii)
    # print( disklst)
    return disklst


def findabspath( 标志文件夹名称="NSBD" ):
    disklst =  getdisks()
    for diskname in disklst:
      # print("----------------",disklst,diskname,标志文件夹名称)
      pathname = os.path.join( diskname+":\\"+标志文件夹名称 )

      # print( pathname)
      if os.path.exists( pathname ):
          return pathname
    return None

# def RestoreFromVobs(VobfolderRela="",sourceFolderRela="",releaseIDPar=None): #releaseIDPar="00139"
#     vobFolder=findabspath( VobfolderRela )
#     sourceFolder=findabspath(sourceFolderRela)
#     oPath = os.path.split( sourceFolder)

#     sourceroot= oPath[0]
#     sourceFolder= oPath[1]
#     print( sourceroot, "---",sourceFolder,";;", vobFolder)
#     releaseID  = releaseIDPar #"00139"
#     c = dirSafeRestore();
#     c.RestoreFromVobs( vobFolder,sourceFolder,sourceroot ,releaseID )

def RestoreFromVobs_Linux(vobFolder="",sourceFolder1="",releaseIDPar=None): #releaseIDPar="00139"
    # vobFolder=findabspath( VobfolderRela )
    # sourceFolder=findabspath(sourceFolderRela)
    oPath = os.path.split( sourceFolder1)

    sourceroot= oPath[0]
    sourceFolder= oPath[1]
    print( sourceroot, "---",sourceFolder,";;", vobFolder)
    releaseID  = releaseIDPar #"00139"
    c = dirSafeRestore();
    c.RestoreFromVobs( vobFolder,sourceFolder1,sourceFolder1 ,releaseID )



# def backup(VobfolderRela="",sourceFolderRela=""):
#     vobFolder=findabspath(VobfolderRela )
#     sourceFolder=findabspath(sourceFolderRela)
#     oPath = os.path.split( sourceFolder)
#     sourceroot= oPath[0]
#     sourceFolder= oPath[1]
#     print( sourceroot, "---",sourceFolder,";;", vobFolder)
#     c = dirSafeBackup();
#     c.Save2Vobs( vobFolder,sourceFolder1,sourceroot ); #.lower()

def haveSubFolder(rootfolder):
    folders = os.listdir(rootfolder)
    for folder in folders:
        if os.path.isdir( folder ) :
            return True;
    return False

def ConfirmRelease():
    print( "::----->",glCls.tgrVobroot )
    a=glCls.tgrVobroot
    base,b= os.path.split( a )
    tem,dtstr ,versionID= b.split("_")
    vob=os.path.join(a,dtstr+".vob")

    if os.path.exists(vob) :#and haveSubFolder( a ):
        tgrReleaseName = os.path.join(base,dtstr+"_"+versionID)
        os.rename( glCls.tgrVobroot,tgrReleaseName)
        print(a, "renamed to" ,tgrReleaseName)
    else:
        print( "deleted Tree:" ,a)
        os.removedirs(a) 
        




def backup_Linux(vobFolder="",sourceFolder1=""):
    oPath = os.path.split( sourceFolder1 )
    sourceroot= oPath[0]
    sourceFolder= oPath[1]
    print( "sourceFolder1 is: ",sourceFolder1 ,"Folder is:",sourceroot, "\nFile in subFolder:",sourceFolder,"\nVobFolder:", vobFolder)
    print()
    c = dirSafeBackup();
    try:
        c.Save2Vobs( vobFolder,sourceFolder,sourceroot );#.lower()
    except:
        pass
    ConfirmRelease(  )

