#-*- coding:utf-8 -*-
#-*- coding: UTF-8
from __future__ import print_function

import clr
import System
import System.IO
import LoadDlls

dllname=r"StdLibALLs.dll";clr.AddReferenceToFileAndPath( LoadDlls.getFullNames(dllname) )

import os       
import datetime
import gzip

class glCls:
  FileNumer=100
  vobFolder= ""
  sourceFolder= ""
  sourceroot = ""
  #print(":{}}}",  sourceroot )
  fullsourceFolder = ""
  curtime= ""
  strcurtime = ""

  VobFolderRoot = ""

  fileCnt= 0
  最新VobFile =""
  recentReleaseID = ""
  # print( "------------",最新VobFile,recentReleaseID )
  CurReleaseStr= ""
  tgrVobroot = ""
  VobFileName= ""
def CompressFile( srcname,tgrname ):
       with open(srcname, 'rb') as f_in:
          with gzip.open(tgrname, 'wb') as f_out:
              #f_out.writelines(f_in)
              while True:
                  chunk = f_in.read(1024)
                  if not chunk:
                      break
                  f_out.write(chunk)
def deCompressFile( srcname,tgrname ):
       with gzip.open(srcname, 'rb') as f_in:
          with open(tgrname, 'wb') as f_out:
              #f_out.writelines(f_in)
              while True:
                  chunk = f_in.read(1024)
                  if not chunk:
                      break
                  f_out.write(chunk)
def isSameID( st , releaseID ):
    if releaseID == None:
         return False
    vobname_1,release1 = st.split("_")
    if releaseID == release1:
        return True
    else:
        return False
def splitDate_ID( st,root ):
    print( "splitDate_ID", st,root)
    vobname_1,release1 = st.split("_")
    st1 = os.path.join( root,st,vobname_1+".vob" )
    release_1 = int( release1)
    return (st1,release_1)

# def __init__(self):
#   self.FileNumer =100
def getFileNumber():
    glCls.FileNumer = glCls.FileNumer+1
    return glCls.FileNumer
def  getRecentVobFile1(VobFolderRoot1,releaseID=None):
    try:
        st=""
        for root,path,f1 in os.walk( VobFolderRoot1):
            # print( "+++",root,VobFolderRoot1)
            if root == VobFolderRoot1:
                for p in path:
                  if True == isSameID( p , releaseID ):
                      print("+++",p,releaseID)
                      return splitDate_ID(p,root)
                  if st < p:
                    st = p
                break
        if releaseID == None:
            return splitDate_ID(st,root)
        else:
            return (None,None)
    except:
        return (None,0)
def  getRecentVobFile(VobFolderRoot1,releaseID=None):
    print( "getRecentVobFile Function :",VobFolderRoot1,releaseID )
    try:
        st = ""
        for root,path,f1 in os.walk( VobFolderRoot1):
            # print( "+++",root,VobFolderRoot1)
            if root == VobFolderRoot1:
                for p in path:
                  if "TemRelease" in p:
                        continue;
                  if  releaseID == p:
                      print("+++",p,releaseID)
                      return splitDate_ID(p,root)
                  if st < p:
                    st = p
                break
        if releaseID == None:
            return splitDate_ID(st,root)
        else:
            return (None,None)
    except:
        return (None,0)    
def init_internal( vobFolder=r"F:\NSBDVOB",sourceFolder= r"NSBD" , sourceroot=r"d:/",doCreateVobPath= True,releaseID=None):
    # print("+++++++++",vobFolder,sourceFolder,sourceroot)
    glCls.FileNumer=100
    glCls.vobFolder= vobFolder
    # print(",,,,", glCls.vobFolder , vobFolder)
    glCls.sourceFolder= sourceFolder
    glCls.sourceroot = os.path.abspath(os.path.join(  sourceroot,"" ))#.lower()
    
    glCls.fullsourceFolder = os.path.abspath( os.path.join(glCls.sourceroot,glCls.sourceFolder) )#.lower()
    print("sourceFolder,sourceroot,Glcls.  --->sourceroot,sourceFolder,fullsourceFolder: ", sourceFolder,sourceroot, glCls.sourceroot ,glCls.sourceFolder, glCls.fullsourceFolder)
    glCls.curtime= datetime.datetime.now()
    glCls.strcurtime = "{0:04d}".format(glCls.curtime.year)+"-"
    glCls.strcurtime += "{0:02d}".format( glCls.curtime.month)+"-"
    glCls.strcurtime += "{0:02d}".format( glCls.curtime.day)+"-"
    glCls.strcurtime += "{0:02d}".format( glCls.curtime.hour)+"-"
    glCls.strcurtime += "{0:02d}".format( glCls.curtime.minute)+"-"
    glCls.strcurtime += "{0:02d}".format( glCls.curtime.second ) 
    # print("wrewrewerewre  erewrew werwrwtr ")
    print(glCls.vobFolder)
    glCls.VobFolderRoot = os.path.join( glCls.vobFolder,"vob")

    glCls.fileCnt= 0
    glCls.最新VobFile,glCls.recentReleaseID = getRecentVobFile(glCls.VobFolderRoot,releaseID)
    print( "------ii------",glCls.最新VobFile,glCls.recentReleaseID )

    glCls.CurReleaseStr= "{0:05d}".format(  glCls.recentReleaseID+1 )

    # glCls.tgrVobroot = os.path.join( glCls.VobFolderRoot,glCls.strcurtime+"_"+glCls.CurReleaseStr)
    glCls.tgrVobroot = os.path.join( glCls.VobFolderRoot,"TemRelease"+"_"+glCls.strcurtime+"_"+glCls.CurReleaseStr )
    glCls.VobFileName = os.path.join( glCls.tgrVobroot,glCls.strcurtime+".vob")
    #print( strcurtime)
    if doCreateVobPath  == True:
      if  not os.path.exists( glCls.VobFolderRoot):
          os.makedirs(glCls.VobFolderRoot  )
      if  not os.path.exists( glCls.tgrVobroot):
          os.makedirs(glCls.tgrVobroot  )
       # os.makedirs(glCls.VobFolderRoot,exist_ok =True )
       # os.makedirs(glCls.tgrVobroot,exist_ok =True )
