#-*- coding:utf-8 -*-
#-*- coding: UTF-8

from __future__ import print_function
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System;
import System.Drawing;
import System.Windows.Forms;
import System.IO;
import BasicWidgets
vobFolderPairs =[]
import os
import sys
pwd1 =os.path.curdir
pwd2= os.path.join( pwd1,"../../")
dlls= os.path.join( pwd2,"Libs")
sys.path.append( os.path.abspath( dlls ) )

import basicLib
from System.Drawing import Point,FontStyle,GraphicsUnit,ContentAlignment 
from System.Drawing import Size
from System.Drawing import Font
from System.Drawing import FontStyle
from System.Drawing import Color
from System.Drawing import Font
vobName="__USBvobFiles"
vobFilder="D:/data/vob/"
docFoldername="D:/data/ipyLog/__worksDocuments.txt"

vobFolderPairs += [ (docFoldername,vobFilder ) ]

glbRecentVobsFileName= "RecentVobsFileName.Lst"

import filesafe

def restoreVobDOcfiles(VobfolderRela,sourceFolderRela,releaseIDPar):
    filesafe.RestoreFromVobs_Linux( VobfolderRela ,sourceFolderRela,releaseIDPar)
    print("restoreVobDOcfiles Completed!!!")

def backDOCfiles(VobfolderRela,sourceFolderRela):
    filesafe.backup_Linux( VobfolderRela,sourceFolderRela )
    print("backDOCfiles Completed!!!")       
def getVOBFile():
    openFileDialog1 = System.Windows.Forms.OpenFileDialog();
    openFileDialog1.DefaultExt = "VOB";
    openFileDialog1.Filter = "vob 登记files (*.VOB)|*.VOB";
    openFileDialog1.InitialDirectory = os.path.abspath(os.path.curdir) ;
    openFileDialog1.FileName = None;    
    result = openFileDialog1.ShowDialog();
    if( result == System.Windows.Forms.DialogResult.OK ):
        folderName = openFileDialog1.FileName;
        return folderName
    return None
def getDocFolder():
    folderBrowserDialog1 =System.Windows.Forms.FolderBrowserDialog();
    folderBrowserDialog1.ShowNewFolderButton = True;
    # folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.Personal;
    # folderBrowserDialog1.RootFolder=System.Environment.Personal 
    folderBrowserDialog1.SelectedPath=os.path.abspath(os.path.curdir)
    folderBrowserDialog1.Description = "选择vob需要归档的： 文档或者数据文件夹";
    result = folderBrowserDialog1.ShowDialog();
    if( result == System.Windows.Forms.DialogResult.OK ):
        folderName = folderBrowserDialog1.SelectedPath;
        folderName = folderName.replace("\\","/").strip()
        return folderName
    return None
def getVobFolder():
    folderBrowserDialog1 =System.Windows.Forms.FolderBrowserDialog();
    folderBrowserDialog1.ShowNewFolderButton = True;
    folderBrowserDialog1.SelectedPath=os.path.abspath(os.path.curdir)
    folderBrowserDialog1.Description = "选择保存vob的文件夹";
    result = folderBrowserDialog1.ShowDialog();
    if( result == System.Windows.Forms.DialogResult.OK ):
        folderName = folderBrowserDialog1.SelectedPath;
        return folderName
    return None
        # print(folderName )
def getVobFolderName(VobFileName ):
    foldername = os.path.dirname(VobFileName )
    folder = os.path.join( foldername ,"vob" )
    return folder

def  getVobVersions(VobFileName):
    folder = getVobFolderName(VobFileName )
    dirs1 = os.listdir( folder )
    dirs=()
    for i in sorted( dirs1,reverse=True ):
        if  "TemRelease" not in i:
            dirs += ( i,)
    return dirs
   
class DialogMain( System.Windows.Forms.Form ):
    def getFont(self):
        fontsize = 13# self.ClientSize.Height*7/10
        # print(self.ClientSize.Height ,self.Size.Height )
        font= Font("仿宋", fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
        return font      
    def __init__(self,title=None,OrgObj=None):
        self.AutoScroll =True
        self.Font= self.getFont()
        screenSize =System.Windows.Forms.Screen.GetWorkingArea(self)
        maxsize1 =System.Drawing.Size(screenSize.Width /3*2 , screenSize.Height  )
        self.Text =title
        # self._Height=60
        # self.OrgObj = OrgObj
        self.Posy=10
        self.Size = System.Drawing.Size(maxsize1.Width,maxsize1.Height)
        self.WindowState = System.Windows.Forms.FormWindowState.Normal
        self.result = None
        self.create_menu()        
        self.body()

        self.bodyShowTree()   
        self.buttonbox()             
        self.init(OrgObj )

    def create_menu(self):
        pass
    def bodyShowTree(self):
        splitPadsize =20
        buttonHeight =40 
        buttonwidth =200

        Width_pad = self.ClientSize.Width - splitPadsize*2 -buttonwidth*2  
        Width_pad /=2
        posx=  Width_pad
        print(self.Posy ,posx )
        # w = System.Windows.Forms.Button()
        w=BasicWidgets.Button_Self( )

        w.Text = "选择VOBPath"
        w.Location = System.Drawing.Point(posx, self.Posy)
        w.Height = buttonHeight
        w.Width  = buttonwidth
        w.ForeColor = System.Drawing.Color.Blue
        w.Click  += self.选择VObPath
        w.Parent = self

        posx += buttonwidth+splitPadsize;     
        # w1 = System.Windows.Forms.Button()
        w1=BasicWidgets.Button_Self( )        
        w1.Text = "选择文档Path" 
        w1.Location = System.Drawing.Point(posx, self.Posy)
        w1.Height = buttonHeight
        w1.Width = buttonwidth
        w1.ForeColor = System.Drawing.Color.Blue
        w1.Click += self.选择文档Path
        w1.Parent =self
        posx += buttonwidth+splitPadsize   ;   

        self.Posy += buttonHeight+10               
        pass
    def getMostPossibleFolder(self,values):
        if values ==None :
            return ""
        if len(values ) ==0 :
            return ""
        for i in range( len(values )-1,-1,-1):
            foldername = values[ i ]
            if os.path.exists( foldername ):
                return foldername
        return "" 

    def body(self):
        posx=0
        posy=10
        hight=40
        width=1110
        title="历史VOB文件名："
        # curvlaue="qweq\were"

        values = self.getDocFiles(glbRecentVobsFileName)
        if values != None and len(values) != 0:
            curvalue =self.getMostPossibleFolder(values )
        else:
            curvalue =""
        name="VobFILE"
        self.vobPathCtrl=BasicWidgets.SelectEditEntry(posx,posy,hight,width,title,curvalue,values,name,colname1 ="NO" )
        self.vobPathCtrl.Parent =self
        self.vobPathCtrl.Notify= self.ProcVOBFILEChanged
        self.Posy = self.vobPathCtrl.Bottom 
        # print("vobPathCtrl",self.Posy )

        # posx=0
        posy = self.Posy = self.Posy  + 10
        # hight=40
        # width=1100
        title="历史文档文件夹："
        curvlaue=""
        values=("",)
        name="DOCpath"
        self.DocPathCtrl =BasicWidgets.SelectEditEntry(posx,posy,hight,width,title,curvlaue,values,name,colname1 ="NO" )
        self.DocPathCtrl.Parent =self
        self.Posy = self.DocPathCtrl.Bottom + 10 
        # print("DocPathCtrl",self.Posy )

        # posx=0
        posy = self.Posy = self.Posy  + 10
        # hight=40
        # width=1100
        title="恢复文档时的版本号，缺省为最新版本："
        curvlaue=""
        values=("",)
        name="Version"
        self.versionsCtrl =BasicWidgets.SelectEditEntry(posx,posy,hight,width,title,curvlaue,values,name,colname1 ="NO" )
        self.versionsCtrl.Parent =self
        self.Posy =self.versionsCtrl.Bottom + 10 
        # print("versionsCtrl",self.Posy )


        self.ProcVOBFILEChanged( self.vobPathCtrl)
        pass
    def init( self,OrgObj ):
        pass
    def buttonbox(self):
        splitPadsize =20
        buttonHeight =40 
        buttonwidth =200

        Width_pad = self.ClientSize.Width - splitPadsize*2 -buttonwidth*3  
        Width_pad /=2
        posx=  Width_pad
        print(self.Posy ,posx )
        # w = System.Windows.Forms.Button()
        w = BasicWidgets.Button_Self( )  
        w.Text = "备份文档"
        w.Location = System.Drawing.Point(posx, self.Posy)
        w.Height = buttonHeight
        w.Width  = buttonwidth
        w.ForeColor = System.Drawing.Color.Blue
        w.Click  += self.备份文档
        w.Parent = self

        posx += buttonwidth+splitPadsize;     
        # w1 = System.Windows.Forms.Button()
        w1 = BasicWidgets.Button_Self( )  
        w1.Text = "恢复文档" 
        w1.Location = System.Drawing.Point(posx, self.Posy)
        w1.Height = buttonHeight
        w1.Width = buttonwidth
        w1.ForeColor = System.Drawing.Color.Blue
        w1.Click += self.恢复文档
        w1.Parent =self


        posx += buttonwidth+splitPadsize   ;     

        # w2 = System.Windows.Forms.Button()
        w2 = BasicWidgets.Button_Self( )          
        w2.Text = "退出系统"
        w2.Location = System.Drawing.Point(posx, self.Posy)
        w2.Height = buttonHeight
        w2.Width = buttonwidth
        w2.ForeColor = System.Drawing.Color.Blue
        w2.Click += self.退出系统
        w2.Parent =self
        
        self.Posy += buttonHeight+10
    def getFolders(self):
        self.VobRegFilename = self.vobPathCtrl.get()
        self.VobfolderRela =None

        if(self.VobRegFilename !=None and self.VobRegFilename !="" ):
            print( "恢复文档:",self.VobRegFilename )
            self. VobfolderRela = os.path.join( getVobFolderName( self.VobRegFilename ),"../" )
            self.VobfolderRela =os.path.abspath( self.VobfolderRela )
            print( "恢复文档1:",self.VobfolderRela )
        print( "恢复文档2:",self.VobfolderRela )
        docfolder = self.DocPathCtrl.get()
        self.sourceFolderRela =None
        if(docfolder !=None and docfolder !="" ):
            self.sourceFolderRela =  docfolder 

        version  = self.versionsCtrl.get()
        self.releaseIDPar =None
        if(version !=None and version !="" ):
            self.releaseIDPar =  version 
    def 备份文档(self, sender,event=None):
        pass
        self.getFolders()
        if self.sourceFolderRela != None and self.VobfolderRela != None:
            backDOCfiles(self.VobfolderRela,self.sourceFolderRela )
        else:
            print("VobfolderRela,sourceFolderRela,releaseIDPar:" ,self.VobfolderRela,self.sourceFolderRela) 
        self.UpdateCtrls()                   
        #backDOCfiles(VobfolderRela,sourceFolderRela):
    def SaveDocFolder2VOBreg(self):
        # self.VobRegFilename
        # sourceFolderRela
        if self.DocFolder == None:
            return 
        values = self.getDocFiles(self.VobRegFilename)

        newFolders=[  self.DocFolder  ]
        for i in range(len(values)-1,-1,-1):
            item  = values[ i ]
            folder = item.replace("\\","/").strip()
            if folder not in newFolders:
                newFolders =  [folder] + newFolders

        fobj = basicLib.openfile(self.VobRegFilename,"wt")
        for i in range(0,len( newFolders) ):
            print( newFolders[i] ,file=fobj )
        fobj.close()

    def 恢复文档(self, sender, event):
        print("恢复文档-----------------------------")
        self.getFolders()
        if self.sourceFolderRela != None and self.VobfolderRela != None:
            pass# 
            message = "你将从备份的文件库中恢复文件到你本地的文件夹中，\n这将有可能覆盖你在本地修改过的文件，请你慎重选择?";
            caption = "慎重选择确认！！";
            buttons = System.Windows.Forms.MessageBoxButtons.OKCancel;
            # result;
            Icon =System.Windows.Forms.MessageBoxIcon.Warning #  Exclamation
            result = System.Windows.Forms.MessageBox.Show(message, caption, buttons,Icon);
            if (result == System.Windows.Forms.DialogResult.OK):
                pass
                print("将恢复并覆盖旧版本的文件:")
                restoreVobDOcfiles(self.VobfolderRela,self.sourceFolderRela,self.releaseIDPar)
        else:
            print("VobfolderRela,sourceFolderRela,releaseIDPar:" ,self.VobfolderRela,self.sourceFolderRela,self.releaseIDPar)

    def 选择VObPath(self, sender, event):
        print("选择VObPath")
        folder = getVOBFile().replace("\\","/")
        if folder != None:
            self.vobPathCtrl.set(folder )
            self.updateVobHIstoryLog(folder )
    def updateVobHIstoryLog(this,vobFolder):
        # glbRecentVobsFileName            
        if vobFolder == None:
            return 
        values = this.getDocFiles(glbRecentVobsFileName)

        newFolders=[ vobFolder ]
        for i in range(len(values)-1,-1,-1):
            item  = values[ i ]
            folder = item.replace("\\","/").strip()
            if folder not in newFolders:
                newFolders =  [folder] + newFolders

        fobj = basicLib.openfile(glbRecentVobsFileName,"wt")
        for i in range(0,len( newFolders) ):
            print( newFolders[i] ,file=fobj )
        fobj.close()  

    def 选择文档Path(self, sender, event):
        print("选择文档Path")
        self.DocFolder = getDocFolder()
        if self.DocFolder != None:
            self.DocPathCtrl.set(self.DocFolder ) 
        self.SaveDocFolder2VOBreg()

    def getDocFiles(self,VOBRegFile ):
        try:
            f1obj =basicLib.openfile(VOBRegFile  ,"rt")
        except:
            f1obj = None
            return ()
        values = []
        for line in f1obj:
            # print(line )
            values += [line.replace("\r","").replace("\n","") ]

        vs =()
        ll = len( values )
        if ll == 0:
             return vs
        for idx in range( ll-1, -1,-1):
            v = values[idx ]
            if v not in vs:
                vs = (v,)+ vs
        for d in vs:
            print(d)

        return vs 
    def UpdateCtrls(self) :
        values = self.getDocFiles(self.VobRegFilename)
        self.DocPathCtrl.setValues( values )

        value = self.getMostPossibleFolder( values )
        self.DocPathCtrl.set( value )

        self.versions = getVobVersions(self.VobRegFilename )
        self.versionsCtrl.setValues(self.versions)
        if len(self.versions) >0 :
            self.versionsCtrl.set(self.versions[0] )
        
    def ProcVOBFILEChanged(self,parent):
        try:
            print( "In ProcVOBFILEChanged:")
            self.VobRegFilename = parent.get()
            if self.VobRegFilename != None and len(self.VobRegFilename) >0:
                print("selected : ",self.VobRegFilename)
                self.UpdateCtrls()
                self.updateVobHIstoryLog(self.VobRegFilename)
            else:
                print("Select NULL values")
        except:
            print("ERROR!! in :","ProcVOBFILEChanged")
            pass            

    def 退出系统(self, sender, event):
        print("exit out ")
        self.Close()


title = "文档备份恢复管理系统" +" : "+os.path.abspath(os.path.curdir)
formMain = DialogMain(  title )
System.Windows.Forms.Application.Run(formMain)
  