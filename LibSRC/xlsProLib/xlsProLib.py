####-*- coding: cp936 -*-
#UTF-8 
ForReading = 'rt' #'以只读方式打开文件。不能写这个文件。 
ForWriting = 'wt'
ForAppending = 'at'

TristateUseDefault = -2 #'使用系统默认值打开文件。 
TristateTrue = -1# '以 Unicode 方式打开文件。 
TristateFalse = 0 #'以 ASCII 方式打开文件。 
import  os

ccc=__file__
c = os.path.split( ccc)
modulePath  = c[0]

moban_rootpath = c[0]+"\\template\\" #r''#E:\\Current\\system\\EclipseRUN\\__Prj\\PyApp\\'
VbNewLine = '\n'
import datetime
def get_BookTail():
  return "</Workbook>" + VbNewLine + VbNewLine

def get_SheetTail():
  return "" + VbNewLine + "</Worksheet>" + VbNewLine



def get_TableHead( ):
    HeadStr=  VbNewLine + '<Table ss:StyleID="Default" ss:DefaultColumnWidth="100" ' \
     + VbNewLine + "" + '   ss:DefaultRowHeight="25"> ' + VbNewLine + ""
    return HeadStr

HiddenStr= '<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"> '+ VbNewLine+\
  ' <Visible>SheetHidden</Visible>' + VbNewLine+\
  '</WorksheetOptions>'+ VbNewLine
def get_TableTail( TableHide  = False ):
  HeadStr = VbNewLine + "</Table>" + VbNewLine
  if TableHide == False:
        return HeadStr
  else:
        return HeadStr + HiddenStr;

def get_RowHead_1():
  return '<Row ss:AutoFitHeight="0"> ' + VbNewLine

def get_StylesHead():
 return "<Styles>" + VbNewLine + "" + VbNewLine + ""
 
def Write_File(filename , Data , flagForWriting):
    a = open(filename, flagForWriting,1,"utf-8")
    #print( a.encoding ,a.newlines,filename)  
    a.write(Data)
    a.close()
def getDataFromFile(filename):
    file1 = open(filename, ForReading,1,"utf-8")
    data = ""
    #print file1.encoding 
    data1 = file1.readlines()

    for line1 in data1:
      data = data + line1 #+VbNewLine
    
    file1.close()
    #print data
    return data
def getDataFromFile_Utf16(filename):
    print(filename)
    file1 = open(filename, ForReading,1,"utf-8")
    file1.read(1)# reading in the coding type filed and discard it
    data = ""
    #print file1.encoding 
    #data1 = file1.readlines()
    #print(data1)
    #===========================================================================
    for line1 in file1:
      data = data + line1 #+VbNewLine
      #print((line1))
    #===========================================================================
    file1.close()
    #print data
    return data
def get_StylesFromFile():
    return getDataFromFile(moban_rootpath + "Excel_style.txt")

def get_StylesTail():
  return  "</Styles>" + VbNewLine

def get_RowTail():
  return "</Row>" + VbNewLine

def  get_SheetHead(sheetName):
 return "" + VbNewLine + '<Worksheet ss:Name="' + sheetName + '"> ' + VbNewLine

def  get_RowHead(RowIndex):
  if RowIndex == -1   :
    return "" + VbNewLine + '<Row  ss:AutoFitHeight="0">' + VbNewLine
  else :
    return "" + VbNewLine + '<Row ss:Index="' + str(RowIndex) + '" ss:AutoFitHeight="0">' + VbNewLine

def get_NullRow(RowCnt):
    return VbNewLine + '<Row ss:AutoFitHeight="0" ss:Span="' + str(RowCnt) + '"' + ' /> ' + VbNewLine

def get_BookHead():
 return  getDataFromFile(moban_rootpath + "ExcelBookHead.txt")

def get_CellDefaultHead(styleIdstr, cellId,MergeDown=0,MergeAcross=0):
 #' styleIdstr 的名称应该与Styles定义的ID相同，否则文件无法打开
  index = ""
  if(cellId != -1) : 
    index = 'ss:Index="' + str( cellId ) + '" '
  fstr=' {2:s}MergeDown="{0:d}"  {2:s}MergeAcross="{1:d}" '.format(MergeDown,MergeAcross,"ss:")
  return ' <Cell ' + index + fstr+ 'ss:StyleID="' + styleIdstr + '"> ' + VbNewLine
#ss:MergeDown="1"  ss:MergeAcross="1"
def get_CellTail():
  return "  </Cell>" + VbNewLine

def get_CellData(datetypeStr , valueStr):
     datastr=strZhuanyiProc( str(valueStr) )
     return '  <Data ss:Type="' +str( datetypeStr )+ '">' + datastr  + "</Data>" + VbNewLine
 #' example as below:
 #' "Number"  1212
 #'DateTime 1900-01-04T00:00:00.000
 #'String  Any Text or line

def get_CellHyperLinkHead(styleIdstr, cellId, HrefAdd):
 #' styleIdstr 的名称应该与Styles定义的ID相同，否则文件无法打开
    index = ""
    if(cellId != -1) :
         index = 'ss:Index="' + cellId + '" '
    
    return  "  <Cell " + index + 'ss:StyleID="' + styleIdstr + '" ss:HRef="' + HrefAdd + '"> ' + VbNewLine

#'    <Cell ss:StyleID="s27" ss:HRef="E:\current\system\mailLog\Send\答复_ 关于Iur-g+现网试点方案的调研_4.msg"><Data
#'     ss:Type="String">答复_ 关于Iur-g+现网试点方案的调研_4.msg</Data></Cell>
   
#'''''''''''''''''''/

def  time2Str( d_t ):
    if (d_t == None ):
        return "1970-01-01"
    timestr  =""
    timestr =str(d_t.year)+"-"+str(d_t.month)+"-"+str(d_t.day) + "T"+str(d_t.hour) \
      +  "-"+ str(d_t.minute) + "-"+ str(d_t.second)+"-"+str(d_t.microsecond /1000)
    return timestr
     
def MakeXml_Head(MailBackPath ) :
    #xmlxlsHead.txt
    filename = MailBackPath + "Head.txt"
    data = getDataFromFile(moban_rootpath+"\\xmlxlsHead.txt") #xmlxlsHead.txt
    Write_File(filename , data + VbNewLine, ForWriting)

#sheetHeadFileName="sheetHeadFileName"
def MakeXmlSheet_Head(MailBackPath, SheetName):
    filename = MailBackPath + SheetName + "head.txt"
    data1=""

    data1 = data1 + get_SheetHead(SheetName)
    data1 = data1 + get_TableHead()

    data1 = data1 + get_NullRow(1)

    data1 = data1 + get_RowHead(-1)  
    data1 = data1 + WriteTitle("Title", "发件人、姓名") #' 1     
    data1 = data1 + WriteTitle("Title", "发送时间")   #' 2
    data1 = data1 + WriteTitle("Title", "主题") #' 3
    data1 = data1 + WriteTitle("Title", "链接地址")  #' 15
    data1 = data1 + get_RowTail() 
    Write_File(filename, data1 ,ForWriting)# ForAppending) 

def  MakeXmlSheetTail(MailBackPath):
    filename = MailBackPath + "sheetTail.txt"
    Data = "</Table>" + VbNewLine + "</Worksheet>" + VbNewLine + ""
    Write_File(filename , Data , ForWriting) 

def MakeXmlTail(MailBackPath) :
    filename = MailBackPath + "tail.txt"
    Data = "</Workbook>"
    Write_File(filename , Data, ForWriting)

#/ entry code here 
def WriteTitle(Style, ColName):
    return  get_CellDefaultHead(Style, -1) + get_CellData("String", ColName) + get_CellTail()

def WriteCEll(Style, datatype, ColText):
    return get_CellDefaultHead(Style, -1) + get_CellData(datatype, ColText) + get_CellTail()


def WriteRefLnk(LinkAdd, content, Style, datatype):
    return get_CellHyperLinkHead(Style, -1, LinkAdd) + get_CellData(datatype, content) + get_CellTail()

def WriteXlcHead_sheethead(filename, WritingFlag, StyleFileName, sheetName):
    str = ""
    str = get_BookHead()
    str = str + get_StylesHead()
    str = str + get_StylesFromFile()
    str = str + get_StylesTail()
    str = str + get_SheetHead(sheetName) + get_TableHead()
    str = str + get_NullRow(1)
    Write_File(filename, str , WritingFlag) 
def WriteXlcHead(filename, WritingFlag, StyleFileName):
    str = ""
    str = get_BookHead()
    str = str + get_StylesHead()
    str = str + get_StylesFromFile()
    str = str + get_StylesTail()
    #str = str + get_SheetHead(sheetName) + get_TableHead()
    #str = str + get_NullRow(1)
    Write_File(filename, str , WritingFlag) 
def WriteSheetHead(filename, WritingFlag, sheetName):
    str = ""
    str = str + get_SheetHead(sheetName) + get_TableHead()
    str = str + get_NullRow(1)
    Write_File(filename, str , WritingFlag) 

def FormatDateTime():
     cur_date=datetime.date.today()
     cur_time=datetime.datetime.now()
     return cur_time.strftime("%Y_%m_%d_%H_%M_%S") 
# '''''''''''''''''''''''''''''''/

class xlsFile:
    def __init__( self,filename ):
         self.tables= dict() #tablename ,rows[]
         self.filename = filename
    def getCell( self ,formatstr,typestr,datastr ):
       return get_CellDefaultHead(  formatstr, -1) + get_CellData(typestr ,  datastr ) + get_CellTail()
    def getCellMerg( self ,formatstr,typestr,datastr ,CellID,DownNum,AcrossNum):
       return get_CellDefaultHead(  formatstr, CellID,DownNum,AcrossNum) + get_CellData(typestr ,  datastr ) + get_CellTail()
       
    def getHyperlinkCell(self ,formatstr ,  LinkAdd, typestr,content):
       return get_CellHyperLinkHead("Title", -1, LinkAdd) + get_CellData("String", content) + get_CellTail()
    def getrow(self , cellsdata,rowID ):
       return  get_RowHead( rowID ) + cellsdata + get_RowTail()
    def  addtable(self, tablename ):
       self.tables[ tablename ] =[]
    def addrow2table(self,tablename , rowCellsStr ,rowID):
         rows = self.tables[tablename]	
         rowStr = self.getrow( rowCellsStr,rowID)
         rows =rows .append ( rowStr)
    def  write2FS(self , Datastr):
        self.filehandle.write(Datastr)   	   
    def writeXlsFile(self):
        #return "dsfdfdsfs"
        #print("dffsdfds ")
        self.filehandle  = open(self.filename, ForWriting,1,"utf-8")
        self.write2FS ( get_BookHead() )
        #print(  get_StylesHead() )
        self.write2FS (  (get_StylesHead()) )
        self.write2FS (  get_StylesFromFile() )
        self.write2FS (  get_StylesTail() )
        #print(tables )
        for (tablename ) in   sorted( list(  self.tables ) ) :
            #print(tablename)
            rows = self.tables[tablename]
            self.write2FS (   get_SheetHead( tablename ) + get_TableHead() )
            rowcnt =0
            for rowcnt in range( 0, len( rows  ) ):
                self.write2FS( rows[ rowcnt ] )
            self.write2FS( get_TableTail() )
            self.write2FS( get_SheetTail() )
        self.write2FS( get_BookTail() )
        self.filehandle.close()
        ##return "ereerrew"

def test(filename ):
    xlsfile1 =xlsFile("d:\\sdfg.xls")
    xlsfile1.addtable( "data")
    cellsDta = xlsfile1.getCell( "Regular","Number","2132132") 
    cellsDta += xlsfile1.getCell( "Date","String","2015-05-12T13:34:23")
    cellsDta += xlsfile1.getCell( "Title","Number","2132132")
    xlsfile1.addrow2table( "data",cellsDta ,4)

    xlsfile1.addtable( "表一" )
    cellsDta = xlsfile1.getCell( "Title","String","2132132")
    cellsDta += xlsfile1.getCell( "Title","String","2015-05-12T13:34:23")
    cellsDta += xlsfile1.getCell( "Title","String","2132132")
    xlsfile1.addrow2table( "表一",cellsDta ,4)

    
    LinkAdd = "E:/current/system/mailLog/Send/答复_ 关于Iur-g+现网试点方案的调研_4.msg"
    content = "OutLookMsg"
    cells = xlsfile1.getHyperlinkCell( "Title",  LinkAdd ,"String", content)
    xlsfile1.addrow2table( "表一",cells ,-1)
    xlsfile1.addrow2table( "表一",cells ,-1)
    xlsfile1.addrow2table( "表一",cells ,-1)
    xlsfile1.addrow2table( "表一",cells ,-1)
    xlsfile1.writeXlsFile()

def ___Testing( filename ):
    str = ""
    str = get_BookHead()
    #print(  get_StylesHead() )
    str = str + (get_StylesHead())
    str = str + get_StylesFromFile()
    str = str + get_StylesTail()
    
    str = str + get_SheetHead("data") + get_TableHead()
    
    str = str + get_NullRow(10)
    str = str + get_RowHead(-1)
    
    str = str + get_CellDefaultHead("Regular", -1) + get_CellData("String", "121432") + get_CellTail()
    
    str = str + get_CellDefaultHead("Date", -1) + get_CellData("String", "1989-09-21T12:30:23.09") + get_CellTail()
    
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "1989-09-21T12:00") + get_CellTail()
    
    str = str + get_RowTail()
    
    str = str + get_TableTail() + get_SheetTail()
        
#    str= str+get_BookTail()
#    #'wSCRIPT.eCHO "str"
#    Write_File(rootpath+"vv.xls",str ,ForWriting)
#    
#    return
    str = str + get_SheetHead("data1") + get_TableHead() #' 增加一个Sheet和Sheet表
    
    str = str + get_NullRow(10)  #' 空行
    str = str + get_RowHead(-1)  #' 启动新行，顺序增加
    
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第一列") + get_CellTail()
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第二列") + get_CellTail()
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第三列") + get_CellTail()
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第四列") + get_CellTail()
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第五列") + get_CellTail()
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第六列") + get_CellTail()
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第七列") + get_CellTail()
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "第八列") + get_CellTail()
    
    str = str + get_RowTail()
    str = str + get_RowHead(-1)  #' 启动新行，顺序增加
    #' 增加一个列单元格。
    str = str + get_CellDefaultHead("Regular", -1) #' 顺序增加一列
    str = str + get_CellData("String", "121432") #' 为新单元格添加数据
    str = str + get_CellTail()#' 结尾新单元格
    
    str = str + get_CellDefaultHead("Date", -1) + get_CellData("String", "1989-09-21T12:30:23.09") + get_CellTail()
    
    str = str + get_CellDefaultHead("Title", -1) + get_CellData("String", "1989-09-21T12:00") + get_CellTail()
#    
    LinkAdd = "E:/current/system/mailLog/Send/答复_ 关于Iur-g+现网试点方案的调研_4.msg"
    content = "OutLookMsg"
    str = str + get_CellHyperLinkHead("Title", -1, LinkAdd) + get_CellData("String", content) + get_CellTail()
    
    
    LinkAdd = "E:/current/system/mailLog/Send/答复_ 关于Iur-g+现网试点方案的调研_4.msg"
    content = "OutLookMsg"
    str = str + get_CellHyperLinkHead("Title", -1, LinkAdd) + get_CellData("String", content) + get_CellTail()
    
    
    LinkAdd = "E:/current/system/mailLog/Send/答复_ 关于Iur-g+现网试点方案的调研_4.msg"
    content = "OutLookMsg"
    str = str + get_CellHyperLinkHead("Title", -1, LinkAdd) + get_CellData("String", content) + get_CellTail()
    
    
    LinkAdd = "E:/current/system/mailLog/Send/答复_ 关于Iur-g+现网试点方案的调研_4.msg"
    content = "OutLookMsg"
    str = str + get_CellHyperLinkHead("Title", -1, LinkAdd) + get_CellData("String", content) + get_CellTail()
    
    str = str + get_RowTail()
    
    str = str + get_TableTail() + get_SheetTail()
    
   # ''''''''''''''''''''''''''''''''
    
    str = str + get_BookTail()
    #'wSCRIPT.eCHO "str"
   # print(str)
    Write_File(filename , str , ForWriting)
def strZhuanyiProc(srcstr):
    temstr= srcstr.replace("&","&amp;");
    temstr= temstr.replace("<","&lt;");
    temstr= temstr.replace(">","&gt;");
    temstr= temstr.replace("'","&apos;");
    temstr= temstr.replace("\"","&quot;");
    return temstr
#Testing()
# MakeXmlSheetData_2_Fs
def MakeXmlSheetData_2_Fs(fs , ItemInfo):#???
    #'&lt;&gt; &quot;
    #'/^(?:Chapter|Section) [1-9][0-9]{0,1}$/
    #'对 VBScript：
    #'"^(?:Chapter|Section) [1-9][0-9]{0,1}$"
    #'/\b([a-z]+) \1\b/gi
    #'等价的 VBScript 表达式为：
    #'"\b([a-z]+) \1\b"
    ItemInfo.SenderName = ItemInfo.SenderName.replace("<", "&lt;")
    ItemInfo.SenderName = ItemInfo.SenderName.replace(">", "&gt;")
    ItemInfo.SenderName = ItemInfo.SenderName.replace('"', "&quot;")
    ItemInfo.title = ItemInfo.title.replace("<", "&lt;")
    ItemInfo.title = ItemInfo.title.replace(">", "&gt;")
    ItemInfo.title = ItemInfo.title.replace('"', "&quot;")

    data=""
    data = data + get_RowHead(-1)  #' 启动新行，顺序增加
    data = data + \
    get_CellDefaultHead("Regular", -1) +get_CellData("String", ItemInfo.SenderName) + get_CellTail()
    data = data + \
    get_CellDefaultHead("Date", -1) + get_CellData("String", time2Str(ItemInfo.MidificationTime)) + get_CellTail()
   # print(data)
    data = data + \
    get_CellDefaultHead("Regular", -1) + get_CellData("String",  ItemInfo.title ) + get_CellTail()
  #  data = data + get_CellDefaultHead("", -1) + get_CellData("String", "1989-09-21T12:30:23.09") + get_CellTail()
  
    LinkAdd = ItemInfo.FileFullName
    content = ItemInfo.FileName
    data = data + get_CellHyperLinkHead("Title", -1, LinkAdd) +\
     get_CellData("String", content) + get_CellTail()
 
    data = data + get_RowTail()
    
    # print(data)
    fs.write(data)
 
#
#___Testing("d:\\aaa.xls")
# test("d:\\eee.xls")