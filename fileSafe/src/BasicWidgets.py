#-*- coding: UTF8
from __future__ import print_function
"""
Test of IronPython's ability to change
individual cell formats in a dot Net
DataGridView control.
Intended behavior of DataGridViewControl:
    1) highlights row and column of mouse
       position simultaneously
    2) mock selection in color scheme
       other than Windows default
    3) selection of individual cells
       and individual rows allowed
       (no multiple selection)
    4) auto copy selections to clipboard
       in a manner that allows direct
       pasting into Excel
Geared toward use with ~1000 rows and
~10 columns.
Intended to make tracking of active row
and data easier.
"""
# XXX - row/column tracking is slow with
# XXX       more than 500 rows
# XXX   could try to optimize by tracking cell indices

# XXX - PageDown, PagUp, Home, and Arrow Keys will
# XXX       still highlight and select cells in the 
# XXX          Windows default fashion.

# XXX - could write more code for column header width
# XXX       adjustment and row width adjustment
# XXX       - columnwidth adjustment clears mock selection
# XXX       - rowwidth adjustment selects row above row
# XXX             boundary

import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import Form,AutoCompleteMode,AutoCompleteSource,AutoCompleteStringCollection 
from System.Windows.Forms import DataGridView
from System.Windows.Forms import DataGridViewContentAlignment
from System.Windows.Forms import Application
from System.Windows.Forms import Control
from System.Windows.Forms import Clipboard
from System.Windows.Forms import DataFormats
from System.Windows.Forms import DataObject,Panel ,MessageBox,DialogResult,MessageBoxButtons

from System.Drawing import Point,FontStyle,GraphicsUnit,ContentAlignment 
from System.Drawing import Size
from System.Drawing import Font
from System.Drawing import FontStyle
from System.Drawing import Color

from System import Text

from System.IO import MemoryStream
# clr.AddReference("System.Windows.Forms")
# clr.AddReference("System.Drawing")

from System.Windows.Forms import Application, Form,Button
from System.Windows.Forms import ComboBox, Label,Panel,TextBox
from System.Drawing import Size, Point ,Font,Color

# import tips
# formatting constants
MDDLCNTR = DataGridViewContentAlignment.MiddleCenter
BOLD = Font(Control.DefaultFont, FontStyle.Bold)
REGL = Font(Control.DefaultFont, FontStyle.Regular)
SELECTCOLOR = Color.LightSkyBlue
ROWCOLOR = Color.Yellow
COLUMNCOLOR = Color.Cyan
MOUSEOVERCOLOR = Color.GreenYellow
REGULARCOLOR = Color.White
ROWHDRWDTH = 65

import re
import datetime
CSV = DataFormats.CommaSeparatedValue

# hack for identifying mouseovered cell
#    mousing over one of the header cells yields
#        an event.RowIndex value of -1       
INDEXERROR = -1

NUMCOLS = 3
HEADERS = ['positive', 'negative', 'flat']
TESTDATA = (['happy', 'sad', 'indifferent'],
            ['ebullient', 'despondent', 'phlegmatic'],
            ['elated', 'depressed', 'apathetic'],
            ['fired up', 'bummed out', "doesn't care"],
            ['psyched', 'uninspired', 'blah'])
NUMROWS = len(TESTDATA)

def getcellidxs(event):
    """
    From a mouse event on the DataGridView,
    returns the row and column indices of 
    the cell as a 2 tuple of integers.
    """

    # this is redundant with DataGridForm.getcell
    #    class methods were not handling unpacking
    #        of tuple with cell in it
    #    trying a separate function
    # print event.RowIndex, event.ColumnIndex
    # print("----?--")
    # print(event.RowIndex, event.ColumnIndex)
    # print("=====>=")
    return event.RowIndex, event.ColumnIndex

def resetcellfmts(gridcontrol, numrows):
    """
    Initialize formatting of all data
    cells in the grid control.
    """
    # need to cycle through all cells individually
    #    to reset them
    for num in xrange(numrows):
        row = gridcontrol.Rows[num]
        for cell in row.Cells:
            # skip over selected cell(s)
            if cell.Style.BackColor != SELECTCOLOR:
                cell.Style.Font = REGL
                cell.Style.BackColor = REGULARCOLOR 

def resetheaderfmts(gridcontrol, rowidx, colidx):
    """
    Reset BackColor on "Header" cells for 
    rows and columns.
    """
    col = gridcontrol.Columns[colidx]
    col.HeaderCell.Style.BackColor = Color.Empty
    # for row header formats, don't clear selected
    row = gridcontrol.Rows[rowidx]
    if row.HeaderCell.Style.BackColor != SELECTCOLOR:
        row.HeaderCell.Style.BackColor = Color.Empty

def clearselection(gridcontrol):
    """
    This works as a separate function,
    but not as part of the form class.
    Clears gridview selection.
    """
    # clear selection
    gridcontrol.ClearSelection()

def mockclearselection(gridcontrol):
    """
    Clears mock selection on custom 
    color scheme for DataGridView.
    """
    # have to cycle through all cells
    rows = gridcontrol.Rows
    for row in rows:
        # deal with selected header, if any
        row.HeaderCell.Style.BackColor = Color.Empty
        cells = row.Cells
        for cell in cells:
            if cell.Style.BackColor == SELECTCOLOR:
                cell.Style.BackColor = REGULARCOLOR
                cell.Style.Font = REGL

def copytoclipboard(args):
    """
    Put data on Windows clipboard 
    in csv format.
    """
    csvx = ""
    if len(args) == 0:
        pass
        # clear clipboard
        # print 'clearing clipboard'
    elif len(args) == 1:
        csvx = args[0]
    else:
        csvx = ""
        csvx = ','.join(args)
    # print "-------------------", csvx
    dobj = DataObject()
    # hack from MSDN PostID 238181
    # this is a bit bizarre,
    #    but it works for getting csv data into Excel
    txt = Text.Encoding.Default.GetBytes(csvx)
    memstr = MemoryStream(txt)
    dobj.SetData('CSV', memstr)
    dobj.SetData(CSV, memstr)
    Clipboard.SetDataObject(dobj)

class DataGridEntry(Panel):
    """
    Container for the DataGridView control
    I'm trying to test.
    DataGridView is customized to have row
    and column of mouse-over'd cell highlighted.
    Also, there is a customized selection 
    color and selection limitations (one cell
    or one row at a time).
    """
    def __init__(self,posx,posy,height,width):
        """
        numcols is the number of columns
        in the grid.
        """
        self.Text = 'DataGridView Cell Format Test'
        # self.ClientSize = Size(400, 175)
        # self.MinimumSize = Size(height, width)
        # self.Height =height
        # self.Width =width
        self.ClientSize=Size(width,height)
        self.Location = Point(posx,posy)

        self.dgv = DataGridView()
        self.numcols = 0
        self.numrows = 0
        self.setupdatagridview()
        self.dgv.Location = Point(0, 0)
        self.dgv.ClientSize=Size(width,height)
        self.dgv.Parent = self
        clearselection(self.dgv)
        self.ToParet=None
    def ClearData(self):
        for i in range(0,self.dgv.Rows.Count):
            self.dgv.Rows.Clear()
        for i in range(0,self.dgv.Columns.Count):
            self.dgv.Columns.Clear()   
    def setupdatagridview(self):
        """
        General construction of DataGridView control.
        Bind mouse events as appropriate.
        """

        # have to have columns defined before inserting rows
        # self.dgv.ColumnCount = self.numcols
        # center all text in all data cells by default
        self.dgv.DefaultCellStyle.Alignment = MDDLCNTR
        # use Mouse events for contingency that actual
        #    position is required
        #    otherwise, can use events without "Mouse"
        #       in them
        # CellMouseEnter event for formatting
        self.dgv.CellMouseEnter += self.onmouseovercell
        # CellMouseLeave event for formatting
        self.dgv.CellMouseLeave += self.onmouseleavingcell
        # another try at MouseClick (avoiding default select color)
        self.dgv.CellMouseUp += self.onmouseclickcell
        self.dgv.CellDoubleClick += self.cellDoubleClick 
        # add empty rows first
        # for num in xrange(self.numrows):
        #     self.dgv.Rows.Add()
        # format empty cells
        resetcellfmts(self.dgv, self.numrows)
        # lock control so user cannot do anything to it datawise
        self.dgv.AllowUserToAddRows = False
        self.dgv.AllowUserToDeleteRows = False
        self.dgv.ReadOnly = True
        self.dgv.ClearSelection()
    def formatheaders(self,headerNames=["column1","column2","column3"]):
        """
        Get header row and left side column 
        populated and formatted.
        """
        self.numcols = len( headerNames)
        self.dgv.ColumnCount = self.numcols 
        for num in xrange(self.numcols):
            col = self.dgv.Columns[num]
            col.Name = headerNames[num]
            col.HeaderCell.Style.Alignment = MDDLCNTR
            col.HeaderCell.Style.Font = BOLD 
            col.HeaderCell.Style.ForeColor = Color.MidnightBlue

        self.dgv.TopLeftHeaderCell.Value = '序号'
        self.dgv.TopLeftHeaderCell.Style.Font = BOLD
        self.dgv.TopLeftHeaderCell.Style.ForeColor = Color.Blue
        self.dgv.RowHeadersWidth = ROWHDRWDTH
    def adddata(self):
        """
        Put data into the grid, 
        row by row, column by column.
        """
        # self.addRowdata( "1",["1","2","3"])
        # self.addRowdata( "2",["11","21","13"])

    def setRowTitle(self,rowobj,rowtitle):
            rowobj.HeaderCell.Value =rowtitle
            rowobj.HeaderCell.Style.Font = BOLD 
            rowobj.HeaderCell.Style.ForeColor = Color.Blue
    def addRowdata(self,rowtitle="1",datas=["1","2","3"]):        
            newidx = self.dgv.Rows.Add()
            row = self.dgv.Rows[newidx]
            self.setRowTitle(row,rowtitle)
            dat = (datax for datax in datas ) 
            for cell in row.Cells:
                cell.Value = dat.next()  
    def getcell(self, event):
        """
        Gets DataGridViewCell that is responding
        to an event.
        Attempt to minimize code duplication by 
        applying method to multiple events.
        """
        colidx = event.ColumnIndex
        rowidx = event.RowIndex
        if rowidx > INDEXERROR and colidx > INDEXERROR:
            row = self.dgv.Rows[rowidx]
            cell = row.Cells[colidx]
            return cell
        else:
            return None

    def onmouseovercell(self, sender, event):
        return
        """
        Change format of data cells
        when mouse passes over them.
        """
        cell = self.getcell(event)
        rowidx, colidx = getcellidxs(event)
        if cell:
            col = self.dgv.Columns[colidx]
            col.HeaderCell.Style.BackColor = COLUMNCOLOR
            row = self.dgv.Rows[rowidx]
            if row.HeaderCell.Style.BackColor != SELECTCOLOR:
                row.HeaderCell.Style.BackColor = ROWCOLOR
            cell.Style.Font = BOLD 
            if cell.Style.BackColor != SELECTCOLOR:
                cell.Style.BackColor = MOUSEOVERCOLOR
            row = self.dgv.Rows[rowidx]
            for cellx in row.Cells:
                if cellx.ColumnIndex != colidx:
                    if cellx.Style.BackColor != SELECTCOLOR:
                        cellx.Style.BackColor = ROWCOLOR
                        cellx.Style.Font = BOLD
            for num in xrange(self.numrows):
                for num2 in xrange(self.numcols):
                    if num != rowidx:
                        if num2 == colidx:
                            row = self.dgv.Rows[num]
                            cell = row.Cells[num2]
                            # 6) color all other cells in column cyan
                            # skip selected cells
                            if cell.Style.BackColor != SELECTCOLOR:
                                cell.Style.BackColor = COLUMNCOLOR
    def onmouseleavingcell(self, sender, event):
        return
        """
        Change format of data cells
        back to "normal" when mouse passes 
        out of the cell.
        """
        cell = self.getcell(event)
        rowidx, colidx = getcellidxs(event)
        if cell:
            resetcellfmts(self.dgv, self.numrows)
            resetheaderfmts(self.dgv, rowidx, colidx)
            clearselection(self.dgv)

    def onmouseclickcell(self, sender, event):
        return
        """
        Attempt to override selection.
        """
        selected = []
        mockclearselection(self.dgv)
        cell = self.getcell(event)
        self.rowidx, self.colidx = getcellidxs(event)
        clearselection(self.dgv)
        if cell:
            cell.Style.Font = BOLD 
            cell.Style.BackColor = SELECTCOLOR
            selected.append(cell.Value)
        if self.colidx == INDEXERROR:
            if self.rowidx != INDEXERROR:
                row = self.dgv.Rows[self.rowidx]
                cells = row.Cells
                for cell in cells:
                    cell.Style.Font = BOLD
                    cell.Style.BackColor = SELECTCOLOR
                    selected.append(cell.Value)
                row.HeaderCell.Style.BackColor = SELECTCOLOR
        self.onmouseovercell(sender, event)
        copytoclipboard(selected)
        
    def cellDoubleClick(self,sender,event):   # CellDoubleClick 
        cell = self.getcell(event)
        # print( cell )
        self.rowidx, self.colidx = getcellidxs(event)
        if  self.ToParet != None  and ( self.rowidx != -1   ):
                rowobj= self.dgv.Rows[self.rowidx ]
                rowtitle = rowobj.HeaderCell.Value 
                self.ToParet(self,rowtitle)
        else:
            pass
    def Clear(self):
        self.dgv.Rows.Clear();
    def getCurRowObj(self):
        return self.dgv.Rows[self.rowidx ]

class Button_Self(Button):
    def getFont(self):
        fontsize = self.ClientSize.Height*9/10
        # print(self.ClientSize.Height ,self.Size.Height )
        font= Font("仿宋", fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
        return font   
    def __init__(self):
        print(self.Font) 
        self.Font= self.getFont( );

class 日期Entry( Panel):
    def getFont(self):
        fontsize = self.ClientSize.Height*7/10
        # print(self.ClientSize.Height ,self.Size.Height )
        font= Font("仿宋", fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
        return font      
    def __init__(self,posx,posy,hight,width,title,curvlaue,name):
        super( 日期Entry,self ).__init__(self);
        self.格式说明="格式：YYYY.MM.DD"
        # self.Font= Font("楷体_GB2312", 24);
        self.Notify=None
        self.curvalue =self.格式说明
        self.ctrlwidth = width /2 ;
        self.Location = Point(posx,posy)
        self.Size = Size(width,hight )
        self.name =name

        # fontsize = self.ClientSize.Height*7/10
        # print(self.ClientSize.Height ,self.Size.Height )
        self.Font= self.getFont( );

        self.label = Label()
        self.label.Location = Point(0, 0)
        self.label.Parent = self
        self.label.Text = title
        self.label.Font= self.Font 
        self.label.TextAlign  = ContentAlignment.MiddleRight
        self.label.Size = Size(self.ctrlwidth,hight )

        self.cb = TextBox()
        self.cb.Parent = self
        self.cb.Location = Point( self.ctrlwidth,0)
        self.cb.Text=self.curvalue
        self.cb.Font= self.Font 
        self.cb.ForeColor = Color.Blue
        self.cb.TextChanged += self.OnChanged
        self.cb.Size = Size(self.ctrlwidth,hight )

    def OnChanged(self, sender, event):
        self.curvalue = sender.Text
        ret = self.validateText()
        if ret == True:
             self.curvalue = sender.Text     
        else:
            self.curvalue = None   
        if self.Notify != None:
                sender = self
                self.Notify(sender)

    def todayStr(self):
        import datetime
        t1 = datetime.date.today()
        str1= "{:04d}.{:02d}.{:02d}".format(t1.year,t1.month,t1.day)
        return str1
    def convert2DT(self,dtstr):
        # print( dtstr )
        a= re.split("[\.年月日/。]",dtstr)
        if len(a) != 3:
            return False
        try:
            # print( a)
            year=int(a[0])
            month=int(a[1])
            day=int(a[2])

            dtObj= datetime.date( year ,month ,day  )
            # print(year,month,day,dtObj)
            if (dtObj.year  != year or dtObj.month  != month or dtObj.day  != day):
                return False
            else:
                return True
        except:
            return False
    def validateText(self):
        t= self.curvalue;
        # print("------timei is------",t)
        if t == None or len(t) ==0 :
            return False
        if self.格式说明 == t :
            return False
        ty= self.convert2DT(t)
        # print(ty)
        if ty== True:
           self.cb.ForeColor = Color.Blue
        else:
            if "TODAY" in t.upper() :
                a1 = self.todayStr()
                # print a1
                self.cb.Text=a1
                self.cb.ForeColor = Color.Blue
                return True

            if( t == "" ):
                self.cb.Text=self.格式说明

            self.cb.ForeColor = Color.Red
            return False        
        return ty 
    def get(self):
        if self.curvalue:
            return self.curvalue
        else:
            return ""
    def set(self,valueStr):
        self.curvalue = valueStr
        self.cb.Text = valueStr
################################
class 金额Entry( Panel):
    def getFont(self):
        fontsize = self.ClientSize.Height*7/10
        # print(self.ClientSize.Height ,self.Size.Height )
        font= Font("仿宋", fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
        return font    
    def __init__(self,posx,posy,hight,width,title,curvlaue,name):
        super( 金额Entry,self ).__init__(self);
        self.格式说明="数值，或者类似 5吨*3元每吨*付款比例30%-扣减金额1000"
        # self.Font= Font("楷体_GB2312", 24);
        self.Notify=None
        self.curvalue =self.格式说明
        self.ctrlwidth = width /2 ;
        self.Location = Point(posx,posy)
        self.Size = Size(width,hight )
        self.title =title
        self.name =name

        # fontsize = self.ClientSize.Height*7/10
        # print(self.ClientSize.Height ,self.Size.Height )
        self.Font= self.getFont();


        self.label = Label()
        self.label.Location = Point(0, 0)
        self.label.Size = Size(self.ctrlwidth,hight )
        self.label.Parent = self
        self.label.Text = title
        self.label.Font= self.Font 
        self.label.TextAlign  = ContentAlignment.MiddleRight

        self.cb = TextBox()
        self.cb.Parent = self
        self.cb.Location = Point( self.ctrlwidth,0)
        self.cb.Size = Size(self.ctrlwidth,hight )
        self.cb.Text=self.curvalue
        self.cb.Font= self.Font 
        self.cb.ForeColor = Color.Blue
        self.cb.TextChanged += self.OnChanged
    def OnChanged(self, sender, event):
        self.curvalue = sender.Text

        ret = self.validateText()
        if ret == True:
             self.curvalue = sender.Text     
        else:
            self.curvalue = None 

        if self.Notify != None:
                sender = self
                self.Notify(sender)    
    def convert2Number(self,金额):
        a=0.0
        try:
            if 金额 == None or 金额 == self.格式说明 or 金额 == "" :
              return None
            b=re.findall("/[^0-9.]*[0-9.]*%",金额)
            for tgr in  b:
                src1= tgr.replace("/","*100/").replace("%","")
                金额 =金额.replace( tgr,src1)
            tgrstr = re.sub("[^0-9.*+-/]","",金额.replace("%","/100")) 
            # print tgrstr
            a=eval(tgrstr)
            self.label.Text = self.title+":"+str(a)
            # print(a,"------",tgrstr , 金额)
        except:
            return None
        return a

    def validateText(self):
        # print("======            ========")        
        t= self.curvalue;
        # print(t,"=======time =======")
        if t == None or len(t) <=0 :
            return False
        ty= self.convert2Number(t)
        if ty != None:
            self.cb.ForeColor = Color.Blue
        else:
            if( t == "" ):
                self.cb.Text = self.格式说明
            self.cb.ForeColor = Color.Red
        return ty != None 
    def get(self):
        if self.curvalue:
            return self.curvalue
        else:
            return "0"
    def set(self,valueStr):
        self.curvalue = valueStr
        self.cb.Text = valueStr        
class generalEntry( Panel):
    def getFont(self):
        fontsize = self.ClientSize.Height*7/10
        # print(self.ClientSize.Height ,self.Size.Height )
        font= Font("仿宋", fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
        return font
    def __init__(self,posx,posy,hight,width,title,curvlaue,name):
        super( generalEntry,self  ).__init__(self);
        self.格式说明="不能为空"

        self.Notify=None
        self.curvalue =self.格式说明
        # self.Width = width
        self.ctrlwidth = width /2 ;
        # self.Height =hight
        self.Size = Size(width,hight )
        self.Location = Point(posx,posy)
        self.title =title
        self.name =name
        
        self.Font= self.getFont()

        self.label = Label()
        self.label.Location = Point(0, 0)
        self.label.Size = Size(self.ctrlwidth,hight )        
        self.label.Parent = self
        self.label.Text = title
        self.label.Font= self.Font 
        self.label.TextAlign  = ContentAlignment.MiddleRight

        self.cb = TextBox()
        self.cb.Parent = self
        self.cb.Location = Point( self.ctrlwidth,0)
        self.cb.Size = Size(self.ctrlwidth,hight )  
        self.cb.Text=self.curvalue
        self.cb.Font= self.Font 
        self.cb.ForeColor = Color.Blue
        self.cb.TextChanged += self.OnChanged

    def OnChanged(self, sender, event):
        self.curvalue = sender.Text

        ret = self.validateText()
        if ret == True:
             self.curvalue = sender.Text     
        else:
            self.curvalue = None 

        if self.Notify != None:
                sender = self
                self.Notify(sender)    

    def validateText(self):
        t= self.curvalue;
        if t == None or len(t) <=0 :
            return False        
       # print(t,"   ++++++++++++++++++")
        ty=False
        if t !="" and t!=self.格式说明:
            ty=True
            self.cb.ForeColor = Color.Blue
        else:
            self.cb.Text = self.格式说明
            self.cb.ForeColor = Color.Red
        return ty 
    def get(self):
        # print self.curvalue
        if self.curvalue:
            return self.curvalue
        else:
            return ""
    def set(self,valueStr):
        self.curvalue = valueStr
        self.cb.Text = valueStr
class generalMultiLineEntry( generalEntry):
    def __init__(self,posx,posy,hight,width,title,curvlaue,name):
        super( generalMultiLineEntry,self  ).__init__(posx,posy,hight,width,title,curvlaue,name);
        self.cb.Multiline =True
    def get(self):
        if self.curvalue:
            return self.curvalue
        else:
            return ""
    def set(self,valueStr):
        self.curvalue = valueStr
        self.cb.Text = valueStr 
    def getFont(self):
        fontsize = self.ClientSize.Height*8/10/10
        # print(self.ClientSize.Height ,self.Size.Height )
        font= Font("仿宋", fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
        return font               
class ComboCtrlview( Panel ):
    def getFont(self):
        fontsize = self.ClientSize.Height*7/10
        # print(">>>>:",self.ClientSize.Height ,self.Size.Height ,fontsize )
        font= Font("仿宋", fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
        return font       
    def __init__(self,posx,posy,hight,width,title,curvlaue,values,name):
        # print( width ,"^^^^^^^^^^^^^^^")
        self._Padding = 3
        super( ComboCtrlview,self ).__init__(self);
        self._Values = values
        self.Notify=None
        self.curvalue = ""
        self.ctrlwidth = (width-self._Padding) /2 ;
        self.Size = Size(width,hight)
        self.Location = Point(posx,posy)
        self.name =name

        self.Font =self.getFont()

        self.label = Label()
        self.label.Parent = self
        self.label.Text = title
        self.label.Font= self.Font 
        self.label.Location = Point(0, 0)        
        self.label.Size = Size(self.ctrlwidth,hight)
        self.label.TextAlign  = ContentAlignment.MiddleRight
        
        self.cb = ComboBox()
        self.cb.Parent = self
        self.cb.Items.AddRange( self._Values )
        if len( curvlaue) >0:
            self.cb.Text=curvlaue
        self.cb.Font= self.Font 
        self.cb.ForeColor = Color.Blue
        self.cb.SelectedValueChanged +=self.OnChanged
        self.cb.DropDown += self.TextUpdate
        self.cb.TextChanged += self.TextUpdate

        self.cb.Location   = Point( self.label.Right +self._Padding,0)        
        self.cb.Size = Size(  self.ctrlwidth,hight)
    def TextUpdate(self, sender, event):
        pass
    def OnChanged(self, sender, event):
        # return 
        # print("OnChanged" )
        self.curvalue = sender.Text
        if self.Notify != None:
                sender = self
                self.Notify(sender)
    def get(self):
        if self.cb.Text:
            return self.cb.Text
        else:
            return ""
    def set(self,valueStr):
        self.curvalue = valueStr
        self.cb.Text = valueStr
    def setValues(self,values):
        self._Values = values
        self.cb.Items.Clear()
        self.cb.Items.AddRange( values )
        ll = len( values )
        # print(ll)
        self.cb.Text =values[ ll -1 ] 
        # print(values,self.cb.Text)
class SelectEditEntry( ComboCtrlview ):
    def validateText(self):
        return  True
    def __init__(self, posx,posy,hight,width,title,curvlaue,values,name,colname1 ="NO"  ):
        self.colname = colname1
        super(SelectEditEntry,self).__init__(posx,posy,hight,width,title,curvlaue,values,name)
        key=""
        a  =self.getTips(key,self.colname )  
        if a  == None or  len(a) == 0:
            # print( "A is NULL ",colname1)
            a = values

        else:
            # for a1 in a :
            #     print("    =",colname1,a1)
            self.cb.Items.Clear()
            self.cb.Items.AddRange( a )
            self._Values = a
        self.cb.AutoCompleteMode  =AutoCompleteMode.Suggest #Append
        self.cb.AutoCompleteSource = AutoCompleteSource.CustomSource
        b= AutoCompleteStringCollection()
        b.AddRange( a )
        self.cb.AutoCompleteCustomSource =  b        
        # for a1 in b:
        #     print("--",a1)

        # self.cb.DropDown += self.检查内容
    def getTips(self,key,Colname):
        return None
        try:
            pass
            # a  = tips.getAllPossible(Colname , key )
            # for a1 in a:
            #      print("--",a1)
            # return a
        except:
            return None
    def 检查内容(self,sender,event):
        key =  sender.Text 
        if self.getTips( self.colname,key ) != None:
            self.cb.Items.Clear()
            self.cb.Items.AddRange( a )
    def TextUpdate(self, sender, event):
        curvalue1 = sender.Text
        values = ()
        for v in self._Values:
            if curvalue1 in v :
                values += ( v, )
                # print("-:",v)
        # print( values,curvalue1 )
        self.cb.Items.Clear()        
        self.cb.Items.AddRange( values ) 
        if self.Notify != None:
                sender = self
                self.Notify(sender)                       
        


  