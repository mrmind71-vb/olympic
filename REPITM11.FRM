VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form itemRep1 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1125
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox xDate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3675
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "Repitm11.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2550
      TabIndex        =   7
      Top             =   600
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1875
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2325
      Width           =   1515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   525
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Repitm11.frx":0014
      DataSource      =   "Data3"
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Top             =   150
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
   End
   Begin MSDBCtls.DBCombo xGroup2 
      Bindings        =   "Repitm11.frx":0028
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   1050
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "Repitm11.frx":003C
      DataSource      =   "Data3"
      Height          =   315
      Left            =   2550
      TabIndex        =   11
      Top             =   1950
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lblStore1 
      BackColor       =   &H8000000B&
      Caption         =   "„‰ „Œ“‰"
      Height          =   240
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2025
      Width           =   1005
   End
   Begin VB.Label LblDate 
      BackColor       =   &H8000000B&
      Caption         =   "«· «—ÌŒ"
      Height          =   240
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1575
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "„‰ „Ã„Ê⁄… "
      Height          =   315
      Index           =   0
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "≈·Ï „Ã„Ê⁄…"
      Height          =   255
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1125
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      Height          =   195
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   1110
   End
End
Attribute VB_Name = "itemRep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StockTable As Recordset, uStockTable As Recordset
Dim storeTable As Recordset
Dim cStore As String
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Select Case publicFlag
Case 1
    RepItem1
Case 2
    RepItem2
Case 3
    RepItem3
End Select
End Sub
Sub RepItem1()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")
If xgroup.BoundText = "" Then
    cString = "SELECT FILE1_10.*, FILE1_50.DESCA as grdesc , FILE1_70.DESCA as MGrdesc , file1_50.M_GROUP" & _
              " FROM (FILE1_50 INNER JOIN FILE1_10 ON FILE1_50.CODE = FILE1_10.GROUP) INNER JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE " & _
              " where file1_70.FLAG = 2 " & _
              " order by FILE1_50.M_GROUP,file1_10.group , file1_10.ITEM "
Else
    cString = "SELECT FILE1_10.*, FILE1_50.DESCA as grdesc , FILE1_70.DESCA as MGrdesc, file1_50.m_group" & _
              " FROM (FILE1_50 INNER JOIN FILE1_10 ON FILE1_50.CODE = FILE1_10.GROUP) INNER JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE " & _
              " where file1_70.code = " & MyParn(xgroup.BoundText) & _
              IIf(xGroup1.BoundText <> "", " and File1_10.Group >= " & MyParn(xGroup1.BoundText), "") & _
              IIf(xGroup2.BoundText <> "", " and File1_10.Group <= " & MyParn(xGroup2.BoundText), "") & _
              " AND file1_70.FLAG = 2 " & _
              " order by FILE1_50.M_GROUP,file1_10.group , file1_10.ITEM "
End If


Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
i = 1
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
            TempTable.AddNew
            TempTable.STR6 = .MGrdesc
            TempTable.STR5 = .M_GROUP
            TempTable.str1 = .Item
            TempTable.str2 = .DESCA
            TempTable.str3 = .Group
            TempTable.str4 = .grdesc
            TempTable.VAL3 = i
            
            TempTable.STR7 = .COST
            TempTable.str9 = .price
            TempTable.str8 = firsttitle
            TempTable.str9 = Secondtitle
            TempTable.Update
            i = i + 1
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm1.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Sub RepItem22()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")

If xStore2.BoundText <> "" Then
    cStoreString = " and cNo Between " & _
                   MyParn(xStore1.BoundText) & " and " & _
                   MyParn(xStore2.BoundText)
    cReportTitle = "„‰ " & xStore1.Text & _
                   "≈·Ï " & xStore2.Text
Else
    cStoreString = IIf(xStore1.BoundText <> "", " and cNo = " & MyParn(xStore1.BoundText), "")
    cReportTitle = IIf(xStore1.BoundText <> "", xStore1.Text, "")
End If


cString = "Select " & _
          "File1_11.item " & " As MidOfItem," & _
          "Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant," & _
          "File1_10.DescA ,File1_10.pack,File1_10.PRICE" & _
          " From File1_11 inner join File1_10 on File1_11.item = file1_10.item " & _
          " where cDate <= DateValue(" & MyParn(xDate.Text) & ")" & _
          cStoreString & _
          " Group By " & _
          "File1_11.item," & _
          "File1_10.DescA,File1_10.Pack,File1_10.PRICE"
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
i = 1
If .RecordCount > 0 Then
    Do While Not .EOF
        If .Quant <> 0 Then
            TempTable.AddNew
            TempTable.str1 = .MidofItem
            TempTable.str2 = .DESCA
            TempTable.str3 = .Quant
            TempTable.str4 = cReportTitle
            
            TempTable.val1 = file1_10.COST
            TempTable.val2 = file1_10.price
            TempTable.VAL3 = i + 1
            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm3.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Sub RepItem33()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")
If xStore2.BoundText <> "" Then
    cStoreString = " and cNo Between " & _
                   MyParn(xStore1.BoundText) & " and " & _
                   MyParn(xStore2.BoundText)
    cReportTitle = "„‰ " & xStore1.Text & _
                   "≈·Ï " & xStore2.Text
Else
    cStoreString = ""
    cReportTitle = ""
End If

cString = "Select File1_11.Item ," & _
          "Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant," & _
          "File1_10.DescA ,File1_10.pack,File1_10.Cost,File1_10.E_date" & _
          " From File1_11 inner join File1_10 on File1_11.item = file1_10.item " & _
          " where E_Date <= DateValue(" & MyParn(xDate.Text) & ")" & _
          cStoreString & _
          " Group By " & _
          "File1_11.item," & _
          "File1_10.DescA ,File1_10.pack,File1_10.Cost,File1_10.E_date"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        If .Quant <> 0 Then
            TempTable.AddNew
            TempTable.str1 = .Item
            TempTable.str2 = .DESCA
            TempTable.str3 = .Quant
            TempTable.str4 = cReportTitle
            TempTable.val1 = .price
            TempTable.val2 = .Quant * .price / .Pack
            TempTable.str10 = IIf(Expire, "» «—ÌŒ «·’·«ÕÌ…", "»œÊ‰  «—ÌŒ «·’·«ÕÌ…")
            TempTable.Date1 = xDate.Text
            TempTable.date2 = .E_Date
            TempTable.str19 = firsttitle
            TempTable.str20 = Secondtitle

            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItem9.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command1_Click()
xDate.Text = ""
xGroup1.Text = ""
xGroup2.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set StockTable = mydb.OpenRecordset("File1_30")
Set storeTable = mydb.OpenRecordset("File1_30")
Select Case publicFlag
Case 2
    LblDate.Visible = True
    xDate1.Visible = True
    xDate2.Visible = True
Case 3
    LblDate.Visible = True
    xDate1.Visible = True
    lblStore1.Visible = True
    xStore1.Visible = True
End Select

storeTable.Index = "nfile1_30_1"

Data2.DatabaseName = MdbPath
Data2.RecordSource = "QFILE1_70_1"
xgroup.BoundColumn = "Code"
xgroup.ListField = "Desca"

Data1.DatabaseName = MdbPath
Data1.RecordSource = "File1_50 "
xGroup1.BoundColumn = "Code"
xGroup1.ListField = "Desca"
xGroup2.BoundColumn = "Code"
xGroup2.ListField = "Desca"

Data3.DatabaseName = MdbPath
Data3.RecordSource = "Select * From File1_70 Where flag = 1"
xStore1.BoundColumn = "Code"
xStore1.ListField = "Desca"
End Sub
Private Sub Store_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xStore1.BoundText = ""
End Sub
Private Function MYVALID()
MYVALID = True
End Function
Private Sub xGroup_Change()
Data1.RecordSource = "Select * From File1_50 where M_GROUP = " & MyParn(xgroup.BoundText)
Data1.Refresh
End Sub
Sub OLDRepItem2()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif("cType = '6'", "[OUT]") & " as Sales,"
cField2 = myiif("cType = '3'", "[IN]") & " as RetSales, "
cField3 = myiif("cType = '6'", "[TOTAL]") & " as SalesValue, "
cField4 = myiif("cType = '3'", "[TOTAL]") & " as RetSalesValue "

If xgroup.Text = "" Then
    cString = "Select File1_11.Item," & _
              "First(File1_10.DescA) as FirstOfDescA," & _
              "File1_50.DescA as GroupDescA,FILE1_50.M_GROUP,File1_70.DescA as MainGroupDescA, " & _
              cField1 & cField2 & cField3 & cField4 & _
              " From ((File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item)  " & _
              "  Inner Join File1_50 on File1_10.Group = File1_50.code ) " & _
              "  inner Join file1_70 on File1_50.m_Group = File1_70.Code " & _
              " where File1_70.flag = 2 " & _
              " and cDate Between DateValue(" & MyParn(xDate1.Text) & ")" & _
              " and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " Group By File1_11.Item,File1_50.DescA,File1_70.DescA,FILE1_50.M_GROUP"
Else
    cString = "Select File1_11.Item," & _
              "First(File1_10.DescA) as FirstOfDescA,FILE1_50.M_GROUP," & _
              "File1_50.DescA as GroupDescA,File1_70.DescA as MainGroupDescA, " & _
              cField1 & cField2 & cField3 & cField4 & _
              " From ((File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item)  " & _
              " Inner Join File1_50 on File1_10.Group = File1_50.code ) " & _
              " inner Join file1_70 on File1_50.m_Group = File1_70.Code " & _
              " where File1_70.flag = 2 " & _
              " and cDate Between DateValue(" & MyParn(xDate1.Text) & ")" & _
              " and DateValue(" & MyParn(xDate2.Text) & ")" & _
              IIf(xgroup.BoundText <> "", " and File1_50.m_Group = " & MyParn(xgroup.BoundText), "") & _
              IIf(xGroup1.BoundText <> "", " and File1_10.Group >= " & MyParn(xGroup1.BoundText), "") & _
              IIf(xGroup2.BoundText <> "", " and File1_10.Group <= " & MyParn(xGroup2.BoundText), "") & _
              " Group By File1_11.Item,File1_50.DescA,File1_70.DescA,FILE1_50.M_GROUP"
End If

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.Item
    TargetTable.str2 = SourceTable.FirstofDescA
    TargetTable.str3 = " ≈Ã„«·Ï „»Ì⁄«  «·√’‰«› ﬂ„Ì… - ﬁÌ„…"
    TargetTable.val1 = .Sales
    TargetTable.val2 = .RetSales
    TargetTable.VAL3 = .Sales - .RetSales
    TargetTable.VAL7 = SourceTable.salesvalue
    TargetTable.VAL8 = SourceTable.retsalesvalue
    TargetTable.VAL9 = SourceTable.salesvalue - SourceTable.retsalesvalue
    
    TargetTable.Date1 = xDate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext

Loop Until SourceTable.EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm7.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Sub RepItem3()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Dim nDoz As Double
Dim nCARTON As Double
Set TempTable = tempdb.OpenRecordset("Temp")

If xgroup.BoundText = "" Then
    cString = "Select File1_11.Item,File1_10.COST,File1_10.group,First(FILE1_10.DOZ) AS F_DOZ , First(FILE1_10.CARTON ) AS F_CARTON," & _
              "First(File1_10.DescA) as FirstOfDescA," & _
              "File1_50.DescA as GroupDescA,FILE1_50.M_GROUP,File1_70.DescA as MainGroupDescA, " & _
              " Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant " & _
              " From ((File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item)  " & _
              " Inner Join File1_50 on File1_10.Group = File1_50.code ) " & _
              " inner Join file1_70 on File1_50.m_Group = File1_70.Code " & _
              " where File1_70.flag = 2 " & _
              IIf(xDate1.Text <> "", " and cDate <= " & DateSql(xDate1.Text), "") & _
              IIf(xStore1.BoundText <> "", " and File1_11.cNO = " & MyParn(xStore1.BoundText), "") & _
              " Group By File1_11.Item,File1_50.DescA,File1_70.DescA,File1_10.COST,File1_10.group,FILE1_50.M_GROUP"
Else
    cString = "Select File1_11.Item,File1_10.COST,File1_10.group,FILE1_50.M_GROUP, First(FILE1_10.DOZ) AS F_DOZ , First(FILE1_10.CARTON ) AS F_CARTON ," & _
              "First(File1_10.DescA) as FirstOfDescA," & _
              "File1_50.DescA as GroupDescA,File1_70.DescA as MainGroupDescA, " & _
              " Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant " & _
              " From ((File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item)  " & _
              " Inner Join File1_50 on File1_10.Group = File1_50.code ) " & _
              " inner Join file1_70 on File1_50.m_Group = File1_70.Code " & _
              " where File1_70.flag = 2 " & _
              IIf(xDate1.Text <> "", " and cDate <= " & DateSql(xDate1.Text), "") & _
              IIf(xgroup.BoundText <> "", " and file1_70.code = " & MyParn(xgroup.BoundText), "") & _
              IIf(xGroup1.BoundText <> "", " and File1_10.Group >= " & MyParn(xGroup1.BoundText), "") & _
              IIf(xGroup2.BoundText <> "", " and File1_10.Group <= " & MyParn(xGroup2.BoundText), "") & _
              IIf(xStore1.BoundText <> "", " and File1_11.cNO = " & MyParn(xStore1.BoundText), "") & _
              " Group By File1_11.Item,File1_50.DescA,File1_70.DescA,File1_10.COST,File1_10.group,FILE1_50.M_GROUP"
End If

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount = 0 Then
    MsgBox "«·  ÊÃœ »Ì‰«  ·⁄—÷Â«"
    Exit Sub
End If
Do While Not .EOF
    If .Quant <> 0 Then
        TempTable.AddNew
        TempTable.str1 = .Item
        TempTable.str2 = .FirstofDescA
        TempTable.str4 = .MainGroupDescA
        TempTable.str3 = .M_GROUP
        TempTable.STR5 = .Group
        TempTable.STR6 = .GroupDescA
                
'       nDoz = .Quant / .DOZ
        TempTable.VAL4 = .COST
        TempTable.VAL5 = .COST * .Quant
        TempTable.val1 = .Quant / .F_DOZ
        TempTable.val2 = (.Quant / .F_DOZ) / .F_CARTON
        TempTable.VAL3 = .Quant
        TempTable.Date1 = IIf(IsDate(xDate1.Text), xDate1.Text, Null)
        TempTable.STR7 = " —’Ìœ «·√’‰«› ( ⁄œœ - œ” … - ﬂ— Ê‰… ) "
        If xDate1.Text <> "" Then
            TempTable.STR7 = TempTable.STR7 & " · «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY")
        End If
        If xStore1.BoundText <> "" Then
            TempTable.STR7 = TempTable.STR7 & " ·„Œ“‰  " & xStore1.Text
        End If
        
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle

        TempTable.Update
    End If
   .MoveNext
Loop
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm_3.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Sub RepItem2()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")

If xgroup.BoundText = "" Then
    cString = "Select File1_11.Item,File1_10.COST,File1_10.group, " & _
              "First(File1_10.DescA) as FirstOfDescA," & _
              "File1_50.DescA as GroupDescA,FILE1_50.M_GROUP,File1_70.DescA as MainGroupDescA, " & _
              " Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant " & _
              " From ((File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item)  " & _
              " Inner Join File1_50 on File1_10.Group = File1_50.code ) " & _
              " inner Join file1_70 on File1_50.m_Group = File1_70.Code " & _
              " where File1_70.flag = 2 " & _
              IIf(xDate1.Text <> "", " and cDate <= " & DateSql(xDate1.Text), "") & _
              IIf(xStore1.BoundText <> "", " and File1_11.cNO = " & MyParn(xStore1.BoundText), "") & _
              " Group By File1_11.Item,File1_50.DescA,File1_70.DescA,File1_10.Cost,File1_10.group,FILE1_50.M_GROUP"
Else
    cString = "Select File1_11.Item,File1_10.COST,File1_10.group,FILE1_50.M_GROUP, " & _
              "First(File1_10.DescA) as FirstOfDescA," & _
              "File1_50.DescA as GroupDescA,File1_70.DescA as MainGroupDescA, " & _
              " Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant " & _
              " From ((File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item)  " & _
              " Inner Join File1_50 on File1_10.Group = File1_50.code ) " & _
              " inner Join file1_70 on File1_50.m_Group = File1_70.Code " & _
              " where File1_70.flag = 2 " & _
              IIf(xDate1.Text <> "", " and cDate <= " & DateSql(xDate1.Text), "") & _
              IIf(xgroup.BoundText <> "", " and file1_70.code = " & MyParn(xgroup.BoundText), "") & _
              IIf(xGroup1.BoundText <> "", " and File1_10.Group >= " & MyParn(xGroup1.BoundText), "") & _
              IIf(xGroup2.BoundText <> "", " and File1_10.Group <= " & MyParn(xGroup2.BoundText), "") & _
              IIf(xStore1.BoundText <> "", " and File1_11.cNO = " & MyParn(xStore1.BoundText), "") & _
              " Group By File1_11.Item,File1_50.DescA,File1_70.DescA,File1_10.Cost,File1_10.group,FILE1_50.M_GROUP"
End If
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount = 0 Then
    MsgBox "«·  ÊÃœ »Ì‰«  ·⁄—÷Â«"
    Exit Sub
End If
Do While Not .EOF
    If .Quant <> 0 Then
        TempTable.AddNew
        TempTable.str1 = .Item
        TempTable.str2 = .FirstofDescA
        TempTable.str4 = .MainGroupDescA
        TempTable.str3 = .M_GROUP
        TempTable.STR5 = .Group
        TempTable.STR6 = .GroupDescA
        
        TempTable.STR7 = " —’Ìœ «·√’‰«› ( ⁄œœ -  ﬂ·›… -  ﬁÌÌ„  ) "
        If xDate1.Text <> "" Then
            TempTable.STR7 = TempTable.STR7 & " · «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY")
        End If
        If xStore1.BoundText <> "" Then
            TempTable.STR7 = TempTable.STR7 & " ·„Œ“‰  " & xStore1.Text
        End If
        
        TempTable.val1 = .COST
        TempTable.val2 = .Quant * .COST
        TempTable.VAL3 = .Quant
        TempTable.Date1 = IIf(IsDate(xDate1.Text), xDate1.Text, Null)
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle

        TempTable.Update
    End If
   .MoveNext
Loop

End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm_2.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub


