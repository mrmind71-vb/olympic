VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form CompSal 
   Caption         =   " ﬁ«—Ì— «·«’‰«›"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1410
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin MSDBCtls.DBCombo xList 
      Bindings        =   "CompSal.frx":0000
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   825
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   540
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   90
      Width           =   1065
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   945
      Width           =   990
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   495
      Width           =   990
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   450
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
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
   Begin VB.Label Lbllist 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ﬁ«∆„… :"
      Height          =   195
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   900
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï ‘Â—"
      Height          =   195
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   510
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ ‘Â—"
      Height          =   195
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "CompSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nMonth1 As Integer
Dim nMonth2 As Integer
Dim nMonth3 As Integer
Dim nMonth4 As Integer
Dim nMonth5 As Integer
Dim nMonth6 As Integer
Private Sub CmdApply_Click()
CompSal
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Sub CompSal()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
Dim listTable As Recordset
Dim GroupTable As Recordset

Set listTable = mydb.OpenRecordset("FILE0_40")
Set GroupTable = mydb.OpenRecordset("File1_50")
listTable.Index = "nFile0_40_1"
GroupTable.Index = "nFile1_50_1"
tempdb.Execute "Delete * From Temp"
Set TargetTable = mydb.CreateDynaset("TEMP")
nMonth1 = Val(Date1.Text)
nMonth2 = Val(Date1.Text) + 1
nMonth3 = Val(Date1.Text) + 2
nMonth4 = Val(Date1.Text) + 3
nMonth5 = Val(Date1.Text) + 4
nMonth6 = Val(Date1.Text) + 5

If nMonth1 > Val(date2.Text) Then nMonth1 = 0
If nMonth2 > Val(date2.Text) Then nMonth2 = 0
If nMonth3 > Val(date2.Text) Then nMonth3 = 0
If nMonth4 > Val(date2.Text) Then nMonth4 = 0
If nMonth5 > Val(date2.Text) Then nMonth5 = 0
If nMonth6 > Val(date2.Text) Then nMonth6 = 0

cString = "Select Mid(File1_11.Item,1,6) as uItem,First(File1_10.DescA) as ItemDesc,pack,[Group]," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth1 & ",[OUT],0))As Sal1  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth2 & " ,[OUT],0))As Sal2  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth3 & " ,[OUT],0))As Sal3  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth4 & " ,[OUT],0))As Sal4  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth5 & " ,[OUT],0))As Sal5  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth6 & " ,[OUT],0))As Sal6  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth1 & " ,[IN] ,0))As Ret1  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth2 & " ,[IN] ,0))As Ret2  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth3 & " ,[IN] ,0))As Ret3  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth4 & " ,[IN] ,0))As Ret4  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth5 & " ,[IN] ,0))As Ret5  ," & _
          "Sum(IIF(MONTH(CDATE)= " & nMonth6 & " ,[IN] ,0))As Ret6  " & _
          " From File1_11 inner Join File1_10 On file1_10.item = file1_11.item " & _
          " WHERE( cType = '6' Or cType = '3' ) " & _
          " Group By Mid(File1_11.Item,1,6),pack,[group]"
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    listTable.Seek ">=", xList.BoundText, SourceTable.uitem
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.[Group]
    TargetTable.str2 = RetSeek(GroupTable, SourceTable.[Group], "Desca")
    TargetTable.str3 = SourceTable.uitem
    TargetTable.str4 = SourceTable.Itemdesc
    TargetTable.STR5 = Date1.Text
    TargetTable.STR7 = date2.Text
    
    TargetTable.val1 = nMonth1
    TargetTable.val2 = nMonth2
    TargetTable.VAL3 = nMonth3
    TargetTable.VAL4 = nMonth4
    TargetTable.VAL5 = nMonth5
    TargetTable.VAL6 = nMonth6
    
    TargetTable.str15 = TurnValue(.SAL1, Null, 0) - TurnValue(.Ret1, Null, 0)
    TargetTable.str16 = TurnValue(.SAL2, Null, 0) - TurnValue(.Ret2, Null, 0)
    TargetTable.str17 = TurnValue(.SAL3, Null, 0) - TurnValue(.Ret3, Null, 0)
    TargetTable.str18 = TurnValue(.SAL4, Null, 0) - TurnValue(.Ret4, Null, 0)
    TargetTable.str19 = TurnValue(.SAL5, Null, 0) - TurnValue(.Ret5, Null, 0)
    TargetTable.str20 = TurnValue(.SAL6, Null, 0) - TurnValue(.Ret6, Null, 0)
    
    If myFound(listTable, Array(xList.BoundText, SourceTable.uitem)) Then
        TargetTable.STR6 = listTable.amountt
    Else
        TargetTable.STR6 = "€Ì— „”Ã·"
    End If
    TargetTable.str8 = xList.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

'    TargetTable.Date1 = Date1.Text
'    TargetTable.date2 = date2.Text
    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\CompSal.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub

Private Sub Form_Load()

    Data1.DatabaseName = MdbPath
    Data1.RecordSource = "Select Doc_No,DescA From File0_40 Group by Doc_No,DescA"
    xList.BoundColumn = "Doc_no"
    xList.ListField = "Desca"

End Sub
Private Sub xList_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
