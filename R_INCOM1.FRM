VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form RepIncome1 
   Caption         =   "ÊÞÇÑíÑ ÇáÇíÑÇÏÇÊ"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
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
   ScaleHeight     =   2445
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -1050
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1125
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   675
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo xComp 
      Bindings        =   "R_incom1.frx":0000
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   975
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "ÚÑÖ"
      Height          =   390
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1950
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ÎÑæÌ"
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   525
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   600
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   1
      BoundReportHeading=   "dddd"
      WindowState     =   2
   End
   Begin MSDBCtls.DBCombo xIncome 
      Bindings        =   "R_incom1.frx":0014
      Height          =   315
      Left            =   1650
      TabIndex        =   3
      Top             =   1425
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.Label LblIncome 
      AutoSize        =   -1  'True
      Caption         =   "ÇáÇíÑÇÏ :"
      Height          =   240
      Left            =   4125
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ÇáÔÑßÉ :"
      Height          =   240
      Index           =   0
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1050
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Çáì ÊÇÑíÎ :"
      Height          =   240
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ãä ÊÇÑíÎ :"
      Height          =   240
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   150
      Width           =   690
   End
End
Attribute VB_Name = "RepIncome1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IncomeTable As Recordset
Private Sub CmdApply_Click()
Select Case publicFlag
Case 5
    RepIncome2
Case 6
    RepIncome3
End Select
End Sub
Private Sub RepIncome1()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "Select Income,vValue,File8_60.DescA," & _
          "FILE8_71.MainGroup,FILE8_71.DescA " & _
          " (From FILE8_60 Inner Join FILE8_71 on FILE8_60.Income = FILE8_71.Code) from file1_11 inner join file3_11 " & _
          " Where vDate >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and vDate <= DateValue(" & MyParn(date2.Text) & ")" & _
          IIf(xcomp.BoundText = "", "", " And Comp = " & MyParn(xcomp.BoundText)) & _
          " and file3_10.vGroup = " & MyParn(xGroup.BoundText) & _
          " Order By MainGroup,Income,vDate"
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ Ýì ÇáÊÞÑíÑ ¿"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .[FILE8_60.INCOME]
    TargetTable.str2 = .DESCA
    If xcomp.BoundText <> "" Then TargetTable.str3 = xcomp.Text
    TargetTable.str4 = .MainGroup
    TargetTable.STR5 = RetIncome(.MainGroup)
    TargetTable.STR6 = .[FILE8_71.DESCA]
    TargetTable.VAL1 = .vValue
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = App.Path & "\Reports\IncRep1.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub RepIncome2()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "Select Sum(Value) as SumofValue," & _
          "FILE8_71.MainGroup,FILE8_71.DescA " & _
          " From FILE8_60 Inner Join FILE8_71 on FILE8_60.Income = FILE8_71.Code " & _
          " Where [Date] >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and [Date] <= DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By Income," & _
          "FILE8_71.MainGroup,FILE8_71.DescA"


Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ Ýì ÇáÊÞÑíÑ ¿"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .INCOME
    TargetTable.str2 = .DESCA
    TargetTable.str4 = .MainGroup
    TargetTable.STR5 = RetIncome(.MainGroup)
    TargetTable.VAL1 = .Sumofvalue
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = App.Path & "\Reports\IncRep1.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub RepIncome3()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select vDate,DescA,vValue" & _
          " From FILE8_60" & _
          " Where vDate >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and vDate <= DateValue(" & MyParn(date2.Text) & ")" & _
          " and Income = " & MyParn(xIncome.BoundText) & _
          IIf(xcomp.BoundText = "", "", " And Comp = " & MyParn(xcomp.BoundText)) & _
          " Order  By vDate"
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ Ýì ÇáÊÞÑíÑ ¿"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .DESCA
    TargetTable.str2 = xIncome.Text
    If xcomp.BoundText <> "" Then TargetTable.str3 = xcomp.Text
    TargetTable.VAL1 = .vValue
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = .vDate
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = App.Path & "\Reports\RepChrg2.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Set IncomeTable = mydb.OpenRecordset("Select * From File1_70 where Flag = 13", dbOpenSnapshot)
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select * From file1_70 where Flag = 5"
xcomp.BoundColumn = "Code"
xcomp.ListField = "DescA"

If publicFlag = 6 Then
    Data2.DatabaseName = MdbPath
    Data2.RecordSource = "File8_71"
    xIncome.BoundColumn = "Code"
    xIncome.ListField = "DescA"
End If
End Sub
Private Function RetIncome(pIncome)
IncomeTable.FindFirst "Code = " & MyParn(pIncome)
If Not IncomeTable.NoMatch Then RetIncome = IncomeTable.DESCA
End Function

Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xcomp.BoundText = ""
End Sub
Function MYVALID()
If Not IsDate(Date1.Text) Then Exit Function
If Not IsDate(date2.Text) Then Exit Function
If DateValue(Date1.Text) > DateValue(date2.Text) Then Exit Function
If xIncome.Visible And xIncome.BoundText = "" Then Exit Function
MYVALID = True
End Function

