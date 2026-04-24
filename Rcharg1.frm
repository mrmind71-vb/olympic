VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RCharg1 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
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
   ScaleHeight     =   2055
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data2"
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
      Top             =   375
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1590
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1590
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   375
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSDBCtls.DBCombo xCharge 
      Bindings        =   "Rcharg1.frx":0000
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xBox 
      Bindings        =   "Rcharg1.frx":0014
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1650
      TabIndex        =   9
      Top             =   840
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483643
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Œ“‰…"
      Height          =   195
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   945
      Width           =   315
   End
   Begin VB.Label LblCharge 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ"
      Height          =   195
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   592
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   225
      Width           =   675
   End
End
Attribute VB_Name = "RCharg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChargeTable As Recordset
Private Sub CmdApply_Click()
Select Case publicFlag
Case 2
    RepCharge2
Case 3
    RepCharge3
Case 5
    RepCharge5
Case 6
    RepCharge6

End Select
End Sub
Private Sub RepCharge2()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select Charge,Sum(Value) as SumofValue," & _
          "File8_70.MainGroup,File8_70.DescA " & _
          " From file8_50 Inner Join File8_70 on File8_50.Charge = File8_70.Code "
If xBox.BoundText <> "" Then
cString = cString & " Where BOX = " & MyParn(xBox.BoundText) & " AND  [Date] >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and [Date] <= DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By Charge," & _
          "File8_70.MainGroup,File8_70.DescA"
Else
cString = cString & " Where [Date] >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and [Date] <= DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By Charge," & _
          "File8_70.MainGroup,File8_70.DescA"
End If
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .CHARGE
    TargetTable.str2 = .DESCA
    TargetTable.str4 = .MainGroup
    TargetTable.STR5 = RetCharge(.MainGroup)
    TargetTable.VAL1 = .Sumofvalue

    TargetTable.str7 = "≈Ã„«·Ï „’«—Ì› "
    If xBox.BoundText <> "" Then TargetTable.str7 = "≈Ã„«·Ï „’«—Ì›  ·Œ“‰… " & xBox.Text
    TargetTable.str8 = "„‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
'   TargetTable.Date1 = Date1.Text
'   TargetTable.date2 = date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\RepChrg1.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub RepCharge3()
If Not MYVALID Then Exit Sub
If xCharge.BoundText = "" Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select Date,DescA,Value,DOC_NO,BOX" & _
          " From file8_50"
If xBox.BoundText <> "" Then
    cString = cString & " Where BOX = " & MyParn(xBox.BoundText) & " AND  Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
              " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
              " and Charge = " & MyParn(xCharge.BoundText) & _
              " Order  By Date"
Else
    cString = cString & " Where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
              " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
              " and Charge = " & MyParn(xCharge.BoundText) & _
              " Order  By Date"
End If
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .DESCA
    TargetTable.str2 = xCharge.Text
    TargetTable.str3 = .doc_no
    TargetTable.str4 = .BOX
    TargetTable.VAL1 = !Value
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = !Date
    TargetTable.str7 = "  ›’Ì·Ï „’—Ê› " & xCharge.Text
    If xBox.BoundText <> "" Then TargetTable.str7 = TargetTable.str7 & "   ·Œ“‰… " & xBox.Text
    TargetTable.str8 = "„‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\RepChrg3.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set ChargeTable = mydb.CreateSnapshot("Select * From File1_70 where Flag = 7")
If publicFlag = 3 Then
    Data2.DatabaseName = MdbPath
    Data2.RecordSource = "select * from file8_70 order by desca "
    xCharge.BoundColumn = "Code"
    xCharge.ListField = "DescA"
End If
If publicFlag = 6 Then
    Data2.DatabaseName = MdbPath
    Data2.RecordSource = "file8_71 "
    xCharge.BoundColumn = "Code"
    xCharge.ListField = "DescA"
End If
Data1.DatabaseName = MdbPath
Data1.RecordSource = "FILE0_50"
xBox.BoundColumn = "CODE"
xBox.ListField = "DESCA"

End Sub
Private Function RetCharge(pCharge)
ChargeTable.FindFirst "Code = " & MyParn(pCharge)
If Not ChargeTable.NoMatch Then RetCharge = ChargeTable.DESCA
End Function
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID()
If Not IsDate(Date1.Text) Then Exit Function
If Not IsDate(date2.Text) Then Exit Function
If DateValue(Date1.Text) > DateValue(date2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub RepCharge5()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select  Charge,Sum(Value) as SumofValue," & _
          "File8_71.MainGroup,File8_71.DescA " & _
          " From file8_60 Inner Join File8_71 on File8_60.Charge = File8_71.Code "
If xBox.BoundText <> "" Then
    cString = cString & " Where BOX = " & MyParn(xBox) & " AND  [Date] >= DateValue(" & MyParn(Date1.Text) & ")" & _
             " and [Date] <= DateValue(" & MyParn(date2.Text) & ")" & _
             " Group By Charge," & _
             "File8_71.MainGroup,File8_71.DescA"
Else
    cString = cString & " Where [Date] >= DateValue(" & MyParn(Date1.Text) & ")" & _
             " and [Date] <= DateValue(" & MyParn(date2.Text) & ")" & _
             " Group By Charge," & _
             "File8_71.MainGroup,File8_71.DescA"
End If
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .CHARGE
    TargetTable.str2 = .DESCA
    TargetTable.str4 = .MainGroup
    TargetTable.STR5 = RetCharge(.MainGroup)
    TargetTable.VAL1 = .Sumofvalue

    TargetTable.str7 = "≈Ã„«·Ï ≈Ì—«œ«  "
    If xBox.BoundText <> "" Then TargetTable.str7 = TargetTable.str7 & "   ·Œ“‰… " & xBox.Text
    TargetTable.str8 = "„‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text

'   TargetTable.Date1 = Date1.Text
'   TargetTable.date2 = date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\RepChrg1.rpt"
Report1.DataFiles(0) = TempPath
Report1.Action = 1
End Sub
Private Sub RepCharge6()
If Not MYVALID Then Exit Sub
If xCharge.BoundText = "" Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select Date,DescA,Value,DOC_NO,BOX" & _
          " From file8_60"
If xBox.BoundText <> "" Then
    cString = cString & " Where BOX = " & MyParn(xBox) & " AND  Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
             " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
             " and Charge = " & MyParn(xCharge.BoundText) & _
             " Order  By Date"
Else
    cString = cString & " Where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
             " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
             " and Charge = " & MyParn(xCharge.BoundText) & _
             " Order  By Date"
End If
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .DESCA
    TargetTable.str2 = xCharge.Text
    TargetTable.str3 = .doc_no
    TargetTable.str4 = .BOX
    TargetTable.VAL1 = !Value
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = !Date
    TargetTable.str7 = "  ›’Ì·Ï ≈Ì—«œ " & xCharge.Text
    If xBox.BoundText <> "" Then TargetTable.str7 = TargetTable.str7 & "   ·Œ“‰… " & xBox.Text
    TargetTable.str8 = "„‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle
    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\RepChrg3.rpt"
Report1.DataFiles(0) = TempPath
Report1.Action = 1
End Sub
