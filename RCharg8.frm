VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RCharg8 
   Caption         =   " ÕÊÌ·«  „‰ Œ“‰… ≈·Ï Œ“‰…"
   ClientHeight    =   1710
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
   ScaleHeight     =   1710
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   750
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1230
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
      Left            =   75
      Top             =   300
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
   Begin MSDBCtls.DBCombo xBox1 
      Bindings        =   "RCharg8.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2475
      TabIndex        =   8
      Top             =   840
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xBox2 
      Bindings        =   "RCharg8.frx":0014
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2475
      TabIndex        =   9
      Top             =   1200
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï Œ“‰…"
      Height          =   195
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1305
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ Œ“‰… "
      Height          =   195
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   945
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
      Height          =   195
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   585
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      Height          =   195
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   675
   End
End
Attribute VB_Name = "RCharg8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChargeTable As Recordset
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset

tempdb.Execute "DELETE * FROM TEMP"

Set TargetTable = tempdb.CreateDynaset("TEMP")

cString = "SELECT FILE0_50.DESCA AS Desc1 , FILE0_50_1.DESCA As Desc2 ,  *  FROM (FILE0_60 LEFT JOIN FILE0_50 ON FILE0_60.NO1 = FILE0_50.CODE) LEFT JOIN FILE0_50 AS  FILE0_50_1 ON FILE0_60.NO2 = FILE0_50_1.CODE " & _
          " Where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and Date <= DateValue(" & MyParn(date2.Text) & ")"

If xBox1.BoundText <> "" Then
    cString = cString & " AND NO1 = " & MyParn(xBox1.BoundText)
End If

If xBox2.BoundText <> "" Then
    cString = cString & " AND NO2 = " & MyParn(xBox2.BoundText)
End If

cString = cString & " ORDER BY DATE , NO1 , NO2 "

Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
'    TargetTable.str1 = .CODEBOX
'    TargetTable.str2 = .DESCA
    TargetTable.str3 = .DESC1
    TargetTable.str4 = .DESC2
    TargetTable.VAL1 = !Value
    TargetTable.date3 = !Date
    TargetTable.str7 = "  ›’Ì·Ï  ÕÊÌ·«  "
    If xBox1.BoundText <> "" Then TargetTable.str7 = TargetTable.str7 & " „‰ Œ“‰… " & xBox1
    If xBox2.BoundText <> "" Then TargetTable.str7 = TargetTable.str7 & " ≈·Ï  Œ“‰… " & xBox2
    TargetTable.str8 = "„‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle
    
    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\RepChrg8.rpt"
Report1.DataFiles(0) = TempPath
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
    End If
End Sub
Function MYVALID()
If Not IsDate(Date1.Text) Then Exit Function
If Not IsDate(date2.Text) Then Exit Function
If DateValue(Date1.Text) > DateValue(date2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "FILE0_50"
xBox1.BoundColumn = "CODE"
xBox1.ListField = "DESCA"
xBox2.BoundColumn = "CODE"
xBox2.ListField = "DESCA"

End Sub
