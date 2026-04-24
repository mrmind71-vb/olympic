VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form R_Chr12 
   Caption         =   " ﬁ«—Ì— «·„’«—Ì›"
   ClientHeight    =   1425
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
   ScaleHeight     =   1425
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
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   75
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   1470
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   510
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   195
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   510
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
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
   Begin MSDBCtls.DBCombo xMan 
      Bindings        =   "R_Chr12.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1650
      TabIndex        =   6
      Top             =   990
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘—Ìﬂ"
      Height          =   195
      Left            =   4470
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1095
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   675
   End
End
Attribute VB_Name = "R_Chr12"
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
cString = "SELECT FILE8_90!DATE as ChargDate , FILE8_90.DESCA as DescCharg, FILE8_90!VALUE AS C_Value, FILE8_90.MAN, FILE8_90.DOC_NO  ," & _
          "FILE1_70.DESCA AS DescCode " & _
          " FROM FILE8_90 LEFT JOIN FILE1_70 ON FILE8_90.MAN = FILE1_70.CODE WHERE FILE1_70.FLAG = 5 "
If IsDate(Date1.Text) Then cString = cString & " AND Date >= " & DateSql(Date1.Text)
If IsDate(date2.Text) Then cString = cString & " AND Date <= " & DateSql(date2.Text)
If xMan.BoundText <> "" Then cString = cString & " AND MAN = " & MyParn(xMan.BoundText)

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.date3 = .ChargDate
    TargetTable.str1 = .DescCharg
    TargetTable.str3 = .DescCode
    TargetTable.VAL1 = .C_Value
    TargetTable.str7 = "  ›’Ì·Ï „”ÕÊ»«  ‘—ﬂ«¡ „‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text & " / " & xMan.Text
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\r_chrg12.rpt"
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
Private Sub Form_Load()
Date1.Text = ""
date2.Text = ""
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 5 ORDER BY DESCA "
xMan.BoundColumn = "CODE"
xMan.ListField = "DESCA"
End Sub
Function MYVALID()
MYVALID = True
End Function
