VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form REP1_22 
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
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
   ScaleHeight     =   1470
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1200
      Width           =   1575
   End
   Begin MSDBCtls.DBCombo xStore 
      Bindings        =   "REP1_22.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   3030
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Height          =   585
      Left            =   120
      Picture         =   "REP1_22.frx":0014
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "⁄—÷ «·—”„"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -1260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   810
      Width           =   1140
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "„Œ“‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "REP1_22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "SELECT Sum([SUMIN]*[COST]) AS SumCostIn, Sum([SUMOUT]*[COST]) AS SumCostOut, FILE1_50.DESCA AS  " & _
          "GrDesc , FILE1_70.CODE AS MGr , BalItem.GrItem AS Gr , FILE1_70.DESCA AS MGrDesc  " & _
          " FROM (FILE1_50 RIGHT JOIN BalItem ON FILE1_50.CODE = BalItem.GrItem) LEFT JOIN FILE1_70 ON  " & _
          " FILE1_50.M_GROUP = FILE1_70.CODE " & _
          "WHERE FILE1_70.FLAG = 2 "
If xStore.BoundText = "" Then
    cString = cString & "GROUP BY FILE1_50.DESCA, FILE1_70.CODE, BalItem.GrItem, FILE1_70.DESCA "
Else
    cString = cString & " AND STORE = " & MyParn(xStore.BoundText) & " GROUP BY FILE1_50.DESCA, FILE1_70.CODE, BalItem.GrItem, FILE1_70.DESCA "
End If
Set SourceTable = mydb.OpenRecordset(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.GR
    TargetTable.str2 = SourceTable.GRDESC
    TargetTable.str3 = SourceTable.MGR
    TargetTable.str4 = SourceTable.MGrdesc
    TargetTable.str7 = " ≈Ã„«·Ï  ﬁÌ„ —’Ìœ "
    If xStore.BoundText <> "" Then
        TargetTable.str7 = " ≈Ã„«·Ï  ﬁÌ„ —’Ìœ " & " ·„Œ“‰ " & xStore.Text
    End If
    If bopt2 Then
    TargetTable.VAL1 = TurnValue(SourceTable.SUMCOSTIN, Null, 0) - TurnValue(SourceTable.SUMCOSTOUT, Null, 0)
    End If
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\R_ITM22.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command1_Click()
    Chart22.Show 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
    xStore.BoundText = ""
    Data1.DatabaseName = MdbPath
    Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
    xStore.ListField = "Desca"
    xStore.BoundColumn = "code"
End Sub
