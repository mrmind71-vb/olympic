VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form REP1_21 
   Caption         =   "ÊÞÇÑíÑ ÇáÇÕäÇÝ"
   ClientHeight    =   1530
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
   ScaleHeight     =   1530
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   585
      Left            =   120
      Picture         =   "REP1_21.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "ÚÑÖ ÇáÑÓã"
      Top             =   855
      Visible         =   0   'False
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
   Begin VB.CommandButton CmdApply 
      Caption         =   "ÚÑÖ"
      Height          =   390
      Left            =   3030
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1050
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ÎÑæÌ"
      Height          =   390
      Left            =   1755
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1050
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1665
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1665
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1200
      Top             =   120
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Çáì ÊÇÑíÎ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ãä ÊÇÑíÎ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "REP1_21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "SELECT FILE1_70.CODE, FILE1_70.DESCA, Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum([QUANT]*[COST]) AS SumCost " & _
          "FROM FILE1_70 LEFT JOIN FILE6_20 ON FILE1_70.CODE = FILE6_20.STORE " & _
          "WHERE FILE1_70.FLAG = 1 " & _
          "AND Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
          "and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
          "GROUP BY FILE1_70.CODE, FILE1_70.DESCA "
Set SourceTable = mydb.OpenRecordset(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ ÈÇáÊÞÑíÑ"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str2 = SourceTable.CODE
    TargetTable.str1 = SourceTable.DESCA
    TargetTable.str7 = " ÅÌãÇáì ãÈíÚÇÊ ãÎÇÒä "
    TargetTable.str8 = " ãä ÊÇÑíÎ " & Date1.Text & " Åáì ÊÇÑíÎ " & date2.Text
    TargetTable.VAL1 = SourceTable.SUMTOTAL
    If bopt2 Then
    TargetTable.VAL3 = SourceTable.SUMTOTAL - SourceTable.SUMCOST
    TargetTable.VAL4 = (SourceTable.SUMTOTAL - SourceTable.SUMCOST) * 100 / SourceTable.SUMTOTAL
    End If
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\R_ITM21.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command1_Click()
    Chart21.Show 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
    Date1.Text = ""
    date2.Text = ""
End Sub
