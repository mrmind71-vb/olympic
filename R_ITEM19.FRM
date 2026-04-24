VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item19 
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
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1050
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ÎÑæÌ"
      Height          =   390
      Left            =   75
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
      Left            =   300
      Top             =   225
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
Attribute VB_Name = "R_Item19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS F_DESCA,  " & _
          "Sum(FILE6_20.TOTAL) AS SumSAL, Sum(FILE6_20.QUANT) AS SumQUANT,  " & _
          "Sum([FILE6_20].[COST]*[FILE6_20].[QUANT]) AS SumCost , First(FILE1_50.DESCA)  " & _
          "AS DescGroup, FILE1_50.CODE AS GroupCode " & _
          "FROM (FILE1_10 INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) LEFT  " & _
          "JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE " & _
          "Where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
          "and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
          "GROUP BY FILE1_10.ITEM, FILE1_50.CODE "
Set SourceTable = mydb.OpenRecordset(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ ÈÇáÊÞÑíÑ"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.Item
    TargetTable.str2 = SourceTable.F_DescA
    TargetTable.str3 = SourceTable.GroupCode
    TargetTable.str4 = SourceTable.DescGroup
    TargetTable.str6 = " ãä ÊÇÑíÎ " & Date1.Text & " Åáì ÊÇÑíÎ " & date2.Text
    
    TargetTable.VAL1 = SourceTable.SUMQUANT
    If SourceTable.SUMSAL <> 0 Then TargetTable.VAL2 = SourceTable.SUMSAL / SourceTable.SUMQUANT
    TargetTable.VAL3 = SourceTable.SUMSAL
    If SourceTable.SUMQUANT <> 0 Then TargetTable.VAL4 = SourceTable.SUMCOST / SourceTable.SUMQUANT
    TargetTable.VAL5 = SourceTable.SUMSAL - SourceTable.SUMCOST
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
Report1.ReportFileName = PublicPath & "\Reports\R_ITM19.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
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

End Sub

