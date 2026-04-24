VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_BANK1 
   Caption         =   "ÊÞÇÑíÑ ÇáÈäæß"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   945
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Report1 
      Left            =   4035
      Top             =   120
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
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "ÎÑæÌ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "ÚÑÖ"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1005
      End
   End
End
Attribute VB_Name = "R_BANK1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim BankBalTable As Recordset
Dim CodeTable As Recordset
Dim nOption As Integer
Dim cBalStr As String
Function MYVALID()
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")
cBalStr = " SELECT Sum(file5_11.[IN]) AS SumIN, Sum(file5_11.OUT) AS SumOUT, FILE5_10.ID, FILE5_10.DESCA " & _
          " FROM file5_11 RIGHT JOIN FILE5_10 ON file5_11.BANK = FILE5_10.ID  GROUP BY FILE5_10.ID, FILE5_10.DESCA "
Set BankBalTable = mydb.OpenRecordset(cBalStr)
Set CodeTable = mydb.OpenRecordset("FILE1_70")
End Sub
Private Sub CmdApply_Click()
tempdb.Execute "Delete * From Temp"
With BankBalTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .Id
    TempTable.STR2 = .DESCA
    TempTable.val1 = TurnValue(.SUMIN, Null, 0) - TurnValue(.SUMOUT, Null, 0)
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle

    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\BANKREP1.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
