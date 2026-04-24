VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ClinSal2 
   Caption         =   "ÊÞÇÑíÑ ÇáÚãáÇÁ"
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
   Begin VB.TextBox xDate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1230
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   360
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ãä ÊÇÑíÎ"
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
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "ClinSal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim nOption As Integer
Function MYVALID()
'If Not IsDate(xdate1.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdundo_Click()
xDate1.Text = ""
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")
End Sub
Private Sub CmdApply_Click()
Dim nBaL As Double
If Not MYVALID Then Exit Sub
tempdb.Execute "Delete * From Temp"
cString = " SELECT FILE3_10.CODE, FILE3_10.DESCA, Sum(FILE3_11.PAY) AS SumPAY, Sum(FILE3_11.SAL) AS " & _
          " SumSAL, Last(FILE8_10.DATE) AS LastDATE " & _
          " FROM (FILE3_10 LEFT JOIN FILE3_11 ON FILE3_10.CODE = FILE3_11.CODE) LEFT JOIN FILE8_10 ON  " & _
          " FILE3_10.CODE = FILE8_10.CODE " & _
          " GROUP BY FILE3_10.CODE, FILE3_10.DESCA ORDER BY Last(FILE8_10.DATE) "
Set ClientTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With ClientTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    nBaL = TurnValue(.sumsal, Null, 0) - TurnValue(.sumpay, Null, 0)
    If .code = "04" Then
       ss = nBaL
    End If
    If (DateValue(xDate1.Text) > DateValue(TurnValue(.lastdate, Null, "1-1-1900")) Or IsNull(.lastdate)) And nBaL <> 0 Then
        TempTable.AddNew
        TempTable.str1 = .code
        TempTable.str2 = .DescA
        TempTable.VAL1 = nBaL
'       TempTable.str4 = .Group
'       TempTable.STR5 = SayCode(CodeTable, 3, .Group)
        TempTable.Date1 = .lastdate
        TempTable.date2 = xDate1.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle
        TempTable.Update
    End If
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\ClSal2.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Function myDateValue(pDate)
myDateValue = "#" & DateValue(pDate) & "#"
End Function
