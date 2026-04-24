VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form repitem4 
   Caption         =   "ÊÞÇÑíÑ ÇáãäÏæÈíä"
   ClientHeight    =   2085
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
   ScaleHeight     =   2085
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xNo2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1650
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   900
      Width           =   1815
   End
   Begin VB.TextBox xNo1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1650
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   525
      Width           =   1815
   End
   Begin VB.TextBox xdate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1650
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   150
      Width           =   1815
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3480
      Top             =   1440
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
      Top             =   1320
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "ÎÑæÌ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "ÚÑÖ"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Åáí ãÎÒä :"
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ãä ãÎÒä :"
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÊÇÑíÎ :"
      Height          =   195
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   195
      Width           =   525
   End
End
Attribute VB_Name = "repitem4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim itemTable As Recordset
Dim CodeTable As Recordset
Dim SubItemTable As Recordset
Dim SubStoreTable As Recordset
Dim nOption As Integer
Function MYVALID()
If Not IsDate(xDate.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xDate.Text = ""
xNo1.Text = ""
xNo2.Text = ""
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")
Set itemTable = mydb.OpenRecordset("FILE1_10")
Set SubItemTable = mydb.OpenRecordset("FILE1_11")
Set SubStoreTable = mydb.OpenRecordset("FILE1_30")
Set CodeTable = mydb.OpenRecordset("FILE1_70")
CodeTable.Index = "nFile1_70_2"
SubItemTable.Index = "nFile1_11_1"
SubStoreTable.Index = "nFile1_30_1"
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "Delete * From Temp"
nStock = 0
With itemTable
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        CodeTable.Seek "=", 1, xNo1.Text
        If myFound(CodeTable, Array(1, xNo1.Text)) Then
            nStock = 0
            Do While CodeTable.flag = 1 And CodeTable.code <= xNo2.Text
                nStock = nStock + BalNoItem(.Item, SubItemTable, CodeTable.code)
                CodeTable.MoveNext
                If CodeTable.EOF Then Exit Do
            Loop
        End If
        If nStock > 0 And .E_Date <= DateValue(xDate.Text) Then
            TempTable.AddNew
            TempTable.str1 = .Item
            TempTable.str2 = .DESCA
            TempTable.Date1 = .E_Date
            TempTable.date2 = xDate.Text
            TempTable.str3 = nStock
            TempTable.val1 = .COST
            TempTable.VAL2 = .COST * nStock / .Pack
            TempTable.str4 = xNo1.Text
            TempTable.STR5 = xNo2.Text
            TempTable.str19 = firsttitle
            TempTable.str20 = Secondtitle
            TempTable.Update
        End If
        .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\REPITEM9.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Function myDateValue(pDate)
myDateValue = "#" & DateValue(pDate) & "#"
End Function


