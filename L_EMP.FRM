VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form L_Emp 
   Caption         =   "ﬂ‘› «·„— »"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   1290
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   780
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   855
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   487
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1485
      Top             =   630
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
      Height          =   690
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   3525
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷ "
         Height          =   390
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1620
      End
   End
   Begin VB.Label xDesc 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·⁄«„·"
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
      Left            =   4260
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "≈·Ï  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   870
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   510
      Width           =   765
   End
End
Attribute VB_Name = "L_Emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim DataTable As Recordset
Dim nOption As Integer
Dim cT_Time As Double
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdClear_Click()
xdate1.Text = ""
xDate2.Text = ""
xCode.Text = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xdate1.Text = ""
xDate2.Text = ""
xCode.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If TypeOf ActiveControl Is DBCombo Then ActiveControl.BoundText = ""
End If
End Sub
Private Sub Form_Load()
xdate1.Text = ""
xDate2.Text = ""
xCode.Text = ""
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("TEMP")
With TempTable
cString = " SELECT FILE9_10.CODE, FILE9_10.NAMEA, FILE9_20.TIME, FILE9_20.DATE,   " & _
          " FILE9_20.DEBIT, FILE9_20.Friday, FILE9_20.DISC  " & _
          " FROM FILE9_10 RIGHT JOIN FILE9_20 ON FILE9_10.CODE = FILE9_20.CODE " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and FILE9_10.CODE = " & MyParn(xCode.Text) & _
          " ORDER BY FILE9_20.DATE "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .str1 = NameOfDay(DataTable.Date)
        .str2 = DataTable.NAMEA
        .val1 = DataTable.Time
        .val2 = DataTable.Friday
        .VAL3 = DataTable.DEBIT
        .VAL4 = DataTable.DISC
        .Date1 = DataTable.Date
        .STR7 = "»Ì«‰ «·Õ÷Ê— ··› —… „‰ ÌÊ„" & xdate1.Text & " ·ÌÊ„ " & xDate2.Text
        .str9 = firsttitle
        .str10 = Secondtitle
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

If TempTable.RecordCount > 0 Then
    Report1.ReportFileName = PublicPath & "\Reports\L_EMP.rpt"
    Report1.DataFiles(0) = App.Path & "\Temp.mdb"
    Report1.Action = 1
Else
    MsgBox "·«  ÊÃœ √ÃÊ— „” Õﬁ… ·Â–… «·› —… "
End If
End With
End Sub

Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Generalarray(3)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select CODE AS [«·—ﬁ„ «·”„·”·],NAMEA as [«”„ «·„ÊŸ›]  From File9_10   "
    Generalarray(3) = "Where NAMEA Like '*cFilter*'"
        
    GrdArray(1) = 2000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
xDesc.Caption = GrdText(Search.Grid1, 1)
Unload Search
End Sub

