VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Rep_501 
   Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
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
   ScaleHeight     =   1515
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3375
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   525
      Width           =   1155
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   975
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1305
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   75
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   1230
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   2925
      Top             =   1050
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
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   525
      Width           =   2295
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Label xDesca 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   2490
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„Ì·"
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
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   525
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      Height          =   225
      Left            =   2535
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   390
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
      Height          =   195
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "Rep_501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChqTable As Recordset
Dim ClinTable  As Recordset
Dim TempTable As Recordset
Dim BoxTable As Recordset
Dim lExit As Boolean
Function MYVALID()
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xdate1.Text = ""
xDate2.Text = ""
xCode.Text = ""
End Sub
Private Sub CmdGroup_Click()
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")
Set ClinTable = mydb.OpenRecordset("FILE3_10", dbOpenDynaset)
Set BoxTable = mydb.OpenRecordset("FILE0_50", dbOpenDynaset)
End Sub
Private Sub CmdApply_Click()
Dim cHead As String
cHead = "  "
lExit = False

tempdb.Execute "Delete * From Temp"
    
cString = "SELECT FILE5_20.DATE_1, FILE5_20.VALUE, FILE5_20.NAME1, FILE5_20.DATE_R, Sum(file5_22.VALUE) AS tpay,max (file5_22.date)  as ldate , FILE5_20.CODE , file5_20.ser_no " & _
          " FROM FILE5_20 LEFT JOIN file5_22 ON FILE5_20.SER_NO = file5_22.ser_no " & _
          " WHERE file5_20.DATE_3 IS NULL "
If IsDate(xdate1.Text) Then cString = cString & " and FILE5_20.DATE_1 >= " & DateSql(xdate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " and FILE5_20.DATE_1 <= " & DateSql(xDate2.Text)
If xCode.Text <> "" Then cString = cString & " and FILE5_20.code = " & MyParn(xCode.Text)
cString = cString & " GROUP BY FILE5_20.DATE_1, FILE5_20.VALUE, FILE5_20.NAME1, FILE5_20.DATE_R, FILE5_20.CODE , file5_20.ser_no"

Set ChqTable = mydb.CreateSnapshot(cString)
             
With ChqTable
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        TempTable.AddNew
        TempTable.str1 = .ser_no
        TempTable.str2 = .name1
        TempTable.VAL1 = .Value
        TempTable.VAL2 = .tpay
        TempTable.VAL3 = .Value - TurnValue(.tpay, Null, 0)
        TempTable.Date1 = .date_1
        TempTable.date2 = .ldate
        TempTable.str6 = "·√Ê—«ﬁ ﬁ»÷ „” Õﬁ… „‰  «—ÌŒ " & Format(xdate1.Text) & " ≈·Ï " & Format(xDate2.Text)
        TempTable.STR19 = firsttitle
        TempTable.Update
        
        .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\Rep_501.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1

End Sub
Private Sub xCode_Change()
ClinTable.FindFirst " code = " & MyParn(xCode.Text)
xDesca.Caption = IIf(ClinTable.NoMatch, "", ClinTable.DESCA)
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„  From FILE3_10   "
    Generalarray(3) = " where  DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 4000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Sub myProc()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = GrdText(Search.Grid1, 0)
End If
Unload Search
End Sub

