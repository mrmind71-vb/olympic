VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item20 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "R_Item20.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   675
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton CmdApply 
      BackColor       =   &H00C0FFFF&
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   1200
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   562
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   562
      Width           =   915
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   525
      Width           =   2040
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   150
      Top             =   75
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
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "R_Item20.frx":000C
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1950
      TabIndex        =   0
      Top             =   75
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      Height          =   195
      Left            =   4875
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   150
      Width           =   1110
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰  «—ÌŒ"
      Height          =   315
      Left            =   4875
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   540
      Width           =   765
   End
End
Attribute VB_Name = "R_Item20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim InvTable As Recordset
Dim itemTable As Recordset
Dim cString As String
Dim nBal As Double
Private Sub CmdApply_Click()
If IsDate(xDate.Text) Then
'Me.MousePointer = 11
cFilter = ""
If xGroup.BoundText <> "" Then
    cString = "SELECT FILE1_10.ITEM, Last(FILE6_20.DATE) AS LastDATE, First(FILE1_10.DESCA) AS F_DESCA, " & _
              " First(FILE1_70.DESCA) AS M_GroupDesc, First(FILE1_70.CODE) AS M_GroupCode, " & _
              " First(FILE1_50.CODE) AS GroupCode, First(FILE1_50.DESCA) AS GroupDesc " & _
              " FROM ((FILE1_10 LEFT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) LEFT JOIN FILE1_50 ON  " & _
              " FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE " & _
              " WHERE FILE1_70.FLAG=2 " & _
              " AND file1_50.M_GROUP = " & MyParn(xGroup.BoundText) & _
              " GROUP BY FILE1_10.ITEM " & _
              " ORDER BY Last(FILE6_20.DATE)"
Set InvTable = mydb.CreateSnapshot(cString)

cString = " SELECT Sum(FILE1_11.OUT) AS SumOUT, Sum(FILE1_11.[IN]) AS SumIN, FILE1_10.ITEM " & _
          " FROM FILE1_11 RIGHT JOIN ((FILE1_10 LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN " & _
          " FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE) ON FILE1_11.ITEM = FILE1_10.ITEM " & _
          " WHERE FILE1_70.FLAG=2 " & _
          " AND file1_50.M_GROUP = " & MyParn(xGroup.BoundText) & _
          " GROUP BY FILE1_10.ITEM " & _
          " ORDER BY Last(FILE1_11.DATE)"
Set itemTable = mydb.CreateSnapshot(cString)


Else
    cString = "SELECT FILE1_10.ITEM, Last(FILE6_20.DATE) AS LastDATE, First(FILE1_10.DESCA) AS F_DESCA, " & _
              " First(FILE1_70.DESCA) AS M_GroupDesc, First(FILE1_70.CODE) AS M_GroupCode, " & _
              " First(FILE1_50.CODE) AS GroupCode, First(FILE1_50.DESCA) AS GroupDesc " & _
              " FROM ((FILE1_10 LEFT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) LEFT JOIN FILE1_50 ON  " & _
              " FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE " & _
              " WHERE FILE1_70.FLAG=2 " & _
              " GROUP BY FILE1_10.ITEM "
Set InvTable = mydb.CreateSnapshot(cString)

cString = " SELECT Sum(FILE1_11.OUT) AS SumOUT, Sum(FILE1_11.[IN]) AS SumIN, FILE1_10.ITEM" & _
          " FROM FILE1_11 RIGHT JOIN ((FILE1_10 LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN " & _
          " FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE) ON FILE1_11.ITEM = FILE1_10.ITEM " & _
          " WHERE FILE1_70.FLAG=2 " & _
          " GROUP BY FILE1_10.ITEM "
Set itemTable = mydb.CreateSnapshot(cString)

End If
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With InvTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    itemTable.FindFirst " ITEM = " & MyParn(.Item)
    If Not itemTable.NoMatch Then
        nBal = TurnValue(itemTable.SUMIn, Null, 0) - TurnValue(itemTable.SUMOUT, Null, 0)
    Else
        nBal = 0
    End If
    If (.LASTDATE <= DateValue(xDate.Text) Or IsNull(.LASTDATE)) And nBal <> 0 Then
    TempTable.AddNew
    TempTable.str1 = .Item
    TempTable.str2 = .F_DescA
    TempTable.Date1 = .LASTDATE
    TempTable.str3 = .GroupCode
    TempTable.str4 = .GroupDesc
    TempTable.str6 = "√’‰«› ·Ì” ·Â« „»Ì⁄«  „‰  «—ÌŒ  " & xDate.Text
    TempTable.VAL1 = nBal
    TempTable.str9 = firsttitle
    TempTable.str10 = Secondtitle
    TempTable.Update
    End If
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\r_item20.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End If
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
Set movetable = mydb.OpenRecordset("File1_11")

Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select Code,DescA From File1_70 Where Flag =  2 "
xGroup.BoundColumn = "Code"
xGroup.ListField = "Desca"
xDate.Text = ""
End Sub
Private Sub Store_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xStore1.BoundText = ""
End Sub
