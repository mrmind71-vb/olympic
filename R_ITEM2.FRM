VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item2 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
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
   ScaleHeight     =   1545
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data1"
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
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "R_Item2.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2550
      TabIndex        =   7
      Top             =   525
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1050
      Width           =   1515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   225
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "R_Item2.frx":0014
      DataSource      =   "Data3"
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Top             =   150
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
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
   Begin MSDBCtls.DBCombo xGroup2 
      Bindings        =   "R_Item2.frx":0028
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   900
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "„‰ „Ã„Ê⁄… "
      Height          =   315
      Index           =   0
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   525
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "≈·Ï „Ã„Ê⁄…"
      Height          =   255
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      Height          =   195
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   1110
   End
End
Attribute VB_Name = "R_Item2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")
cString = "SELECT FILE1_10.ITEM, First(FILE1_10.PRICE) AS ItemPRICE, " & _
          " First(FILE1_10.COST) AS ItemCOST, First(FILE1_10.DESCA) AS ItemDESCA,  " & _
          "Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT, " & _
          " First(FILE1_50.M_GROUP) AS MGrCode, First(FILE1_50.DESCA) AS GrDESC,  " & _
          " First(FILE1_10.GROUP) AS GrCode, First(FILE1_70.DESCA) AS MGrDesc" & _
          " FROM ((FILE1_10 RIGHT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON " & _
          " FILE1_50.M_GROUP = FILE1_70.CODE" & _
          " where file1_70.FLAG = 2 "

If xGroup1.BoundText <> "" Then
    cString = cString & " AND file1_50.M_GROUP = " & xGroup1.BoundText
    If xGroup2.Text <> "" Then cString = cString & " AND file1_10.GROUP >= " & xGroup2.BoundText
    If xGroup3.Text <> "" Then cString = cString & " AND file1_10.GROUP <= " & xGroup3.BoundText
End If
cString = cString & " GROUP BY File1_10.ITEM "
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

i = 1
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
            TempTable.AddNew
            TempTable.STR6 = .MGrdesc
            TempTable.STR5 = .MGrCode
            TempTable.str1 = .Item
            TempTable.str2 = .ItemDescA
            TempTable.str3 = .GrCode
            TempTable.str4 = .grdesc
            
            TempTable.val1 = TurnValue(.SumIn, Null, 0) - TurnValue(.SUMout, Null, 0)
            TempTable.VAL2 = .ItemCOST
            TempTable.VAL3 = .ItemCOST * (TurnValue(.SumIn, Null, 0) - TurnValue(.SUMout, Null, 0))
            TempTable.STR7 = " ﬁ—Ì— «—’œ… «·‘—ﬂ…"
            TempTable.str8 = firsttitle
            TempTable.str9 = Secondtitle
            TempTable.Update
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm_2.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "File1_50 "
xGroup1.BoundColumn = "Code"
xGroup1.ListField = "Desca"
xGroup2.BoundColumn = "Code"
xGroup2.ListField = "Desca"

Data3.DatabaseName = MdbPath
Data3.RecordSource = "Select * From File1_70 Where flag = 1"
End Sub
Private Function MYVALID()
MYVALID = True
End Function
Private Sub xGroup_Change()
Data1.RecordSource = "Select * From File1_50 where M_GROUP = " & MyParn(xGroup.BoundText)
Data1.Refresh
End Sub
