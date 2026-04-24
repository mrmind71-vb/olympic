VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form CLIENTREPORT4 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   1980
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
   ScaleHeight     =   1980
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   915
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2355
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   915
      Width           =   1230
   End
   Begin Crystal.CrystalReport Report1 
      Bindings        =   "ClRep4.frx":0000
      Left            =   3825
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -975
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   405
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -945
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   45
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1275
      Width           =   3390
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   195
         Width           =   1035
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   360
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   195
         Width           =   990
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   360
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   195
         Width           =   1035
      End
   End
   Begin MSDBCtls.DBCombo xClient 
      Bindings        =   "ClRep4.frx":0014
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   105
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xItem 
      Bindings        =   "ClRep4.frx":0028
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   510
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      Height          =   300
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   945
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
      Left            =   3705
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   765
   End
   Begin VB.Label iLabel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4485
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·’‰› :"
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
      Left            =   3705
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«·⁄„Ì· :"
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
      Left            =   3705
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   195
      Width           =   720
   End
End
Attribute VB_Name = "CLIENTREPORT4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cardtable As Recordset
Dim nOption As Integer
Function CountBalance(Client)
Dim ClientMove As Recordset
NRETURN = 0
Set ClientMove = mydb.CreateDynaset("select * from file3_11 where code = " & MyParn(Client))
If ClientMove.RecordCount = 0 Then
    CountBalance = 0
    Exit Function
End If
ClientMove.MoveFirst
Do
    NRETURN = NRETURN + TurnValue(ClientMove.sal, Null, 0) - TurnValue(ClientMove.PAY, Null, 0)
    ClientMove.MoveNext
Loop Until ClientMove.EOF
CountBalance = NRETURN
End Function
Function MYVALID()
If Not IsDate(xdate1.Text) And xdate1.Text <> "" Then Exit Function
If Not IsDate(xDate2.Text) And xDate2.Text <> "" Then Exit Function
If xdate1.Text > xDate2.Text Then Exit Function
MYVALID = True
End Function

Private Sub CmdApply_Click()
cSelection = ""
If Not MYVALID Then Exit Sub
Select Case publicFlag
Case 1
    Report1.ReportFileName = PublicPath & "\Reports\report7.rpt"
Case 2
    Report1.ReportFileName = PublicPath & "\Reports\report8.rpt"
End Select
Report1.DataFiles(0) = App.Path & "\temp.mdb"
If xClient.BoundText <> "" Then cSelection = MakeQry(cSelection, "{Client} = " & MyParn(xClient.BoundText))
If xItem.BoundText <> "" Then cSelection = MakeQry(cSelection, "{ItemCode} = " & MyParn(xItem.BoundText))
If xdate1.Text <> "" Then cSelection = MakeQry(cSelection, "{InvDate} >= " & myDateString(xdate1.Text))
If xDate2.Text <> "" Then cSelection = MakeQry(cSelection, "{InvDate} <= " & myDateString(xDate2.Text))
If cSelection <> "" Then Report1.SelectionFormula = cSelection
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdUndo_Click()
xGroup1.BoundText = ""
xGroup2.BoundText = ""
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select Code,DescA From File4_10 Order by DescA"
xClient.ListField = "Desca"
xClient.BoundColumn = "Code"
Data1.Refresh

Data2.DatabaseName = MdbPath
Data2.RecordSource = "Select Item,DescA from File1_10 Order by DescA"
xItem.ListField = "Desca"
xItem.BoundColumn = "Item"
Data2.Refresh
End Sub

Private Sub iLabel_Change()
If iLabel.Caption = "" Then
    xGroup2.BoundText = ""
    xGroup2.Enabled = False
Else
    Data2.RecordSource = "Select * From ClGroup where Code >= " & MyParn(iLabel.Caption)
    Data2.Refresh
    If Not xGroup2.Enabled Then xGroup2.Enabled = True
End If
End Sub

Private Sub xGroup1_Click(Area As Integer)
iLabel.Caption = xGroup1.BoundText
End Sub


