VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item0 
   Caption         =   " ﬁ«—Ì— «·√’‰«›"
   ClientHeight    =   2685
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
   ScaleHeight     =   2685
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3840
   End
   Begin VB.TextBox xItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   850
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
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   2175
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
      Top             =   500
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4800
      Top             =   1650
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
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1875
      Width           =   3765
      Begin VB.CommandButton CmdClear 
         Caption         =   " ÃœÌœ"
         Height          =   390
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰"
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label xDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   855
      Width           =   2445
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·’‰›"
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   830
      Width           =   525
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   490
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "R_Item0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim SourceTable As Recordset
Dim CodeTable As Recordset
Dim ItemTable As Recordset
Dim nOption As Integer
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function

MYVALID = True
End Function
Private Sub CmdClear_Click()
xdate1.Text = ""
xDate2.Text = ""
xItem.Text = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xItem.Text = ""
xdate1.Text = ""
xDate2.Text = ""
xDescA.Text = ""
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
xDescA.Text = ""
xItem.Text = ""
xdate1.Text = ""
xDate2.Text = ""
Set TempTable = tempdb.OpenRecordset("Temp")
Set CodeTable = mydb.OpenRecordset("SELECT * FROM FILE1_70 ")
Set ItemTable = mydb.OpenRecordset("SELECT * FROM FILE1_10 ")
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
cString = "SELECT * FROM FILE1_60" & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")"

If xItem.Text <> "" Then
    cString = cString & " and ITEM = " & MyParn(xItem.Text)
End If
If xDescA.Text <> "" Then
    cString = cString & " AND FILE1_60.DESCA Like '*" & xDescA.Text & "*' "
End If
cString = cString & " ORDER BY DATE  "

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        TempTable.AddNew
        TempTable.Date1 = .Date
        TempTable.str1 = .DOC_NO
        TempTable.str2 = .DESCA
        TempTable.str3 = Say2Code(CodeTable, 1, .store1)
        TempTable.str4 = Say2Code(CodeTable, 1, .Store2)
        TempTable.val1 = .Quant
        
        TempTable.STR7 = " „ «»⁄…  ÕÊÌ·«  «·’‰› " & xDesc.Caption
        TempTable.str8 " „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle
        TempTable.Update
        .MoveNext
    Loop
End If
End With
REPORT1.ReportFileName = PublicPath & "\Reports\R_ITEM0.rpt"
REPORT1.DataFiles(0) = App.Path & "\Temp.mdb"
REPORT1.Action = 1
End Sub
Sub ItemsLookup()
ActiveControl.Text = ""
Dim Generalarray(4)
Dim GrdArray(3)
    
Set Generalarray(1) = Me
Generalarray(2) = "Select Item as «·’‰›,DescA,pack as [«”„ «·’‰›] From file1_10 as [«·»⁄Ê…] "
Generalarray(3) = " Where DescA Like('*cFilter*')"
Generalarray(4) = "Order by Item"
    
GrdArray(1) = 1000
GrdArray(2) = 3500
GrdArray(3) = 1500

Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Private Sub xItem_Change()
ItemTable.FindFirst " ITEM = " & MyParn(xItem.Text)
If ItemTable.NoMatch Then
    xDesc.Caption = ItemTable.DESCA
End If
End Sub
Private Sub xItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookup
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
xDesc.Caption = GrdText(Search.Grid1, 1)
Unload Search
End Sub
