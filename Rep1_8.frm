VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Rep1_8 
   Caption         =   " ﬁ—Ì— «· ÕÊÌ·« "
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
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   -225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   600
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
      Height          =   390
      Left            =   -360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   165
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1125
      Top             =   225
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
      Height          =   690
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1875
      Width           =   3765
      Begin VB.CommandButton CmdClear 
         Caption         =   " ÃœÌœ"
         Height          =   390
         Left            =   1350
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1065
      End
   End
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "Rep1_8.frx":0000
      Height          =   315
      Left            =   825
      TabIndex        =   2
      Top             =   1050
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xStore2 
      Bindings        =   "Rep1_8.frx":0014
      Height          =   315
      Left            =   825
      TabIndex        =   3
      Top             =   1500
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï „Œ“‰ :"
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
      TabIndex        =   11
      Top             =   1575
      Width           =   885
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ „Œ“‰ :"
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
      Top             =   1125
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "Rep1_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim SourceTable As Recordset
Dim CodeTable As Recordset
Dim nOption As Integer
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
If xStore1.BoundText = xStore2.BoundText Then
    MsgBox "·« Ì„ﬂ‰ «· ÕÊÌ· „‰ „Œ“‰ «·Ï ‰›” „Œ“‰"
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdClear_Click()
xdate1.Text = ""
xDate2.Text = ""
xStore1.BoundText = ""
xStore2.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xStore1.BoundText = ""
xdate1.Text = ""
xDate2.Text = ""
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
Set TempTable = tempdb.OpenRecordset("Temp")
Set CodeTable = mydb.OpenRecordset("SELECT * FROM FILE1_70 ")
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
Data2.DatabaseName = MdbPath
Data2.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore1.ListField = "Desca"
xStore1.BoundColumn = "code"
xStore2.ListField = "Desca"
xStore2.BoundColumn = "code"
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
    cString = "SELECT FILE1_60.DOC_NO, FILE1_60.date, FILE1_60.DESCA, FILE1_60.quant, FILE1_60.item, " & _
              "FILE1_10.DESCA AS ItemDesc , FILE1_60.store1, FILE1_60.store2 " & _
              "FROM FILE1_60 LEFT JOIN FILE1_10 ON FILE1_60.item = FILE1_10.ITEM " & _
              " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
              " and DateValue(" & MyParn(xDate2.Text) & ")" & _
              IIf(xStore1.Text <> "", " and Store1 = " & MyParn(xStore1.BoundText), "") & _
              IIf(xStore2.Text <> "", " and Store2 = " & MyParn(xStore2.BoundText), "")

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        TempTable.AddNew
        TempTable.str1 = .doc_no
        TempTable.str2 = .DESCA
        TempTable.str3 = .Item
        TempTable.str4 = .ItemDesc
        TempTable.STR5 = Say2Code(CodeTable, 1, .store1)
        TempTable.str6 = Say2Code(CodeTable, 1, .Store2)
        TempTable.VAL1 = .Quant
        TempTable.Date1 = .[Date]
        TempTable.str7 = " ≈Ã„«·Ï  ÕÊÌ·«  „‰ „Œ“‰ " & xStore1.Text & " ≈·Ï „Œ“‰ " & xStore2.Text
        TempTable.str8 = " „‰ " & xdate1.Text & " ≈·Ï " & xDate2.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle

        TempTable.Update
        .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\r_item8.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
