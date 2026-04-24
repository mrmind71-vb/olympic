VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form REP1_26 
   Caption         =   " ﬁ«—Ì— «·«’‰«›"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
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
   ScaleHeight     =   1905
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -1260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   810
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   1590
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1290
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1290
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   435
      Width           =   1665
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   1665
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   300
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
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "REP1_26.frx":0000
      Height          =   315
      Left            =   990
      TabIndex        =   6
      Top             =   840
      Width           =   3315
      _ExtentX        =   5847
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
      Caption         =   "«·„Œ“‰"
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
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ï  «—ÌŒ :"
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
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   555
      Width           =   825
   End
   Begin VB.Label Label1 
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
      Height          =   195
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   765
   End
End
Attribute VB_Name = "REP1_26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FlagTable As Recordset
Private Sub CmdApply_Click()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
If xStore1.BoundText = "" Then
    cString = "SELECT FILE6_20.DOC_NO,FILE6_20.ITEM, FILE6_20.DATE, FILE6_20.CODE, FILE6_20.cost,Store," & _
              "FILE6_20.QUANT, FILE6_20.TOTAL, FILE6_20.PRICE, FILE3_10.DESCA AS ClientName ,FILE1_10.DESCA " & _
              "FROM FILE1_10 RIGHT JOIN (FILE6_20 LEFT JOIN FILE3_10 ON FILE6_20.CODE = " & _
              " FILE3_10.CODE) ON FILE1_10.ITEM = FILE6_20.ITEM " & _
              " Where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
              " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
              " AND STORE  <> 'SS' " & _
              " Order by File6_20.Doc_No,Store asc"
Else
    cString = "SELECT FILE6_20.DOC_NO,FILE6_20.ITEM, FILE6_20.DATE, FILE6_20.CODE, FILE6_20.cost,Store," & _
              "FILE6_20.QUANT, FILE6_20.TOTAL, FILE6_20.PRICE, FILE3_10.DESCA AS ClientName ,FILE1_10.DESCA " & _
              "FROM FILE1_10 RIGHT JOIN (FILE6_20 LEFT JOIN FILE3_10 ON FILE6_20.CODE = " & _
              " FILE3_10.CODE) ON FILE1_10.ITEM = FILE6_20.ITEM " & _
              " Where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
              " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
              " AND STORE = " & MyParn(xStore1.BoundText) & _
              " Order by File6_20.Doc_No,Store asc"
End If
Set SourceTable = mydb.OpenRecordset(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    If SourceTable.COST >= SourceTable.price Then
    TargetTable.str1 = SourceTable.doc_no
    TargetTable.str2 = SourceTable.CODE
    TargetTable.str3 = SourceTable.ClientName
    TargetTable.str4 = IIf(SourceTable!Store = "SS", RetFind(FlagTable, "Code", "DescA", SourceTable.Item), SourceTable!DESCA)
    TargetTable.STR5 = SourceTable.[Date] & SourceTable.doc_no
    TargetTable.str6 = " ›« Ê—… " & SourceTable.doc_no & " » «—ÌŒ " & Format(SourceTable.[Date], "DD-MM-YYYY") & " ··⁄„Ì· " & SourceTable.ClientName
    TargetTable.str7 = SourceTable.Item
    TargetTable.str8 = "„‰ " & Date1.Text & " ≈·Ï " & date2.Text
    If xStore1.Text <> "" Then TargetTable.str8 = "„‰ " & Date1.Text & " ≈·Ï " & date2.Text & " ·„Œ“‰ " & xStore1.Text
    TargetTable.VAL1 = SourceTable.Quant
    TargetTable.VAL2 = SourceTable.price
    TargetTable.VAL3 = SourceTable.total
    If bopt2 Then
    If SourceTable!Store <> "ss" Then
        TargetTable.VAL4 = SourceTable.COST
        TargetTable.VAL5 = SourceTable.total - (SourceTable.COST * SourceTable.Quant)
    Else
        TargetTable.VAL4 = 0
        TargetTable.VAL5 = 0
    End If
    End If
    TargetTable.Date1 = SourceTable.[Date]
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle
    TargetTable.Update
    End If
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\R_ITM26.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
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
    Set FlagTable = mydb.OpenRecordset("File8_70", dbOpenDynaset)
    Date1.Text = ""
    date2.Text = ""
    Data1.DatabaseName = MdbPath
    Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
    xStore1.ListField = "Desca"
    xStore1.BoundColumn = "code"

End Sub

