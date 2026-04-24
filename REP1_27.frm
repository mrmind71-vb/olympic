VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form REP1_27 
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
      Bindings        =   "REP1_27.frx":0000
      Height          =   315
      Left            =   990
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
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
      Visible         =   0   'False
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
Attribute VB_Name = "REP1_27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FlagTable As Recordset
Private Sub CmdApply_Click()
Dim SourceTable As Recordset
Dim SubSaleTable As Recordset

Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
CFIELD1 = myiif(" STORE = 'zz' " _
           , "val(Format([TOTAL ]))") & _
           " As T_DISC "

CFIELD2 = myiif(" STORE <> 'zz' " _
           , "val(Format([TOTAL ]))") & _
           " As T_ITEM "

cString = "SELECT FILE6_20.DATE, Sum(FILE6_20.TOTAL) AS TSum  , DOC_NO , FILE3_10.DESCA  , " & _
            CFIELD1 & " ," & CFIELD2 & _
          " FROM FILE6_20 LEFT JOIN FILE3_10 ON FILE6_20.CODE = FILE3_10.CODE " & _
          " where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
          " GROUP BY FILE6_20.DATE, DOC_NO , FILE3_10.DESCA "

Set SourceTable = mydb.OpenRecordset(cString)
Set SubSaleTable = mydb.OpenRecordset(" SELECT * FROM FILE6_22 ")

If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.Date1 = SourceTable.[Date]
    TargetTable.VAL1 = SourceTable.T_ITEM
    TargetTable.VAL2 = SourceTable.T_DISC
    TargetTable.str1 = SourceTable.doc_no
    TargetTable.str2 = SourceTable.DESCA
    TargetTable.str3 = NameOfDay(SourceTable.[Date]) & " «·„Ê«›ﬁ " & Format(SourceTable.[Date], "DD-MM-YYYY")
    SubSaleTable.FindFirst " DOC_NO = " & MyParn(SourceTable.doc_no)
    If Not SubSaleTable.NoMatch Then
        TargetTable.VAL3 = SubSaleTable.CASH
        TargetTable.VAL4 = SubSaleTable.VISA
        TargetTable.VAL6 = TurnValue(SourceTable.TSUM, Null, 0) - TurnValue(SubSaleTable.VISA, Null, 0) - TurnValue(SubSaleTable.CASH, Null, 0)
    Else
        TargetTable.VAL5 = SourceTable.TSUM
    End If
    TargetTable.str9 = " ≈Ã„«·Ï „»Ì⁄«  ÌÊÌ„… „‰  «—ÌŒ " & Format(Date1.Text, "DD-MM-YYYY") & " ≈«·Ï  «—ÌŒ " & Format(date2.Text, "DD-MM-YYYY")
    TargetTable.STR19 = firsttitle
'    ' TargetTable.str20 = Secondtitle
    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
If MsgBox("≈Ã„«·Ï ÌÊÌ„… ", vbOKCancel) = vbOK Then
    Report1.ReportFileName = PublicPath & "\Reports\R0_ITM27.rpt"
Else
    Report1.ReportFileName = PublicPath & "\Reports\R_ITM27.rpt"
End If
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

