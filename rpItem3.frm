VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpItem3 
   Caption         =   " ﬁ—Ì— «· ÕÊÌ·« "
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
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
   ScaleHeight     =   2370
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   5010
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1290
      End
      Begin MSDataListLib.DataCombo xStore1 
         Height          =   315
         Left            =   75
         TabIndex        =   2
         Top             =   945
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xStore2 
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   1305
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label4 
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   915
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1305
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Width           =   825
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4275
      Top             =   1800
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
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1710
      Width           =   3615
      Begin VB.CommandButton CmdClear 
         Caption         =   " ÃœÌœ"
         Height          =   390
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   135
         Width           =   1140
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -1845
      Top             =   1170
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   5100
      Top             =   2550
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "rpItem3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim nOption As Integer
Function MYVALID()
If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then Exit Function
If (Not IsDate(xDate2.Text)) And Trim(xDate2.Text) <> "" Then Exit Function
If (Trim(xStore1.BoundText) = Trim(xStore2.BoundText) And xStore2.BoundText <> "") Then
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
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xStore1.BoundText = ""
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If TypeOf ActiveControl Is DataCombo Then ActiveControl.BoundText = ""
End If
End Sub
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "SELECT CODE , DESCA FROM FILE0_40"
Set xStore1.RowSource = data1
xStore1.ListField = "Desca"
xStore1.BoundColumn = "Code"
Set xStore2.RowSource = data1
xStore2.ListField = "Desca"
xStore2.BoundColumn = "Code"
End Sub
Private Sub cmdApply_Click()
Dim aHeader(1)
If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_60H.DOC_NO, FILE1_60H.date, FILE1_60.quant, FILE1_60.item,FILE1_60.COST,FILE0_40.DESCA AS STORE1DESCA,FILE0_40_1.DESCA AS STORE2DESCA, " & _
          "FILE1_10.DESCA AS ItemDesc , FILE1_60H.store1, FILE1_60H.store2 " & _
          "FROM (((FILE1_60H INNER JOIN FILE1_60 ON FILE1_60H.DOC_NO = FILE1_60.DOC_NO) INNER JOIN FILE1_10 ON FILE1_60.item = FILE1_10.ITEM) INNER JOIN FILE0_40 ON FILE1_60H.STORE1 = FILE0_40.CODE) INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE1_60H.STORE2 = FILE0_40_1.CODE "

If xStore1.BoundText <> "" Then
    cString = cString & turnFound(cString) & " FILE1_60H.store1 = " & MyParn(xStore1.BoundText)
    aHeader(0) = "[" & BetweenString(" „Œ“‰ " & xStore1.Text, " „Œ“‰ " & xStore2.Text, , " ≈·Ì ") & "]"
End If

If xStore2.BoundText <> "" Then
    cString = cString & turnFound(cString) & " FILE1_60H.store2 = " & MyParn(xStore2.BoundText)
    aHeader(0) = "[" & BetweenString(" „Œ“‰ " & xStore1.Text, " „Œ“‰ " & xStore2.Text, , " ≈·Ì ") & "]"
End If

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "FILE1_60H.DATE >= " & DateSq(xdate1.Text)
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If
If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "FILE1_60H.DATE <= " & DateSq(xDate2.Text)
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If


Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    GoTo lastsub
End If
Do Until .EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!Str3 = !Item
    temptable!str4 = !itemDesc
    temptable!str5 = !STORE1DESCA
    temptable!str6 = !STORE2DESCA
    temptable!val1 = !Quant
    temptable!Date1 = ![Date]
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 2))
    temptable!str8 = " „‰ " & xdate1.Text & " ≈·Ï " & xDate2.Text
    temptable!val2 = !cost
    temptable!Val3 = !Quant * Val(!cost & "")
    'temptable!str9 = firstTitle & IIf(xBranch.BoundText <> "", "-" & " »Ì«‰«  " & xBranch.Text & " ›ﬁÿ ", "")
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\item3.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
lastsub:
    temptable.Close
    sourcetable.Close
    Set temptable = Nothing
    Set sourcetable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
