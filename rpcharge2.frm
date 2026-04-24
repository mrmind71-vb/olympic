VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpCharge2 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
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
   ScaleHeight     =   2670
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   -45
      Width           =   5145
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1665
         Width           =   3345
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   315
         TabIndex        =   2
         Top             =   945
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xCharge 
         Height          =   315
         Left            =   315
         TabIndex        =   3
         Top             =   1305
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»Ì«‰ «·„’—Ê› :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1755
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "«·„’—Ê› :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label Label1 
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Œ“‰… :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1035
         Width           =   465
      End
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2070
      Width           =   1320
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1320
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3555
      Top             =   2295
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3330
      Top             =   1800
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "rpCharge2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
If publicFlag = 2 Then
    doprint1
Else
    doprint2
End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub doprint1()
Dim sourcetable As New ADODB.Recordset
Dim temptable As New ADODB.Recordset
Dim aHeader(2)
If Not MYVALID Then Exit Sub

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select FILE8_50H.Date,FILE8_50.DescA,FILE0_50.DESCA AS FILE0_50DESCA,FILE8_51.DESCA AS CHARGEDESCA,Value,FILE8_50H.DOC_NO,BOX" & _
          " From ((FILE8_50 INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO)Left join file8_51 on file8_50.charge = file8_51.code) LEFT JOIN FILE0_50 ON FILE8_50.BOX = FILE0_50.CODE "

If xBox.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  FILE8_50.BOX = " & MyParn(xBox.BoundText)
     aHeader(1) = "[" & xBox.Text & "]"
End If

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & " FILE8_50H.date >= " & DateSq(xdate1.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & " FILE8_50H.date <= " & DateSq(xDate2.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If
    

If XCHARGE.BoundText <> "" Then
    cString = cString & turnFound(cString) & " FILE8_50.charge like " & MyParn(XCHARGE.BoundText)
End If

If Trim(xDesca.Text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xDesca.Text, "file8_50.desca")
    aHeader(2) = "[" & "«·»Ì«‰ : " & xDesca.Text & "]"
End If


sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!str2 = !FILE0_50DESCA
    'temptable!Str3 = TurnValue(ArbString(!desca & turn(!chargeDesca & "", turn(!desca, "-") & !chargeDesca) & ""))
    temptable!Str3 = !Desca
    If Not IsNull(!chargeDesca) Then
        temptable!Str3 = TurnValue(ArbString(temptable!Str3 & turn(temptable!Str3 & "", "-") & !chargeDesca & ""))
    End If
    temptable!val1 = !Value
    temptable!Date1 = !Date
    temptable!str21 = "  ›’Ì·Ï „’—Ê› " & XCHARGE.Text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 2))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\charge2.rpt"
main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.REPORT1.Action = 1
End Sub
Private Sub doprint2()
Dim sourcetable As New ADODB.Recordset
Dim temptable As New ADODB.Recordset
Dim aHeader(1)
If Not MYVALID Then Exit Sub

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select FILE8_60H.Date,FILE8_60.DescA,FILE0_50.DESCA AS FILE0_50DESCA,FILE8_60.Value,FILE8_60H.DOC_NO,FILE8_60.BOX" & _
          " From (FILE8_60 INNER JOIN FILE8_60H ON FILE8_60.DOC_NO = FILE8_60H.DOC_NO) INNER JOIN FILE0_50 ON FILE8_60.BOX = FILE0_50.CODE "

If xBox.BoundText <> "" Then
    cString = cString & turnFound(cString) & "FILE8_60.BOX = " & MyParn(xBox.BoundText)
     aHeader(1) = "[" & xBox.Text & "]"
End If

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "FILE8_60H.date >= " & DateSq(xdate1.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "FILE8_60H.date <= " & DateSq(xDate2.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If
    

If XCHARGE.BoundText <> "" Then
    cString = cString & turnFound(cString) & "FILE8_60.charge = " & MyParn(XCHARGE.BoundText)
End If

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!str2 = !FILE0_50DESCA
    temptable!Str3 = TurnValue(Trim(!Desca))
    temptable!val1 = !Value
    temptable!Date1 = !Date
    temptable!str21 = "  ›’Ì·Ï «Ì—«œ " & XCHARGE.Text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 2))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\income2.rpt"
main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.REPORT1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
Set sourcetable = New ADODB.Recordset
Set temptable = New ADODB.Recordset
data1.ConnectionString = strCon

data1.RecordSource = "Select * FROM " & IIf(publicFlag = 2, "file8_51", "file8_61") & " ORDER BY DESCA"
Set XCHARGE.RowSource = data1
XCHARGE.BoundColumn = "Code"
XCHARGE.ListField = "DescA"

data2.ConnectionString = strCon
data2.RecordSource = "SELECT * FROM FILE0_50 ORDER BY DESCA"
Set xBox.RowSource = data2
xBox.BoundColumn = "CODE"
xBox.ListField = "DESCA"
If publicFlag <> 2 Then Label4 = "«·«Ì—«œ :"
End Sub
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
'If xCharge.BoundText = "" Then
'    MsgBox IIf(publicFlag = 2, "»Ì«‰ «·„’—Ê› „ÿ·Ê»", "»Ì«‰ «·«Ì—«œ „ÿ·Ê»")
'    Exit Function
'End If
If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xDate2.Text)) And Trim(xDate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xbox_Validate(Cancel As Boolean)
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub

Private Sub xCharge_Validate(Cancel As Boolean)
If Not XCHARGE.MatchedWithList Then XCHARGE.BoundText = ""
End Sub
