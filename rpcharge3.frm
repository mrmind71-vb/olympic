VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpCharge3 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
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
   ScaleHeight     =   2355
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   5775
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
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
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   315
         TabIndex        =   9
         Top             =   1305
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„’—Ê› —∆Ì”Ì :"
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
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1350
         Width           =   1335
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
      TabIndex        =   3
      Top             =   1845
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1845
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
      Top             =   2160
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
Attribute VB_Name = "rpCharge3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
If publicFlag = 3 Then
    doprint1
Else
    doprint2
End If
End Sub
Private Sub doprint1()
ReDim aHeader(2)
If Not MYVALID Then Exit Sub
Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT FILE8_50H.DATE, FILE8_50.DESCA as DescCharg, FILE8_50.VALUE, FILE8_50.CHARGE, FILE8_50H.DOC_NO ,FILE8_50.BOX," & _
          "FILE8_51.DESCA AS DescCode,FILE0_50.DESCA AS BOXDESCA " & _
          " FROM ((FILE8_50 INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO) LEFT JOIN FILE8_51 ON FILE8_50.CHARGE = FILE8_51.CODE) LEFT JOIN FILE0_50 ON FILE8_50.BOX = FILE0_50.CODE "

If xGroup.BoundText <> "" Then
    cString = cString & turnFound(cString) & " FILE8_51.[GROUP] = " & MyParn(xGroup.BoundText)
    aHeader(2) = "[" & xGroup.Text & "]"
End If
If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "FILE8_50H.Date >= " & DateSq(xdate1.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If
If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "FILE8_50H.Date <= " & DateSq(xDate2.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If xBox.Text <> "" Then
    cString = cString & turnFound(cString) & "FILE8_50.BOX = " & MyParn(xBox.BoundText)
    aHeader(1) = "[" & xBox.Text & "]"
End If

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!date3 = !Date
    temptable!str1 = !doc_no
    temptable!str2 = !BOXDESCA
    temptable!Str3 = !DescCode
    temptable!str4 = !DescCharg
      
    temptable!val1 = !Value
    temptable!str8 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.ReportFileName = App.Path & "\Reports\charge3.rpt"
    main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    main.REPORT1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub doprint2()
ReDim aHeader(1)
If Not MYVALID Then Exit Sub
Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT FILE8_60H.DATE, FILE8_60.DESCA as DescCharg, FILE8_60.VALUE, FILE8_60.CHARGE, FILE8_60H.DOC_NO ,FILE8_60.BOX ," & _
          "FILE8_61.DESCA AS DescCode,FILE0_50.DESCA AS BOXDESCA " & _
          " FROM ((FILE8_60 INNER JOIN FILE8_60H ON FILE8_60.DOC_NO = FILE8_60H.DOC_NO) LEFT JOIN FILE8_61 ON FILE8_60.CHARGE = FILE8_61.CODE) LEFT JOIN FILE0_50 ON FILE8_60.BOX = FILE0_50.CODE "

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "FILE8_60H.Date >= " & DateSq(xdate1.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If
If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "FILE8_60H.Date <= " & DateSq(xDate2.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If xBox.Text <> "" Then
    cString = cString & turnFound(cString) & "FILE8_60.BOX = " & MyParn(xBox.BoundText)
    aHeader(1) = "[" & xBox.Text & "]"
End If

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!date3 = !Date
    temptable!str1 = !doc_no
    temptable!str2 = !BOXDESCA
    temptable!Str3 = !DescCode
    temptable!str4 = !DescCharg
    temptable!val1 = !Value
    temptable!str8 = TurnValue(retHeader(aHeader, 0, 2))
    temptable.Update
    sourcetable.MoveNext
Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.ReportFileName = App.Path & "\Reports\income3.rpt"
    main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    main.REPORT1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "FILE0_50"
Set xBox.RowSource = data1
xBox.BoundColumn = "CODE"
xBox.ListField = "DESCA"

data2.ConnectionString = strCon
data2.RecordSource = "SELECT * FROM FILE8_52"
Set xGroup.RowSource = data2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
End Sub
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
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
