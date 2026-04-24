VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SendDataFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁ· «·»Ì«‰« "
   ClientHeight    =   1560
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   6915
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   4065
      Begin VB.TextBox xDrive 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   810
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton cmdGetPhoto 
         Caption         =   "”Õ» ﬂ· «·’Ê—"
         Height          =   450
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhotoNew 
         Caption         =   "”Õ» «·’Ê—… «·ÕœÌÀ…"
         Height          =   450
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetData2 
         Caption         =   "‰ﬁ· «·»Ì«‰« "
         Height          =   450
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Drive "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CheckBox chkNoPhoto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "»œÊ‰ «·’Ê—"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   855
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2490
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "send_data2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   2310
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "SendDataFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim WithEvents myZip As ChilkatZip
Attribute myZip.VB_VarHelpID = -1
Dim sFolder As String
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGetData2_Click()
Dim sError As String
If getData(sError) Then
    MsgBox "‰„ ‰ﬁ· «·„·›«  »‰Ã«Õ"
Else
    MsgBox sError
End If
End Sub
Private Function getData(pError As String) As Boolean
Dim sSource As String
Dim sTarget As String
Dim sCaption As String

sSource = App.Path
sTarget = xDrive.text & ":\gate_Olympic"

On Error GoTo myerror
If Not MYVALID(sSource, sTarget, pError) Then Exit Function

Me.MousePointer = vbHourglass

Dim fs As New FileSystemObject
fs.CopyFile sSource & "\mdb\data_empty.mdb", sSource & "\mdb\data_trans.mdb"

Dim conMdb As New ADODB.Connection
openConMdb conMdb, sSource & "\mdb\data_trans.mdb"

SendDataMember conMdb
Inform " „ ‰ﬁ· »Ì«‰«  «·«⁄÷«¡ «·⁄«„·Ì‰ »‰Ã«Õ"

SendDataMember_I conMdb
Inform " „ ‰ﬁ· »Ì«‰«  «·«⁄÷«¡ «·„ﬁ”ÿÌ‰ »‰Ã«Õ"

closeCon conMdb


Me.Caption = sCaption & " " & "‰ﬁ· »Ì«‰«  «·«⁄÷«¡"
fs.CopyFile sSource & "\mdb\data_trans.mdb", sTarget & "\data_Trans.mdb"
Inform " „ ‰ﬁ· „·› «·»Ì«‰«  »‰Ã«Õ"


Set myZip = New ChilkatZip
myZip.SetCompressionLevel 0

Me.Caption = sCaption & " " & "÷€ÿ »Ì«‰«  «·«⁄÷«¡ «·⁄«„·Ì‰"
If ZipFolder(myZip, sSource & "\photo1", sSource & "\gate\photo1.zip", True, pError, Me) Then
    Inform " „ ÷€ÿ ’Ê— «·«⁄÷«¡ «·⁄«„·Ì‰ »‰Ã«Õ"
Else
    GoTo myExit
End If

Me.Caption = sCaption & " " & "‰”Œ »Ì«‰«  «·«⁄÷«¡ «·⁄«„·Ì‰"
fs.CopyFile sSource & "\gate\photo1.zip", sTarget & "\photo1.zip"
Inform " „ ‰ﬁ· ’Ê— «·«⁄÷«¡ «·⁄«„·Ì‰ »‰Ã«Õ"

Me.Caption = sCaption & " " & "÷€ÿ »Ì«‰«  «·«⁄÷«¡ «·„ﬁ”ÿÌ‰"
prog1.Visible = True
If ZipFolder(myZip, sSource & "\photo_I", sSource & "\gate\photo_i.zip", True, pError, Me) Then
    Inform " „ ÷€ÿ ’Ê— «·«⁄÷«¡ «· ﬁÌ”ÿ »‰Ã«Õ"
Else
    GoTo myExit
End If

Me.Caption = sCaption & " " & "‰”Œ »Ì«‰«  «·«⁄÷«¡ «·„ﬁ”ÿÌ‰"
fs.CopyFile sSource & "\gate\photo_i.zip", sTarget & "\photo_i.zip"
Inform " „ ‰ﬁ· ’Ê— «·«⁄÷«¡ «·„ﬁ”ÿÌ‰ »‰Ã«Õ"

prog1.Visible = False

getData = True

myExit:
Me.Caption = sCaption
Me.MousePointer = vbNormal
prog1.Visible = False
Exit Function
myerror:
Err = Err.Description
Err.Clear
GoTo myExit
End Function
Private Sub Form_Load()
Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")


xDrive.text = RetSetting(xDrive.Name, TempSave(Me))
openCon con
End Sub
Private Sub Form_Unload(Cancel As Integer)
addSetting xDrive.Name, xDrive.text, TempSave(Me)
Set SendDataFrm = Nothing
Unload Me
End Sub

Private Sub xDrive_Change()
xDrive.text = UCase(xDrive.text)
End Sub
Private Function SendDataMember(conMdb As ADODB.Connection) As Boolean
Dim cString As String
Dim loctable As ADODB.Recordset
Dim i As Long, nRecordcount As Long

Set loctable = myCmd("select * from file1_10 order by code", con)
nRecordcount = loctable.RecordCount

prog1.Visible = True
prog1.Value = 0
sCaption = Me.Caption

Do Until loctable.EOF
    i = i + 1
    Me.Caption = "Record " & i & " from " & nRecordcount & " - " & "«⁄÷«¡ ⁄«„·Ì‰"
    
    prog1.Value = mRound((i / nRecordcount) * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!code)
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    
    aInsert = AddFlag(aInsert, "TITLE", addstring(loctable!Title))
    
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷Ê ⁄«„·"))
    aInsert = AddFlag(aInsert, "[GENDER]", addvalue(loctable!GENDER))
    
    aPaid = Member_Paid(loctable!code, , con)
    
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(retFlag(aPaid, "DATE2"))))
    aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(retFlag(aPaid, "YEAR_CODE")))
    aInsert = AddFlag(aInsert, "[DOC_LAST]", addvalue(retFlag(aPaid, "DOC_NO")))
    aInsert = AddFlag(aInsert, "[DATE_LAST]", addstring(myFormat(retFlag(aPaid, "DATE"))))
    aInsert = AddFlag(aInsert, "[DIED]", IIf(loctable!died, "TRUE", "FALSE"))
    
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring("«·⁄÷Ê «·«”«”Ì"))
    conMdb.Execute addInsert(aInsert, "FILE1_10")
    loctable.MoveNext
Loop
prog1.Visible = False
prog1.Value = 0

Set loctable = Nothing
Set loctable = New ADODB.Recordset
cString = "select FILE1_10.DATE_BEGIN,FILE1_11.RELATION,FILE1_11.DESCA,FILE1_11.TITLE," & _
               "FILE1_11.MEMBER,FILE1_11.HANDI,FILE1_11.CODE,RELATION_CODES.DESCA AS RELATION_DESCA," & _
               "FILE1_11.DATE_BIRTH,FILE1_11.GENDER, FILE1_11.ID from FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE LEFT JOIN RELATION_CODES ON FILE1_11.RELATION = RELATION_CODES.CODE"

Set loctable = myCmd(cString, con)
               
nRecordcount = loctable.RecordCount
prog1.Visible = True
prog1.Value = 0
i = 0
Do Until loctable.EOF
    i = i + 1
    Me.Caption = "Record " & i & " from " & nRecordcount & " - " & " Ê«»⁄ «⁄÷«¡ ⁄«„·Ì‰"
    prog1.Value = mRound(i / nRecordcount * 100, 2)

    aInsert = AddFlag(Empty, "code", loctable!code)
    aInsert = AddFlag(aInsert, "MEMBER", loctable!MEMBER)
    aInsert = AddFlag(aInsert, "TITLE", addstring(loctable!Title))
    
    aInsert = AddFlag(aInsert, "DATE_BIRTH", addstring(myFormat(loctable!DATE_BIRTH)))
    
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷ÊÌ… ⁄«„·…"))
    aInsert = AddFlag(aInsert, "[GENDER]", addvalue(loctable!GENDER))
    aInsert = AddFlag(aInsert, "[HANDI]", IIf(loctable!HANDI, "TRUE", "FALSE"))

    aPaid = Member_Paid(loctable!MEMBER, , con)
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(retFlag(aPaid, "DATE2"))))
    
    If loctable!RELATION = 1 Then
        sDesca = "⁄÷Ê ⁄«„·"
    ElseIf loctable!RELATION = 2 Then
        'sDesca = ageSonString(myFormat(loctable!DATE_BIRTH), myFormat(IIf(bOverEnd, sDate_Season2, sDate_Season)), con)
        sDesca = "«»‰«¡"
    Else
        sDesca = " «»⁄Ì‰"
    End If
    aInsert = AddFlag(aInsert, "[RELATION]", addvalue(loctable!RELATION))
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring(sDesca & ""))
    aInsert = AddFlag(aInsert, "[ID]", loctable!ID)

    conMdb.Execute addInsert(aInsert, "FILE1_11")
    loctable.MoveNext
Loop
prog1.Visible = False
prog1.Value = 0
End Function
Private Function SendDataMember_I(conMdb As ADODB.Connection) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select   * from file2_10 where (status = 1 or status = 2) order by code", con)
nRecordcount = loctable.RecordCount
prog1.Visible = True
prog1.Value = 0
i = 0
Do Until loctable.EOF
    i = i + 1
    Me.Caption = "Record " & i & " from " & nRecordcount & " - " & "«⁄÷«¡  ﬁ”Ìÿ"
    prog1.Value = mRound((i / nRecordcount) * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!code)
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷Ê œ⁄Ê…"))
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(loctable!DATE_END)))
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring("«·⁄÷Ê «·«”«”Ì"))
    conMdb.Execute addInsert(aInsert, "FILE2_10")
    loctable.MoveNext
Loop

prog1.Visible = False
prog1.Value = 0

Set loctable = Nothing
Set loctable = New ADODB.Recordset
Set loctable = myCmd("select FILE2_10.DATE_END,FILE2_11.RELATION,FILE2_11.DESCA,FILE2_11.MEMBER,FILE2_11.CODE,RELATION_CODES.DESCA AS RELATION_DESCA from file2_11 INNER JOIN FILE2_10 ON FILE2_11.MEMBER = FILE2_10.CODE LEFT JOIN RELATION_CODES ON FILE2_11.RELATION = RELATION_CODES.CODE where (file2_10.status = 1 or file2_10.status = 2)", con)

nRecordcount = loctable.RecordCount
prog1.Visible = True
prog1.Value = 0
i = 0
Do Until loctable.EOF
    i = i + 1
    Me.Caption = "Record " & i & " from " & nRecordcount & " - " & " Ê«»⁄ «⁄÷«¡  ﬁ”Ìÿ"
    
    prog1.Value = mRound(i / nRecordcount * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!code)
    aInsert = AddFlag(aInsert, "MEMBER", loctable!MEMBER)
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷Ê œ⁄Ê…"))
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(loctable!DATE_END)))
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring(loctable!RELATION_DESCA & ""))
    conMdb.Execute addInsert(aInsert, "FILE2_11")
    loctable.MoveNext
Loop
prog1.Visible = False
prog1.Value = 0
End Function
Private Function SendDataMember_H(conMdb As ADODB.Connection) As Boolean
Dim loctable As New ADODB.Recordset
loctable.Open "select  * from file3_10 WHERE (NOT [NO] IS NULL)  order by code", con, adOpenStatic, adLockReadOnly, adCmdText
nRecordcount = loctable.RecordCount
prog1.Visible = True
prog1.Value = 0
i = 0
Do Until loctable.EOF
    i = i + 1
    Me.Caption = "Record " & i & " from " & nRecordcount & " - " & "«⁄÷«¡ ‘—›ÌÌ‰"
    prog1.Value = mRound((i / nRecordcount) * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!code)
    aInsert = AddFlag(aInsert, "[NO]", mRound(loctable!NO))
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷ÊÌ… ‘—›Ì…"))
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(loctable!DATE_END)))
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring("«·⁄÷Ê «·«”«”Ì"))
    conMdb.Execute addInsert(aInsert, "FILE3_10")
    loctable.MoveNext
Loop

prog1.Visible = False
prog1.Value = 0
End Function
Private Function MYVALID(pSource As String, pTarget As String, pError As String) As Boolean
If Trim(xDrive.text) = "" Then
    pError = "«·ﬁ—’ €Ì— „”Ã·"
    Exit Function
End If

If Not MyCreateFolder(pTarget) Then
    pError = "„‘ﬂ·… ›Ï «‰‘«¡ „Ã·œ «·›·«‘…"
    Exit Function
End If

If Not MyCreateFolder(pSource & "\gate") Then
    pError = "„‘ﬂ·… ›Ï «‰‘«¡ „Ã·œ «·„’œ—"
    Exit Function
End If

MYVALID = True
End Function
Private Sub myZip_PercentDone(ByVal pctDone As Long, abort As Long)
prog1.Value = pctDone
End Sub

