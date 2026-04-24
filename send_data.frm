VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SendDataFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "‰ﬁ· «·»Ì«‰« "
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6675
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkNoPhoto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "»œÊ‰ «·’Ê—"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   945
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   2490
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "send_data.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   2310
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   0
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
         TabIndex        =   4
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton cmdGetData 
         Caption         =   "‰ﬁ· «·»Ì«‰« "
         Height          =   450
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhotoNew 
         Caption         =   "”Õ» «·’Ê—… «·ÕœÌÀ…"
         Height          =   450
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhoto 
         Caption         =   "”Õ» ﬂ· «·’Ê—"
         Height          =   450
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
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
         TabIndex        =   5
         Top             =   360
         Width           =   555
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   1290
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
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
Dim sFolder As String
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGetData_Click()
Dim fs As New FileSystemObject, conMdb As New ADODB.Connection, aInsert As Variant

If Trim(xDrive.text) = "" Then
    MsgBox "«·ﬁ—’ €Ì— „”Ã·"
    Exit Sub
End If

sFolder = xDrive.text & ":\Olympic_door_sql"


If Not MyCreateFolder(sFolder) Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·»—‰«„Ã"
    Exit Sub
End If

If chkNoPhoto.Value = 0 Then
    If Not MyCreateFolder(sFolder & "\photo1") Then
        MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·’Ê—"
        Exit Sub
    End If
End If

If Not MyCreateFolder(sFolder & "\photo_i", True) Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— ’Ê— «·«⁄÷«¡ «· ﬁ”Ìÿ"
    Exit Sub
End If

If Not MyCreateFolder(sFolder & "\photo_h", True) Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— ’Ê— «·«⁄÷«¡ «·‘—›ÌÌ‰"
    Exit Sub
End If

Dim sTarget As String, sÚSource As String, sSourceEmpty As String
sTarget = sFolder & "\data_trans.mdb"
sSource = App.Path & "\mdb\data_trans.mdb"
sSourceEmpty = App.Path & "\mdb\data_empty.mdb"

fs.CopyFile sSourceEmpty, sSource
openConMdb conMdb, sSource

SendDataMember conMdb
SendDataMember_I conMdb
SendDataMember_H conMdb


closeCon conMdb

fs.CopyFile sSource, sTarget

Me.Caption = sCaption

Inform " „ «—”«· «·»Ì«‰«  »‰Ã«Õ"

If chkNoPhoto.Value = 0 Then
    If fs.FileExists(sSource) Then
        If SendPhotos Then Inform " „ «—”«·  ’Ê— «·«⁄÷«¡ «·⁄«„·Ì‰ »‰Ã«Õ"
        If sendPhotos_I Then Inform " „ «—”«·  ’Ê— «·«⁄÷«¡ «·„ﬁ”ÿÌ‰ »‰Ã«Õ"
        If sendPhotos_h Then Inform " „ «—”«·  ’Ê— «·«⁄÷«¡ «·‘—›Ì‰‰ »‰Ã«Õ"
        MsgBox " „ «—”«· «·»Ì«‰«  »‰Ã«Õ"
    Else
        MsgBox "„·› «·»Ì«‰«  €Ì— „ÊÃÊœ"
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Prog1.Visible = False
Prog1.Value = 0
End Sub
Private Sub Form_Load()
xDrive.text = RetSetting(xDrive.Name, TempSave(Me))
openCon con
End Sub
Private Function sendPhotos_h() As Boolean
Dim fs As New FileSystemObject, sSource As String, nRecordcount As Double, I As Long

Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset

loctable.Open "select code as Member,NULL as Serial from file3_10 where (not [NO] IS NULL)", con, adOpenStatic, adLockReadOnly
nRecordcount = loctable.RecordCount

Dim sCation As String
Me.Caption = sCaption

On Error GoTo myerror
Prog1.Visible = True
Prog1.Value = 0

Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount
    Prog1.Value = mRound(I / nRecordcount * 100, 2)
    sCode = loctable!MEMBER & turn(loctable!Serial & "", "-" & loctable!Serial)
    sSource = RetPhotoh(sCode)
    sTarget = RetPhotoh(sCode, sFolder)
    If fs.FileExists(sSource) Then
        bCopy = True
        If fs.FileExists(sTarget) Then
           If myFormat(fs.GetFile(sTarget).DateLastModified) >= myFormat(fs.GetFile(sSource).DateLastModified) Then
               bCopy = False
           End If
        End If
        If bCopy Then fs.CopyFile sSource, sTarget
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Prog1.Visible = False
Prog1.Value = 0
sendPhotos_h = True
Exit Function
myerror:
Me.Caption = sCaption
MsgBox Err.Description
Err.Clear
Prog1.Visible = False
Prog1.Value = 0
End Function
Private Function sendPhotos_I() As Boolean
Dim fs As New FileSystemObject, sSource As String, nRecordcount As Double, I As Long
Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset

Set loctable = myCmd("select code as Member,NULL as Serial from file2_10 where (status = 1 or status = 2)  union all select member,file2_11.code as Serial from file2_11 inner join file2_10 on file2_11.member = file2_10.code", con)
nRecordcount = loctable.RecordCount
On Error GoTo myerror
Prog1.Visible = True
Prog1.Value = 0

Dim sCaption As String
Me.Caption = sCaption

Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount
    Prog1.Value = mRound(I / nRecordcount * 100, 2)
    sCode = loctable!MEMBER & turn(loctable!Serial & "", "-" & loctable!Serial)
    sSource = RetPhoto_I(sCode)
    sTarget = RetPhoto_I(sCode, sFolder)
    If fs.FileExists(sSource) Then
        bCopy = True
        If fs.FileExists(sTarget) Then
           If myFormat(fs.GetFile(sTarget).DateLastModified) >= myFormat(fs.GetFile(sSource).DateLastModified) Then
               bCopy = False
           End If
        End If
        If bCopy Then fs.CopyFile sSource, sTarget
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Me.Caption = sCaption
Prog1.Visible = False
Prog1.Value = 0
sendPhotos_I = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
Me.Caption = sCaption
Prog1.Visible = False
Prog1.Value = 0
End Function
Private Function SendPhotos() As Boolean
Dim fs As New FileSystemObject, sSource As String, nRecordcount As Double, I As Long
Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset

Set loctable = myCmd("select  code as Member,NULL as Serial from file1_10  union all select member,code as Serial from file1_11 order by member ", con)
nRecordcount = loctable.RecordCount
On Error GoTo myerror
Dim sCation As String
Me.Caption = sCaption

Prog1.Visible = True
Prog1.Value = 0
Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount
    Prog1.Value = mRound(I / nRecordcount * 100, 2)
    sCode = loctable!MEMBER & turn(loctable!Serial & "", "-" & loctable!Serial)
    sSource = retPhoto(sCode)
    sTarget = retPhoto(sCode, , , sFolder)
    If fs.FileExists(sSource) Then
        bCopy = True
        If fs.FileExists(sTarget) Then
           If myFormat(fs.GetFile(sTarget).DateLastModified) >= myFormat(fs.GetFile(sSource).DateLastModified) Then
               bCopy = False
           End If
        End If
        If bCopy Then
            fs.CopyFile sSource, sTarget
        End If
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Prog1.Visible = False
Me.Caption = sCaption
Prog1.Value = 0
SendPhotos = True
Exit Function
myerror:
Me.Caption = sCaption
MsgBox Err.Description
Err.Clear
Prog1.Visible = False
Prog1.Value = 0
End Function
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
Dim I As Long, nRecordcount As Long

Set loctable = myCmd("select * from file1_10 order by code", con)
nRecordcount = loctable.RecordCount

Prog1.Visible = True
Prog1.Value = 0
sCaption = Me.Caption

Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount & " - " & "«⁄÷«¡ ⁄«„·Ì‰"
    
    Prog1.Value = mRound((I / nRecordcount) * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!CODE)
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    
    aInsert = AddFlag(aInsert, "TITLE", addstring(loctable!Title))
    
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷Ê ⁄«„·"))
    aInsert = AddFlag(aInsert, "[GENDER]", addvalue(loctable!GENDER))
    
    aPaid = Member_Paid(loctable!CODE, , con)
    
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(retFlag(aPaid, "DATE2"))))
    aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(retFlag(aPaid, "YEAR_CODE")))
    aInsert = AddFlag(aInsert, "[DOC_LAST]", addvalue(retFlag(aPaid, "DOC_NO")))
    aInsert = AddFlag(aInsert, "[DATE_LAST]", addstring(myFormat(retFlag(aPaid, "DATE"))))
    aInsert = AddFlag(aInsert, "[DIED]", IIf(loctable!died, "TRUE", "FALSE"))
    
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring("«·⁄÷Ê «·«”«”Ì"))
    conMdb.Execute addInsert(aInsert, "FILE1_10")
    loctable.MoveNext
Loop
Prog1.Visible = False
Prog1.Value = 0

Set loctable = Nothing
Set loctable = New ADODB.Recordset
cString = "select FILE1_10.DATE_BEGIN,FILE1_11.RELATION,FILE1_11.DESCA,FILE1_11.TITLE," & _
               "FILE1_11.MEMBER,FILE1_11.HANDI,FILE1_11.CODE,RELATION_CODES.DESCA AS RELATION_DESCA," & _
               "FILE1_11.DATE_BIRTH,FILE1_11.GENDER, FILE1_11.ID from FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE LEFT JOIN RELATION_CODES ON FILE1_11.RELATION = RELATION_CODES.CODE"

Set loctable = myCmd(cString, con)
               
nRecordcount = loctable.RecordCount
Prog1.Visible = True
Prog1.Value = 0
I = 0
Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount & " - " & " Ê«»⁄ «⁄÷«¡ ⁄«„·Ì‰"
    Prog1.Value = mRound(I / nRecordcount * 100, 2)

    aInsert = AddFlag(Empty, "code", loctable!CODE)
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
Prog1.Visible = False
Prog1.Value = 0
End Function
Private Function SendDataMember_I(conMdb As ADODB.Connection) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select   * from file2_10 where (status = 1 or status = 2) order by code", con)
nRecordcount = loctable.RecordCount
Prog1.Visible = True
Prog1.Value = 0
I = 0
Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount & " - " & "«⁄÷«¡  ﬁ”Ìÿ"
    Prog1.Value = mRound((I / nRecordcount) * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!CODE)
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷Ê œ⁄Ê…"))
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(loctable!DATE_END)))
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring("«·⁄÷Ê «·«”«”Ì"))
    conMdb.Execute addInsert(aInsert, "FILE2_10")
    loctable.MoveNext
Loop

Prog1.Visible = False
Prog1.Value = 0

Set loctable = Nothing
Set loctable = New ADODB.Recordset
Set loctable = myCmd("select FILE2_10.DATE_END,FILE2_11.RELATION,FILE2_11.DESCA,FILE2_11.MEMBER,FILE2_11.CODE,RELATION_CODES.DESCA AS RELATION_DESCA from file2_11 INNER JOIN FILE2_10 ON FILE2_11.MEMBER = FILE2_10.CODE LEFT JOIN RELATION_CODES ON FILE2_11.RELATION = RELATION_CODES.CODE where (file2_10.status = 1 or file2_10.status = 2)", con)

nRecordcount = loctable.RecordCount
Prog1.Visible = True
Prog1.Value = 0
I = 0
Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount & " - " & " Ê«»⁄ «⁄÷«¡  ﬁ”Ìÿ"
    
    Prog1.Value = mRound(I / nRecordcount * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!CODE)
    aInsert = AddFlag(aInsert, "MEMBER", loctable!MEMBER)
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷Ê œ⁄Ê…"))
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(loctable!DATE_END)))
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring(loctable!RELATION_DESCA & ""))
    conMdb.Execute addInsert(aInsert, "FILE2_11")
    loctable.MoveNext
Loop
Prog1.Visible = False
Prog1.Value = 0
End Function
Private Function SendDataMember_H(conMdb As ADODB.Connection) As Boolean
Dim loctable As New ADODB.Recordset
loctable.Open "select  * from file3_10 WHERE (NOT [NO] IS NULL)  order by code", con, adOpenStatic, adLockReadOnly, adCmdText
nRecordcount = loctable.RecordCount
Prog1.Visible = True
Prog1.Value = 0
I = 0
Do Until loctable.EOF
    I = I + 1
    Me.Caption = "Record " & I & " from " & nRecordcount & " - " & "«⁄÷«¡ ‘—›ÌÌ‰"
    Prog1.Value = mRound((I / nRecordcount) * 100, 2)
    aInsert = AddFlag(Empty, "code", loctable!CODE)
    aInsert = AddFlag(aInsert, "[NO]", mRound(loctable!NO))
    aInsert = AddFlag(aInsert, "desca", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[TYPE_DESCA]", addstring("⁄÷ÊÌ… ‘—›Ì…"))
    aInsert = AddFlag(aInsert, "[CARD_END]", addstring(myFormat(loctable!DATE_END)))
    aInsert = AddFlag(aInsert, "[RELATION_DESCA]", addstring("«·⁄÷Ê «·«”«”Ì"))
    conMdb.Execute addInsert(aInsert, "FILE3_10")
    loctable.MoveNext
Loop

Prog1.Visible = False
Prog1.Value = 0
End Function


