VERSION 5.00
Begin VB.Form maindoorfrmold 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«” ⁄·«„ «·«⁄÷«¡"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main_door_old.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   19290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "maindoorfrmold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conMdb As New ADODB.Connection, oSearchMember As New Search_mdb
Dim CardTable As ADODB.Recordset

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdGetData_Click()
Dim fs As New FileSystemObject
If Trim(xDrive.Text) = "" Then
    MsgBox "«·ﬁ—’ €Ì— „”Ã·"
    Exit Sub
End If

Dim sSource As String, sTarget As String
sSource = xDrive.Text & ":\etahad_door_sql\data_trans.mdb"
On Error GoTo myerror
If fs.FileExists(sSource) Then
    CloseData
    fs.CopyFile sSource, App.Path & "\mdb_door\data.mdb"
    OpenData
    GetPhotos
    MsgBox " „ ”Õ» «·»Ì«‰«  »‰Ã«Õ"
Else
    MsgBox "„·› «·»Ì«‰«  €Ì— „ÊÃÊœ"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub CmdGo_Click()
mydefine
myload
'xBarCode.Text = ""
End Sub

Private Sub Command1_Click()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select code,[desca] from file1_10"
Generalarray(2) = "Order by FILE1_10.CODE"
Generalarray(3) = 7000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-—ﬁ„ «·⁄÷ÊÌ…"
listarray(0, 1) = "(VAL('cFilter') = CODE OR %%DESCA%%)"


GrdArray(0, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 9000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchMember.Caption = "«” ⁄·«„ «·«⁄÷«¡"
oSearchMember.Show 1
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdInform_Click()
MemberLookup_I
End Sub

Private Sub Form_Load()
SetKbLayout Lang_AR
'xDrive.Text = RetSetting(xDrive.Name, TempSave(Me))
'OpenData
openConMdb conMdb, App.Path & "\MDB\DATA_TRANS.MDB"
mydefine
End Sub
Private Function openCardTable()
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
cString = "SELECT MEMBERS_INV.* " & _
           " FROM MEMBERS_INV"

cFilter = ""
cFilter = "MEMBERS_INV.MEMBER = " & xCode.Caption
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
cString = cString & " order by MEMBERS_INV.CODE"

Set CardTable = New ADODB.Recordset
CardTable.Open cString, conMdb, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If XCODE_REL.Caption <> "" Then
        CardTable.Find "CODE = " & addvalue(XCODE_REL.Caption), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveFirst
    Else
        CardTable.MoveFirst
    End If
    myload
End If
End Sub
Private Sub myload()
Dim acode As Variant, loctable As New ADODB.Recordset, nIndex As Long, sPhoto As String, sPhotoRecord

If Not MYVALID(acode) Then
    On Error Resume Next
    Photo1(nIndex).Import.FromFile MainPath & "\error.jpg"
    Err.Clear
    Exit Sub
End If
'MyLoadMember acode

'If retFlag(acode, "TYPE") = "1" Then
'    MyLoadMember acode
'End If
'    If retFlag(aCode, "CODE") = "" And IsNull(LOCTABLE!code) Then
'        If validPhoto(RetPhoto(LOCTABLE!Member)) Then
'            Photo1(0).Picture = LoadPicture(RetPhoto(LOCTABLE!Member))
'        End If
'    ElseIf validInt(retFlag(aCode, "CODE") & "") And Val(retFlag(aCode, "CODE") & "") = Val(LOCTABLE!code & "") Then
'        If validPhoto(RetPhoto(LOCTABLE!Member & "-" & LOCTABLE!code)) Then
'            Photo1(0).Picture = LoadPicture(RetPhoto(LOCTABLE!Member & "-" & LOCTABLE!code))
'        End If
'    Else
'        nIndex = nIndex + 1
'        sPhoto = RetPhoto(LOCTABLE!Member & turn(LOCTABLE!code & "", "-" & LOCTABLE!code))
'        photo1(nIndex) =
'    End If

End Sub
Function aUnMyCodeBar(sCode) As Variant
Dim nVal1 As Integer, nVal2 As Integer, nValSub1 As Integer, nValSub2 As Integer
Dim nNumber1 As String, nNumber2 As String, nNumber3 As String


If Trim(sCode) = "" Then Exit Function
Dim aRet As Variant, aSplit As Variant
aSplit = Split(sCode, "-")
If IsEmpty(aSplit) Then Exit Function
If UBound(aSplit) > 1 Then Exit Function
If Not ValidInt(Val(aSplit(0))) Then Exit Function
If UBound(aSplit) = 1 Then
    If Not ValidInt(Val(aSplit(1))) Then Exit Function
End If

'If Not ValidInt(Val(aret(1))) Then Exit Function
'If Not ValidInt(Val(aret(2))) Then Exit Function

'nVal1 = 74: nVal2 = 11: nValSub1 = 71: nValSub2 = 4

'nNumber1 = aret(0)
'nNumber2 = aret(1)
'nNumber3 = aret(2)
'
'nNumber1 = StrReverse(nNumber1)
'nNumber1 = Val(nNumber1) + Val(nVal2)
'nNumber1 = nNumber1 * 2
'nNumber1 = Val(Left(nNumber1, Len(nNumber1) - 1))
'nNumber1 = nNumber1 - nVal1
'
'nNumber2 = Val(nNumber2) + Val(nValSub2)
'nNumber2 = nNumber2 * 2
'nNumber2 = Val(Left(nNumber2, Len(nNumber2) - 1))
'nNumber2 = nNumber2 - nValSub1
'nNumber2 = nNumber2 - Right(nNumber1, 1)
aRet = AddFlag(Empty, "MEMBER", aSplit(0))
If UBound(aSplit) = 1 Then aRet = AddFlag(aRet, "CODE", aSplit(1))
'aret = AddFlag(aret, "TYPE", IIf(Val(nNumber3) = 0, "", Val(nNumber3)))
aUnMyCodeBar = aRet
End Function
Private Function MYVALID(acode) As Boolean
If IsEmpty(acode) Then
    n = Beep(1000, 1000)
   MsgBox "«·ﬂÊœ €Ì— „ÊÃÊœ «Ê Œÿ√ ›Ï «·»«—ﬂÊœ", vbCritical
    Exit Function
End If

If Not ValidInt(retFlag(acode, "MEMBER")) Then
    n = Beep(1000, 1000)
    MsgBox "«·ﬂÊœ €Ì— „ÊÃÊœ «Ê Œÿ√ ›Ï «·»«—ﬂÊœ", vbCritical
    Exit Function
End If

If (Not ValidInt(retFlag(acode, "CODE"))) And Trim(retFlag(acode, "CODE")) <> "" Then
    n = Beep(1000, 1000)
    MsgBox "«·ﬂÊœ €Ì— „ÊÃÊœ «Ê Œÿ√ ›Ï «·»«—ﬂÊœ", vbCritical
    Exit Function
End If

If Val(retFlag(acode, "CODE")) > 20 Then
    n = Beep(1000, 1000)
    MsgBox "Œÿ√ ›Ï «·»«—ﬂÊœ", vbCritical
    Exit Function
End If

'If Not ValidInt(retFlag(acode, "TYPE")) Then
'    n = Beep(1000, 1000)
'    MsgBox "Œÿ√ ›Ï ‰Ê⁄ «·»«—ﬂÊœ", vbCritical
'    Exit Function
'End If

If ValidInt(retFlag(acode, "MEMBER")) And Trim(retFlag(acode, "CODE")) = "" Then
    If IsEmpty(GetField("SELECT CODE FROM FILE1_10 WHERE CODE = " & retFlag(acode, "MEMBER"))) Then
        MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „ÊÃÊœ ", vbCritical
        n = Beep(1000, 1000)
        Exit Function
    End If
End If

If ValidInt(retFlag(acode, "MEMBER")) And ValidInt(retFlag(acode, "CODE")) Then
    If IsEmpty(GetField("SELECT CODE FROM FILE1_10 WHERE CODE = " & retFlag(acode, "MEMBER"))) Then
        MsgBox "ﬂÊœ «·⁄÷Ê «·«”«”Ì €Ì— „ÊÃÊœ ", vbCritical
        n = Beep(1000, 1000)
        Exit Function
    End If

    If IsEmpty(GetField("SELECT MEMBER FROM FILE1_11 WHERE MEMBER = " & retFlag(acode, "MEMBER") & " AND CODE = " & retFlag(acode, "CODE"))) Then
        MsgBox "ﬂÊœ «·⁄÷Ê «· «»⁄ €Ì— „ÊÃÊœ Ê«·«”«”Ì „ÊÃÊœ ", vbCritical
        n = Beep(1000, 1000)
        Exit Function
    End If
End If

MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
closeCon conMdb
addSetting xDrive.Name, xDrive.Text, TempSave(Me)
Set maindoorfrm = Nothing
End Sub

Private Sub Label9_Click()

End Sub

Private Sub Photo1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If Source.Tag <> "" Then
    xBarCode.Text = Source.Tag
    CmdGo_Click
End If
End Sub

Private Sub xBarCode_Change()
cmdGo.Enabled = Trim(xBarCode.Text) <> ""
End Sub
Private Function MyLoadMember(ByVal acode As Variant) As Boolean
Dim aMember As Variant, nCaption As Long
xMember.Caption = retFlag(acode, "MEMBER") & turn(retFlag(acode, "CODE") & "", "-" & retFlag(acode, "CODE"))
xDateLast.Caption = PaidString(retFlag(acode, "MEMBER"))
If retFlag(acode, "CODE") = "" Then
    xdesca.Caption = GetField("select desca from file1_10 where code = " & retFlag(acode, "MEMBER"))
    xType.Caption = "«·⁄÷Ê «·«”«”Ì"
Else
    Dim aRet As Variant
    aRet = GetFields("select FILE1_10.DESCA AS MEMBER_DESCA,FILE1_11.desca ,FILE0_00.DESCA AS REL_DESCA from (FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE) LEFT JOIN FILE0_00 ON (FILE1_11.RELATION = FILE0_00.CODE AND FILE0_00.FLAG = 0) where FILE1_11.member = " & retFlag(acode, "MEMBER") & " and FILE1_11.code = " & retFlag(acode, "code"))
    If Not IsEmpty(aRet) Then
        xdesca.Caption = retFlag(aRet, "MEMBER_DESCA") & vbCrLf & retFlag(aRet, "REL_DESCA") & " " & "(" & retFlag(aRet, "DESCA") & ")"
    End If
End If

Dim loctable As New ADODB.Recordset
cString = "SELECT FILE1_10.CODE AS MEMBER,FILE1_10.DESCA,'«·⁄÷Ê ‰›”Â' AS REL_DESCA,NULL AS CODE" & _
          " FROM FILE1_10 "
cString = cString & turn(cString) & "FILE1_10.CODE = " & retFlag(acode, "MEMBER")
cString = cString & " UNION ALL "
cString = cString & "SELECT FILE1_10.CODE AS MEMBER,FILE1_11.DESCA ,FILE0_00.DESCA AS REL_DESCA,FILE1_11.CODE" & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.CODE = FILE1_11.MEMBER) LEFT JOIN FILE0_00 ON (FILE1_11.RELATION = FILE0_00.CODE AND FILE0_00.FLAG = 0) "
cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE = " & retFlag(acode, "MEMBER")
cString = cString & turn(cWhere) & cWhere
cString = cString & " ORDER BY CODE"

On Error GoTo myerror
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
sPhoto = retFlag(acode, "MEMBER") & turn(retFlag(acode, "CODE") & "", "-" & retFlag(acode, "CODE"))
Do Until loctable.EOF
    sPhotoRecord = loctable!MEMBER & turn(loctable!CODE & "", "-" & loctable!CODE)
    If sPhotoRecord = sPhoto Then
        If validPhoto(RetPhoto(sPhotoRecord)) Then
            Photo1(0).Tag = sPhotoRecord
            Photo1(0).Import.FromFile RetPhoto(sPhoto)
         End If
    Else
        nIndex = nIndex + 1
        If validPhoto(RetPhoto(sPhotoRecord)) Then
            Photo1(nIndex).Visible = True
            Photo1(nIndex).Import.FromFile RetPhoto(sPhotoRecord)
            Photo1(nIndex).Tag = sPhotoRecord
        End If
        xdesca1(nIndex).Caption = loctable!desca & ""
    End If
    nCaption = nCaption + 1
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub mydefine()
Dim i As Long
'For i = 1 To Photo1.UBound
'    Photo1(i).Images.Clear
'    Photo1(i).Tag = ""
'    'xdesca1(i).Caption = ""
'Next
'xdesca.Caption = ""
'xMember.Caption = ""
'xCode.Caption = ""
'xType_desca.Caption = ""
'xCard_end.Caption = ""
End Sub
Private Sub xBarCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If cmdGo.Enabled Then CmdGo_Click
End If
End Sub
Sub myProc()
xCode.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
oSearchMember.Hide
openCardTable
myUndo
End Sub
Private Sub CloseData()
closeCon con
End Sub
Private Sub GetPhotos()
Dim fs As New FileSystemObject, sSource As String, nRecordCount As Double, i As Long
Dim conMdb As New ADODB.Connection
openConMdb pCon
Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset

loctable.Open "select code as Member,NULL as Serial from file1_10  union all select member,code as Serial from file1_11 ", con, adOpenStatic, adLockReadOnly
If Not loctable.EOF Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
    Prog1.Visible = True
    Prog1.Value = 0
End If
On Error GoTo myerror

Do Until loctable.EOF
    i = i + 1
    bCopy = True
    sCode = loctable!MEMBER & turn(loctable!Serial & "", "-" & loctable!Serial)
    sSource = RetPhotoNew(sCode, , , xDrive & ":\etahad_door")
    sTarget = RetPhoto(sCode)
    
    If fs.FileExists(sSource) Then
        If fs.FileExists(sTarget) Then
           If myFormat(fs.GetFile(sTarget).DateLastModified) >= myFormat(fs.GetFile(sSource).DateLastModified) Then
               bCopy = False
           End If
        End If
        If bCopy Then fs.CopyFile sSource, sTarget
    End If
    Me.Caption = i
    If Prog1.Value <> Int(i / nRecordCount * 100) Then Prog1.Value = Int(i / nRecordCount * 100)
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Prog1.Visible = False
Prog1.Value = 0

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Prog1.Visible = False
Prog1.Value = 0
End Sub

Private Sub xDrive_Change()
xDrive.Text = UCase(xDrive.Text)
End Sub
Private Sub MemberLookup_I(Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE1_50.CODE,FILE1_50.DESCA,FORMAT(CARD_END,'DD/MM/YYYY')" & _
                  " From FILE1_50 "

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE1_50.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE1_50.DESCA%% OR **FILE1_50.CODE**)"

GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «·⁄÷Ê"
GrdArray(1, 1) = 5500

GrdArray(2, 0) = " «—ÌŒ «·«‰ Â«¡"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchMember.pMdbPath = App.Path & "\MDB\DATA_TRANS.MDB"
oSearchMember.Caption = "≈” ⁄·«„ «⁄÷«¡ «·œ⁄Ê…"
oSearchMember.Show 1
End Sub


