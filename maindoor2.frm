VERSION 5.00
Begin VB.Form maindoorfrm2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«” ⁄·«„ «⁄÷«¡ «·‰«œÌ"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15885
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "maindoor2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "maindoorfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, oSearchMember As New Search, oSearchRel As New Search
Dim CardTable As ADODB.Recordset, cFilter As String, cFile As String, cFile_Rel As String, oSearchType As New Search_empty
Const LoadMode = 1, DefineMode = 2
Private Sub cmdType_Click()
Set oSearchType = New Search_empty
TypeLookUp Me, oSearchType
End Sub
Private Sub CmdDel_Click()
xCode.Caption = ""
XCODE_REL.Caption = ""
openCardTable
myUndo
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
mydefine
myload
End Sub
Private Sub CmdInform_Click()
If cmdType.Tag = 1 Then
    MemberLookupAll Me, oSearchMember
ElseIf cmdType.Tag = 2 Then
    Member_InLookupAll Me, oSearchMember
End If
End Sub

Private Sub cmdInformRel_Click()
If cmdType.Tag = 1 Then
    relLookupAll Me, oSearchRel
ElseIf cmdType.Tag = 2 Then
    relLookupAll_I Me, oSearchRel
End If
End Sub
Private Sub Form_Load()
SetKbLayout Lang_AR

openCon con
'mydefine
cmdType.Tag = "1"
cmdType.Caption = "⁄÷ÊÌ… ⁄«„·…"
cFile = "file1_10"
cFile_Rel = "file1_11"
myUndo
End Sub
Private Function openCardTable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 * FROM " & cFile
If pCode <> "" Then cWhere = "CODE " & pSign & addvalue(pCode)

cFilter = ""
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by CODE desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by CODE ASC"
End If
CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Function
Private Sub myUndo()
'On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xCode.Caption) Then
    openCardTable xCode.Caption
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub CmdNext_Click()
openCardTable xCode.Caption, ">"
If CardTable.EOF Then openCardTable xCode.Caption, "="
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xCode.Caption, "<"
If CardTable.EOF Then openCardTable xCode.Caption, "="
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub myload()
Dim acode As Variant, loctable As New ADODB.Recordset, nIndex As Long, sPhoto As String, sPhotoRecord
Dim nUnPaid As Integer
Dim aPaid As Variant

ClearText
xCode.Caption = CardTable!code & ""
xDesca.Caption = CardTable!desca
If Not IsNull(CardTable!Type) Then
    xType_desca.Caption = GetField("SELECT DESCA FROM TYPE_CODES WHERE CODE = " & addvalue(CardTable!Type), con) & ""
Else
    xType_desca.Caption = ""
End If
xDate_Birth.Caption = myFormat_p(CardTable!DATE_BIRTH)
Photo_main.Images.Clear
If cmdType.Tag = 1 Then
    If validPhoto(RetPhoto(xCode.Caption)) Then
        Set Photo_main.Picture = LoadPicture(RetPhoto(xMember.Caption))
    End If
    xLast_date.Caption = ""
        aPaid = Member_Paid(xCode.Caption, , con)
    If Not IsEmpty(aPaid) Then
        xUnPaid.Caption = unpaid_years(retFlag(aPaid, "year_code"), sSeason, con)
        xLast_date.Caption = myFormat_p(retFlag(aPaid, "date"))
        If mRound(xUnPaid.Caption) = 0 Then xUnPaid.Caption = "·«  ÊÃœ ”‰Ê«  €Ì— „”œœ…"
        If retFlag(aPaid, "is_save") Then
            XLAST_PAID.Caption = "Õ«›Ÿ ⁄÷ÊÌ… Õ Ì " & retFlag(aPaid, "year_desca") & ""
        Else
            XLAST_PAID.Caption = "„”œœ Õ Ì " & retFlag(aPaid, "year_desca") & ""
        End If
    Else
        XLAST_PAID.Caption = "·„ Ì”œœ „‰ ﬁ»·"
        xUnPaid.Caption = unpaid_years_count(xCode.Caption, sSeason, con)
    End If
ElseIf cmdType.Tag = 2 Then
    If validPhoto(RetPhoto_I(xCode.Caption)) Then
        Set Photo_main.Picture = LoadPicture(RetPhoto_I(xCode.Caption))
    End If
    cWhere = "INSTALL_BALANCE.CODE = " & xCode.Caption
    cWhere = cWhere & " AND " & "INSTALL_BALANCE.Value - INSTALL_BALANCE.VALUE_PAID > 0"
    cWhere = cWhere & " AND " & "INSTALL_BALANCE.DATE_DUE <= " & DateSq(Date)
    nUnPaid = mRound(GetField("SELECT Sum(INSTALL_BALANCE.INS_COUNT) AS Ins_Count FROM INSTALL_BALANCE  WHERE " & cWhere, con))
    xUnPaid_Install.Caption = IIf(nUnPaid = 0, "·«  ÊÃœ «ﬁ”«ÿ „ √Œ—…", nUnPaid)
    xDate_Birth.Caption = myFormat_p(CardTable!DATE_BIRTH)
    xLast_Date_I.Caption = myFormat_p(GetField("select dbo.f_last_year_date_install(" & xCode.Caption & ")", con))
    xFirst_Install.Caption = myFormat_p(GetField("select top 1 date_due from INSTALL_BALANCE WHERE VALUE - VALUE_PAID  > 0 AND CODE = " & xCode.Caption & " ORDER BY DATE_DUE ASC", con))
    xInstall_desca.Caption = CardTable!install_desca & ""
    xRelation_Desca.Caption = CardTable!RELATION_DESCA & ""
End If
Handlecontrols LoadMode
End Sub
Sub Handlecontrols(nMode)
aRecords = retRecords(xCode.Caption)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")
If nMode = LoadMode Then
    xRecord_No.Caption = ArbString("”Ã· " & nRecord & " „‰ " & nRecords)
Else
    xRecord_No.Caption = "·«  ÊÃœ ”Ã·« "
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2
End Sub
Private Function retRecords(pCode) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If ValidNum(pCode) Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN CODE <= " & pCode & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM FILE1_10 " & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If ValidNum(pCode) Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set oSearchMember = Nothing
'addSetting xDrive.Name, xDrive.Text, TempSave(Me)
closeCon con
Set maindoorfrm = Nothing
Err.Clear
End
End Sub
Private Sub Photo_main_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Tag <> "" Then
    XCODE_REL.Caption = Source.Tag
    myUndo
ElseIf Source.Tag = -1 Then
    XCODE_REL.Caption = ""
    myUndo
End If
End Sub
Private Function MyLoadPhotos() As Boolean
Dim loctable As New ADODB.Recordset, cString As String, cWhere As String, I As Long, sPhotoRecord As String
For I = 1 To Photo1.UBound
    Photo1(I).Images.Clear
    Photo1(I).Tag = ""
    xDesca(I).Caption = ""
    Photo1(I).Visible = False
    xdesca1(I).Visible = False
Next
cString = "SELECT " & cFile_Rel & ".* " & _
           " FROM " & cFile_Rel

cWhere = "MEMBER = " & xCode.Caption
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " order by CODE"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
I = 0
Do Until loctable.EOF
    I = I + 1
    If I <= Photo1.UBound Then
        Photo1(I).Visible = True
        xdesca1(I).Visible = True
        xdesca1(I).Caption = loctable!desca & ""
        sPhotoRecord = loctable!MEMBER
        If loctable!flag <> 1 Then
            sPhotoRecord = sPhotoRecord & "-" & loctable!code
            Photo1(I).Tag = loctable!code
        Else
            Photo1(I).Tag = "-1"
        End If
        
        If cmdType.Tag = 1 Then
            If validPhoto(RetPhoto(sPhotoRecord)) Then
                Photo1(I).Import.FromFile RetPhoto(sPhotoRecord)
             End If
        ElseIf cmdType.Tag = 2 Then
            If validPhoto(RetPhoto_I(sPhotoRecord)) Then
                Photo1(I).Import.FromFile RetPhoto_I(sPhotoRecord)
             End If
        End If
    End If
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
Dim I As Long
For I = 1 To Photo1.UBound
    Photo1(I).Images.Clear
    Photo1(I).Tag = ""
    xdesca1(I).Caption = ""
    Photo1(I).Visible = False
    xdesca1(I).Visible = False
Next
Photo_main.Images.Clear
ClearText
xRecord_No.Caption = "·«  ÊÃœ »Ì«‰« "
Handlecontrols DefineMode
End Sub
Private Sub xBarCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If cmdGo.Enabled Then cmdGo_Click
End If
End Sub
Sub myProc()
If ActiveControl.Name = Me.cmdInform.Name Then
    xCode.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
    XCODE_REL.Caption = ""
    oSearchMember.Hide
ElseIf ActiveControl.Name = cmdInformRel.Name Then
    xCode.Caption = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.Row, 0)
    XCODE_REL.Caption = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.Row, 1)
    oSearchRel.Hide
ElseIf ActiveControl.Name = cmdType.Name Then
    cmdType.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
    cmdType.Caption = IIf(cmdType.Tag = "", "‰Ê⁄ «·⁄÷ÊÌ…", oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1))
    If cmdType.Tag = 1 Then
        cFile = "MEMBERS"
        FrameInstall.Visible = False
        FrameEmp.Visible = True
    ElseIf cmdType.Tag = 2 Then
        cFile = "MEMBERS_INV"
        FrameInstall.Visible = True
        FrameEmp.Visible = False
    End If
    oSearchType.Hide
End If
myUndo
End Sub
Private Sub CloseData()
'closeCon con
End Sub
Private Sub TypeLookUp(oForm As Form, oSearch As Form, Optional bAddRow As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT TYPE_CODES_SPORT.CODE,TYPE_CODES_SPORT.DESCA" & _
                  " From TYPE_CODES_SPORT "

Generalarray(2) = "Order by CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·‰Ê⁄"
listarray(0, 1) = "(%%TYPE_CODES_SPORT.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·‰Ê⁄"
GrdArray(1, 1) = 5500

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "‰Ê⁄ «·⁄÷ÊÌ…")
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.sCaption = "≈” ⁄·«„ «‰Ê«⁄ «·⁄÷ÊÌ…"
oSearch.Show 1
End Sub
Private Sub ClearText()
Photo_main.Images.Clear
xLast_date.Caption = ""
xUnPaid.Caption = ""
xInstall_desca.Caption = ""
xUnPaid_Install.Caption = ""

xDesca.Caption = ""
xType_desca.Caption = ""
xLast_Date_I.Caption = ""
xFirst_Install.Caption = ""
xDate_Birth.Caption = ""
End Sub

