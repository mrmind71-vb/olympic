Attribute VB_Name = "abd"
Public bedit As Boolean
Public sboxSales As String
Public bSupermode As Boolean
Public searchArray
Function Newflag(cTable, cField, Optional pCon As ADODB.Connection, Optional cFilter As String = "") As Double
Dim loctable As New ADODB.Recordset
If pCon Is Nothing Then
    loctable.Open "Select Max(" & cField & ") as Maxof From " & cTable & turn(cFilter) & cFilter, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    loctable.Open "Select Max(" & cField & ") as Maxof From " & cTable & turn(cFilter) & cFilter, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If

If Not (loctable.EOF And loctable.BOF) Then
    Newflag = Val(loctable!maxOf & "") + 1
End If

If Newflag = 0 Then Newflag = 1
loctable.Close
Set loctable = Nothing
End Function
Function DateSq(ByVal X As Variant, Optional bShort As Boolean = False) As String
If Not IsDate(X) Then
    DateSq = X
    Exit Function
End If
If bShort Then
    X = Format(X, "YYYY-MM-DD")
Else
    X = Format(X, "YYYY-MM-DD HH:NN")
End If
DateSq = MyParn(X)
End Function
Function DateSq_mdb(ByVal X As Variant) As String
If Not IsDate(X) Then
    DateSq_mdb = X
    Exit Function
End If
DateSq_mdb = "#" & myFormat(X) & "#"
End Function
Function DateConv(ByVal X As Variant)
If Not IsDate(X) Then
    DateConv = ""
    Exit Function
End If
X = Format(X, "YYYY-MM-DD")
DateConv = Right(X, 4) & "-" & Mid(X, 4, 2) & "-" & Left(X, 2)
End Function
Function DateFix(dDate) As Variant
If Not IsDate(dDate) Then
    DateFix = Null
    Exit Function
End If
DateFix = DateValue(Format(dDate, "YYYY-MM-DD"))
End Function
Function MyParnAnd(cSearch, cField) As String
Dim aString, cString2
aString = Split(Trim(cSearch), " ")
For i2 = 0 To UBound(aString)
    If Trim(aString(i2)) <> "" Then cString2 = cString2 & IIf(cString2 = "", "", " and ") & cField & " Like " & "'%" & aString(i2) & "%'"
Next
MyParnAnd = cString2
End Function
Function aGetDesca(pString, Optional pCon As ADODB.Connection) As Variant
Dim loctable As New ADODB.Recordset
If pCon Is Nothing Then
    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If
ReDim aRet(0)
If Not (loctable.BOF And loctable.EOF) Then
    ReDim aRet(loctable.Fields.Count)
    For i = 0 To loctable.Fields.Count - 1
        aRet(i + 1) = loctable.Fields(i).Value
    Next
End If
aGetDesca = aRet
loctable.Close
Set loctable = Nothing
End Function
Function GetDesca(pString, Optional pCon As ADODB.Connection) As String
Dim loctable As New ADODB.Recordset
loctable.CursorLocation = adUseClient
If pCon Is Nothing Then
     loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If
If Not (loctable.BOF And loctable.EOF) Then GetDesca = loctable(0) & ""
loctable.Close
Set loctable = Nothing
End Function
Function GetBoolean(pString, Optional pCon As ADODB.Connection) As Integer
Dim loctable As New ADODB.Recordset
loctable.CursorLocation = adUseClient
If pCon Is Nothing Then
    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If
If Not (loctable.BOF And loctable.EOF) Then
    GetBoolean = IIf(loctable(0), 1, 0)
Else
    GetBoolean = -1
End If
loctable.Close
Set loctable = Nothing
End Function
Public Sub Inform(Mcaption As String, Optional mCaption2 As String, Optional nInterval As Integer = 900)
On Error Resume Next
Informfrm.sLabel1 = Mcaption
Informfrm.sLabel2 = mCaption2
Informfrm.nInterval = nInterval
Informfrm.Show 1
DoEvents
Err.Clear
End Sub
Public Sub InformOk(Mcaption As String)
On Error Resume Next
Load Informfrm
InformfrmOK.lbl_inform.Caption = Mcaption
InformfrmOK.Show 1
Err.Clear
End Sub
Function addDate(pDate As Variant, Optional bShort As Boolean = False) As String
If Not myIsdate(pDate & "") Then
    addDate = "NULL"
Else
    addDate = DateSq(pDate & "", bShort)
End If
End Function
Function myIsdate(pDate As Variant) As Boolean
If Not IsDate(pDate & "") Then Exit Function
If Not validYear(Year(pDate)) Then Exit Function
myIsdate = True
End Function
Function validYear(nYear As String) As Boolean
If Not ValidNum(nYear) Then Exit Function
validYear = Val(nYear) >= 1800 And Val(nYear) <= 2100
End Function
Function validMonth(nMonth As String) As Boolean
If Not ValidInt(nMonth) Then Exit Function
validMonth = Val(nMonth) >= 1 And Val(nMonth) <= 12
End Function
Function BetweenString(ByVal pValue1 As String, ByVal pValue2 As String, Optional pString1 = "„‰", Optional pString2 = "Õ Ì")
If Trim(pValue1) = "" And Trim(pValue2) = "" Then Exit Function
If Trim(pValue1) <> "" Then BetweenString = Trim(pString1) & turn(pString1, " ") & pValue1
If Trim(pValue2) <> "" Then BetweenString = BetweenString & turn(BetweenString, " ") & Trim(pString2) & turn(pString2, " ") & Trim(pValue2)
End Function
Function retHeader(aHeader, nBegin, pCount, Optional pSep As String = "  ") As String
Dim nFound As Integer, i As Integer, nCount As Integer
If IsEmpty(aHeader) Then Exit Function
If EmptyArray(aHeader) Then Exit Function
For i = 0 To UBound(aHeader)
    If aHeader(i) <> "" Then
        If nFound >= nBegin Then
            retHeader = retHeader + IIf(retHeader = "", "", pSep) & aHeader(i)
            nCount = nCount + 1
            If nCount = pCount Then Exit For
        End If
        nFound = nFound + 1
    End If
Next
End Function
Function retDef(cTable, Optional cField As String = "code", Optional pWhere As String) As String
Dim loctable As New ADODB.Recordset, cString As String
cString = "select Min(" & cField & ") as myField from " & cTable
cString = cString & turn(pWhere) & pWhere
cString = cString & " having count(*) = 1"
loctable.Open cString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then retDef = loctable!myField & ""
LastFunction:
loctable.Close
Set loctable = Nothing
End Function
Function myDef(cTable, Optional cField As String = "code", Optional cWhere As String) As String
Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT MIN(" & cField & ")  FROM " & cTable
If cWhere <> "" Then cString = cString & turn(cString) & cWhere
cString = cString & " HAVING COUNT(*) = 1"
Dim aRet As Variant
aRet = GetField(cString)
If Not IsEmpty(aRet) Then myDef = aRet & ""
End Function

Sub grdMake(pString As String, pFieldBound, pFieldList, pCon As ADODB.Connection, grid1 As VSFlexGrid, Optional pRows As Integer = 10)
Dim rstLocal As ADODB.Recordset, cString As String
Set rstLocal = New ADODB.Recordset
grid1.rows = 0
grid1.rows = pRows
rstLocal.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
cString = "#" & ";"
Do Until rstLocal.EOF
    cString = cString & "|#" & rstLocal(pFieldBound) & ";" & rstLocal(pFieldList)
    rstLocal.MoveNext
Loop
rstLocal.Close
Set rstLocal = Nothing
grid1.ColComboList(0) = cString
End Sub
Function GrdQry(pGrid As VSFlexGrid, pField, Optional isString As Boolean) As String
Dim cString As String
For i = 0 To pGrid.rows - 1
    If pGrid.TextMatrix(i, 0) <> "" Then cString = cString & IIf(cString = "", "(", " or ") & pField & " = " & IIf(isString, "'", "") & pGrid.TextMatrix(i, 0) & IIf(isString, "'", "")
Next
If cString <> "" Then cString = cString & ")"
GrdQry = Trim(cString)
End Function
Function GrdTitle(pGrid As VSFlexGrid) As String
Dim cString As String
For i = 0 To pGrid.rows - 1
    If pGrid.TextMatrix(i, 0) <> "" Then cString = cString & IIf(cString = "", "", " - ") & pGrid.Cell(flexcpTextDisplay, i, 0, i, 0)
Next
GrdTitle = cString
End Function
Function retFilter(pTable As ADODB.Recordset, pFilter)
Dim aFilter
ReDim aFilter(pTable.Fields.Count - 1)
pTable.Filter = pFilter
If Not (pTable.EOF And pTable.BOF) Then
    For i = 0 To pTable.Fields.Count - 1
        aFilter(i) = pTable.Fields(i).Value
    Next
Else
    For i = 0 To pTable.Fields.Count - 1
        aFilter(i) = Null
    Next
End If
retFilter = aFilter
End Function
Function MyCreateFolder(pDir As String, Optional bMsg As Boolean = False) As Boolean
On Error GoTo myerror
Dim fs As FileSystemObject
Set fs = CreateObject("Scripting.FileSystemObject")
aString = Split(pDir, "\")
cString = aString(0)
For i = 1 To UBound(aString)
    On Error Resume Next
    cString = cString & "\" & aString(i)
    If Not fs.FolderExists(cString) Then fs.CreateFolder (cString)
    Err.Clear
Next
MyCreateFolder = fs.FolderExists(cString)
Exit Function
myerror:
If bMsg Then MsgBox Err.Description
Err.Clear
End Function
Function LastDrive(Optional bLetter As Boolean = False)
Dim fs, d, DC, letter
Set fs = CreateObject("Scripting.FileSystemObject")
Set DC = fs.Drives
For Each d In DC
    If d.DriveType = 1 Or d.DriveType = 2 Then
        On Error Resume Next
        letter = IIf(bLetter, d.DriveLetter, d.SerialNumber)
    End If
Next
LastDrive = letter
End Function
Function CreateInsert(ByVal aInsert, ByVal cTable) As String
Dim cString1 As String, cString2 As String
For i = 0 To UBound(aInsert)
    If aInsert(i, 0) <> "" Then
        cString1 = cString1 & IIf(cString1 = "", "", ",") & aInsert(i, 0)
        cString2 = cString2 & IIf(cString2 = "", "", ",") & aInsert(i, 1)
    End If
Next
CreateInsert = "Insert into " & cTable & " (" & _
                cString1 & _
                ")"
CreateInsert = CreateInsert & " values(" & _
                cString2 & _
                ")"
End Function
Function CreateUpdate(ByVal aInsert, ByVal cTable, ByVal cCondition, Optional ByVal aIg) As String
Dim bUpDate As Boolean, cString As String
If IsMissing(aIg) Then aIg = Array(0)

For i = 0 To UBound(aIg)
    If aIg(i) >= 0 And aIg(i) < UBound(aInsert) Then
        aInsert(aIg(i), 0) = ""
    End If
Next

For i = 0 To UBound(aInsert)
    If aInsert(i, 0) <> "" Then
        CreateUpdate = CreateUpdate & IIf(CreateUpdate = "", "", ",") & aInsert(i, 0) & _
                       " = " & aInsert(i, 1)
    End If
Next
CreateUpdate = "UPDATE " & cTable & " SET " & _
               CreateUpdate & _
               cCondition
End Function
Function DefAdd(sFlag, sFlagDesca, sFlagValue)
On Error Resume Next
condef.Execute "insert into DEFTABLE(Flag,FlagDesca,FlagValue)" & _
               "Values(" & _
               addstring(sFlag) & "," & _
               addstring(sFlagDesca) & "," & _
               addstring(sFlagValue) & _
               ")"
If Err.Number = -2147467259 Then
    Err.Clear
    condef.Execute "update defTable Set " & _
                   " FlagValue  = " & addstring(sFlagValue) & _
                   " where Flag = " & MyParn(sFlag) & _
                   " and FlagDesca = " & MyParn(sFlagDesca)
End If
Err.Clear
End Function
Function DefGet(sFlag, sFlagDesca) As String
Dim loctable As New ADODB.Recordset
cString = "Select * From defTable " & _
          " where Flag = " & MyParn(sFlag) & _
          " and FlagDesca = " & MyParn(sFlagDesca)
loctable.Open cString, condef, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.BOF And loctable.EOF) Then DefGet = loctable!FlagVAlue & ""
loctable.Close
Set loctable = Nothing
End Function
Function StrList(cString)
Dim listTable As New ADODB.Recordset
listTable.Open cString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until listTable.EOF
    StrList = StrList & "|#" & listTable.Fields(0) & ";" & listTable.Fields(1)
    listTable.MoveNext
Loop
End Function
Function StrList2(cString) As String
Dim listTable As New ADODB.Recordset
listTable.Open cString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
StrList2 = "#" & ";"
Do Until listTable.EOF
    StrList2 = StrList2 & "|#" & listTable.Fields(0) & ";" & listTable.Fields(1)
    listTable.MoveNext
Loop
End Function
Function StrListArray(aString) As String
If IsEmpty(aString) Then Exit Function
StrListArray = "#" & ";"
For i = 0 To UBound(aString)
    StrListArray = StrListArray & "|#" & retFlag(aString(i), "CODE") & ";" & retFlag(aString(i), "DESCA")
Next
End Function

Sub FilterGrd(pGrid, pString As String, Optional pCol As Integer = 1)
Dim aString
If Trim(pString) = "" Then
    For i = 1 To pGrid.rows - 2:         pGrid.RowHidden(i) = False:    Next
End If
aString = Split(Trim(pString))
For nRow = 1 To pGrid.rows - 1
    pGrid.RowHidden(nRow) = False
    For i = 0 To UBound(aString)
        If Trim(aString(i)) <> "" Then
            pGrid.RowHidden(nRow) = InStr(1, pGrid.TextMatrix(nRow, pCol), Trim(aString(i))) = 0
           If pGrid.RowHidden(nRow) = True Then Exit For
        End If
    Next
Next
End Sub
Function RetSetting(cSearch As String, Optional cFile As String) As String
Dim TextLine As String
On Error GoTo myerror
If cFile = "" Then cFile = App.Path & "\conf.txt"
Open cFile For Input As #1   ' Open file.
Do While Not EOF(1)   ' Loop until end of file.
   Line Input #1, TextLine   ' Read line into variable.
   If InStr(1, LCase(TextLine), LCase(cSearch) & "=") > 0 Then
       RetSetting = Mid(TextLine, Len(cSearch) + 2)
       Exit Do
   End If
Loop
Close #1   ' Close file.
Exit Function
myerror:
Err.Clear
RetSetting = ""
End Function
Function addSetting(cField, cValue, Optional cFile As String) As Boolean
Dim TextLine As String, cText As String, aLocal, nFoundTimes As Integer
Dim fs As New FileSystemObject
If cFile = "" Then cFile = App.Path & "\conf.txt"
On Error GoTo myerror
If fs.FileExists(cFile) Then
    Open cFile For Input As #1   ' Open file.
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
       cText = cText & turn(cText, vbCrLf) & TextLine
    Loop
    Close #1   ' Close file.
End If

aLocal = Split(cText, vbCrLf)
Open cFile For Output As #2   ' Open file.
For i = 0 To UBound(aLocal)
    nFound = InStr(1, LCase(Trim(aLocal(i))), LCase(Trim(cField)) & "=")
    If nFound > 0 Then
        nFoundTimes = nFoundTimes + 1
        If nFoundTimes = 1 Then
            If Trim(cValue) <> "" Then
                Print #2, UCase(Trim(cField)) & "=" & Trim(cValue)
            End If
        End If
    Else
        Print #2, aLocal(i)
    End If
Next
If nFoundTimes = 0 Then
    If Trim(cValue) <> "" Then
        Print #2, Trim(UCase(cField)) & "=" & Trim(cValue)
    End If
End If
Close #2
addSetting = True
Exit Function
myerror:
Err.Clear
End Function
Function turn(ByVal cString, Optional ByVal strFind, Optional ByVal caseFound, Optional ByVal CaseNotfound) As String
If IsMissing(strFind) And IsMissing(caseFound) And IsMissing(CaseNotfound) Then
    If Trim(cString) <> "" Then
        turn = IIf(InStr(1, LCase(cString), " where ") > 0, " AND ", " WHERE ")
    End If
ElseIf (Not IsMissing(strFind)) And IsMissing(caseFound) And IsMissing(CaseNotfound) Then
    If Trim(cString) <> "" Then turn = strFind
ElseIf (Not IsMissing(strFind)) And (Not IsMissing(caseFound)) And IsMissing(CaseNotfound) Then
    If Trim(strFind) <> "" Then
        turn = IIf(InStr(1, LCase(cString), LCase(strFind)) > 0, caseFound, strFind)
    Else
        turn = IIf(Trim(cString) = "", caseFound, strFind)
    End If
ElseIf (Not IsMissing(strFind)) And (Not IsMissing(caseFound)) And (Not IsMissing(CaseNotfound)) Then
    If Trim(strFind) <> "" Then
        turn = IIf(InStr(1, UCase(cString), UCase(strFind)) > 0, caseFound, CaseNotfound)
    Else
        turn = IIf(Trim(cString) = "", caseFound, CaseNotfound)
    End If
End If
End Function
Function ArbString(ByVal pString) As String
Dim aLocal As Variant
pString = Trim(pString & "")
If pString = "" Then Exit Function
aLocal = Split(pString & "")
For i = 0 To UBound(aLocal)
    If Trim(aLocal(i)) <> "" Then
        ArbString = ArbString & turn(ArbString, Chr(254) & " ") & Trim(aLocal(i))
    End If
Next
ArbString = Chr(254) & ArbString & Chr(254)
'ArbString = Trim(Chr(254) & Replace(pString & "", " ", Chr(254) & " " & Chr(254)))
End Function
Function ArbStr(ByVal pString) As String
'Dim aLocal
'pString = Trim(pString)
If Trim(pString & "") <> "" Then
    ArbStr = Chr(254) & pString & Chr(254)
End If
End Function
Function Myvalue(ByVal pValue As Variant, Optional pFormat As String = "", Optional nRound As Integer = 2) As String
Myvalue = IIf(mRound(pValue, nRound) = 0, "", mRound(pValue, nRound))
If pFormat <> "" Then Myvalue = Format(Myvalue, pFormat)
End Function
Function NumSql(cField As String) As String
NumSql = "case when " & cField & " = 0 then Null else " & cField & " end "
End Function
Function mySplit(pString, Optional nPos As Integer = 1, Optional pchr As String = ",") As String
Dim aString
aString = Split(pString, pchr)
If UBound(aString) >= nPos - 1 Then
    mySplit = aString(nPos - 1)
End If
End Function
Function CpuId() As String
On Error GoTo myerror
Dim computer As String
Dim wmi As Variant
Dim processors As Variant
Dim cpu As Variant
Dim cpu_ids As String

    computer = "."
    Set wmi = GetObject("winmgmts:" & _
        "{impersonationLevel=impersonate}!\\" & _
        computer & "\root\cimv2")
    Set processors = wmi.ExecQuery("Select * from " & _
        "Win32_Processor")

    For Each cpu In processors
        cpu_ids = cpu_ids & ", " & cpu.ProcessorId
    Next cpu
    If Len(cpu_ids) > 0 Then cpu_ids = Mid$(cpu_ids, 3)

    CpuId = cpu_ids
Exit Function
myerror:
End Function
Function RetMenu(cCode, cControl, pCon) As Variant
Dim obj As New ADODB.Recordset, aRet(2) As Variant
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc

cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, 15, cCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("control", adVarWChar, adParamInput, 50, cControl)
cmdTable.Parameters.Append cmdTable.CreateParameter("bVisible", adInteger, adParamOutput)
cmdTable.Parameters.Append cmdTable.CreateParameter("bEditable", adInteger, adParamOutput)
cmdTable.Parameters.Append cmdTable.CreateParameter("nSR", adInteger, adParamOutput)

cmdTable.CommandText = "retMenu"
Set obj = cmdTable.Execute
aRet(0) = cmdTable.Parameters(2).Value
aRet(1) = cmdTable.Parameters(3).Value
aRet(2) = cmdTable.Parameters(4).Value
Set obj = Nothing
RetMenu = aRet
End Function
Sub myGotFocus(ByRef pControl As Variant)
With pControl
.SelStart = 0
.SelLength = Len(pControl.text)
.BackColor = &HC0FFFF
End With
End Sub
Sub myLostFocus(ByRef pControl As Variant)
pControl.BackColor = &H80000005
End Sub
Sub myValidDate(ByRef pControl As Variant)
With pControl
If IsDate(.text) Then
    .text = myFormat_p(.text)
Else
    .text = ""
End If
End With
End Sub
Sub LoadText(myForm As Form, Optional bForAll As Boolean = False, Optional sFlag As String)
Dim cFileSave As String, cText As String
If bForAll Then cFileSave = App.Path & "\" & myForm.Name & turn(sFlag, "_") & sFlag & ".txt" Else cFileSave = tempPath & "\" & myForm.Name & turn(sFlag, "_") & sFlag & ".txt"
For i = 0 To myForm.Count - 1
    If TypeOf myForm(i) Is TextBox Then
        myForm(i).text = RetSetting(myForm(i).Name, cFileSave)
    ElseIf TypeOf myForm(i) Is DataCombo Then
        myForm(i).BoundText = RetSetting(myForm(i).Name, cFileSave)
    ElseIf TypeOf myForm(i) Is CheckBox Then
        If RetSetting(myForm(i).Name, cFileSave) <> "" Then myForm(i).Value = RetSetting(myForm(i).Name, cFileSave)
    ElseIf TypeOf myForm(i) Is OptionButton Then
'        If RetSetting(myForm(i).Name, cFilesave) <> "" Then myForm(i).Value = IIf(RetSetting(myForm(i).Name, cFilesave) = "TRUE", True, False)
    ElseIf TypeOf myForm(i) Is Label Then
        If LCase(myForm(i).Tag) = "t" Then
            If RetSetting(myForm(i).Name, cFileSave) <> "" Then myForm(i).Caption = RetSetting(myForm(i).Name, cFileSave)
        End If
    ElseIf TypeOf myForm(i) Is SSCommand Then
        If RetSetting(myForm(i).Name, cFileSave) <> "" Then myForm(i).Tag = RetSetting(myForm(i).Name, cFileSave)
    End If
Next
End Sub
Sub DefineText(myForm As Form)
Dim cFileSave As String, cText As String
If bForAll Then cFileSave = App.Path & "\" & myForm.Name & ".txt" Else cFileSave = cTempDir & "\" & myForm.Name & ".txt"
For i = 0 To myForm.Count - 1
    If TypeOf myForm(i) Is TextBox Then
        myForm(i).text = ""
    ElseIf TypeOf myForm(i) Is DataCombo Then
        myForm(i).BoundText = ""
    ElseIf TypeOf myForm(i) Is CheckBox Then
        myForm(i).Value = 0
    ElseIf TypeOf myForm(i) Is Label Then
        If LCase(myForm(i).Tag) = "t" Then
            myForm(i).Caption = ""
        End If
    End If
Next
End Sub
Sub SaveText(myForm As Form, Optional bForAll As Boolean = False, Optional aText As Variant, Optional sFlag As String = "")
Dim cFileSave As String
If bForAll Then cFileSave = App.Path & "\" & myForm.Name & turn(sFlag, "_") & sFlag & ".txt" Else cFileSave = tempPath & "\" & myForm.Name & turn(sFlag, "_") & sFlag & ".txt"
If Not IsMissing(aText) Then
    For i = 0 To UBound(aText)
        cText = cText & turn(cText, "", "@") & LCase(aText(i)) & "@"
    Next
End If
For i = 0 To myForm.Count - 1
    If InStr(1, cText, "@" & LCase(myForm(i).Name) & "@") > 0 Or IsMissing(aText) Then
        If TypeOf myForm(i) Is TextBox Then
            addSetting myForm(i).Name, myForm(i).text, cFileSave
        ElseIf TypeOf myForm(i) Is DataCombo Then
            addSetting myForm(i).Name, myForm(i).BoundText, cFileSave
        ElseIf TypeOf myForm(i) Is CheckBox Then
            addSetting myForm(i).Name, myForm(i).Value, cFileSave
        ElseIf TypeOf myForm(i) Is OptionButton Then
 '           addSetting myForm(i).Name, IIf(myForm(i).Value, "TRUE", "FALSE"), cFilesave
        ElseIf TypeOf myForm(i) Is Label Then
            If LCase(myForm(i).Tag) = "t" Then
                addSetting myForm(i).Name, myForm(i).Caption, cFileSave
            End If
        ElseIf TypeOf myForm(i) Is SSCommand And (Not IsMissing(aText)) Then
            If Not IsEmpty(aText) Then
                addSetting myForm(i).Name, myForm(i).Tag, cFileSave
            End If
        End If
    End If
Next
End Sub
Function TempSave(pform As Variant, Optional sFlag As String = "", Optional sExt As String = "txt") As String
TempSave = tempPath & turn(tempPath, "\") & pform.Name & turn(sFlag, "_" & sFlag) & "." & sExt
End Function
Function noReadOnly(sFilePath As String) As String
Dim fs As New FileSystemObject
Dim fil As File
On Error GoTo myerror
If Not fs.FileExists(sFilePath) Then
    noReadOnly = "notfound"
    Exit Function
End If
Set fil = fs.GetFile(sFilePath)
If (fil.Attributes And ReadOnly) Then
  fil.Attributes = fil.Attributes - ReadOnly
End If
noReadOnly = "ok"
Set fs = Nothing
Exit Function
myerror:
noReadOnly = Err.Description
Err.Clear
End Function
Public Sub ToFileExelOld(MyGrid, Optional aIg As Variant, Optional nRowHead As Long = 0, Optional aRowMerge As Variant = Empty, Optional aCol As Variant = Empty, Optional nRate As Double = 1, Optional aWidth As Variant = Empty, Optional arowHeight As Variant = Empty, Optional aSetUp As Variant = Empty, Optional nSize As Integer = 12, Optional acolSplit As Variant = Empty, Optional myForm As Form, Optional nTopMargin As Integer = 30)
    Dim irow As Integer, i As Long, i2 As Long, nCols As Long
    Dim icol As Integer
    Dim objExcl As Excel.Application
    Dim objWk As Excel.Workbook
    Dim objSht As Excel.Worksheet
    Dim iHead As Integer
    Dim vHead As Variant
    On Error Resume Next
    Set objExcl = Excel.Application
    objExcl.Application.Visible = False
    Set objWk = objExcl.Workbooks.Add
    Set objSht = objWk.Sheets(1)
    objExcl.Application.DisplayAlerts = False
    Dim nRows As Long, nFixed As Long
        
    objSht.PageSetup.TopMargin = nTopMargin
    objSht.PageSetup.LeftMargin = 10
    objSht.PageSetup.HeaderMargin = nTopMargin
    objSht.PageSetup.CenterHeader = "&B &" & nSize
    
    objSht.Cells.NumberFormat = "@"
    
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If
    
    For irow = 0 To MyGrid.rows - 1
        If Not MyGrid.RowHidden(irow) Then
            nRows = nRows + 1
            nCols = 0
            If Not myForm Is Nothing Then
                If myForm.prog1.Value <> Int((irow / MyGrid.rows - 1) * 100) Then
                    'myForm.prog1.Value = IIf(Int(irow / (MyGrid.rows - 1) * 100) > 100, 100, Int(I / (MyGrid.rows - 1) * 100))
                    myForm.prog1.Value = IIf(Round(irow / (MyGrid.rows), 2) > 1, 1, Round(irow / (MyGrid.rows), 2)) * 100
                End If
            End If
            For icol = 0 To MyGrid.Cols - 1
                If Not MyGrid.ColHidden(icol) Then
                    nCols = nCols + 1
                    If MyGrid.ColDataType(icol) = flexDTDate And irow > MyGrid.FixedRows - 1 Then
                        objSht.Cells(nRows, nCols) = myFormat_p(MyGrid.Cell(flexcpTextDisplay, irow, icol))
                    ElseIf MyGrid.ColDataType(icol) = flexDTBoolean And irow > MyGrid.FixedRows - 1 Then
                        objSht.Cells(nRows, nCols) = IIf(Val(MyGrid.TextMatrix(irow, icol)) = 0, "·«", "‰⁄„")
                    ElseIf MyGrid.ColDataType(nCols) = flexDTDouble And irow > MyGrid.FixedRows - 1 Then
                        objSht.Cells.NumberFormat = ""
                    Else
                        objSht.Cells(nRows, nCols) = MyGrid.Cell(flexcpTextDisplay, irow, icol) & ""
                    End If
                End If
            Next icol
        End If
    Next irow

    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If

    nFixed = 0
    For i = 0 To MyGrid.FixedRows - 1
        If Not MyGrid.RowHidden(i) Then
            nFixed = nFixed + 1
        End If
    Next
                        
    nFixedCols = 0
    For i = 0 To MyGrid.FixedCols - 1
        If Not MyGrid.ColHidden(i) Then
            nFixedCols = nFixedCols + 1
        End If
    Next
                    
            
    Dim nRow2 As Long
    If Not IsEmpty(aCol) Then
        For nCol = 0 To UBound(aCol)
            nValue = 0
            For nRow2 = 1 To nRows
                If Trim(objSht.Cells(nRow2, aCol(nCol))) <> Trim(cValue & "") Then
                    If nValue <> 0 Then
                        objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
                    End If
                    cValue = Trim(objSht.Cells(nRow2, aCol(nCol)))
                    nValue = 0
                    nBegin = nRow2
                Else
                    nValue = nValue + 1
                End If
            Next
            If nValue <> 0 Then
                objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
            End If
        Next
    End If
 
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixed, nCol)).Interior.ColorIndex = 40
'        objSht.Range(objSht.Cells(nFixed, 1), objSht.Cells(nFixed, nCol)).Font.bold = True
    
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Borders.ColorIndex = 0
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).AutoFit = True
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).VerticalAlignment = xlCenter
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedrows + 1, nCols)).HorizontalAlignment = xlCenter
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedrows + 1, nCols)).Interior.ColorIndex = 40
 
 
    If Not IsEmpty(aRowMerge) Then
        For i = 0 To UBound(aRowMerge)
            If Not IsEmpty(retFlag(aRowMerge(i), "cols")) Then
                If Not IsEmpty(retFlag(aRowMerge(i), "text")) Then
                    objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))) = retFlag(aRowMerge(i), "text")
                End If
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))).Merge
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = 19
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Borders.ColorIndex = 0
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Size = nSize
            End If
            
            If retFlag(aRowMerge(i), "split") Then
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).PageBreak = xlPageBreakManual
            End If
            
            If retFlag(aRowMerge(i), "word_wrap") Then
                'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).WrapText = True
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).WrapText = True
            End If
            If Not IsEmpty(retFlag(aRowMerge(i), "height")) Then
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).RowHeight = retFlag(arowHeight(i), "height")
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).RowHeight = retFlag(aRowMerge(i), "height")
            End If
             If Not IsEmpty(retFlag(aRowMerge(i), "back_color")) Then
               objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = retFlag(aRowMerge(i), "back_color")
            End If
            
            If retFlag(aRowMerge(i), "bold") Then
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).Font.Bold = True
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
            End If
        Next
    End If

    If Not IsEmpty(acolSplit) Then
        For i = 0 To UBound(acolSplit)
            objSht.Columns(retFlag(acolSplit(i), "col")).PageBreak = xlPageBreakManual
        Next
    End If


'    If Not IsEmpty(arowHeight) Then
'        For i = 0 To UBound(arowHeight)
'            objSht.Range(objSht.Cells(retFlag(arowHeight(i), "row") + 1, 1), objSht.Cells(retFlag(arowHeight(i), "row") + 1, NCOLS)).RowHeight = retFlag(arowHeight(i), "height")
'            If retFlag(arowHeight(i), "word_wrap") Then objSht.Range(objSht.Cells(retFlag(arowHeight(i), "row") + 1, 1), objSht.Cells(retFlag(arowHeight(i), "row") + 1, NCOLS)).WrapText = True
'            If retFlag(arowHeight(i), "bold") Then objSht.Range(objSht.Cells(retFlag(arowHeight(i), "row") + 1, 1), objSht.Cells(retFlag(arowHeight(i), "row") + 1, NCOLS)).WrapText = True
'        Next
'    End If

    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).VerticalAlignment = xlCenter
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixed, nCols)).HorizontalAlignment = xlCenter
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixed, nCols)).Interior.ColorIndex = 40
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Borders.ColorIndex = 0
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Size = nSize
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Bold = True
    
    nCol = 0
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCol = nCol + 1
            If MyGrid.ColFormat(icol) = "(##,##.##" Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).NumberFormat = "_(#,###.00_);[Red](#,###.00);0.00"
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
            ElseIf MyGrid.ColAlignment(icol) = flexAlignLeftCenter Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
            ElseIf MyGrid.ColAlignment(icol) = flexAlignRightCenter Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
            ElseIf MyGrid.ColAlignment(icol) = flexAlignCenterCenter Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlCenter
            Else
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
            End If
            If nRate > 0 Then objSht.Columns(nCol).ColumnWidth = (MyGrid.ColWidth(icol) / 100) * nRate
        End If
    Next icol
        
    If Not IsEmpty(aWidth) Then
        For i = 0 To UBound(aWidth)
            If Val(retFlag(aWidth(i), "width")) = 0 Then
                objSht.Range(objSht.Cells(1, retFlag(aWidth(i), "col")), objSht.Cells(nRows, retFlag(aWidth(i), "col"))).Columns.AutoFit
            Else
                objSht.Columns(retFlag(aWidth(i), "col")).ColumnWidth = Val(retFlag(aWidth(i), "width")) / 100
            End If
        Next
    End If
    
    objSht.PageSetup.Orientation = xlLandscape
    If Not IsEmpty(aSetUp) Then
        
        If Not IsEmpty(retFlag(aSetUp, "title_col")) Then
            objSht.PageSetup.PrintTitleColumns = objSht.Columns(retFlag(aSetUp, "title_col")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "title_row")) Then
            objSht.PageSetup.PrintTitleRows = objSht.rows(retFlag(aSetUp, "title_row")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "freeze")) Then
            ActiveWindow.SplitColumn = retFlag(aSetUp, "freeze")
            ActiveWindow.SplitRow = 0
            ActiveWindow.FreezePanes = True
        End If
        
        If retFlag(aSetUp, "autoSize") Then
            objSht.PageSetup.FitToPagesWide = True
        End If
                
        If retFlag(aSetUp, "center_header") <> "" Then
            objSht.PageSetup.CenterHeader = "&" & nSize & "&B" & retFlag(aSetUp, "center_header")
            objSht.PageSetup.CenterHeader = "&" & nSize & "&B" & retFlag(aSetUp, "center_header")
            objSht.PageSetup.CenterHeader = "&" & nSize & "&B" & retFlag(aSetUp, "center_header")
        End If
    End If
    
    'objSht.Range(objSht.Cells(0, 1), objSht.Cells(0, NCOLS)).WrapText = True
    
    'objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Columns.AutoFit
    If Not IsMissing(aIg) Then
        For i = 0 To UBound(aIg)
            objSht.rows(aIg(i)).Hidden = True
        Next
    End If
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If
    objExcl.Application.Visible = True
'    If Not IsEmpty(aRowMerge) Then
'        For i = 0 To UBound(aRowMerge)
'            objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols") + 1)).Merge
'        Next
'    End If

'''''''''''''''''''''''''''''
'›Ì Õ«·  —Ìœ Õ›Ÿ Ê—ﬁ… «·√ﬂ”·
'objWk.SaveAs "c:\Book1.xls"
'objWk.Close
'«·”ÿ— «· «·Ì Ì€·ﬁ »—‰«„Ã «·√ﬂ”·
'objExcl.Quit
''''''''''''''''''''''''''''
Set objSht = Nothing
Set objWk = Nothing
Set objExcl = Nothing
If Not myForm Is Nothing Then myForm.prog1.Visible = False
End Sub
Function myRecordSet(cString As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim rdTable As New ADODB.Recordset
rdTable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Set myRecordSet = rdTable
Set rdTable = Nothing
End Function
Function ReplaceStr(TextIn, ByVal SearchStr As String, _
                        ByVal replacement As String, _
                        ByVal CompMode As Integer)
   Dim WorkText As String, Pointer As Integer
     If IsNull(TextIn) Then
       ReplaceStr = Null
     Else
       WorkText = TextIn
       Pointer = InStr(1, WorkText, SearchStr, CompMode)
       Do While Pointer > 0
         WorkText = Left(WorkText, Pointer - 1) & replacement & _
                    Mid(WorkText, Pointer + Len(SearchStr))
         Pointer = InStr(Pointer + Len(replacement), WorkText, _
                         SearchStr, CompMode)
       Loop
       ReplaceStr = WorkText
     End If
   End Function
   Function SQLFixup(TextIn)
     SQLFixup = ReplaceStr(TextIn, "'", "''", 0)
   End Function

   Function JetSQLFixup(TextIn)
   Dim temp
     temp = ReplaceStr(TextIn, "'", "''", 0)
     JetSQLFixup = ReplaceStr(temp, "|", "' & chr(124) & '", 0)
   End Function
   Function FindFirstFixup(TextIn)
   Dim temp
     temp = ReplaceStr(TextIn, "'", "' & chr(39) & '", 0)
     FindFirstFixup = ReplaceStr(temp, "|", "' & chr(124) & '", 0)
   End Function
Function RetPrinter(pName) As Variant
Dim printer As printer, aRet As Variant
For Each printer In Printers
    If LCase(Trim(printer.DeviceName)) = LCase(Trim(pName)) Then
        aRet = AddFlag(aRet, "name", pName)
        aRet = AddFlag(aRet, "port", printer.Port)
        aRet = AddFlag(aRet, "driver", printer.DriverName)
        RetPrinter = aRet
        Exit For
    End If
Next
End Function
Sub FixPrinter(pReport As CrystalReport, Optional pType As String = "1")
Dim aRet As Variant, cPrinterName, cPort As String, cDriver As String
cPrinterName = RetSetting("printer" & pType, tempPath & turn(tempPath, "\") & "printers.txt")
If cPrinterName <> "" Then
    aRet = RetPrinter(cPrinterName)
    If Not IsEmpty(aRet) Then
        pReport.PrinterDriver = retFlag(aRet, "driver")
        pReport.PrinterPort = retFlag(aRet, "port")
        pReport.PrinterName = retFlag(aRet, "name")
    End If
End If
End Sub
Function RetPrinterByType(Optional sType As String = "1") As String
Dim sPrinter As String, aRet As Variant
sPrinter = RetSetting("printer" & sType, tempPath & turn(tempPath, "\") & "printers.txt")
aRet = RetPrinter(sPrinter)
If Not IsEmpty(aRet) Then RetPrinterByType = sPrinter
End Function
Sub FixRpImage(myForm As Form)
On Error Resume Next
With myForm
.CmdApply.Picture = LoadPicture(App.Path & "\sys_img\preview.jpg")
.cmdExit.Picture = LoadPicture(App.Path & "\sys_img\exit.jpg")
.cmdClear.Picture = LoadPicture(App.Path & "\sys_img\clear.jpg")
Err.Clear
If .CmdApply.Picture = 0 Then .CmdApply.Caption = "⁄—÷ «· ﬁ—Ì—"
If .cmdExit.Picture = 0 Then .cmdExit.Caption = "Œ—ÊÃ"
If .cmdClear.Picture = 0 Then .cmdClear.Caption = "„”Õ"
End With
End Sub
Function AddFlag(ByVal aString As Variant, ByVal cFlag As Variant, Optional ByVal cFlagValue As Variant, Optional bedit As Boolean = False) As Variant
If IsEmpty(aString) Then aString = Array()
If IsMissing(cFlagValue) Then
    ReDim Preserve aString(UBound(aString) + 1)
    aString(UBound(aString)) = cFlag
    AddFlag = aString
    Exit Function
ElseIf bedit Then
    If UBound(aString) > 0 Then
        Dim i As Long
        For i = 0 To UBound(aString) Step 2
            If Trim(LCase(aString(i))) = Trim(LCase(cFlag)) Then
                aString(i + 1) = cFlagValue
                AddFlag = aString
                Exit Function
            End If
        Next
    End If
End If
ReDim Preserve aString(UBound(aString) + 2)
aString(UBound(aString) - 1) = cFlag
aString(UBound(aString)) = cFlagValue
AddFlag = aString
End Function
Function retFlag(aString As Variant, cFlag As String, Optional pEmptyValue As Variant = Empty) As Variant
If IsEmpty(aString) Then Exit Function
Dim nPos As Long
If UBound(aString) > 0 Then
    For nPos = 0 To UBound(aString) Step 2
        If Trim(LCase(aString(nPos))) = Trim(LCase(cFlag)) Then
            retFlag = aString(nPos + 1)
            Exit For
        End If
    Next
End If
If IsEmpty(retFlag) Then
    retFlag = pEmptyValue
ElseIf IsNull(retFlag) And (Not IsEmpty(pEmptyValue)) Then
    retFlag = pEmptyValue
End If
End Function
Function addInsert(ByVal aInsert, ByVal cTable, Optional ByVal pWhere As String) As String
Dim cString1 As String, cString2 As String, i As Long
For i = 0 To UBound(aInsert) Step 2
    cString1 = cString1 & IIf(cString1 = "", "", ",") & aInsert(i)
    cString2 = cString2 & IIf(cString2 = "", "", ",") & aInsert(i + 1)
Next
addInsert = "Insert into " & cTable & " (" & _
                cString1 & _
                ")"
If pWhere = "" Then
    addInsert = addInsert & " values(" & _
                    cString2 & _
                    ")"
ElseIf pWhere <> "" Then
    addInsert = addInsert & " Select " & _
                    cString2
    addInsert = addInsert & " WHERE " & pWhere
End If
End Function
Function addInsertTb(ByVal aInsert, ByVal sTable, ByVal sWhere As String) As String
Dim cString1 As String, cString2 As String, i As Long
For i = 0 To UBound(aInsert) Step 2
    cString1 = cString1 & IIf(cString1 = "", "", ",") & aInsert(i)
    cString2 = cString2 & IIf(cString2 = "", "", ",") & aInsert(i + 1)
Next

addInsertTb = "Insert into " & sTable & " (" & _
                cString1 & _
                ")"

addInsertTb = addInsertTb & " SELECT " & _
                cString2 & _
                turn(sWhere, " WHERE ") & sWhere
End Function
Function addInsertTable(ByVal aInsert, ByVal sTable_insert As String, sTable_select As String, sCondition As String) As String
Dim cString1 As String, cString2 As String
For i = 0 To UBound(aInsert) Step 2
    cString1 = cString1 & IIf(cString1 = "", "", ",") & aInsert(i)
    cString2 = cString2 & IIf(cString2 = "", "", ",") & aInsert(i + 1)
Next
addInsertTable = "Insert into " & sTable_insert & " (" & _
                cString1 & _
                ")"

addInsertTable = addInsertTable & " Select " & _
                cString2 & _
                " from " & sTable_select & turn(sCondition, " WHERE ") & sCondition
End Function
Function addUpdate(ByVal aInsert, ByVal cTable, ByVal cCondition) As String
Dim i As Long
For i = 0 To UBound(aInsert) Step 2
    addUpdate = addUpdate & turn(addUpdate, ",") & aInsert(i) & _
                   " = " & aInsert(i + 1)
Next
addUpdate = "UPDATE " & cTable & " SET " & _
               addUpdate
If cCondition <> "" Then addUpdate = addUpdate & turn(addUpdate) & cCondition
End Function
Function GetFields(pString, Optional pCon As ADODB.Connection) As Variant
Dim loctable As New ADODB.Recordset
If pCon Is Nothing Then
    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If
If Not (loctable.BOF And loctable.EOF) Then
   Dim i As Long
    For i = 0 To loctable.Fields.Count - 1
        GetFields = AddFlag(GetFields, LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
    Next
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function GetField(pString, Optional pCon As ADODB.Connection, Optional pEmptyValue As Variant = Empty) As Variant
Dim loctable As New ADODB.Recordset
If pCon Is Nothing Then
    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If
If Not (loctable.BOF And loctable.EOF) Then
    GetField = loctable(0).Value
Else
    GetField = pEmptyValue
End If
loctable.Close
Set loctable = Nothing
End Function
Function GetRows(pString, Optional pCon As ADODB.Connection) As Variant
Dim loctable As New ADODB.Recordset
If pCon Is Nothing Then
    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If

If Not (loctable.BOF And loctable.EOF) Then
    Dim aRet, i As Long
    aRet = Array()
    Do Until loctable.EOF
        ReDim Preserve aRet(UBound(aRet) + 1)
        For i = 0 To loctable.Fields.Count - 1
            aRet(UBound(aRet)) = AddFlag(aRet(UBound(aRet)), LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
        Next
        loctable.MoveNext
    Loop
    GetRows = aRet
End If
loctable.Close
Set loctable = Nothing
End Function
Function GetRowsNew(pString As String, Optional pCon As ADODB.Connection) As Variant
Dim loctable As ADODB.Recordset
Dim aRet As Variant
If pCon Is Nothing Then
    Set loctable = myCmd(pString, GetCon, adText, , 300)
Else
    Set loctable = myCmd(pString, pCon, adText, , 300)
End If

If Not (loctable.BOF And loctable.EOF) Then
    Dim i As Long
    aRet = Array()
    Do Until loctable.EOF
        ReDim Preserve aRet(UBound(aRet) + 1)
        For i = 0 To loctable.Fields.Count - 1
            aRet(UBound(aRet)) = AddFlag(aRet(UBound(aRet)), LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
        Next
        loctable.MoveNext
    Loop
    GetRowsNew = aRet
End If
loctable.Close
Set loctable = Nothing
End Function

Function ItemFields(ByVal pItem As String, pCon As ADODB.Connection) As Variant
If Not validItemCode(pItem) Then Exit Function
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 20, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
If Not (rdTable.EOF And rdTable.BOF) Then
    Dim i As Long
    For i = 0 To rdTable.Fields.Count - 1
        ItemFields = AddFlag(ItemFields, rdTable.Fields(i).Name, rdTable.Fields(i).Value)
    Next
End If
Set rdTable = Nothing
End Function
Function ItemField(pItem As String, pField As String, pCon As ADODB.Connection) As Variant
If Not validItemCode(pItem) Then Exit Function
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 20, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
If Not (rdTable.EOF And rdTable.BOF) Then ItemField = rdTable(pField)
Set rdTable = Nothing
End Function
Function ValidInt(pString As Variant, Optional pNoFirstZero As Boolean = False) As Boolean
Dim i As Long, nChr As String
If Trim(pString) = "" Then Exit Function
For i = 1 To Len(Trim(pString))
    nChr = Mid(pString, i, 1)
    If Not IsNumeric(nChr) Then
        Exit Function
    ElseIf pNoFirstZero And i = 1 Then
        If nChr = "0" Then Exit Function
    End If
Next
ValidInt = True
'ValidInt = Int(Abs(Val(pString & ""))) & "" = Trim(pString)
End Function
Function validItemCode(pCode As String) As Boolean
If Trim(pCode) = "" Then Exit Function
validItemCode = True
End Function
Function validItem(pItem As String, pCon As ADODB.Connection, Optional sType As String = "1") As Boolean
On Error GoTo myerror
Dim aRet As Variant
aRet = ItemFields(pItem, pCon)
If IsEmpty(aRet) Then Exit Function
If sType <> "" Then
    If retFlag(aRet, "TYPE") <> sType Then Exit Function
End If
validItem = True
'Dim rdTable As New ADODB.Recordset
'Dim cmdTable As New ADODB.Command
'Set cmdTable.ActiveConnection = pcon
'cmdTable.CommandType = adCmdStoredProc
'cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarChar, adParamInput, 6, pItem)
'cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adBoolean, adParamOutput)
'cmdTable.CommandText = "validItem"
'Set rdTable = cmdTable.Execute
'validItem = cmdTable.Parameters(1).Value
'Set rdTable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function

Public Sub UnloadAllForms(Optional FormToUnload _
As String = "")
Dim f As Form
For Each f In Forms
  If Trim(UCase(f.Name)) = Trim(UCase(FormToUnload)) Then
    Unload f
    Set f = Nothing
  End If
Next f
End Sub
Function RetIndex(pOption As Variant)
For i = 0 To pOption.Count - 1
    If pOption(i).Value Then
        RetIndex = i
        Exit Function
    End If
Next
End Function
Function IsTime(ByVal sTime As String) As Boolean
If Not (Len(IsTime) >= 1 And Len(IsTime) <= 5) Then Exit Function
aRet = Split(sTime, ":")
If UBound(aRet) <> 1 Then Exit Function
For i = 0 To 1
    nLen = Len(aRet(i))
    If Not (nLen >= 1 And nLen <= 2) Then Exit Function
    For i2 = 1 To nLen
        If Not IsNumeric(Mid(aRet(i), i2, 1)) Then Exit Function
    Next
Next
If Val(aRet(0)) > 23 Or Val(aRet(1)) > 59 Then Exit Function
IsTime = True
End Function
Function RetTime(ByVal sTime) As String
If sTime = "" Then Exit Function
If IsTime(sTime) Then
    aRet = Split(sTime, ":")
    For i = 0 To UBound(aRet)
        sTime = RetZero(aRet(0), 2) + ":" & RetZero(aRet(1), 2)
    Next
End If

RetTime = sTime
If Len(sTime) = 3 And Right(sTime, 1) = ":" Then sTime = Left(sTime, 2)

For i = 1 To Len(sTime)
    If Not IsNumeric(sTime) Then Exit Function
    If i > 2 Then Exit Function
Next
If Val(sTime) > 23 Then Exit Function
RetTime = RetZero(sTime, 2) + ":" & "00"
End Function
Sub myUnLoad(pform As Form)
On Error Resume Next
Unload pform
Set pform = Nothing
Err.Clear
End Sub
Function MinuteToTime(nMinutes) As String
If Val(nMinutes) = 0 Then Exit Function
nMinutes = Abs(nMinutes)
If Fix(nMinutes / 60) > 0 Then
   MinuteToTime = Fix(nMinutes / 60) & " ”«⁄… "
End If

If nMinutes Mod 60 > 0 Then
   MinuteToTime = MinuteToTime & IIf(Fix(nMinutes / 60) > 0, " Ê ", "") & nMinutes Mod 60 & " œﬁÌﬁ… "
End If
End Function
Function MinuteToTimeDg(nMinutes) As String
If nMinutes = 0 Then Exit Function
nMinutes = Abs(nMinutes)
If Fix(nMinutes / 60) > 0 Then
   MinuteToTimeDg = RetZero(Fix(nMinutes / 60), 2)
Else
   MinuteToTimeDg = "00"
End If
MinuteToTimeDg = MinuteToTimeDg & ":" & RetZero(nMinutes Mod 60, 2)
End Function
Function retDiffMinutes(ByVal sDate1 As String, ByVal sdate2 As String, Optional bMintues As Boolean = False) As Double
If bMintues Then
    If TimeValue(RetTime(sdate2)) >= TimeValue(RetTime(sDate1)) Then
        sdate2 = Format("16-04-1971", "yyyy-mm-dd") & " " & sdate2
    Else
        sdate2 = Format("17-04-1971", "yyyy-mm-dd") & " " & sdate2
    End If
    sDate1 = Format("16-04-1971", "yyyy-mm-dd") & " " & sDate1
End If
retDiffMinutes = DateDiff("n", sDate1, sdate2)
End Function
Function myFormat(sDate As Variant, Optional bLong As Boolean = False) As String
myFormat = Format(sDate, "YYYY-MM-DD") & IIf(bLong, Format(sDate & "", " HH:NN"), "")
End Function
Function myFormat_sp(sDate As Variant, Optional bLong As Boolean = False) As Variant
myFormat_sp = TurnValue(myFormat(sDate, bLong))
End Function
Public Function myFormat_l(sDate As Variant) As String
myFormat_l = Format(sDate & "", "YYYY-MM-DD HH:NN")
End Function
Function retPhoto(ByVal cMember As String, Optional nMax As Integer = 10000, Optional nMaxDir As Integer = 1, Optional pFolder As String = "") As String
Dim cString, cFileName, cFileDir

If cMember = "" Then Exit Function

nPos = InStr(1, cMember, "-")
If nPos <> 0 Then mainmember = Mid(cMember, 1, nPos - 1) Else mainmember = cMember

nDir = Val(mainmember) / nMax
If nDir <> Fix(nDir) Then nDir = Fix(nDir) + 1

If nDir = 0 Then nDir = 1
If nDir > nMaxDir Then nDir = nMaxDir
cDir = "\photo" & nDir & "\"

If bCrypt Then
    cFileName = MyCodeString(cMember, nDir)
Else
    cFileName = cMember & ".jpg"
End If

If cFileName <> "" Then
    If pFolder = "" Then
        retPhoto = sPath_App & cDir & cFileName
    Else
        retPhoto = pFolder & cDir & cFileName
    End If
End If
End Function
Function RetAppendPhoto(ByVal cMember As String, ByVal cAppend) As String
If Not (ValidInt(cMember) And ValidInt(cAppend)) Then Exit Function
RetAppendPhoto = retPhoto(cMember & "-" & cAppend)
End Function
Function MemPhoto(ByVal cMember As String, Optional ByVal cAppend As String = "") As String
If (Not ValidNum(cMember)) Then Exit Function
If (Not ValidNum(cAppend)) And Trim(cAppend) <> "" Then Exit Function
If ValidNum(cAppend) Then
    MemPhoto = RetAppendPhoto(cMember, cAppend)
Else
    MemPhoto = retPhoto(cMember)
End If
End Function
Function RetPhotoNew(ByVal cMember As String, Optional nMax As Integer = 10000, Optional nMaxDir As Integer = 3, Optional pPath As String) As String
Dim cString, cFileName, cFileDir

If cMember = "" Then Exit Function

nPos = InStr(1, cMember, "-")
If nPos <> 0 Then mainmember = Mid(cMember, 1, nPos - 1) Else mainmember = cMember

nDir = Val(mainmember) / nMax
If nDir <> Fix(nDir) Then nDir = Fix(nDir) + 1

If nDir = 0 Then nDir = 1
If nDir > nMaxDir Then nDir = nMaxDir
cDir = "\photo" & nDir & "\"

If Not bCrypt Then
    cFileName = cMember & ".jpg"
    If cFileName = "" Then Exit Function
    RetPhotoNew = pPath & "\photo\" & cFileName
Else
    cFileName = MyCodeString(cMember, nDir)
    If cFileName = "" Then Exit Function
    RetPhotoNew = pPath & cDir & cFileName
End If
End Function
Function MyCodeString(nCode, Optional nMethod = 1) As String
Select Case nMethod
Case 1
   nVal1 = 74: nVal2 = 11: nValSub1 = 71: nValSub2 = 4
Case 2
   nVal1 = 71: nVal2 = 12: nValSub1 = 65: nValSub2 = 5
Case 3
   nVal1 = 69: nVal2 = 13: nValSub1 = 73: nValSub2 = 6
End Select

nPos = InStr(1, nCode, "-")

If nPos > 0 Then
    nNumber1 = Left(nCode, nPos - 1)
    nNumber2 = Mid(nCode, nPos + 1)
    If Not myValidInt(nNumber1) Then Exit Function
    If Not myValidInt(nNumber2) Then Exit Function
Else
    nNumber1 = nCode
    If Not myValidInt(nNumber1) Then Exit Function
    nNumber2 = 0
End If

mycode1 = Val(nNumber1) + Val(nVal1)
mycode1 = mycode1 & "2"
mycode1 = mycode1 / 2
mycode1 = mycode1 - nVal2
mycode1 = StrReverse(mycode1)

nNumber2 = Val(nNumber2) + Val(Right(nNumber1, 1))
myCode2 = Val(nNumber2) + Val(nValSub1)
myCode2 = myCode2 & "2"
myCode2 = myCode2 / 2
myCode2 = myCode2 - nValSub2
MyCodeString = mycode1 & "." & myCode2
End Function
Function myPhoto_Path(ByVal cMember As String, Optional pPath As String) As String
Dim aSplit As Variant
If cMember = "" Then Exit Function
aSplit = Split(Trim(cMember), "-")
If UBound(aSplit) > 1 Then Exit Function
If Not ValidNum(aSplit(0)) Then Exit Function
If UBound(aSplit) = 1 Then
    If Not ValidNum(aSplit(1)) Then Exit Function
End If
myPhoto_Path = Trim(aSplit(0))
If UBound(aSplit) = 1 Then myPhoto_Path = myPhoto_Path & "-" & Trim(aSplit(1))
myPhoto_Path = pPath & "\" & myPhoto_Path & ".jpg"
End Function
Function myPhoto(ByVal cMember As String, Optional pFlag As String) As String
Dim aSplit As Variant
If cMember = "" Then Exit Function
aSplit = Split(Trim(cMember), "-")
If UBound(aSplit) > 1 Then Exit Function
If Not ValidNum(aSplit(0)) Then Exit Function
If UBound(aSplit) = 1 Then
    If Not ValidNum(aSplit(1)) Then Exit Function
End If
myPhoto = Trim(aSplit(0))
If UBound(aSplit) = 1 Then myPhoto = myPhoto & "-" & Trim(aSplit(1))
myPhoto = sPath_App & "\PHOTO" & turn(pFlag, "_") & pFlag & "\" & myPhoto & ".jpg"
End Function
Function MemPhoto_I(ByVal cMember As String, Optional ByVal cAppend As String = "") As String
If (Not ValidNum(cMember)) Then Exit Function
If (Not ValidNum(cAppend)) And Trim(cAppend) <> "" Then Exit Function
If ValidNum(cAppend) Then
   MemPhoto_I = RetAppendPhoto_i(cMember, cAppend)
Else
    MemPhoto_I = RetPhoto_I(cMember)
End If
End Function
Function validPhoto(sPhoto As String) As Boolean
On Error GoTo myerror
If sPhoto = "" Then Exit Function
If Dir(sPhoto) = "" Then Exit Function
'validPhoto = LoadPicture(sPhoto)
validPhoto = True
Exit Function
myerror:
Err.Clear
End Function
Function RetPhoto_I(ByVal cMember As String, Optional pFolder As String = "") As String
Dim cMainMember, nDir As Double, acode As Variant
If pFolder = "" Then
    RetPhoto_I = sPath_App & "\photo_i\" & cMember & ".jpg"
Else
    RetPhoto_I = pFolder & "\photo_i\" & cMember & ".jpg"
End If
End Function
Function RetAppendPhoto_i(ByVal cMember As String, ByVal cAppend) As String
If Not (ValidNum(cMember) And ValidNum(cAppend)) Then Exit Function
RetAppendPhoto_i = sPath_App & "\photo_i\" & cMember & "-" & cAppend & ".jpg"
End Function
Function RetPhotoh(ByVal cMember As String, Optional pFolder As String = "") As String
If pFolder = "" Then
    RetPhotoh = sPath_App & "\photo_H\" & cMember & ".jpg"
Else
    RetPhotoh = pFolder & "\photo_H\" & cMember & ".jpg"
End If
End Function
Function RetPhotow(ByVal cMember As String) As String
Dim cMainMember, nDir As Double, acode As Variant
If Not ValidInt(cMember) Then Exit Function
RetPhotow = sPath_App & "\photo_w\" & cMember & ".jpg"
End Function
Function RetPhoto_s(ByVal cMember As String) As String
Dim cMainMember, nDir As Double, acode As Variant
If Not ValidInt(cMember) Then Exit Function
RetPhoto_s = sPath_App & "\photo_s\" & cMember & ".jpg"
End Function
Function RetPhoto_s_old(ByVal cMember As String) As String
Dim cMainMember, nDir As Double, acode As Variant
RetPhoto_s_old = sPath_App & "\PHOTO_S_OLD\" & cMember & ".jpg"
End Function

Function RetPhoto_v(ByVal cMember As String) As String
Dim cMainMember, nDir As Double, acode As Variant
If Not ValidInt(cMember) Then Exit Function
RetPhoto_v = sPath_App & "\photo_v\" & cMember & ".jpg"
End Function

Function aUnMyCodeBar(sCode) As Variant
Dim nVal1 As Integer, nVal2 As Integer, nValSub1 As Integer, nValSub2 As Integer
Dim nNumber1 As String, nNumber2 As String, nNumber3 As String


If Trim(sCode) = "" Then Exit Function
Dim aRet As Variant
aRet = Split(sCode, "_")
If IsEmpty(aRet) Then Exit Function
If UBound(aRet) <> 2 Then Exit Function
If Not ValidInt(Val(aRet(0))) Then Exit Function
If Not ValidInt(Val(aRet(1))) Then Exit Function
If Not ValidInt(Val(aRet(2))) Then Exit Function

nVal1 = 74: nVal2 = 11: nValSub1 = 71: nValSub2 = 4

nNumber1 = aRet(0)
nNumber2 = aRet(1)
nNumber3 = aRet(2)

nNumber1 = StrReverse(nNumber1)
nNumber1 = Val(nNumber1) + Val(nVal2)
nNumber1 = nNumber1 * 2
nNumber1 = Val(Left(nNumber1, Len(nNumber1) - 1))
nNumber1 = nNumber1 - nVal1

nNumber2 = Val(nNumber2) + Val(nValSub2)
nNumber2 = nNumber2 * 2
nNumber2 = Val(Left(nNumber2, Len(nNumber2) - 1))
nNumber2 = nNumber2 - nValSub1
nNumber2 = nNumber2 - Right(nNumber1, 1)

aRet = AddFlag(Empty, "MEMBER", Val(nNumber1))
aRet = AddFlag(aRet, "CODE", IIf(Val(nNumber2) = 0, "", Val(nNumber2)))
aRet = AddFlag(aRet, "TYPE", IIf(Val(nNumber3) = 0, "", Val(nNumber3)))
aUnMyCodeBar = aRet
End Function
Function MyCodeBar(sCode As String, Optional sType As String = "1") As String
Dim nMethod As Long, sMember As String
Dim nVal1 As Long, nVal2 As Long, nValSub1 As Long, nValSub2 As Long
Dim nNumber1 As Double, nNumber2 As Double

If Trim(sCode) = "" Then Exit Function
Dim aRet As Variant
aRet = Split(sCode, "-")
If IsEmpty(aRet) Then Exit Function
If UBound(aRet) < 0 Or UBound(aRet) > 1 Then Exit Function
If Not ValidInt(aRet(0)) Then Exit Function
If UBound(aRet) = 1 Then
    If Not ValidInt(aRet(1)) Then Exit Function
End If

sMember = aRet(0)
nVal1 = 74: nVal2 = 11: nValSub1 = 71: nValSub2 = 4

nPos = InStr(1, nCode, "-")
nNumber1 = Val(aRet(0))
If UBound(aRet) = 1 Then nNumber2 = Val(aRet(1))


mycode1 = Val(nNumber1) + Val(nVal1)
mycode1 = mycode1 & "2"
mycode1 = mycode1 / 2
mycode1 = mycode1 - nVal2
mycode1 = StrReverse(mycode1)

nNumber2 = Val(nNumber2) + Val(Right(nNumber1, 1))
myCode2 = Val(nNumber2) + Val(nValSub1)
myCode2 = myCode2 & "2"
myCode2 = myCode2 / 2
myCode2 = myCode2 - nValSub2
MyCodeBar = mycode1 & "_" & myCode2 & "_" & sType
End Function
Function unMyCodeBar(sCode, Optional ntype As Long = 0) As String
Dim nVal1 As Integer, nVal2 As Integer, nValSub1 As Integer, nValSub2 As Integer
Dim nNumber1 As String, nNumber2 As String, nNumber3 As String

If Trim(sCode) = "" Then Exit Function
Dim aRet As Variant
aRet = Split(sCode, "_")
If IsEmpty(aRet) Then Exit Function
If UBound(aRet) <> 2 Then Exit Function
If Not ValidInt(Val(aRet(0))) Then Exit Function
If Not ValidInt(Val(aRet(1))) Then Exit Function
If Not ValidInt(Val(aRet(2))) Then Exit Function


nVal1 = 74: nVal2 = 11: nValSub1 = 71: nValSub2 = 4

nNumber1 = aRet(0)
nNumber2 = aRet(1)
nNumber3 = aRet(2)

nNumber1 = StrReverse(nNumber1)
nNumber1 = Val(nNumber1) + Val(nVal2)
nNumber1 = nNumber1 * 2
nNumber1 = Val(Left(nNumber1, Len(nNumber1) - 1))
nNumber1 = nNumber1 - nVal1

nNumber2 = Val(nNumber2) + Val(nValSub2)
nNumber2 = nNumber2 * 2
nNumber2 = Val(Left(nNumber2, Len(nNumber2) - 1))
nNumber2 = nNumber2 - nValSub1
nNumber2 = nNumber2 - Right(nNumber1, 1)
If ntype = 0 Then
    unMyCodeBar = nNumber1 & IIf(nNumber2 > 0, "-" & nNumber2, "")
Else
    unMyCodeBar = nNumber3
End If
End Function
Function AgeString(dtDOB As Date, _
    Optional ByVal dtToday As Date = "12:00:00 AM") As String
    '---------------------------------------
    '     -------------------------------
    ' Name: AgeString
    ' Author:Bil Becker
    'iguanasoftware@yahoo.com
    ' Created: August 5,1999
    ' Modified: August 10,1999
    '
    ' Description: Function to format the ag
    '     e of something
    'example:
    'Input = 05/24/1997
    'Today = 08/05/1999
    'Output = 28 Years 2 Months 12 Days
    '
    'Input: dtDOB = date of birth or date th
    '     at you want age of
    'dtToday = starting date (todays date is
    '     default)
    '
    ' Output: Age formated ## Years ## Month
    '     s ## Days
    '
    ' Changes: August 10,1999
    'added dtToday input so fuction can get
    '     age between
    'two dates
    '---------------------------------------
    '     -------------------------------
    Dim intYear As Integer
    Dim intMonth As Integer
    Dim intDay As Integer
    Dim intCurrentYear As Integer
    Dim intCurrentMonth As Integer
    Dim intCurrentDay As Integer
    Dim intAgeYear As Integer
    Dim intAgeMonth As Integer
    Dim intAgeDay As Integer
    Dim stYear As String
    Dim stMonth As String
    Dim stDay As String
    Dim sngLeap As Single
    If dtToday = "12:00:00 AM" Then dtToday = Now
    intYear = Format(dtDOB, "yyyy")
    intMonth = Format(dtDOB, "mm")
    intDay = Format(dtDOB, "dd")
    intCurrentYear = Format(dtToday, "yyyy")
    intCurrentMonth = Format(dtToday, "mm")
    intCurrentDay = Format(dtToday, "dd")


    If intDay > intCurrentDay Then
        Select Case intCurrentMonth 'intMonth
            Case Is = 1
            intCurrentDay = intCurrentDay + 31
            intCurrentMonth = 12
            intCurrentYear = intCurrentYear - 1
            Case Is = 2
            sngLeap = intCurrentYear / 4


            If InStr(Str$(sngLeap), ".") Then
                intCurrentDay = intCurrentDay + 28
                intCurrentMonth = intCurrentMonth - 1
            Else
                intCurrentDay = intCurrentDay + 29
                intCurrentMonth = intCurrentMonth - 1
            End If
            Case Is = 4, 6, 9, 11
            intCurrentDay = intCurrentDay + 30
            intCurrentMonth = intCurrentMonth - 1
            Case Is = 3, 5, 7, 8, 10, 12
            intCurrentDay = intCurrentDay + 31
            intCurrentMonth = intCurrentMonth - 1
        End Select
End If


If intMonth > intCurrentMonth Then

    Select Case intCurrentMonth
        Case Is = 1
        intCurrentMonth = 13
        intCurrentYear = intCurrentYear - 1
        Case Else
        intCurrentMonth = intCurrentMonth + 12
        intCurrentYear = intCurrentYear - 1
    End Select
End If
intAgeYear = intCurrentYear - intYear
intAgeMonth = intCurrentMonth - intMonth
intAgeDay = intCurrentDay - intDay


Select Case intAgeYear
Case Is = 0
stYear = ""
Case Is = 1
stYear = Trim$(Str$(intAgeYear)) & " ”‰…"
Case Else
stYear = Trim$(Str$(intAgeYear)) & " ”‰…"
End Select


Select Case intAgeMonth
Case Is = 0
stMonth = ""
Case Is = 1
stMonth = Trim$(Str$(intAgeMonth)) & " ‘Â— "
Case Else
stMonth = Trim$(Str$(intAgeMonth)) & " ‘Â— "
End Select


Select Case intAgeDay
Case Is = 0
stDay = ""
Case Is = 1
stDay = Trim$(Str$(intAgeDay)) & " ÌÊ„ "
Case Else
stDay = Trim$(Str$(intAgeDay)) & " ÌÊ„ "
End Select
AgeString = Trim$(stYear & " " & stMonth & " " & stDay)
End Function
Function Age(dtDOB As Date, _
    Optional ByVal dtToday As Date = "12:00:00 AM") As Long
    '---------------------------------------
    '     -------------------------------
    ' Name: AgeString
    ' Author:Bil Becker
    'iguanasoftware@yahoo.com
    ' Created: August 5,1999
    ' Modified: August 10,1999
    '
    ' Description: Function to format the ag
    '     e of something
    'example:
    'Input = 05/24/1997
    'Today = 08/05/1999
    'Output = 28 Years 2 Months 12 Days
    '
    'Input: dtDOB = date of birth or date th
    '     at you want age of
    'dtToday = starting date (todays date is
    '     default)
    '
    ' Output: Age formated ## Years ## Month
    '     s ## Days
    '
    ' Changes: August 10,1999
    'added dtToday input so fuction can get
    '     age between
    'two dates
    '---------------------------------------
    '     -------------------------------
    Dim intYear As Integer
    Dim intMonth As Integer
    Dim intDay As Integer
    Dim intCurrentYear As Integer
    Dim intCurrentMonth As Integer
    Dim intCurrentDay As Integer
    Dim intAgeYear As Integer
    Dim intAgeMonth As Integer
    Dim intAgeDay As Integer
    Dim stYear As String
    Dim stMonth As String
    Dim stDay As String
    Dim sngLeap As Single
    If dtToday = "12:00:00 AM" Then dtToday = Now
    intYear = Format(dtDOB, "yyyy")
    intMonth = Format(dtDOB, "mm")
    intDay = Format(dtDOB, "dd")
    intCurrentYear = Format(dtToday, "yyyy")
    intCurrentMonth = Format(dtToday, "mm")
    intCurrentDay = Format(dtToday, "dd")


    If intDay > intCurrentDay Then
        Select Case intCurrentMonth 'intMonth
            Case Is = 1
            intCurrentDay = intCurrentDay + 31
            intCurrentMonth = 12
            intCurrentYear = intCurrentYear - 1
            Case Is = 2
            sngLeap = intCurrentYear / 4


            If InStr(Str$(sngLeap), ".") Then
                intCurrentDay = intCurrentDay + 28
                intCurrentMonth = intCurrentMonth - 1
            Else
                intCurrentDay = intCurrentDay + 29
                intCurrentMonth = intCurrentMonth - 1
            End If
            Case Is = 4, 6, 9, 11
            intCurrentDay = intCurrentDay + 30
            intCurrentMonth = intCurrentMonth - 1
            Case Is = 3, 5, 7, 8, 10, 12
            intCurrentDay = intCurrentDay + 31
            intCurrentMonth = intCurrentMonth - 1
        End Select
End If


If intMonth > intCurrentMonth Then
    Select Case intCurrentMonth
        Case Is = 1
        intCurrentMonth = 13
        intCurrentYear = intCurrentYear - 1
        Case Else
        intCurrentMonth = intCurrentMonth + 12
        intCurrentYear = intCurrentYear - 1
    End Select
End If
intAgeYear = intCurrentYear - intYear
intAgeMonth = intCurrentMonth - intMonth
intAgeDay = intCurrentDay - intDay
Age = intAgeYear
End Function
Function createCommand(pString As String, pCon As ADODB.Connection) As Boolean
Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = pCon
FS1.CommandText = pString
FS1.Execute
Set FS1 = Nothing
End Function
Function myFormat_p(sDate As Variant, Optional bLong As Boolean = False) As String
If bLong Then
    myFormat_p = Format(sDate, "YYYY/M/D " & "HH:NN")
Else
    myFormat_p = Format(sDate, "YYYY/M/D")
End If
End Function
Public Function mySplitSep(pString As String, Optional pSep As String = "-")
Dim aString As Variant
If Trim(pString) = "" Then Exit Function
aString = Split(pString, pSep)
If UBound(aString) < 1 Then Exit Function
For i = UBound(aString) To 0 Step -1
    mySplitSep = mySplitSep & IIf(mySplitSep = "", "", pSep) & aString(i)
Next
End Function
Public Function mRound(ByVal nValue As Variant, Optional nRound As Integer = 2) As Double
Dim cString As String
cString = "##"
If nRound > 0 Then cString = cString & "." & String(nRound, "#")
mRound = Val(Format(nValue, cString))
End Function
Function findRows(aRows, Optional pFieldFind As String = "code", Optional pValue As Variant, Optional pFieldRet As String, Optional pType As String = "n", Optional pDef As Variant) As Variant
If Not IsEmpty(aRows) Then
    For i = 0 To UBound(aRows)
        If pType = "n" Then
            If Val(retFlag(aRows(i), pFieldFind) & "") = pValue Then
                findRows = retFlag(aRows(i), pFieldRet)
                Exit For
            End If
        ElseIf pType = "s" Then
            If Trim(UCase(retFlag(aRows(i), pFieldFind)) & "") = Trim(pValue) Then
                findRows = retFlag(aRows(i), pFieldRet)
                Exit For
            End If
        ElseIf pType = "d" Then
            If myFormat(retFlag(aRows(i), pFieldFind)) = myFormat(pValue) Then
                findRows = retFlag(aRows(i), pFieldRet)
                Exit For
            End If
        ElseIf pType = "b" Then
            If retFlag(aRows(i), pFieldFind) = pValue Then
                findRows = retFlag(aRows(i), pFieldRet)
                Exit For
            End If
        End If
    Next
End If
End Function
Public Function grdRate(pGrid, Optional pWidth As Integer = 11500) As Double
Dim nwidth As Double, nRate As Double
With pGrid
For i = 0 To .Cols - 1
    If Not .ColHidden(i) Then
        nwidth = .ColWidth(i) + nwidth
    End If
Next
grdRate = pWidth / nwidth
End With
End Function
Function dateSql(ByVal pDate As Variant, Optional ByVal pTime As String = "") As String
If Not IsDate(pDate) Then
    dateSql = ""
Else
    If pTime <> "" Then
        dateSql = MyParn(Format(pDate, "YYYY-MM-DD") & " " & Format(pTime, "HH:NN"))
    ElseIf Format(pDate, "HH:NN") = "00:00" Then
        dateSql = MyParn(Format(pDate, "YYYY-MM-DD"))
    Else
        dateSql = MyParn(Format(pDate, "YYYY-MM-DD HH:NN"))
    End If
End If
End Function
Public Function DeletePhoto(sMember As String, Optional sAppend As String = "") As Boolean
Dim fs As New FileSystemObject
'On Error GoTo myerror
If fs.FileExists(MemPhoto(sMember, sAppend)) Then
    MyCreateFolder (sPath_App & "\photo_tmp")
    Dim cPhotoTemp As String
    cPhotoTemp = sPath_App & "\photo_tmp\" & sMember & turn(sAppend, "-") & sAppend & ".jpg"
    If fs.FileExists(cPhotoTemp) Then
        Dim i As Long
        Do
            i = i + 1
            cPhotoTemp = sPath_App & "\photo_tmp\" & sMember & turn(sAppend, "-") & sAppend & "_" & i & ".jpg"
        Loop Until Not fs.FileExists(cPhotoTemp)
    End If
    fs.CopyFile MemPhoto(sMember, sAppend), cPhotoTemp
    fs.DeleteFile MemPhoto(sMember, sAppend)
End If
Set fs = Nothing
DeletePhoto = True
Exit Function
'myerror:
'If bMsg Then MsgBox Err.Description
'Err.Clear
End Function
Public Function DeletePhoto_I(sMember As String, Optional sAppend As String = "") As Boolean
Dim fs As New FileSystemObject
If fs.FileExists(MemPhoto_I(sMember, sAppend)) Then
    MyCreateFolder (sPath_App & "\photoI_tmp")
    Dim cPhotoTemp As String
    cPhotoTemp = sPath_App & "\photoI_tmp\" & sMember & turn(sAppend, "-") & sAppend & ".jpg"
    If fs.FileExists(cPhotoTemp) Then
        Dim i As Long
        Do
            i = i + 1
            cPhotoTemp = sPath_App & "\photoI_tmp\" & sMember & turn(sAppend, "-") & sAppend & "_" & i & ".jpg"
        Loop Until Not fs.FileExists(cPhotoTemp)
    End If
    fs.CopyFile MemPhoto_I(sMember, sAppend), cPhotoTemp
    fs.DeleteFile MemPhoto_I(sMember, sAppend)
End If
Set fs = Nothing
DeletePhoto_I = True
Exit Function
'myerror:
'If bMsg Then MsgBox Err.Description
'Err.Clear
End Function
Function NextVisible(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
Dim nLast
For i = IIf(nBegincol = -1, 0, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, IIf(nEndCol > pGrid.Cols - 1, pGrid.Cols - 1, nEndCol))
    If pGrid.ColHidden(i) = False Or pGrid.ColWidth(i) = 0 Then
        NextVisible = i
        Exit Function
    End If
Next
NextVisible = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
Function myValidInt(ByVal nNumber) As Boolean
If Not IsNumeric(nNumber) Then Exit Function
If Val(nNumber) = 0 Then Exit Function
If Fix(Val(nNumber)) <> Val(nNumber) Then Exit Function
If Val(nNumber) < 0 Then Exit Function
myValidInt = True
End Function
Function mySet(pString As String, pCon As ADODB.Connection, Optional nTimeOut As Integer = 300) As ADODB.Recordset
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Dim loctable As New ADODB.Recordset
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdText
cmdTable.CommandText = pString
cmdTable.CommandTimeout = nTimeOut
Set obj = cmdTable.Execute
Set mySet = obj
Set obj = Nothing
End Function

Public Sub ToFileExel22(MyGrid, Optional aIg As Variant, Optional nRowHead As Long = 0, Optional aRowMerge As Variant = Empty, Optional aCol As Variant = Empty, Optional nRate As Double = 0, Optional aWidth As Variant = Empty, Optional arowHeight As Variant = Empty, Optional aSetUp As Variant = Empty, Optional nSize As Integer = 12, Optional acolSplit As Variant = Empty, Optional myForm As Form, Optional pHeader As Variant = Empty)
    Dim irow As Long, i As Long, i2 As Long, nCols As Long, nFixedCols As Long, nFixedRows As Long, n As Long
    Dim icol As Long
    Dim objExcl As Excel.Application
    Dim objWk As Excel.Workbook
    Dim objSht As Excel.Worksheet
    Dim iHead As Long
    Dim vHead As Variant
    On Error Resume Next
    Set objExcl = Excel.Application
    objExcl.Application.Visible = False
    Set objWk = objExcl.Workbooks.Add
    Set objSht = objWk.Sheets(1)
    objExcl.Application.DisplayAlerts = False
    Dim nRows As Long
        
    objSht.PageSetup.TopMargin = 10
    objSht.PageSetup.LeftMargin = 10
    objSht.PageSetup.HeaderMargin = 20
    objSht.PageSetup.CenterHeader = "&B &14"
    
        
    For i = 0 To MyGrid.FixedRows - 1
        If Not MyGrid.RowHidden(i) Then
            nFixedRows = nFixedRows + 1
        End If
    Next
                        
    For i = 0 To MyGrid.FixedCols - 1
        If Not MyGrid.ColHidden(i) Then
            nFixedCols = nFixedCols + 1
        End If
    Next
            
    
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCols = nCols + 1
            If nFixedRows > 0 Then
                objSht.Range(objSht.Cells(1, nCols), objSht.Cells(nFixedRows, nCols)).NumberFormat = "@"
            End If
            If MyGrid.rows > nFixedRows Then
                If Not (MyGrid.ColDataType(icol) = flexDTDouble) Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCols), objSht.Cells(MyGrid.rows, nCols)).NumberFormat = "General"
                ElseIf (MyGrid.ColDataType(icol) = flexDTDate) Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCols), objSht.Cells(MyGrid.rows, nCols)).NumberFormat = "dd-mm-yyyy"
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCols), objSht.Cells(MyGrid.rows, nCols)).NumberFormat = ""
                End If
            End If
        End If
    Next icol
    
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If
    
    For irow = 0 To MyGrid.rows - 1
        If (Not myForm Is Nothing) And MyGrid.rows > 1 Then
            myForm.prog1.Value = IIf((irow / (MyGrid.rows - 1)) * 100 > 100, 100, (irow / (MyGrid.rows - 1)) * 100)
        End If
        If Not MyGrid.RowHidden(irow) Then
            nRows = nRows + 1
            nCols = 0
            For icol = 0 To MyGrid.Cols - 1
                If Not MyGrid.ColHidden(icol) Then
                    nCols = nCols + 1
                    If MyGrid.ColDataType(icol) = flexDTDate Then
                        objSht.Cells(nRows, nCols) = myFormat_p(MyGrid.Cell(flexcpTextDisplay, irow, icol))
                    Else
                        objSht.Cells(nRows, nCols) = MyGrid.Cell(flexcpTextDisplay, irow, icol)
                    End If
                End If
            Next icol
        End If
    Next irow
                                
    Dim nRow2 As Long
    If Not IsEmpty(aCol) Then
        For nCol = 0 To UBound(aCol)
            nValue = 0
            For nRow2 = 1 To nRows
                If Trim(objSht.Cells(nRow2, aCol(nCol))) <> Trim(cValue & "") Then
                    If nValue <> 0 Then
                        objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
                    End If
                    cValue = Trim(objSht.Cells(nRow2, aCol(nCol)))
                    nValue = 0
                    nBegin = nRow2
                Else
                    nValue = nValue + 1
                End If
            Next
            If nValue <> 0 Then
                objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
            End If
        Next
    End If
  
    If Not IsEmpty(acolSplit) Then
        For i = 0 To UBound(acolSplit)
            objSht.Columns(retFlag(acolSplit(i), "col")).PageBreak = xlPageBreakManual
        Next
    End If
    
'    If nFixedRows > 0 Then
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, nCols)).HorizontalAlignment = xlCenter
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, nCols)).Interior.ColorIndex = 40
'    End If
    If nRows > 0 Then
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).VerticalAlignment = xlCenter
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, nCols)).Borders.ColorIndex = 0
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Size = nSize
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Bold = True
    End If
    
    nCol = 0
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCol = nCol + 1
            If nRows > 0 Then
                If MyGrid.ColFormat(icol) = "(##,##.##" Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).NumberFormat = "_(#,###.00_);[Red](#,###.00);0.00"
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignLeftCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
                ElseIf MyGrid.ColAlignment(icol) = flexAlignRightCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignCenterCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlCenter
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
                End If
            End If
            If nRate > 0 Then objSht.Columns(nCol).ColumnWidth = (MyGrid.ColWidth(icol) / 100) * nRate
        End If
    Next icol
               
    If Not IsEmpty(aWidth) Then
        For i = 0 To UBound(aWidth)
            If Val(retFlag(aWidth(i), "width")) = 0 Then
                objSht.Range(objSht.Cells(1, retFlag(aWidth(i), "col")), objSht.Cells(nRows, retFlag(aWidth(i), "col"))).Columns.AutoFit
            Else
                objSht.Columns(retFlag(aWidth(i), "col")).ColumnWidth = Val(retFlag(aWidth(i), "width")) / 100
            End If
        Next
    End If
    
    If Not IsEmpty(aRowMerge) Then
        For i = 0 To UBound(aRowMerge)
            If Not IsEmpty(retFlag(aRowMerge(i), "cols")) Then
                If Not IsEmpty(retFlag(aRowMerge(i), "text")) Then
                    objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))) = retFlag(aRowMerge(i), "text")
                End If
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))).Merge
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = 19
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Borders.ColorIndex = 0
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Size = nSize
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).HorizontalAlignment = xlCenter
            End If
            
            If retFlag(aRowMerge(i), "split") Then
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).PageBreak = xlPageBreakManual
            End If
            
            If retFlag(aRowMerge(i), "word_wrap") Then
                'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).WrapText = True
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).WrapText = True
            End If
            If Not IsEmpty(retFlag(aRowMerge(i), "height")) Then
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).RowHeight = retFlag(arowHeight(i), "height")
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).RowHeight = retFlag(aRowMerge(i), "height")
            End If
             If Not IsEmpty(retFlag(aRowMerge(i), "back_color")) Then
               objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = retFlag(aRowMerge(i), "back_color")
            End If
            
            If retFlag(aRowMerge(i), "bold") Then
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).Font.Bold = True
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
            End If
        Next
    End If
    
    objSht.PageSetup.Orientation = xlLandscape
    If Not IsEmpty(aSetUp) Then
        If Not IsEmpty(retFlag(aSetUp, "title_col")) Then
            objSht.PageSetup.PrintTitleColumns = objSht.Columns(retFlag(aSetUp, "title_col")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "title_row")) Then
            objSht.PageSetup.PrintTitleRows = objSht.rows(retFlag(aSetUp, "title_row")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "freeze")) Then
            ActiveWindow.SplitColumn = retFlag(aSetUp, "freeze")
            ActiveWindow.SplitRow = 0
            ActiveWindow.FreezePanes = True
        End If
        
        If retFlag(aSetUp, "autoSize") Then
            objSht.PageSetup.FitToPagesWide = True
        End If
                
        If retFlag(aSetUp, "center_header") <> "" Then
            objSht.PageSetup.CenterHeader = "&14&B" & retFlag(aSetUp, "center_header")
        End If
    End If
    
    
    If Not IsMissing(aIg) Then
        For i = 0 To UBound(aIg)
            objSht.rows(aIg(i)).Hidden = True
        Next
    End If
    
    
    If Not IsEmpty(pHeader) Then
        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                objSht.Range("A1", "S1").Insert
                objSht.Range("A1", Chr(64 + nCols) & "1").Merge
                objSht.Range("A1", Chr(64 + nCols) & "1").Font.Size = nSize + 1
                objSht.Range("A1", Chr(64 + nCols) & "1").Font.Bold = True
                objSht.Range("A1", Chr(64 + nCols) & "1").VerticalAlignment = xlCenter
                objSht.Range("A1", Chr(64 + nCols) & "1").HorizontalAlignment = xlCenter
            End If
        Next
        
        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                n = n + 1
                objSht.Cells(n, 1) = pHeader(i)
            End If
        Next
    
    End If
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If
    objExcl.Application.Visible = True
'    If Not IsEmpty(aRowMerge) Then
'        For i = 0 To UBound(aRowMerge)
'            objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols") + 1)).Merge
'        Next
'    End If

'''''''''''''''''''''''''''''
'›Ì Õ«·  —Ìœ Õ›Ÿ Ê—ﬁ… «·√ﬂ”·
'objWk.SaveAs "c:\Book1.xls"
'objWk.Close
'«·”ÿ— «· «·Ì Ì€·ﬁ »—‰«„Ã «·√ﬂ”·
'objExcl.Quit
''''''''''''''''''''''''''''
Set objSht = Nothing
Set objWk = Nothing
Set objExcl = Nothing
If Not myForm Is Nothing Then myForm.prog1.Visible = False
End Sub
Public Function addSelect(Optional pField1 As String = " CODE", Optional pField2 As String = "DESCA", Optional pType As String = "VARCHAR(10)") As String
addSelect = "SELECT '' AS " & pField1 & ",'' AS " & pField2 & " UNION ALL "
End Function
Function EmptyArray(pArray) As Boolean
If Not IsEmpty(pArray) Then
     Dim nBound As Long
     On Error Resume Next
     nBound = UBound(pArray)
     If Err.Number = 0 Then Exit Function
     Err.Clear
End If
EmptyArray = True
End Function
Public Sub SendKeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.SendKeys CStr(text), wait
  Set WshShell = Nothing
End Sub
Public Function Tr(pString As Variant, Optional pReturn As String = " AND ") As String
Tr = IIf(Trim(pString & "") = "", "", pReturn)
End Function
Public Function DefUser() As Boolean
DefUser = RetSetting("DEFAULT") = "1"
End Function
Public Function CopyFile(pSource As String, pTarget As String, Optional bOverRide As Boolean = True, Optional ByRef pError As String = "") As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
fs.CopyFile pSource, pTarget, bOverRide
CopyFile = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function

