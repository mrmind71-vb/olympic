Attribute VB_Name = "Module1"
Public Lookupdata, lManger As Boolean, cPathTemp As String
Public MdbPath, PublicPath, tempPath As String, tempFile As String, aPublic()
Public PublicVar As String, publicFlag As Variant, lCust As Integer
Public Firsttitle As String
Public nCountPrint As Byte
Function aTurnValue(pSource, aOld, pNew)
For I = 0 To UBound(aOld)
    pOld = aOld(I)
    If pSource = pOld Or (IsNull(pSource) And IsNull(pOld)) Then
         aTurnValue = pNew
         Exit Function
    End If
Next
aTurnValue = pSource
End Function
Function aDel(aTarget, nitem)
Dim aTemp
ReDim aTemp(UBound(aTarget) - 1, UBound(aTarget, 2))
For I = 0 To UBound(aTarget, 1) - 1
    If I <> nitem Then
        For i2 = 0 To UBound(aTarget, 2) - 1
            aTemp(IIf(I < nitem, I, I - 1), i2) = aTarget(I, i2)
        Next
    End If
Next
aDel = aTemp
End Function
Function MyParn(pValue)
MyParn = "'" & Trim(SQLFixup(pValue)) & "'"
End Function
Function myDateString(pDate)
myDateString = "Date(" & Year(pDate) & "," & Month(pDate) & "," & Day(pDate) & ")"
End Function
Function ReturnIndexed(aTarget, aIndex, nDim)
Dim aReturn()
If nDim = 1 Then
    ReDim aReturn(UBound(aTarget))
Else
    ReDim aReturn(UBound(aTarget, 1), UBound(aTarget, 2))
End If
If nDim = 1 Then
    For I = 0 To UBound(aTarget) - 1
        aReturn(aIndex(I)) = aTarget(I)
    Next
ElseIf nDim = 2 Then
    For I = 0 To UBound(aTarget, 1) - 1
        For i2 = 0 To UBound(aTarget, 2) - 1
            aReturn(aIndex(I), i2) = aTarget(I, i2)
        Next
    Next
End If
ReturnIndexed = aReturn
End Function
Function IndexArray(myArray)
Dim aReturn()
ReDim aReturn(UBound(myArray))
For I = 0 To UBound(myArray) - 1
    aReturn(I) = itemIndex(myArray, I, nBegin)
Next
IndexArray = aReturn
End Function
Function OneDimArray(ParArray, nDim)
Dim myArray()
ReDim myArray(UBound(ParArray))
For I = LBound(myArray) To UBound(myArray) - 1
    myArray(I) = ParArray(I, nDim)
Next
OneDimArray = myArray
End Function
Function itemIndex(myArray, nitem, nBegin)
Item = myArray(nitem)
NRETURN = nBegin
For I = nBegin To UBound(myArray) + nBegin - 1
    If I < nitem Then
        NRETURN = IIf(Item >= myArray(I), NRETURN + 1, NRETURN)
    Else
        NRETURN = IIf(Item > myArray(I), NRETURN + 1, NRETURN)
    End If
Next
itemIndex = NRETURN
End Function
Function aSearch(myArray, xSearch, Optional mydim, Optional nBegin)
nBegin = IIf(IsMissing(nBegin), 0, nBegin)
If IsMissing(mydim) Then
    For I = nBegin To UBound(myArray)
        If myArray(I) = xSearch Then
            aSearch = I
            Exit Function
        End If
    Next
Else
    For I = nBegin To UBound(myArray)
        If myArray(I, mydim) = xSearch Then
            aSearch = I
            Exit Function
        End If
    Next
End If
aSearch = Null
End Function
Function aScan(myArray, xSearch, Optional nBegin)
For I = nBegin To UBound(myArray)
    If myArray(I) = xSearch Then
    aScan = I
    Exit Function
    End If
Next
aScan = Null
End Function
Function aSearch2(myArray, xSearch, Optional mydim, Optional nBegin)
nBegin = IIf(IsMissing(nBegin), 0, nBegin)
If IsMissing(mydim) Then
    For I = nBegin To UBound(myArray)
        If myArray(I) = xSearch Then
'            aSearch = i
            Exit Function
        End If
    Next
Else
    For I = nBegin To UBound(myArray)
        If myArray(I, mydim) = xSearch Then
            Exit Function
        End If
    Next
End If
End Function
Function aAdd(aTarget, xItem)
Dim aTemp
ReDim aTemp(UBound(aTarget) + 1)
For I = 1 To UBound(aTarget)
      aTemp(I) = aTarget(I)
Next
aTemp(UBound(aTemp)) = xItem
aAdd = aTemp
End Function

Function LastInStr(cStr1, cStr2)
    If InStr(1, cStr1, cStr2) = 0 Then
        LastInStr = 0: Exit Function
    Else
        nFound = nFound = InStr(1, cStr1, cStr2)
        Do While True
        nloop = InStr(nloop + 1, cStr1, cStr2)
        If nloop = 0 Then
            Exit Do
        Else
            nFound = nloop
        End If
        Loop
    End If
    LastInStr = nFound
End Function
Function TurnValue(pSource, Optional pOld = "", Optional pNew = Null)
  TurnValue = IIf(pSource = pOld Or (IsNull(pSource) And IsNull(pOld)), pNew, pSource)
End Function
Function myNull(pSource As Variant, pDefault As Variant)
myNull = IIf(IsNull(pSource), pDefault, pSource)
End Function

Function RetField(xTable, cField, cSearchStr As String, xReturn As String)
cSearchStr = cField & " = " & "'" & cSearchStr & "'"
xTable.FindFirst cSearchStr
If xTable.NoMatch Then Exit Function
RetField = xTable(xReturn)
End Function
Function StrDel(cSource, myArray)
mycounter = UBound(myArray)
For I = 0 To mycounter
cSource = StrTran(cSource, myArray(I), "")
Next
StrDel = cSource
End Function

Function StrTran(cSource, cStr1, cStr2)
    nLen = Len(cStr1)
    Do Until InStr(1, cSource, cStr1) = 0
        nPosition = InStr(1, cSource, cStr1)
        cSource = Mid(cSource, 1, nPosition - 1) & cStr2 & _
                  Mid(cSource, nPosition + nLen)
    Loop
    StrTran = cSource
End Function
Function IncRec(cString)
Dim G As Double
Dim cChr As String
nLen = Len(cString)
For G = 0 To nLen - 1
    cChr = Mid(cString, nLen - G, 1)
    If IsNumeric(cChr) And cChr <> "9" Then
        cChr = cChr + 1
        cString = Mid(cString, 1, nLen - (G + 1)) & cChr & Mid(cString, nLen - (G - 1))
        Exit For
    Else
       cChr = IIf(cChr = "9", "0", cChr)
       cString = Mid(cString, 1, nLen - (G + 1)) & cChr & Mid(cString, nLen - (G - 1))
    End If
Next
IncRec = cString
End Function
Function RetNumber(pNumber, myDec As Boolean)
If (pNumber >= 48 And pNumber <= 57) _
    Or pNumber = 8 Or pNumber = 45 _
    Or (myDec = True And pNumber = 46) Then
    RetNumber = pNumber
Else
    RetNumber = 0
End If
End Function
Function ValidQuant(nQuant, Pack)
On Error Resume Next
ValidQuant = Abs(nQuant) / Pack < 1
If Err.Number > 0 Then
ValidQuant = 0
End If
End Function
Function myiif(cCondition, cField, Optional cField2 As String = "0", Optional cFunction As String = "sum")
If cCondition = "" Then
    myiif = cFunction & "(" & cField & ")"
Else
    myiif = cFunction & "( case when (" & cCondition & ") THEN " & _
         cField & " else " & cField2 & " end" & ")"
End If
End Function
Function myiif2(cCondition, cField, Optional cField2 As String = "0")
If cCondition = "" Then
    myiif2 = cField
Else
    myiif2 = "case when (" & cCondition & ") THEN " & _
         cField & " else " & cField2 & " end"
End If
End Function


Function RetFind(rTable, cFieldFind, cFieldRet, cSearch)
'rTable.FindFirst cFieldFind & "=" & MyParn(cSearch)
'RetFind = IIf(rTable.NoMatch, "", rTable(cFieldRet))
End Function
Function Units(nPart1, nPart2, nPack)
If VarType(nPart1) = vbString Then nPart1 = Val(nPart1)
If VarType(nPart2) = vbString Then nPart2 = Val(nPart2)
If VarType(nPack) = vbString Then nPack = Val(nPack)
Units = (TurnValue(nPart1, Null, 0) * nPack) + TurnValue(nPart2, Null, 0)
End Function
Function retRev(cString) As String
For I = 1 To Len(cString)
    retRev = retRev & Mid(cString, Len(cString) - I + 1, 1)
Next
End Function
Function NameOfDay(xPass)
    cDay = Weekday(xPass)
    Select Case cDay
        Case 1
            NameOfDay = "«·√Õœ"
        Case 2
            NameOfDay = "«·√À‰Ì‰"
        Case 3
            NameOfDay = "«·À·«À«¡"
        Case 4
            NameOfDay = "«·√—»⁄«¡"
        Case 5
            NameOfDay = "«·Œ„Ì”"
        Case 6
            NameOfDay = "«·Ã„⁄…"
        Case 7
            NameOfDay = "«·”» "
   End Select
End Function
Function myQuery(sWhere)
myQuery = sWhere & IIf(sWhere = "", " Where ", " and ")
End Function
Function RetNumber2(cString, pNumber)
cString = cString & Format(Chr(pNumber))
RetNumber2 = IIf(IsNumeric(cString), pNumber, 0)
End Function
Function Unit1Q(Quant, Pack, cType)
If Not IsNull(Quant) And TurnValue(Pack, Null, 0) <> 0 Then
    If cType = "1" Then
        Unit1Q = Fix(TurnValue(Quant, Null, 0) / Pack)
    Else
        Unit1Q = TurnValue(Quant, Null, 0) Mod Pack
    End If
End If
If TurnValue(Pack, Null, 0) = 0 And cType = "2" Then Unit1Q = Quant
End Function
Function QtyToString(Quant, Pack)
If Not IsNull(Quant) And TurnValue(Pack, Null, 0) <> 0 Then
    QtyToString = TurnValue(Quant, Null, 0)
    If Len(QtyToString) = 3 Then QtyToString = " " & QtyToString
    If Len(QtyToString) = 2 Then QtyToString = "  " & QtyToString
    If Len(QtyToString) = 1 Then QtyToString = "   " & QtyToString
    QtyToString = QtyToString & "/" & TurnValue(Quant, Null, 0) Mod Pack
End If
If TurnValue(Pack, Null, 0) = 0 Then QtyToString = Quant
End Function
Function Chr254(cPass)
    Do While True
        nPos = MyInStr(cPass)
        If nPos > 0 Then
             cPass = Mid(cPass, 1, nPos - 1) & " " & Chr(254) & Mid(cPass, nPos + 1)
        Else
            Exit Do
        End If
    Loop
    Chr254 = Chr(254) & cPass & Chr(254)
End Function
Function RetZero(ByVal cString, Optional ByVal nLen As Integer = 6)
cString = Trim(cString)
If Len(cString) >= nLen Then
    RetZero = cString
    Exit Function
End If
nLen = nLen - Len(cString)
RetZero = String(nLen, "0") & cString
End Function
Function DelZero(cString)
Dim cStr1 As String
If IsNull(cString) Then Exit Function
cStr1 = cString
For I = 1 To Len(cStr1)
    If Mid(cStr1, I, 1) <> "0" Then Exit For
Next I
If I > 0 Then DelZero = Mid(cString, I)
If I = 0 Then DelZero = cString
End Function

Function addstring(pValue)
addstring = IIf(Trim(pValue & "") = "", "null", "'" & Trim(SQLFixup(pValue)) & "'")
End Function
Function addvalue(pValue, Optional bZero As Boolean = False) As String
addvalue = IIf(ValidNum(pValue, , bZero), pValue & "", "null")
End Function
Function addVal(pValue) As String
If ValidNum(pValue) Then
    addVal = IIf(Val(pValue) <> 0, pValue, "NULL")
Else
    addVal = "NULL"
End If
End Function
Public Function ValidNum(ByVal pCode As Variant, Optional nLen As Integer = 0, Optional pZero As Boolean = False, Optional nMax As Integer = 18, Optional nMin As Integer = 1) As Boolean
Dim sNumber As String, I As Long
If Trim(pCode & "") = "" Then Exit Function
pCode = Trim(pCode & "")

If nLen <> 0 Then
    If Len(pCode) <> nLen Then Exit Function
Else
    If Len(pCode & "") < nMin Then Exit Function
    If Len(pCode & "") > nMax Then Exit Function
    If Mid(pCode, 1, 1) = "0" And (pCode <> "0" Or (Not pZero)) Then Exit Function
End If

For I = 1 To Len(pCode)
    If Not IsNumeric(Mid(pCode, I, 1)) Then
        Exit Function
    End If
Next
ValidNum = True
End Function
Public Function ValidNumAny(ByVal pCode As Variant) As Boolean
Dim sNumber As String, I As Long
If Trim(pCode & "") = "" Then Exit Function
pCode = Trim(pCode & "")
For I = 1 To Len(pCode)
    If Not IsNumeric(Mid(pCode, I, 1)) Then Exit Function
Next
ValidNumAny = True
End Function
Function MyParnAll(parStr, Optional bAsOne As Boolean = False)
Dim aString
If bAsOne Then
    MyParnAll = "'%" & parStr & "%'"
Else
    aString = Split(parStr, " ")
    If Not IsEmpty(aString) Then
        For I = 0 To UBound(aString)
            MyParnAll = IIf(MyParnAll = "", "%", "") & MyParnAll & aString(I) & "%"
        Next
    End If
    MyParnAll = "'" & MyParnAll & "'"
End If
End Function


Function MyInStr(cPass)
    Do While True
        MyInStr = InStr(cPass, " 0")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 1")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 2")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 3")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 4")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 5")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 6")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 7")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 8")
        If MyInStr <> 0 Then Exit Do
        MyInStr = InStr(cPass, " 9")
        If MyInStr <> 0 Then Exit Do
        Exit Do
    Loop
End Function
Function TwoLine(cPass)
nPos = InStr(cPass, "-")
If nPos = 0 Then
    TwoLine = cPass
Else
    TwoLine = Mid(cPass, 1, nPos - 1) & String(30 - nPos, " ") & "-" & Mid(cPass, nPos)
End If
End Function
Public Function D_RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
cString = "Select sum(val([IN] & '') - VAL([OUT] & '')) as Balance From file1_11 where item = " & MyParn(cItem) & _
          " and Store = " & MyParn(cStore) & " and Date <= " & DateConv(dDate)
D_RetItemBalance = Val(GetDesca(cString) & "")
End Function


Function myNear(nNumber, nMode) As Double
Dim nNum1, nNum2, nNum3
If Val(nNumber) = 0 Then
    myMynear = 0
    Exit Function
ElseIf nMode = 0 Then
    myNear = nNumber
    Exit Function
End If
nNum1 = (nNumber - Fix(nNumber)) * 100
nNum1 = nNum1 Mod (nMode * 100)
myNear = (nNumber) - (nNum1 / 100) + IIf(nNum1 > 0, nMode, 0)
End Function

Sub MyEditItem(pGrid As Variant, Row As Long, Col As Long, Optional isCode As Boolean = False)
With pGrid
If Trim(.TextMatrix(Row, Col)) <> "" And Row > 0 Then
    .TextMatrix(.rows - 1, Col) = .TextMatrix(Row, Col)
    If isCode Then .TextMatrix(.rows - 1, Col + 1) = .TextMatrix(Row, Col + 1)
End If
End With
End Sub

Function myDateDiffString(pDate1, pDate2)
nYears = myDateDiff(pDate1, pDate2)
If nYears > 0 Then
     myDateDiffString = nYears & " ”‰… "
End If
pDate1 = DateAdd("yyyy", nYears, pDate1)

If Month(pDate2) > Month(pDate1) Then
    nMonth = Month(pDate2) - Month(pDate1)
    If Day(DDate2) < Day(DDate1) Then nMonth = nMonth - 1
End If

If Month(pDate2) < Month(pDate1) Then
    nMonth = Month(pDate2) + (12 - Month(pDate1))
    If Day(DDate2) < Day(DDate1) Then nMonth = nMonth - 1
End If

If nMonth > 0 Then
     myDateDiffString = myDateDiffString & nMonth & "‘Â—"
End If


If Day(pDate2) > Day(pDate1) Then
    nDays = Day(pDate2) - Day(pDate1)
ElseIf Day(pDate2) < Day(pDate1) Then
    nDays = Day(pDate2) + Monthdays(Month(pDate1)) - Day(pDate1)
End If
If nDays > 0 Then myDateDiffString = myDateDiffString & nDays & " ÌÊ„ "
End Function
Function myDateDiff(pDate1, pDate2)
myDateDiff = Year(pDate2) - Year(pDate1)
If Month(pDate2) > Month(pDate1) Then Exit Function
If Month(pDate2) < Month(pDate1) Then
    myDateDiff = myDateDiff - 1
    Exit Function
End If
If Day(pDate2) < Day(pDate1) Then myDateDiff = myDateDiff - 1
End Function
Private Function Monthdays(nMonth)
Select Case Monthdays
Case 1, 3, 5, 7, 8, 10, 12
    Monthdays = 31
Case 2
    Monthdays = 28
Case Else
    Monthdays = 30
End Select
End Function
Public Sub FixTotals(grd As Variant, Row As Long, myArray As Variant)
Dim I As Long
For I = 0 To UBound(myArray)
    grd.TextMatrix(Row, myArray(I)) = mRound(grd.TextMatrix(Row, myArray(I)), 2)
Next
End Sub
