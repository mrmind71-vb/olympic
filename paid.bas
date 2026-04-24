Attribute VB_Name = "Paid2"
Public cWhere_rel As String
Public Function validMemberPay(pCode As String, pCon As ADODB.Connection) As String
Dim aMember As Variant
aMember = Member_Load(pCode, , pCon)

If IsEmpty(aMember) Then
    validMemberPay = "·« ÌÊÃœ ⁄÷Ê »Â–« «·«”„"
    Exit Function
End If

If retFlag(aMember, "DROP") Then
    validMemberPay = "«·⁄÷Ê ”«ﬁÿ ⁄÷ÊÌ…"
    Exit Function
End If

If Not ValidNum(retFlag(aMember, "TYPE")) Then
    validMemberPay = "·Ì” ··⁄÷Ê ‰Ê⁄ ⁄÷ÊÌ…"
    Exit Function
End If

If Not IsDate(retFlag(aMember, "DATE_BEGIN") & "") Then
    validMemberPay = "·Ì” ··⁄÷Ê  «—ÌŒ »œ«Ì… ⁄÷ÊÌ…"
    Exit Function
End If

If Not IsDate(retFlag(aMember, "DATE_BIRTH") & "") Then
    validMemberPay = "·Ì” ··⁄÷Ê  «—ÌŒ „Ì·«œ"
    Exit Function
End If

validMemberPay = "ok"
End Function
Function validClaim(pCode As String, pDate As String, pType As String, pCon As ADODB.Connection) As String
Dim atype As Variant, nYear_code As Variant
atype = Claim_Type_Load(pType, , pCon)

If IsEmpty(atype) Then
    validClaim = "‰Ê⁄ «·„ÿ«·»… €Ì— „”Ã·"
    Exit Function
End If

nYear_code = Ret_Year(pDate, "code", pCon)
If IsEmpty(nYear_code) Then
    validClaim = "”‰… «·„ÿ«·»… €Ì— „”Ã·…"
    Exit Function
End If

aPaid = Member_Paid(pCode, , pCon)
nYears = unpaid_years_count(pCode, nYear_code & "", pCon)

If mRound(retFlag(atype, "TYPE")) = 1 Then
    If IsEmpty(aPaid) Then
        validClaim = "·Ì” ··⁄÷Ê ”œ«œ ”«»ﬁ"
        Exit Function
    ElseIf nYears = 0 Then
        validClaim = "·Ì” ⁄·Ì «·⁄÷Ê ”‰Ê«  ”«»ﬁ…"
        Exit Function
    End If
ElseIf mRound(retFlag(atype, "TYPE")) = 2 Then
    If IsEmpty(aPaid) Then
        validClaim = "·Ì” ··⁄÷Ê ”œ«œ ”«»ﬁ"
    ElseIf nYears > 1 Then
        validClaim = "⁄·Ì «·⁄÷Ê " & nYears & " ”‰Ê« "
        'Exit Function
    End If
ElseIf mRound(retFlag(atype, "TYPE")) = 4 Or mRound(retFlag(atype, "TYPE")) = 5 Then
    If Not IsEmpty(aPaid) Then
        validClaim = "··⁄÷Ê ”œ«œ ”«»ﬁ"
        Exit Function
    End If
ElseIf mRound(retFlag(atype, "TYPE")) = 6 Then
    If IsEmpty(aPaid) Then
        validClaim = "·Ì” ··⁄÷Ê ”œ«œ ”«»ﬁ"
        Exit Function
    ElseIf mRound(retFlag(aPaid, "type")) <> 2 Then
        validClaim = "«·⁄÷Ê ·Ì” Õ«›Ÿ ⁄÷ÊÌ…"
        Exit Function
    End If
Else
    If IsEmpty(aPaid) Then
        validClaim = "·Ì” ··⁄÷Ê ”œ«œ ”«»ﬁ"
        Exit Function
    ElseIf nYears > 0 Then
        validClaim = "⁄·Ì «·⁄÷Ê " & nYears & " ”‰Ê« "
        Exit Function
    End If
End If

If retFlag(atype, "over_age") Then
    Dim nOverAge As Integer
    nOverAge = OverAge(pCode, nYear_code, pCon)
    If nOverAge > 0 Then
        validClaim = "··⁄÷Ê " & nOverAge & " «»‰«¡ ›Êﬁ «·”‰"
        Exit Function
    End If
End If

validClaim = "ok"
End Function
Private Function unpaid_years_count_byDate(pCode As String, pDate As String, pCon As ADODB.Connection) As Variant
Dim nYear_code As Long
nYear_code = Ret_Year(pDate, "code", pCon)
If Not IsEmpty(nYear_code) Then
    unpaid_years_count_byDate = GetField("Select dbo.unpaid_years_count(" & pCode & "," & nYear_code, pCon)
End If
End Function
Public Function unpaid_years_count(pCode As String, nYear_code As String, pCon As ADODB.Connection) As Variant
unpaid_years_count = GetField("Select dbo.unpaid_years_count(" & pCode & "," & nYear_code & ")", pCon)
End Function
Public Function unpaid_years(nYear_Paid As String, nYear_code As String, pCon As ADODB.Connection) As Variant
unpaid_years = GetField("Select dbo.unpaid_years(" & nYear_Paid & "," & nYear_code & ")", pCon)
End Function
Public Function addPayment(pCode As String, pDate As String, pType As String, pCon As ADODB.Connection) As Variant
Dim aMember As Variant, aPaid As Variant, aYear As Variant, aUnPaid As Variant
Dim nFirst_year As Integer, nFirst_year1 As Integer, nFirst_year2 As Integer, nFirst_year3 As Integer
Dim nYear_code As String, nYear_Code1 As Long, nYear_code2 As Long, nYear_code3 As Long
Dim sDoc_no As String, aInsert As Variant, sError As String
Dim cInsertHeader As String
Dim I As Integer, cInsert As String

sReturn = validMemberPay(pCode, pCon)
If sReturn <> "ok" Then
    addPayment = AddFlag(addPayment, "error", sReturn)
    Exit Function
End If

sReturn = validClaim(pCode, pDate, pType, pCon)
If sReturn <> "ok" Then
    addPayment = AddFlag(addPayment, "error", sReturn)
    Exit Function
End If

nYear_code = Ret_Year(pDate, "CODE", pCon)

cWhere_rel = ""
If pType = 3 Then
    addrelfrm.sCode = pCode
    addrelfrm.pType = pType
    addrelfrm.pYear_code = nYear_code
    addrelfrm.Show 1
End If

If cWhere_rel = "" And pType = 3 Then
    addPayment = AddFlag(addPayment, "error", "·„   „ «÷«›… «Ì «ﬁ«—»")
    Exit Function
End If

Dim nOverAge As Integer
nOverAge = OverAgeAll(pCode, nYear_code, pCon)
If nOverAge > 0 Then
    addPayment = AddFlag(addPayemnt, "msg", ArbString("··⁄÷Ê „‰ „ ÕœÌ «·«⁄«ﬁ… «Ê „ Êﬁ›Ì «·⁄÷ÊÌ…" & nOverAge & " «»‰«¡ ›Êﬁ «·”‰"))
End If

sDoc_no = Newflag("FILE6_20H", "DOC_NO", pCon)
aInsert = AddFlag(Empty, "DOC_NO", addVal(sDoc_no))
aInsert = AddFlag(aInsert, "[DATE]", addDate(pDate))
aInsert = AddFlag(aInsert, "[DATE_ISSUE]", addDate(pDate))
aInsert = AddFlag(aInsert, "[CODE]", addvalue(pCode))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(pType))
aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(nYear_code))
aInsert = AddFlag(aInsert, "[YEARS]", "1")
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
aInsert = AddFlag(aInsert, "[TIME]", "getdate()")
cInsert_header = addInsert(aInsert, "FILE6_20H")
aSql = AddFlag(aSql, cInsert_header)

Dim nType As Integer
nType = mRound(Claim_Type_Load(pType, "TYPE", pCon))

Dim rsYears As New ADODB.Recordset
If nType = 1 Or nType = 2 Then
    Set rsYears = UnPaidYears(pCode, nYear_code, pCon)
    Do Until rsYears.EOF
        If rsYears.RecordCount = 1 Then
            cInsert = addPaidItems(sDoc_no, pCode, rsYears!CODE, nYear_code, pType, pDate, 0, cWhere_rel, pCon)
        ElseIf rsYears!CODE = nYear_code Then
            I = I + 1
            cInsert = addPaidItems(sDoc_no, pCode, rsYears!CODE, nYear_code, pType, pDate, 0, cWhere_rel, pCon) & ";"
        ElseIf Not rsYears!NO_LATE Then
            I = I + 1
            If pType = 2 Then
                cInsert = addPaidItems(sDoc_no, pCode, rsYears!CODE, nYear_code, 1, pDate, I, cWhere_rel, pCon) & ";"
            Else
                cInsert = addPaidItems(sDoc_no, pCode, rsYears!CODE, nYear_code, 1, pDate, I, cWhere_rel, pCon) & ";"
            End If
        Else
            cInsert = addPaidItems(sDoc_no, pCode, rsYears!CODE, nYear_code, 1, pDate, 0, cWhere_rel, pCon) & ";"
        End If
        aSql = AddFlag(aSql, cInsert)
        rsYears.MoveNext
    Loop
Else
    'Set rsYears = UnPaidYears(pCode, nYear_code, pCon)
    'If Not rsYears.EOF Then
        cInsert = addPaidItems(sDoc_no, pCode, nYear_code, nYear_code, pType, pDate, 0, cWhere_rel, pCon)
        aSql = AddFlag(aSql, cInsert)
    'End If
End If

If UBound(aSql) > 0 Then
    'cInsert = cInsert_header & ";" & cInsert
    aSql = AddFlag(aSql, "update file6_20h set years_desca = dbo.f_get_years(" & sDoc_no & ") where file6_20h.doc_no = " & sDoc_no & ";")
    aSql = AddFlag(aSql, fixClaim(sDoc_no))
    addPayment = AddFlag(addPayment, "sql", aSql)
    addPayment = AddFlag(addPayment, "doc_no", sDoc_no)
End If
End Function
Public Function addPaidItems(pDoc_no As String, pCode As String, pYear_code As String, pYear_code_Doc As String, pType As String, pDate_Paid, pLate As Integer, pWhere_Rel As String, pCon As ADODB.Connection) As String
Dim cString As String, nAge As Long, bMemberAdd As Boolean
Dim pSection As String, pDate1 As String, pDate2 As String, aMeetMem As Double, aMeetRel As Variant, aMeetSub As Variant
Dim nAll As Long, aPaid As Variant, bApg As Boolean
Dim nLate As Long

aYear = GetFields("select * from years_codes where code = " & pYear_code)

pDate1 = myFormat(retFlag(aYear, "date1"))
pDate2 = myFormat(retFlag(aYear, "date2"))

pSection = Member_Load(pCode, "TYPE")
aPaid = Member_Paid(pCode, , pCon)
atype = Claim_Type_Load(pType, , pCon)
'nAll = retAll(pCode, pDate, atype, pCon)
bApg = Member_Load(pCode, "apg", pCon)
aRates = Array(0, 50, 100, 200)
nLate = aRates(IIf(pLate > 3, 3, pLate))

cString = "SELECT FILE6_11.ITEM,FILE6_11.TYPE,FILE6_11.YEAR_CODE,FILE6_10.AGE1,FILE6_10.AGE2 ,FILE6_10.DESCA, FILE6_10.ALLMEMBER, FILE6_10.LATE, FILE6_10.RELATION," & _
      " FILE6_10.ISMEMBER, COALESCE(FILE6_10.AGE1,0), COALESCE(FILE6_10.AGE2,0), FILE6_10.GENDER, " & _
      " FILE6_10.BASICDIED,FILE6_10.LATE_DAYS, FILE6_10.BASICNEW,FILE6_10.BASICOLD, FILE6_10.MEETING,FILE6_11.TAX, FILE6_10.DAYS, FILE6_10.NORATE, " & _
      " FILE6_11.value, FILE6_11.Discount " & _
      " FROM FILE6_10 INNER JOIN FILE6_11 ON FILE6_10.ITEM = FILE6_11.item " & _
      " WHERE FILE6_11.TYPE = " & pType & _
      " AND FILE6_11.BASIC = 1 " & _
      " AND FILE6_11.YEAR_CODE = " & IIf(mRound(pYear_code) < 9, "9", pYear_code) & _
      " AND [SECTION] =  " & pSection
cString = cString & " ORDER BY FILE6_10.ITEM"

Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText

bMemberAdd = Member_Load(pCode, "Died")
Do Until loctable.EOF
    If loctable!MEETING Then
        nValue = 0
        If Not bApg Then
            If loctable!isMember Then
                nValue = AddMeetingMem(loctable, pCode, pYear_code, aMeetMem, pCon)
            ElseIf loctable!RELATION = 1 Then
                If Not IsEmpty(aMeetRel) Then
                    nValue = AddMeetingRel(loctable, pCode, pYear_code, aMeetRel, pCon)
                End If
            End If
            If nValue <> 0 Then
                aInsert = AddFlag(Empty, "doc_no", pDoc_no)
                aInsert = AddFlag(aInsert, "item", loctable!Item)
                aInsert = AddFlag(aInsert, "value", nValue)
                aInsert = AddFlag(aInsert, "quant", 1)
                aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
                aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!tax))
                If loctable!late Then
                    'aInsert = AddFlag(aInsert, "late_rate", mRound(retFlag(aYear, "late_rate")))
                    aInsert = AddFlag(aInsert, "late_rate", nLate)
                End If
                aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
                cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
            End If
        End If
    ElseIf loctable!LATE_DAYS Then
        If IsDate(myFormat(retFlag(aYear, "DATE_LATE"))) Then
            If myFormat(pDate_Paid) >= myFormat(retFlag(aYear, "DATE_LATE")) Then
                aInsert = AddFlag(Empty, "doc_no", pDoc_no)
                aInsert = AddFlag(aInsert, "item", loctable!Item)
                aInsert = AddFlag(aInsert, "value", nValue)
                aInsert = AddFlag(aInsert, "quant", 1)
                aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
                aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!tax))
                If loctable!late Then
                    aInsert = AddFlag(aInsert, "late_rate", nLate)
                End If
                aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
                cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
            End If
        End If
    ElseIf loctable!isMember Then
        If (Not bMemberAdd) Then
            If AddMemberData(pCode, mRound(loctable!Age1), mRound(loctable!Age2), pDate2, pCon) Then
                nAll = nAll + 1
                aInsert = AddFlag(Empty, "doc_no", pDoc_no)
                aInsert = AddFlag(aInsert, "item", loctable!Item)
                aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
                aInsert = AddFlag(aInsert, "quant", "1")
                aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
                If loctable!late Then
                    aInsert = AddFlag(aInsert, "late_rate", nLate)
                End If
                aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!tax))
                aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
                cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
                aMeetMem = mRound(loctable!Value * mRound(1 - (loctable!discount / 100)))
                bMemberAdd = True
            End If
        End If
    ElseIf (Not IsNull(loctable!RELATION)) Then
'        If loctable!Type = 3 Then
'            cSql = AddRelationJoin(loctable, pCode, nYear_code, nRelation, pCon)
'            If cSql <> "" Then cInsert = cInsert & cSql & ";"
'        Else
        aMeetSub = Empty
        nRelation = addRelation(loctable, pCode, pDate1, pDate2, aMeetSub, loctable!RELATION, atype, pWhere_Rel, pYear_code_Doc, pCon)
        If nRelation > 0 Then
            nAll = nAll + nRelation
            aInsert = AddFlag(Empty, "doc_no", pDoc_no)
            aInsert = AddFlag(aInsert, "item", loctable!Item)
            aInsert = AddFlag(aInsert, "value", mRound(loctable!Value))
            aInsert = AddFlag(aInsert, "quant", nRelation)
            aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
            aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!tax))
            If loctable!late Then
                aInsert = AddFlag(aInsert, "late_rate", nLate)
            End If
            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
            cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
            If Not IsEmpty(aMeetSub) Then aMeetRel = AddFlag(aMeetRel, aMeetSub)
        End If
'        End If
'    ElseIf loctable!BasicNew Or loctable!basicOld Then
'        If (loctable!BasicNew And IsEmpty(aPaid)) Or (loctable!basicOld And Not IsEmpty(aPaid)) Then
'            aInsert = AddFlag(Empty, "doc_no", pDoc_no)
'            aInsert = AddFlag(aInsert, "item", loctable!Item)
'            aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
'            aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
'            aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
'            If loctable!late Then
'                aInsert = AddFlag(aInsert, "late_rate", mRound(retFlag(aYear, "late_rate")))
'            End If
'            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
'            cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
'        End If
    ElseIf mRound(loctable!Age2) <> 0 Then
        nAge = GetField("SELECT dbo.f_age(DATE_BIRTH, " & addstring(myFormat(pDate2)) & ") FROM FILE1_10 WHERE CODE = " & pCode)
        If nAge <= mRound(loctable!Age2) Then
            aInsert = AddFlag(Empty, "doc_no", pDoc_no)
            aInsert = AddFlag(aInsert, "item", loctable!Item)
            aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
            aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
            aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
            aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!tax))
            If loctable!late Then
                aInsert = AddFlag(aInsert, "late_rate", nLate)
            End If
            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
            cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
        End If
    Else
        aInsert = AddFlag(Empty, "doc_no", pDoc_no)
        aInsert = AddFlag(aInsert, "item", loctable!Item)
        aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
        aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
        aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
        aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!tax))
        If loctable!late Then
            aInsert = AddFlag(aInsert, "late_rate", nLate)
        End If
        aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
        cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
    End If
    loctable.MoveNext
Loop
addPaidItems = cInsert
End Function
Private Function retAll(pMember As Variant, pDate As String, atype As Variant, pCon As ADODB.Connection) As Integer
Dim cString As String
cString = "SELECT SUM(1) FROM FILE1_11"
cString = cString & " WHERE FILE1_11.MEMBER = " & pMember
'If retFlag(atype, "over_age") Then cString = cString & " AND FILE1_11.PENDING = 0"
If retFlag(atype, "over_age") Then
    cString = cString & " AND (NOT(GENDER = 1 AND HANDI = 0 AND RELATION = 2 AND dbo.f_age(FILE1_11.DATE_BIRTH ," & addstring(myFormat(pDate)) & ") > 24))"
End If
If IsDate(pDate) Then cString = cString & turn(cString) & "FILE1_11.DATE_BEGIN <= " & DateSq(pDate)
retAll = IIf(Member_Load(pMember, "died", pCon), 0, 1) + mRound(GetField(cString, pCon))
End Function
Private Function AddMemberData(pCode As String, pAge1 As Integer, pAge2 As Integer, pDate As String, pCon As ADODB.Connection) As Boolean
Dim nAge As Integer, nGender As Integer
aMember = Member_Load(pCode, , pCon)
If IsDate(retFlag(aMember, "DATE_BIRTH") & "") Then
    nAge = Age(myFormat(retFlag(aMember, "DATE_BIRTH")), myFormat(pDate))
Else
   nAge = 1
End If
If pAge1 > nAge And pAge1 <> 0 Then Exit Function
If pAge2 < nAge And pAge2 Then Exit Function
'If (Not IsNull(loctable!GENDER)) Then
'    If retFlag(aMember, "Gender", 1) <> loctable!GENDER Then Exit Function
'End If
AddMemberData = True
End Function
Private Function addRelation(ByRef loctable As ADODB.Recordset, pCode, pDate1, pDate2, ByRef aMeet As Variant, nRelation As Integer, atype As Variant, pWhere_Rel As String, pYear_code_Doc As String, pCon As ADODB.Connection) As Integer
Dim myRecordSet As New ADODB.Recordset
Dim nAge As Integer, nGender As Integer
cString = " SELECT [CODE],[DATE_BIRTH],COALESCE(GENDER,1) From FILE1_11"
cString = cString & " where relation = " & nRelation
cString = cString & " AND MEMBER = " & pCode
If loctable!year_code = pYear_code_Doc Then
    cString = cString & " AND PENDING = 0"
End If
If retFlag(atype, "over_age") Then
    cString = cString & " AND (NOT(GENDER = 1 AND HANDI = 0 AND RELATION = 2 AND dbo.f_age(FILE1_11.DATE_BIRTH ," & addstring(myFormat(pDate2)) & ") > 24))"
End If

If loctable!Type = 3 Then
    cString = cString & " AND (DATE_BEGIN IS NULL OR DATE_BEGIN >= " & addDate(pDate1) & ")"
Else
    cString = cString & " AND DATE_BEGIN <= " & addDate(pDate2)
End If

If pWhere_Rel <> "" Then
    cString = cString & " AND FILE1_11.CODE IN (" & pWhere_Rel & ")"
End If

If Not IsNull(loctable!GENDER) Then cString = cString & " AND COALESCE(GENDER,1) = " & loctable!GENDER
myRecordSet.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until myRecordSet.EOF
    If IsDate(myRecordSet!DATE_BIRTH & "") Then
       nAge = Age(myFormat(myRecordSet!DATE_BIRTH), myFormat(IIf(nRelation = 2, pDate2, pDate2)))
    Else
       nAge = 99
    End If
    If (nAge >= mRound(loctable!Age1) Or Val(loctable!Age1 & "") = 0) And (nAge <= Val(loctable!Age2 & "") Or mRound(loctable!Age2 & "") = 0) Then
        addRelation = addRelation + 1
        If nRelation = 1 Then
            aMeet = AddFlag(Empty, "CODE", myRecordSet!CODE)
            aMeet = AddFlag(aMeet, "VALUE", mRound(loctable!Value * mRound(1 - (loctable!discount / 100))))
        End If
    End If
    myRecordSet.MoveNext
Loop
myRecordSet.Close
Set myRecordSet = Nothing
End Function
Private Function AddMeetingMem(ByRef loctable As ADODB.Recordset, pCode As String, pYear_code, pValue As Double, pCon As ADODB.Connection) As Double
Dim cString As String
cField = "SUM(ROUND((RATE/100) * " & pValue & ",2)) AS TOTAL"
cString = "SELECT " & cField & " FROM  MEETING_H INNER JOIN  MEETING ON MEETING_H.CODE = MEETING.meeting"
cString = cString & " WHERE MEETING.MEMBER = " & pCode & " AND (RELATION IS NULL) AND ABSENT = 1"
cString = cString & " AND year_code_collect = " & pYear_code
AddMeetingMem = mRound(GetField(cString, pCon))
End Function
Private Function AddMeetingRel(ByRef loctable As ADODB.Recordset, pCode As String, pYear_code, aMeet, pCon As ADODB.Connection) As Double
Dim cString As String, I As Long
For I = 0 To UBound(aMeet)
    cField = "SUM(ROUND((RATE/100) * " & mRound(retFlag(aMeet(I), "value")) & ",2)) AS TOTAL"
    cString = "SELECT  " & cField & " FROM  MEETING_H INNER JOIN  MEETING ON MEETING_H.CODE = MEETING.meeting"
    cString = cString & " WHERE MEETING.MEMBER = " & pCode & " AND (RELATION = " & retFlag(aMeet(I), "code") & ") AND ABSENT = 1"
    cString = cString & " AND year_code_collect = " & pYear_code
    AddMeetingRel = AddMeetingRel + mRound(GetField(cString, pCon))
Next
End Function
Private Function OverAge(ByRef pCode As String, pYear_code As Variant, pCon As ADODB.Connection) As Double
Dim cString As String, aYear As Variant
sDate = GetField("select date1 from years_codes where code = " & pYear_code, pCon)
OverAge = mRound(GetField("SELECT SUM(1) FROM FILE1_11 WHERE member = " & pCode & " AND GENDER = 1 AND RELATION = 2 AND dbo.f_age(FILE1_11.DATE_BIRTH ," & addstring(myFormat(sDate)) & ") > 24 AND FILE1_11.HANDI = 0 AND FILE1_11.PENDING = 0", pCon))
End Function
Private Function OverAgeAll(ByRef pCode As String, pYear_code As Variant, pCon As ADODB.Connection) As Double
Dim cString As String, aYear As Variant
sDate = GetField("select date1 from years_codes where code = " & pYear_code, pCon)
OverAgeAll = mRound(GetField("SELECT SUM(1) FROM FILE1_11 WHERE member = " & pCode & " AND GENDER = 1 AND RELATION = 2 AND dbo.f_age(FILE1_11.DATE_BIRTH ," & addstring(myFormat(sDate)) & ") > 24", pCon))
End Function
Public Function DocSameDay(pCode As String, pType As String, pDate, pCon As ADODB.Connection) As Variant
DocSameDay = GetField("select top 1 doc_no from file6_20h where code = " & addvalue(pCode) & " and type = " & addvalue(pType) & " and date = " & DateSq(pDate), pCon)
End Function


