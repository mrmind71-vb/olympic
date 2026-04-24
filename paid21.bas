Attribute VB_Name = "Paid2"
Public cWhere_rel As String
Public Const MAX_YEARS = 4
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
Function validClaim(pCode As String, pDate As String, pType As String, pCon As ADODB.Connection, Optional pMaxYears As Integer = MAX_YEARS, Optional bReNew As Boolean = False, Optional bAdmit As Boolean) As String
Dim atype As Variant, nYear_code As Variant
Dim nValueTax As Double, nValueInstall As Double

Dim ntype
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


If (nYears > pMaxYears And (Not bReNew)) Then
    validClaim = "⁄œœ «·„Ê«”„ «ﬂ»— „‰ " & pMaxYears & " „Ê«”„"
    Exit Function
End If

'nValueTax = mRound(MyFuncValue("[dbo].[fn_tax_member_all]", pCon, "null", pCode, "null"))
'nValueTax = nValueTax - MyFuncValue("[dbo].[fn_tax_paid]", pCon, pCode)

'nValueTax = mRound(MyFuncValue("[dbo].[fn_tax_member_all]", pCon, "null", pCode, "null"))
nValueTax = mRound(MyFuncValue("[dbo].[fn_tax_member_bal]", pCon, "null", pCode, addstring(myFormat(pDate))))

Dim cmd As New ADODB.Command
aInsert = AddFlag(Empty, "code", pCode)
aInsert = AddFlag(aInsert, "year_code", nYear_code)
Set cmd = myCmdEx("[dbo].[Member_Admit_value]", pCon, aInsert)
nValueInstall = mRound(cmd.Parameters("@value").Value)

If mRound(retFlag(atype, "TYPE")) = 100 Then
    If nValueTax = 0 And nValueInstall = 0 Then
        validClaim = "·Ì” ⁄·Ì «·⁄÷Ê ›—Êﬁ ﬁÌ„… „÷«›… «Ê ›—Êﬁ «ﬁ”«ÿ"
        Exit Function
    End If
    If Not bAdmit Then
        If cmd.Parameters("@must_admit").Value Then
            validClaim = "«·⁄÷Ê ⁄·ÌÂ ›—Êﬁ ÷—Ì»… „÷«›… ⁄·Ì «·«ﬁ”«ÿ »ﬁÌ„… " & cmd.Parameters("@value").Value & " Ê·„ ÌﬁÊ„ »⁄„· «ﬁ—«—"
            Exit Function
        End If
    End If
    validClaim = "ok"
    Exit Function
End If


If cmd.Parameters("@must_admit").Value Then
    validClaim = "«·⁄÷Ê ⁄·ÌÂ ›—Êﬁ ÷—Ì»… „÷«›… ⁄·Ì «·«ﬁ”«ÿ »ﬁÌ„… " & cmd.Parameters("@value").Value & " Ê·„ ÌﬁÊ„ »⁄„· «ﬁ—«—"
    Exit Function
End If

If nValueTax <> 0 Then
    validClaim = "⁄·Ì «·⁄÷Ê ›—Êﬁ ﬁÌ„… „÷«›… »ﬁÌ„… " & nValueTax
    If MyFuncValue("dbo.fn_tax_paid", pCon, pCode) = 0 Then
        Exit Function
    End If
End If

If retFlag(atype, "ONCE") Then
    Dim pDoc_Before As Variant
    pDoc_Before = paid_once(pType, nYear_code, pCode, pCon)
    If Not IsNull(pDoc_Before) Then
        validClaim = "«·⁄÷Ê ÿ»⁄ „ÿ«·»… " & retFlag(atype, "Desca") & " »—ﬁ„ " & pDoc_Before
        Exit Function
    End If
End If

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
ElseIf mRound(retFlag(atype, "TYPE")) = 200 Or mRound(retFlag(atype, "TYPE")) = 300 Then
    Dim nReturn As Variant
    nReturn = GetField("select year_code from vw_last_paid_printed where code = " & pCode, pCon)
    If IsEmpty(nReturn) Then
        validClaim = "«·⁄÷Ê ·„ Ì”œœ „‰ ﬁ»·"
        Exit Function
    End If
    If Val(nReturn) < nYear_code Then
        validClaim = "«·⁄÷Ê ·„ Ì”œœ «·”‰… «·Õ«·Ì…"
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
Public Function addPayment(pCode As String, pDate As String, pType As String, pCon As ADODB.Connection, Optional pFile As String = "FILE6_20", Optional pFileHeader As String = "FILE6_20H", Optional pMaxYears As Integer = MAX_YEARS, Optional bFawry As Boolean, Optional bReNew As Boolean = False, Optional bAdmit As Boolean = False, Optional pItem As String) As Variant
Dim aMember As Variant, aPaid As Variant, aYear As Variant, aUnPaid As Variant
Dim nFirst_year As Integer, nFirst_year1 As Integer, nFirst_year2 As Integer, nFirst_year3 As Integer
Dim nYear_code As String, nYear_Code1 As Long, nYear_code2 As Long, nYear_code3 As Long
Dim sDoc_no As String, aInsert As Variant, sError As String
Dim cInsertHeader As String
Dim nValue As Double
Dim I As Integer, cInsert As String

sReturn = ValidPay(pCode, pDate, pType, pCon, pMaxYears, bReNew, bAdmit)

If sReturn <> "ok" Then
    addPayment = AddFlag(addPayment, "error", sReturn)
    Exit Function
End If

nYear_code = Ret_Year(pDate, "CODE", pCon)

Dim ntype As Integer
ntype = mRound(Claim_Type_Load(pType, "TYPE", pCon))

cWhere_rel = ""
If ntype = 3 Then
    addrelfrm.sCode = pCode
    addrelfrm.pType = pType
    addrelfrm.pYear_code = nYear_code
    addrelfrm.Show 1
End If

If cWhere_rel = "" And pType = 3 Then
    addPayment = AddFlag(addPayment, "error", "·„   „ «÷«›… «Ì «ﬁ«—»")
    Exit Function
End If

nOverAge = OverAgeAll(pCode, nYear_code, pCon)
If nOverAge > 0 Then
    addPayment = AddFlag(addPayemnt, "msg", ArbString("··⁄÷Ê „‰ „ ÕœÌ «·«⁄«ﬁ… «Ê „ Êﬁ›Ì «·⁄÷ÊÌ…" & nOverAge & " «»‰«¡ ›Êﬁ «·”‰"))
End If

Dim cmd As New ADODB.Command
aInsert = AddFlag(Empty, "code", pCode)
aInsert = AddFlag(aInsert, "year_code", nYear_code)
Set cmd = myCmdEx("[dbo].[Member_Admit_value]", pCon, aInsert)
If cmd.Parameters("@value").Value > 0 Then
    If mRound(cmd.Parameters("@value").Value) = mRound(cmd.Parameters("@tax_diff").Value) Then
        addPayment = AddFlag(addPayemnt, "msg2", ArbString("«·⁄÷Ê ⁄·ÌÂ ›—Êﬁ ÷—Ì»… „÷«›… ⁄·Ì «·«ﬁ”«ÿ »ﬁÌ„… " & cmd.Parameters("@value").Value & " Êﬁ«„ »⁄„· «ﬁ—«—"))
'    ElseIf Not IsNull(cmd.Parameters("@year_code")) Then
'        If cmd.Parameters("@year_code") < nYear_code Then
'            addPayment = AddFlag(addPayemnt, "msg2", ArbString("«·⁄÷Ê ⁄·ÌÂ ›—Êﬁ ÷—Ì»… „÷«›… ⁄·Ì «·«ﬁ”«ÿ »ﬁÌ„… " & cmd.Parameters("@value").Value & " „‰ «·„Ê”„ «·”«»ﬁ"))
'        End If
    End If
End If


sDoc_no = Newflag(pFileHeader, "DOC_NO", pCon)
aInsert = AddFlag(Empty, "DOC_NO", addVal(sDoc_no))
aInsert = AddFlag(aInsert, "[DATE]", addDate(pDate))
aInsert = AddFlag(aInsert, "[DATE_ISSUE]", addDate(pDate))
aInsert = AddFlag(aInsert, "[CODE]", addvalue(pCode))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(pType))
aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(nYear_code))
aInsert = AddFlag(aInsert, "[YEARS]", "1")
aInsert = AddFlag(aInsert, "[RENEW]", IIf(bReNew, 1, 0))
aInsert = AddFlag(aInsert, "[ADMIT]", IIf(bAdmit, 1, 0))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName & " [" & GetComputerName & "]"))
aInsert = AddFlag(aInsert, "[TIME]", "getdate()")
If bFawry Then aInsert = AddFlag(aInsert, "IsFawry", "1")

cInsert_Header = addInsert(aInsert, pFileHeader)
aSql = AddFlag(aSql, cInsert_Header)

Dim rsYears As New ADODB.Recordset
If ntype = 1 Or ntype = 2 Then
    Set rsYears = UnPaidYears(pCode, nYear_code, pCon)
    Do Until rsYears.EOF
        If rsYears.RecordCount = 1 Then
            cInsert = addPaidItems(sDoc_no, pCode, rsYears!code, nYear_code, pType, pDate, 0, cWhere_rel, pCon, pFile, pDate)
        ElseIf rsYears!code = nYear_code Then
            I = I + 1
            cInsert = addPaidItems(sDoc_no, pCode, rsYears!code, nYear_code, pType, pDate, 0, cWhere_rel, pCon, pFile, pDate) & ";"
        ElseIf Not rsYears!NO_LATE Then
            I = I + 1
            If pType = 2 Then
                cInsert = addPaidItems(sDoc_no, pCode, rsYears!code, nYear_code, 1, pDate, I, cWhere_rel, pCon, pFile, pDate) & ";"
            Else
                cInsert = addPaidItems(sDoc_no, pCode, rsYears!code, nYear_code, 1, pDate, I, cWhere_rel, pCon, pFile, pDate) & ";"
            End If
        Else
            cInsert = addPaidItems(sDoc_no, pCode, rsYears!code, nYear_code, 1, pDate, 0, cWhere_rel, pCon, pFile, pDate) & ";"
        End If
        aSql = AddFlag(aSql, cInsert)
        rsYears.MoveNext
    Loop
Else
    cInsert = addPaidItems(sDoc_no, pCode, nYear_code, nYear_code, pType, pDate, 0, cWhere_rel, pCon, pFile, pDate, pItem)
    aSql = AddFlag(aSql, cInsert)
End If

If UBound(aSql) > 0 Then
    If pFileHeader = "FILE6_20H" Then
        aSql = AddFlag(aSql, fixYears(sDoc_no))
        aSql = AddFlag(aSql, fixClaim(sDoc_no))
    Else
        aSql = AddFlag(aSql, fixYears2(sDoc_no))
        aSql = AddFlag(aSql, fixClaimOther(sDoc_no, pFileHeader))
    End If
    addPayment = AddFlag(addPayment, "sql", aSql)
    addPayment = AddFlag(addPayment, "doc_no", sDoc_no)
End If
End Function
Public Function addPaidItems(pDoc_No As String, pCode As String, pYear_code As String, pYear_code_Doc As String, pType As String, pDate_Paid, pLate As Integer, pWhere_Rel As String, pCon As ADODB.Connection, Optional pFile As String = "FILE6_20", Optional pDate As String, Optional pItem As String = "") As String
Dim cString As String, nAge As Long, bMemberAdded As Boolean
Dim pSection As String, pDate1 As String, pDate2 As String, aMeetMem As Double, aMeetRel As Variant, aMeetSub As Variant
Dim nAll As Long, aPaid As Variant, bApg As Boolean
Dim nMeetingValueMember As Double, nMeetingValueRelation As Double
Dim nLate As Long, nLate_Tax As Double, nValueTax As Double
Dim nMeetWife As Integer
Dim cmd As ADODB.Command

aYear = GetFields("select * from years_codes where code = " & pYear_code)

pDate1 = myFormat(retFlag(aYear, "date1"))
pDate2 = myFormat(retFlag(aYear, "date2"))

pSection = Member_Load(pCode, "TYPE")
aPaid = Member_Paid(pCode, , pCon)
atype = Claim_Type_Load(pType, , pCon)
'nAll = retAll(pCode, pDate, atype, pCon)
bApg = Member_Load(pCode, "apg", pCon)


'aRates = Array(0, 50, 100, 200)
'nLate = aRates(IIf(pLate > 3, 3, pLate))

nLate = mRound(retFlag(aYear, "Late_Rate"))

If pYear_code <> pYear_code_Doc Then
    nLate_Tax = mRound(MyFuncValue("dbo.fn_TaxRate", pCon, IIf(mRound(pYear_code) < 9, "9", pYear_code), addstring(myFormat(pDate_Paid))) / 100, 2)
End If

If pItem = "" Then
    cString = "SELECT FILE6_11.ITEM,FILE6_11.TYPE,FILE6_11.YEAR_CODE,FILE6_10.AGE1,FILE6_10.AGE2 ,FILE6_10.DESCA, FILE6_10.ALLMEMBER, FILE6_10.LATE, FILE6_10.RELATION," & _
          " FILE6_10.ISMEMBER,FILE6_10.TAX_LATE, COALESCE(FILE6_10.AGE1,0), COALESCE(FILE6_10.AGE2,0), FILE6_10.GENDER, " & _
          " FILE6_10.BASICDIED,FILE6_10.LATE_DAYS, FILE6_10.BASICNEW,FILE6_10.BASICOLD, FILE6_10.MEETING,FILE6_11.TAX, FILE6_10.DAYS, FILE6_10.NORATE, " & _
          " FILE6_11.value, FILE6_11.Discount,TAX_LATE_INSTALL " & _
          " FROM FILE6_10 INNER JOIN FILE6_11 ON FILE6_10.ITEM = FILE6_11.item " & _
          " WHERE FILE6_11.TYPE = " & pType & _
          " AND FILE6_11.BASIC = 1 " & _
          " AND FILE6_11.YEAR_CODE = " & IIf(mRound(pYear_code) < 9, "9", pYear_code) & _
          " AND [SECTION] =  " & pSection
    cString = cString & " ORDER BY FILE6_10.ITEM"
Else
    cString = "SELECT FILE6_11.ITEM,FILE6_11.TYPE,FILE6_11.YEAR_CODE,FILE6_10.AGE1,FILE6_10.AGE2 ,FILE6_10.DESCA, FILE6_10.ALLMEMBER, FILE6_10.LATE, FILE6_10.RELATION," & _
          " FILE6_10.ISMEMBER,FILE6_10.TAX_LATE, COALESCE(FILE6_10.AGE1,0), COALESCE(FILE6_10.AGE2,0), FILE6_10.GENDER, " & _
          " FILE6_10.BASICDIED,FILE6_10.LATE_DAYS, FILE6_10.BASICNEW,FILE6_10.BASICOLD, FILE6_10.MEETING,FILE6_11.TAX, FILE6_10.DAYS, FILE6_10.NORATE, " & _
          " FILE6_11.value, FILE6_11.Discount,TAX_LATE_INSTALL " & _
          " FROM FILE6_10 INNER JOIN FILE6_11 ON FILE6_10.ITEM = FILE6_11.item " & _
          " WHERE FILE6_11.TYPE = " & pType & _
          " AND FILE6_11.YEAR_CODE = " & IIf(mRound(pYear_code) < 9, "9", pYear_code) & _
          " AND [SECTION] =  " & pSection & _
          " AND FILE6_11.ITEM = " & pItem
    cString = cString & " ORDER BY FILE6_10.ITEM"
End If

Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText

Dim bDied As Boolean
bDied = Member_Load(pCode, "Died")
If bDied Then bMemberAdded = True

Do Until loctable.EOF
    If loctable!MEETING Then
        If nMeetingValueMember <> 0 And loctable!isMember Then
            aInsert = AddFlag(Empty, "doc_no", pDoc_No)
            aInsert = AddFlag(aInsert, "item", loctable!Item)
            aInsert = AddFlag(aInsert, "value", mRound(nMeetingValueMember / 2, 2))
            aInsert = AddFlag(aInsert, "quant", 1)
            aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
            aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
            If loctable!late Then
                aInsert = AddFlag(aInsert, "late_rate", nLate)
            Else
                aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
            End If
            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
            cInsert = cInsert & addInsert(aInsert, pFile) & ";"
        ElseIf nMeetingValueRelation <> 0 And Not IsNull(loctable!RELATION) Then
            aInsert = AddFlag(Empty, "doc_no", pDoc_No)
            aInsert = AddFlag(aInsert, "item", loctable!Item)
            aInsert = AddFlag(aInsert, "value", mRound(nMeetingValueRelation / 2, 2))
            aInsert = AddFlag(aInsert, "quant", 1)
            aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
            aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
            If loctable!late Then
                aInsert = AddFlag(aInsert, "late_rate", nLate)
            Else
                aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
            End If
            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
            cInsert = cInsert & addInsert(aInsert, pFile) & ";"
        End If
    ElseIf loctable!LATE_DAYS Then
        If IsDate(myFormat(retFlag(aYear, "DATE_LATE"))) Then
            If myFormat(pDate_Paid) >= myFormat(retFlag(aYear, "DATE_LATE")) Then
                aInsert = AddFlag(Empty, "doc_no", pDoc_No)
                aInsert = AddFlag(aInsert, "item", loctable!Item)
                aInsert = AddFlag(aInsert, "value", nValue)
                aInsert = AddFlag(aInsert, "quant", 1)
                aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
                aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
                If loctable!late Then
                    aInsert = AddFlag(aInsert, "late_rate", nLate)
                Else
                    aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
                End If
                aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
                cInsert = cInsert & addInsert(aInsert, pFile) & ";"
            End If
        End If
    ElseIf loctable!isMember Then
        If (Not bMemberAdded) Then
            If AddMemberData(pCode, mRound(loctable!Age1), mRound(loctable!Age2), pDate2, pCon) Then
                nAll = nAll + 1
                aInsert = AddFlag(Empty, "doc_no", pDoc_No)
                aInsert = AddFlag(aInsert, "item", loctable!Item)
                aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
                aInsert = AddFlag(aInsert, "quant", "1")
                aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
                aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
                If loctable!late Then
                    aInsert = AddFlag(aInsert, "late_rate", nLate)
                Else
                    aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
                End If
                aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
                cInsert = cInsert & addInsert(aInsert, pFile) & ";"
                If RetMeet(pCode, , pYear_code, "0", pCon) Then
                    nMeetingValueMember = nMeetingValueMember + mRound(loctable!Value * mRound(1 - (loctable!discount / 100)))
                End If
                bMemberAdded = True
            End If
        End If
    ElseIf (Not IsNull(loctable!RELATION)) Then
        nMeetWife = 0
        nRelation = addRelation(loctable, pCode, pDate1, pDate2, aMeetSub, loctable!RELATION, atype, pWhere_Rel, pYear_code_Doc, pCon, nMeetWife)
        If nRelation > 0 Then
            nAll = nAll + nRelation
            aInsert = AddFlag(Empty, "doc_no", pDoc_No)
            aInsert = AddFlag(aInsert, "item", loctable!Item)
            aInsert = AddFlag(aInsert, "value", mRound(loctable!Value))
            aInsert = AddFlag(aInsert, "quant", nRelation)
            aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
            aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
            If loctable!late Then
                aInsert = AddFlag(aInsert, "late_rate", nLate)
            Else
                aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
            End If
            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
            cInsert = cInsert & addInsert(aInsert, pFile) & ";"
            
            If loctable!RELATION = "1" Then
                If nMeetWife > 0 Then
                    nMeetingValueRelation = nMeetingValueRelation + mRound((nMeetWife * loctable!Value) * mRound(1 - (loctable!discount / 100)))
                End If
            End If
        End If
    ElseIf mRound(loctable!Age1) <> 0 Then
        nAge = GetField("SELECT dbo.f_age(DATE_BIRTH, " & addstring(myFormat(pDate2)) & ") FROM FILE1_10 WHERE CODE = " & pCode)
        If nAge >= mRound(loctable!Age1) Then
            aInsert = AddFlag(Empty, "doc_no", pDoc_No)
            aInsert = AddFlag(aInsert, "item", loctable!Item)
            aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
            aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
            aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
            aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
            If loctable!late Then
                aInsert = AddFlag(aInsert, "late_rate", nLate)
            Else
                aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
            End If
            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
            cInsert = cInsert & addInsert(aInsert, pFile) & ";"
        End If
    ElseIf mRound(loctable!Age2) <> 0 Then
        nAge = GetField("SELECT dbo.f_age(DATE_BIRTH, " & addstring(myFormat(pDate2)) & ") FROM FILE1_10 WHERE CODE = " & pCode)
        If nAge <= mRound(loctable!Age2) Then
            aInsert = AddFlag(Empty, "doc_no", pDoc_No)
            aInsert = AddFlag(aInsert, "item", loctable!Item)
            aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
            aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
            aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
            aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
            If loctable!late Then
                aInsert = AddFlag(aInsert, "late_rate", nLate)
            Else
                aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
            End If
            aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
            cInsert = cInsert & addInsert(aInsert, pFile) & ";"
        End If
    ElseIf loctable!TAX_LATE Then
        nValueTax = mRound(MyFuncValue("[dbo].[fn_tax_member_bal]", pCon, "null", pCode, addstring(myFormat(pDate))))
            If nValueTax > 0 Then
                aInsert = AddFlag(Empty, "doc_no", pDoc_No)
                aInsert = AddFlag(aInsert, "item", loctable!Item)
                aInsert = AddFlag(aInsert, "value", nValueTax)
                aInsert = AddFlag(aInsert, "quant", 1)
                aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
                aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
                aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
                cInsert = cInsert & addInsert(aInsert, pFile) & ";"
            End If
    ElseIf loctable!TAX_LATE_INSTALL Then
        Set cmd = New ADODB.Command
        aInsert = AddFlag(Empty, "code", pCode)
        aInsert = AddFlag(aInsert, "year_code", pYear_code)
        Set cmd = myCmdEx("[dbo].[Member_Admit_value]", pCon, aInsert)
        nValue = cmd.Parameters("@value").Value
        If nValue <> 0 Then
                aInsert = AddFlag(Empty, "doc_no", pDoc_No)
                aInsert = AddFlag(aInsert, "item", loctable!Item)
                aInsert = AddFlag(aInsert, "value", nValue)
                aInsert = AddFlag(aInsert, "quant", 1)
                aInsert = AddFlag(aInsert, "discount_rate", mRound(loctable!discount))
                aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
                aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
                cInsert = cInsert & addInsert(aInsert, pFile) & ";"
        End If
    Else
        aInsert = AddFlag(Empty, "doc_no", pDoc_No)
        aInsert = AddFlag(aInsert, "item", loctable!Item)
        aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
        aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
        aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
        'aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX) + IIf(mRound(loctable!TAX) <> 0, mRound(loctable!TAX) * nLate_Tax, 0))
        aInsert = AddFlag(aInsert, "tax_rate", mRound(loctable!TAX))
        If loctable!late Then
            aInsert = AddFlag(aInsert, "late_rate", nLate)
        Else
            aInsert = AddFlag(aInsert, "late_rate", mRound(loctable!TAX) * nLate_Tax)
        End If
        aInsert = AddFlag(aInsert, "YEAR_CODE", pYear_code)
        cInsert = cInsert & addInsert(aInsert, pFile) & ";"
    End If
    loctable.MoveNext
Loop
addPaidItems = cInsert
End Function
Public Function addInstall(pCode As String, pDate As String, pCon As ADODB.Connection, Optional ByVal pDoc_No As String, Optional nTotalMax As Double = -1, Optional pFawry As Boolean = False, Optional nCard_value As Double, Optional nOther As Double = 0) As Variant
Dim cString As String, sDoc_no As String

Dim cInsert As String, cInsert_Header As String
Dim nLate As Long

If Not ValidNum(pDoc_No) Then
    cString = "SELECT  " & IIf(nTotalMax = -1, " TOP 1 ", "") & " ID,dbo.f_install_rest(file6_21.ID) AS REST,FILE6_21.RATE_TAX,FILE6_21.CHARGE,FILE6_21.TAX_DIFF FROM FILE6_21 WHERE FILE6_21.CODE = " & pCode & _
              " AND dbo.f_install_rest(file6_21.ID) > 0 ORDER BY FILE6_21.DATE_DUE"
Else
    cString = "SELECT " & IIf(nTotalMax = -1, " TOP 1 ", "") & " ID,dbo.f_install_rest(file6_21.ID) AS REST,FILE6_21.RATE_TAX,FILE6_21.CHARGE,FILE6_21.TAX_DIFF FROM FILE6_21 WHERE FILE6_21.CODE = " & pCode & _
              " AND dbo.f_install_rest_doc(file6_21.ID," & pDoc_No & ") > 0 ORDER BY FILE6_21.DATE_DUE"
End If

Dim bInterest As Boolean
bInterest = myNull(Member_Load_install(pCode, "interest", pCon), False)

Dim loctable As New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    If pDoc_No = "" Then
        sDoc_no = Newflag("FILE6_30H", "DOC_NO", pCon)
        aInsert = AddFlag(Empty, "DOC_NO", addVal(sDoc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(pDate))
        aInsert = AddFlag(aInsert, "[DATE_ISSUE]", addDate(pDate))
        aInsert = AddFlag(aInsert, "[CODE]", addvalue(pCode))
        aInsert = AddFlag(aInsert, "[Card_value]", nCard_value)
        aInsert = AddFlag(aInsert, "[other]", nOther)
        aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
        aInsert = AddFlag(aInsert, "isFawry", IIf(pFawry, 1, 0))
        aInsert = AddFlag(aInsert, "[TIME]", "getdate()")
        cInsert_Header = addInsert(aInsert, "FILE6_30H")
    Else
        sDoc_no = pDoc_No
    End If
    
    
    Dim nTotal As Double, nRate As Double, nTotalRow As Double
    Dim nRateInt As Double
    
    Do Until loctable.EOF
        nTotalRow = mRound(loctable!Rest)
        nTotalRow = nTotalRow * (1 + (mRound(loctable!Rate_Tax / 100, 2)))
        nTotalRow = nTotalRow + loctable!TAX_DIFF
        nTotal = nTotal + nTotalRow
        
        Dim cmdInt As ADODB.Recordset
        Set cmdInt = myCmdInterest(loctable!ID, pDate, pCon)
        If Not cmdInt.EOF Then
            nRate = mRound(cmdInt!Rate, 2)
        Else
            nRate = 0
        End If
        
        
        If nTotalMax >= nTotal Or nTotalMax = -1 Then
            aInsert = AddFlag(Empty, "doc_no", sDoc_no)
            aInsert = AddFlag(aInsert, "late_id", loctable!ID)
            aInsert = AddFlag(aInsert, "value", mRound(loctable!Rest))
            aInsert = AddFlag(aInsert, "tax_rate", loctable!Rate_Tax)
            aInsert = AddFlag(aInsert, "tax_diff", loctable!TAX_DIFF)
            aInsert = AddFlag(aInsert, "charge", loctable!CHARGE)
            If bInterest Then aInsert = AddFlag(aInsert, "Rate_interest", cmdInt!Rate)
            cInsert = cInsert & addInsert(aInsert, "FILE6_30") & ";"
        End If
        If nTotal >= nTotalMax Then Exit Do
        loctable.MoveNext
    Loop
    loctable.Close
    Set loctable = Nothing
    
    If cInsert <> "" Then
        If cInsert_Header <> "" Then
            aSql = AddFlag(aSql, cInsert_Header)
        End If
        aSql = AddFlag(aSql, cInsert)
        addInstall = AddFlag(addInstall, "doc_no", sDoc_no)
        addInstall = AddFlag(addInstall, "sql", aSql)
    End If
Else
    addInstall = AddFlag(Empty, "error", "·« ÌÊÃœ «ﬁ”«ÿ ··”œ«œ")
End If
End Function
Public Function myCmdInterest(pId As String, pDate As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "id", pId)
aPrm = AddFlag(aPrm, "date", myFormat(pDate))
Set myCmdInterest = myCmdProc("dbo.sp_interest_total", pCon, aPrm)
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
AddMemberData = True
End Function
Private Function addRelation(ByRef loctable As ADODB.Recordset, pCode As String, pDate1, pDate2, ByRef aMeet As Variant, nRelation As Integer, atype As Variant, pWhere_Rel As String, pYear_code_Doc As String, pCon As ADODB.Connection, ByRef nMeetWife As Integer) As Integer
Dim myRecordSet As New ADODB.Recordset
Dim nAge As Integer, nGender As Integer
cString = " SELECT [CODE],[DATE_BIRTH],COALESCE(GENDER,1) From FILE1_11"
cString = cString & " where relation = " & nRelation
cString = cString & " AND MEMBER = " & pCode

If loctable!year_code = pYear_code_Doc Then
    cString = cString & " AND PENDING = 0"
End If

If nRelation = 2 Then
    If retFlag(atype, "over_age") Then
        cString = cString & " AND (NOT(GENDER = 1 AND HANDI = 0  AND dbo.f_age(FILE1_11.DATE_BIRTH ," & addstring(myFormat(pDate2)) & ") > 24))"
    End If
End If
    
If loctable!Type = 3 Then
    cString = cString & " AND (DATE_BEGIN IS NULL OR DATE_BEGIN >= " & addDate(pDate1) & ")"
Else
    cString = cString & " AND (DATE_BEGIN IS NULL OR  DATE_BEGIN <= " & addDate(pDate2) & ")"
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
       nAge = 55
    End If
    If (nAge >= mRound(loctable!Age1) Or Val(loctable!Age1 & "") = 0) And (nAge <= Val(loctable!Age2 & "") Or mRound(loctable!Age2 & "") = 0) Then
        addRelation = addRelation + 1
        If nRelation = 1 Then
            nMeetWife = nMeetWife + IIf(RetMeet(pCode, myRecordSet!code, loctable!year_code, "0", pCon), 1, 0)
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
Public Function DocSameDay_i(pCode As String, pDate, pCon As ADODB.Connection) As Variant
DocSameDay_i = GetField("select top 1 doc_no from file6_30h where code = " & addvalue(pCode) & " and date = " & DateSq(pDate), pCon)
End Function
Public Function ValidPay(pCode As String, pDate As String, pType As String, pCon As ADODB.Connection, Optional pMaxYears As Integer = MAX_YEARS, Optional bReNew As Boolean = False, Optional bAdmit As Boolean) As String
Dim sReturn As String
ValidPay = validMemberPay(pCode, pCon)
If ValidPay = "ok" Then
    ValidPay = validClaim(pCode, pDate, pType, pCon, pMaxYears, bReNew, bAdmit)
End If
End Function
Private Function RetMeet(pCode As String, Optional pRel As String = "0", Optional pYear As String = "null", Optional pType As String = "null", Optional pCon As ADODB.Connection) As Boolean
RetMeet = GetField("select dbo.f_meeting_type(" & pCode & "," & pRel & "," & pYear & "," & pType & ")", pCon, False)
End Function
