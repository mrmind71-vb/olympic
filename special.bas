Attribute VB_Name = "Special"
Public aAppPath As String
Public aAddress As Variant
Public aPaidTypes As Variant
Public sDateFormat As String
Public cUserName As String
Public Const DATE_TAX1 = "2016-09-01", DATE_TAX2 = "2017-12-31"
Public Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public nUsercode As String, aSec As Variant, sSeason As String, aSeason As Variant, sDate_Season As String
Sub MemberLookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(4, 1)
Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE1_10.CODE,FILE1_10.DESCA,TYPE_CODES.DESCA,Ses_no,file1_10.code_main " & _
                  " From FILE1_10 INNER JOIN TYPE_CODES ON FILE1_10.[TYPE] = TYPE_CODES.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE1_10.DESCA"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE1_10.DESCA%% OR **FILE1_10.CODE**)"

listarray(1, 0) = "‰Ê⁄ «·⁄÷ÊÌ…-—ﬁ„ «·„Ê«›ﬁ…"
listarray(1, 1) = "(%%TYPE_CODES.DESCA%% OR SES_NO LIKE '%cFilter' or SES_NO LIKE 'cFilter%')"

listarray(2, 0) = "›«’· „‰"
listarray(2, 1) = "(**file1_10.code_main**)"

GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «·⁄÷Ê"
GrdArray(1, 1) = 3500

GrdArray(2, 0) = "‰Ê⁄ «·⁄÷ÊÌ…"
GrdArray(2, 1) = 2500

GrdArray(3, 0) = "—ﬁ„ «·„Ê«›ﬁ…"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "›«’· „‰"
GrdArray(4, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.nMax_records = 1000
oSearch.sCaption = "≈” ⁄·«„ «·«⁄÷«¡"
oSearch.Show 1
End Sub
Sub MemberLookupInstall(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE2_10.CODE,FILE2_10.DESCA,TYPE_CODES.DESCA " & _
                  " From FILE2_10 INNER JOIN TYPE_CODES ON FILE2_10.[TYPE] = TYPE_CODES.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE2_10.DESCA"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE2_10.DESCA%% OR **FILE2_10.CODE**)"

listarray(1, 0) = "‰Ê⁄ «·⁄÷ÊÌ…"
listarray(1, 1) = "(%%TYPE_CODES.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «·⁄÷Ê"
GrdArray(1, 1) = 4500

GrdArray(2, 0) = "‰Ê⁄ «·⁄÷ÊÌ…"
GrdArray(2, 1) = 2500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.nMax_records = 2000
oSearch.sCaption = "≈” ⁄·«„ «·«⁄÷«¡ «· ﬁ”Ìÿ"
oSearch.Show 1
End Sub
Sub Member_InLookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE2_10.CODE,FILE2_10.DESCA,TYPE_CODES.DESCA,NOTES " & _
                  " From FILE2_10 INNER JOIN TYPE_CODES ON FILE2_10.[TYPE] = TYPE_CODES.CODE"
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE2_10.DESCA"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE2_10.DESCA%% OR **FILE2_10.CODE**)"

listarray(1, 0) = "«·„·ÕÊŸ…"
listarray(1, 1) = "(%%NOTES%%)"

GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «·⁄÷Ê"
GrdArray(1, 1) = 3500

GrdArray(2, 0) = "‰Ê⁄ «·⁄÷ÊÌ…"
GrdArray(2, 1) = 0

GrdArray(3, 0) = "„·ÕÊŸ…"
GrdArray(3, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE2_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.nMax_records = 2000
oSearch.sCaption = "≈” ⁄·«„ «·«⁄÷«¡"
oSearch.Show 1
End Sub
Sub MemberLookupAll_I(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE1_50.CODE,FILE1_50.DESCA,TYPE_CODES.DESCA" & _
                  " From FILE1_50 INNER JOIN TYPE_CODES ON FILE1_50.[TYPE] = TYPE_CODES.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE1_50.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE1_50.DESCA%% OR **FILE1_50.CODE**)"

listarray(1, 0) = "‰Ê⁄ «·⁄÷ÊÌ…"
listarray(1, 1) = "(%%TYPE_CODES.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «·⁄÷Ê"
GrdArray(1, 1) = 3500

GrdArray(2, 0) = "‰Ê⁄ «·⁄÷ÊÌ…"
GrdArray(2, 1) = 2500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_50.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "≈” ⁄·«„ «⁄÷«¡ «·œ⁄Ê…"
oSearch.Show 1
End Sub
Sub MemberHonorLookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE2_10.CODE,FILE2_10.DESCA,SHARE_CODES.DESCA,FILE2_10.JOB,CONVERT(VARCHAR(10),FILE2_10.DATE_PRINT,111)" & _
                  " From FILE2_10 LEFT JOIN SHARE_CODES ON FILE2_10.[SHARE] = SHARE_CODES.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE2_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·ÊŸÌ›…-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE2_10.DESCA%%  OR **FILE2_10.CODE**)"

listarray(1, 0) = "«·Õ’…"
listarray(1, 1) = "(%%SHARE_CODES.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «·⁄÷Ê"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·Õ’…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·ÊŸÌ›…"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = " «—ÌŒ «·ÿ»«⁄…"
GrdArray(4, 1) = 1400


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE2_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "≈” ⁄·«„ «·«⁄÷«¡ «·›Œ—ÌÌ‰"
oSearch.Show 1
End Sub
Sub WorkerLookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE8_10.CODE,FILE8_10.DESCA,FILE8_10.JOB,CONVERT(VARCHAR(10),FILE8_10.DATE_PRINT,111)" & _
                  " From FILE8_10"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE8_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·ÊŸÌ›…-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE8_10.DESCA%% OR **FILE8_10.CODE**)"


GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «·⁄÷Ê"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·ÊŸÌ›…"
GrdArray(2, 1) = 4000

GrdArray(3, 0) = " «—ÌŒ «·ÿ»«⁄…"
GrdArray(3, 1) = 1400


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE8_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "≈” ⁄·«„ «·«⁄÷«¡ «·›Œ—ÌÌ‰"
oSearch.Show 1
End Sub
Sub BadgeLookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE1_50.CODE,FILE1_50.DESCA,FILE1_50.UNION_NO,FILE1_50.LIC,YES_NO.DESCA" & _
                  " From FILE1_50 LEFT JOIN YES_NO ON FILE1_50.ISENG = YES_NO.CODE "

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE1_50.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·ÊŸÌ›…-—ﬁ„ «·⁄÷Ê-—ﬁ„ «·‰ﬁ«»…-—ﬁ„ «·—Œ’…"
listarray(0, 1) = "(%%FILE1_50.DESCA%% OR %%LIC%% OR %%UNION_NO%% OR **FILE1_50.CODE**)"

listarray(1, 0) = "„Â‰œ”"
listarray(1, 1) = "(**YES_NO.CODE**)"
listarray(1, 2) = "SELECT CODE,DESCA FROM YES_NO"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·ﬁÌœ"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "—ﬁ„ «·—Œ’…"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "„Â‰œ”"
GrdArray(4, 1) = 1400


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_50.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "≈” ⁄·«„ «·«⁄÷«¡ «·›Œ—ÌÌ‰"
oSearch.Show 1
End Sub
Sub ServiceLookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE7_10.CODE,FILE7_10.DESCA,ENG_CODES.DESCA,CONVERT(VARCHAR(10),FILE7_10.DATE_PRINT,111),RELATION" & _
                  " From FILE7_10 LEFT JOIN ENG_CODES ON FILE7_10.DEGREE = ENG_CODES.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE7_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„”·”·"
listarray(0, 1) = "(%%FILE7_10.DESCA%% OR **FILE7_10.CODE**)"

listarray(2, 0) = "«·›∆…"
listarray(2, 1) = "(**FILE7_10.DEGREE**)"
listarray(2, 2) = "SELECT CODE,DESCA FROM ENG_CODES"
listarray(2, 3) = "CODE"
listarray(2, 4) = "DESCA"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·›∆…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = " «—ÌŒ «·ÿ»«⁄…"
GrdArray(3, 1) = 1400

GrdArray(4, 0) = "’·… «·ﬁ—«»…"
GrdArray(4, 1) = 1400

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE2_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "≈” ⁄·«„ ÿ·»… ‰Â«∆Ì Â‰œ”…"
oSearch.Show 1
End Sub
Sub relLookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE1_10.CODE,FILE1_11.CODE,FILE1_11.DESCA,FILE1_10.DESCA,RELATION_CODES.DESCA" & _
                  " From FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE LEFT JOIN RELATION_CODES ON FILE1_11.[RELATION] = RELATION_CODES.CODE"
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE1_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "√”„ «·⁄÷Ê «· «»⁄-—ﬁ„ «·⁄÷Ê «·—∆Ì”Ï"
listarray(0, 1) = "(%%FILE1_11.DESCA%%  OR **FILE1_10.CODE**)"

listarray(1, 0) = "√”„ «·⁄÷Ê «·—∆Ì”Ì"
listarray(1, 1) = "(%%FILE1_10.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê «·—∆Ì”Ì"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ﬂÊœ «· «»⁄"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«”„ «· «»⁄"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "«”„ «·—∆Ì”Ì"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "œ—Ã… «·ﬁ—«»…"
GrdArray(4, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "«” ⁄·«„ «· Ê«»⁄"
oSearch.Show 1
End Sub
Sub relLookupAll2(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE1_11.CODE,FILE1_11.DESCA,RELATION_CODES.DESCA" & _
                  " From FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE LEFT JOIN RELATION_CODES ON FILE1_11.[RELATION] = RELATION_CODES.CODE"
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE1_11.CODE"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "√”„ «·⁄÷Ê «· «»⁄"
listarray(0, 1) = "(%%FILE1_11.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «· «»⁄"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «· «»⁄"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "œ—Ã… «·ﬁ—«»…"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "«” ⁄·«„ «· Ê«»⁄"
oSearch.Show 1
End Sub
Sub relLookupAll_i2(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE2_11.CODE,FILE2_11.DESCA,RELATION_CODES.DESCA" & _
                  " From FILE2_11 INNER JOIN FILE2_10 ON FILE2_11.MEMBER = FILE2_10.CODE LEFT JOIN RELATION_CODES ON FILE2_11.[RELATION] = RELATION_CODES.CODE"
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE2_11.CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "√”„ «·⁄÷Ê «· «»⁄"
listarray(0, 1) = "(%%FILE2_11.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «· «»⁄"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«”„ «· «»⁄"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "œ—Ã… «·ﬁ—«»…"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE1_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "«” ⁄·«„ «· Ê«»⁄"
oSearch.Show 1
End Sub
Sub relLookupAll_I(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE2_10.CODE,FILE2_11.CODE,FILE2_11.DESCA,FILE2_10.DESCA,RELATION_CODES.DESCA" & _
                  " From FILE2_11 INNER JOIN FILE2_10 ON FILE2_11.MEMBER = FILE2_10.CODE LEFT JOIN RELATION_CODES ON FILE2_11.[RELATION] = RELATION_CODES.CODE"
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE2_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "√”„ «·⁄÷Ê «· «»⁄-—ﬁ„ «·⁄÷Ê «·—∆Ì”Ï"
listarray(0, 1) = "(%%FILE2_11.DESCA%%  OR **FILE2_10.CODE**)"

listarray(1, 0) = "√”„ «·⁄÷Ê «·—∆Ì”Ì"
listarray(1, 1) = "(%%FILE2_10.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê «·—∆Ì”Ì"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ﬂÊœ «· «»⁄"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«”„ «· «»⁄"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "«”„ «·—∆Ì”Ì"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "œ—Ã… «·ﬁ—«»…"
GrdArray(4, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE2_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "«” ⁄·«„ «· Ê«»⁄"
oSearch.Show 1
End Sub
Sub Years_LookupAll(oForm As Form, oSearch As Form, Optional sControl As String = "", Optional bAddRow As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select YEARS_CODES.CODE,YEARS_CODES.DESCA_R,CONVERT(VARCHAR(10),YEARS_CODES.DATE1,111),CONVERT(VARCHAR(10),YEARS_CODES.DATE2,111),YEARS_CODES.[YEAR]  " & _
                  " FROM YEARS_CODES"
Generalarray(2) = " ORDER BY DATE1 DESC"
Generalarray(3) = 3000
Generalarray(5) = True

listarray(0, 0) = "«·”‰…"
listarray(0, 1) = "(**[YEAR]**)"

GrdArray(1, 0) = "«·„Ê”„"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "„‰  «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "Õ Ì  «—ÌŒ"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "«·”‰…"
GrdArray(4, 1) = 1500

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "ﬂ· «·„Ê«”„")
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.sCaption = "≈” ⁄·«„ «·„Ê«”„"
oSearch.sControl = sControl
oSearch.Show 1
End Sub
Sub Company_Look(oForm As Form, oSearch As Form, Optional sControl As String = "", Optional bAddRow As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select COMPANY_CODES.CODE,COMPANY_CODES.DESCA " & _
                  " FROM COMPANY_CODES_CODES"
Generalarray(2) = " ORDER BY DESCA DESC"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%COMPANY_CODES.[DESCA]%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 7000

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "ﬂ· «·‘—ﬂ« ")
    aRow = AddFlag(aRow, "col", 1)
End If
1
searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.sCaption = "≈” ⁄·«„ «·‘—ﬂ« "
oSearch.sControl = sControl
oSearch.Show 1
End Sub
Sub Job_Lookup(oForm As Form, oSearch As Form, Optional sControl As String = "", Optional bAddRow As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select JOB_CODES.CODE,JOB_CODES.DESCA " & _
                  " FROM JOB_CODES"
Generalarray(2) = " ORDER BY DESCA DESC"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%JOB_CODES.[DESCA]%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 7000

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "ﬂ· «·ÊŸ«∆›")
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.sCaption = "≈” ⁄·«„ «·ÊŸ«∆›"
oSearch.sControl = sControl
oSearch.Show 1
End Sub
Sub Install_type_Lookup(oForm As Form, oSearch As Form, Optional sControl As String = "", Optional pWhere As String, Optional bAddRow As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select Status_codes.CODE,Status_codes.DESCA " & _
                  " FROM Status_codes"
If pWhere <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & pWhere
End If
Generalarray(2) = " ORDER BY DESCA DESC"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%Status_codes.[DESCA]%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "Õ«·… «·⁄÷Ê"
GrdArray(1, 1) = 7000

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "ﬂ· Õ«·«  «·«⁄÷«¡")
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.sCaption = "≈” ⁄·«„ Õ«·«  «⁄÷«¡ «·ﬁ”ÿ"
oSearch.sControl = sControl
oSearch.Show 1
End Sub
Sub region_Lookup(oForm As Form, oSearch As Form, Optional sControl As String = "", Optional bAddRow As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select region_CODES.CODE,region_CODES.DESCA " & _
                  " FROM region_CODES"
Generalarray(2) = " ORDER BY DESCA DESC"
Generalarray(3) = 4000
Generalarray(5) = True

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%region_CODES.[DESCA]%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 7000

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "ﬂ· «·„‰«ÿﬁ")
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.sCaption = "≈” ⁄·«„ «·ÊŸ«∆›"
oSearch.sControl = sControl
oSearch.Show 1
End Sub

Sub claim_LookupAll(oForm As Form, oSearch As Form, Optional sControl As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select PAID_TYPES.CODE,PAID_TYPES.DESCA" & _
                  " FROM PAID_TYPES"
Generalarray(2) = " ORDER BY CODE"
Generalarray(3) = 3000
Generalarray(5) = True

listarray(0, 0) = "«·‰Ê⁄"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·‰Ê⁄"
GrdArray(1, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "≈” ⁄·«„ «‰Ê«⁄ «·„ÿ«·»« "
oSearch.sControl = sControl
oSearch.Show 1
End Sub
Function ValidDate() As Boolean
ValidDate = Format("01-12-2000", "DD-MM-YYYY") = "01-12-2000"
If Not ValidDate Then
    cString = "‰Ÿ«„ «· «—ÌŒ €Ì— ’«·Õ" & vbCrLf & _
            "«Ã⁄· «·œÊ·… ›Ï ·ÊÕ… «· Õﬂ„" & vbCrLf & _
            "control Panel" & vbCrLf & _
            "Regional and language options" & vbCrLf & _
            "Egypt"
            
    MsgBox ArbString(cString), vbCritical
    End
End If
End Function
Function RetItemBalance(cItem, cStore, dDate, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Store", adVarWChar, adParamInput, 3, cStore)
cmdTable.Parameters.Append cmdTable.CreateParameter("date", adVarWChar, adParamInput, 12, DateConv(dDate))
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 20, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "retitemBalance"
Set obj = cmdTable.Execute
RetItemBalance = Val(cmdTable.Parameters(3).Value & "")
Set obj = Nothing
End Function
Function Inv_Paid(sInv_no As String, pCon As ADODB.Connection) As Double
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("INV_NO", adVarChar, adParamInput, 6, sInv_no)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "INV_PAID"
Set obj = cmdTable.Execute
Inv_Paid = Val(cmdTable.Parameters(1).Value & "")
Set obj = Nothing
End Function
Function LastDoc_card(pMember, pCon As ADODB.Connection, Optional pMember_sub As String = "", Optional bAll As Boolean = False, Optional pSeason As String = "") As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Member", adInteger, adParamInput, , pMember)
If ValidInt(pMember_sub) Then
    cmdTable.Parameters.Append cmdTable.CreateParameter("Member_sub", adInteger, adParamInput, , pMember_sub)
End If
If pSeason <> 0 Then cmdTable.Parameters.Append cmdTable.CreateParameter("Season", adInteger, adParamInput, , pSeason)
cmdTable.CommandText = "MEMBER_DOC" & IIf(ValidInt(pMember_sub), "_SUB", "") & IIf(bAll, "_all", "") & IIf(pSeason <> 0, "_season", "")
Set obj = cmdTable.Execute
If Not obj.EOF Then
    For I = 0 To obj.Fields.Count - 1
        LastDoc_card = AddFlag(LastDoc_card, obj.Fields(I).Name, obj.Fields(I).Value)
    Next
End If
Set obj = Nothing
End Function
Function LastDoc(pCode, pCon As ADODB.Connection, Optional pSeason As Integer = 0) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pCode)
If pSeason <> 0 Then cmdTable.Parameters.Append cmdTable.CreateParameter("Season", adInteger, adParamInput, , pSeason)
cmdTable.CommandText = "LAST_PAID" & IIf(pSeason <> 0, "_season", "")
Set obj = cmdTable.Execute
If Not obj.EOF Then
    For I = 0 To obj.Fields.Count - 1
        LastDoc = AddFlag(LastDoc, obj.Fields(I).Name, obj.Fields(I).Value)
    Next
End If
Set obj = Nothing
End Function
Function Member_Paid(pCode, Optional pField As String = "", Optional pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adInteger, adParamInput, , pCode)
cmdTable.CommandText = "MEMBER_PAID"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Member_Paid = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Member_Paid = AddFlag(Member_Paid, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
End If
Set obj = Nothing
End Function
Function Member_Paid_Install(pCode, Optional pField As String = "", Optional pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adInteger, adParamInput, , pCode)
cmdTable.CommandText = "Member_Paid_Install"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Member_Paid_Install = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Member_Paid_Install = AddFlag(Member_Paid_Install, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
End If
Set obj = Nothing
End Function
Function Member_Load(pCode, Optional pField As String = "", Optional pCon As ADODB.Connection, Optional pProcName As String = "MEMBER_LOAD") As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pCode)
cmdTable.CommandText = pProcName
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Member_Load = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Member_Load = AddFlag(Member_Load, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
End If
Set obj = Nothing
End Function
Function Member_Load_install(pCode, Optional pField As String = "", Optional pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pCode)
cmdTable.CommandText = "MEMBER_LOAD_INSTALL"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Member_Load_install = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Member_Load_install = AddFlag(Member_Load_install, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
End If
Set obj = Nothing
End Function
Function Doc_Totals(pDoc_No As String, Optional pField As String = "", Optional pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("DOC_NO", adInteger, adParamInput, , pDoc_No)
cmdTable.CommandText = "DOC_TOTALS"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Doc_Totals = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Doc_Totals = AddFlag(Doc_Totals, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
End If
Set obj = Nothing
End Function
Function Year_Load(pCode, Optional pField As String = "", Optional pCon As ADODB.Connection, Optional pEmpty As Variant = Empty) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pCode)
cmdTable.CommandText = "YEAR_LOAD"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Year_Load = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Year_Load = AddFlag(Year_Load, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
ElseIf Not IsEmpty(pEmpty) Then
    Year_Load = pEmpty
End If
Set obj = Nothing
End Function
Function Ret_Year(pDate, Optional pField As String = "", Optional pCon As ADODB.Connection, Optional pEmpty As Variant = Empty) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Date", adVarChar, adParamInput, 10, myFormat(pDate))
cmdTable.CommandText = "RET_YEAR"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Ret_Year = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Ret_Year = AddFlag(Ret_Year, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
ElseIf Not IsEmpty(pEmpty) Then
    Ret_Year = pEmpty
End If
Set obj = Nothing
End Function
Function Claim_Type_Load(pCode, Optional pField As String = "", Optional pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
If pCon Is Nothing Then Set cmdTable.ActiveConnection = GetCon Else Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pCode)
cmdTable.CommandText = "claim_type_load"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    If pField <> "" Then
        Claim_Type_Load = obj.Fields(pField).Value
    Else
        For I = 0 To obj.Fields.Count - 1
            Claim_Type_Load = AddFlag(Claim_Type_Load, obj.Fields(I).Name, obj.Fields(I).Value)
        Next
    End If
End If
Set obj = Nothing
End Function
Public Function retUnPaid(pCode As String, pYear, pCon As ADODB.Connection, Optional ByVal pPaid As Variant = Empty, Optional ByVal pMember As Variant = Empty) As Variant
Dim sDate_Begin As Variant, nYear_Paid As Variant, nYears As Long
If IsEmpty(pPaid) Then
    pPaid = Member_Paid(pCode, , pCon)
End If

'If IsEmpty(pPaid) Then
'    If IsEmpty(pMember) Then
'        sDate_Begin = Member_Load(pCode, "date_Begin") & ""
'    Else
'        sDate_Begin = retFlag(pMember, "date_Begin") & ""
'    End If
'
'    If Not IsDate(sDate_Begin) Then
'        retUnPaid = AddFlag(retUnPaid, "Years", 0)
'        retUnPaid = AddFlag(retUnPaid, "Desca", "·« ÌÊÃœ ”œ«œ ”«»ﬁ - ·« ÌÊÃœ  «—ÌŒ »œ«Ì… ⁄÷ÊÌ…")
'        retUnPaid = AddFlag(retUnPaid, "Error", True)
'    Else
'        nYear_Paid = Ret_Year(sDate_Begin, "Year", pCon, Year(sDate_Begin)) - 1
'        nYears = IIf(pYear - nYear_Paid >= 0, pYear - nYear_Paid, 0)
'        retUnPaid = AddFlag(retUnPaid, "Years", nYears)
'        For i = 1 To nYears
'            If i = 1 Then
'                sDesca = " „‰ " & nYear_Paid + i
'            ElseIf i = nYears Then
'                sDesca = sDesca & " Õ Ï " & nYear_Paid + i
'            End If
'            'sDesca = sDesca & turn(sDesca, ",") & (nYear_Paid + i)
'        Next
'        retUnPaid = AddFlag(retUnPaid, "Desca", sDesca)
'        retUnPaid = AddFlag(retUnPaid, "Error", False)
'    End If
'Else
'    nYear_Paid = retFlag(pPaid, "Year_code") + (retFlag(pPaid, "Years") - 1)
'    nYears = IIf(pYear - nYear_Paid >= 0, pYear - nYear_Paid, 0)
'    retUnPaid = AddFlag(retUnPaid, "Years", nYears)
'    For i = 1 To nYears
''        sDesca = sDesca & turn(sDesca, ",") & (nYear_Paid + i)
'        If i = 1 Then
'            sDesca = " „‰ " & nYear_Paid + i
'        ElseIf i = nYears Then
'            sDesca = sDesca & " Õ Ï " & nYear_Paid + i
'        End If
'    Next
'    retUnPaid = AddFlag(retUnPaid, "desca", sDesca)
'End If
End Function
Function Printed(pMember As String, pCode As String, pYear As String, pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("member", adInteger, adParamInput, , pMember)
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , IIf(ValidNum(pCode), pCode, Null))
cmdTable.Parameters.Append cmdTable.CreateParameter("year", adInteger, adParamInput, , pYear)
cmdTable.CommandText = "PRINTED"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    For I = 0 To obj.Fields.Count - 1
        Printed = AddFlag(Printed, obj.Fields(I).Name, obj.Fields(I).Value)
    Next
End If
Set obj = Nothing
End Function
Function Printed_I(pMember As String, pCode As String, pYear As String, pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("member", adInteger, adParamInput, , pMember)
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , IIf(ValidNum(pCode), pCode, Null))
cmdTable.Parameters.Append cmdTable.CreateParameter("year", adInteger, adParamInput, , pYear)
cmdTable.CommandText = "PRINTED_I"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    For I = 0 To obj.Fields.Count - 1
        Printed_I = AddFlag(Printed_I, obj.Fields(I).Name, obj.Fields(I).Value)
    Next
End If
Set obj = Nothing
End Function
Function Printed_h(pCode As String, pYear As String, pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , IIf(ValidNum(pCode), pCode, Null))
cmdTable.Parameters.Append cmdTable.CreateParameter("year", adInteger, adParamInput, , pYear)
cmdTable.CommandText = "PRINTED_H"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    For I = 0 To obj.Fields.Count - 1
        Printed_h = AddFlag(Printed_h, obj.Fields(I).Name, obj.Fields(I).Value)
    Next
End If
Set obj = Nothing
End Function
Function UnPaidYears(pMember As String, pYear_code As String, pCon As ADODB.Connection) As ADODB.Recordset
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pMember)
cmdTable.Parameters.Append cmdTable.CreateParameter("year_code", adInteger, adParamInput, , pYear_code)
cmdTable.CommandText = "YEARS_UNPAID"
Set obj = cmdTable.Execute
Set UnPaidYears = obj
Set obj = Nothing
End Function
Function Printed_w(pCode As String, pYear As String, pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("year", adInteger, adParamInput, , pYear)
cmdTable.CommandText = "PRINTED_w"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    For I = 0 To obj.Fields.Count - 1
        Printed_w = AddFlag(Printed_w, obj.Fields(I).Name, obj.Fields(I).Value)
    Next
End If
Set obj = Nothing
End Function
Function printed_s(pCode As String, pYear As String, pCon As ADODB.Connection) As Variant
Dim obj As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, , pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("year", adInteger, adParamInput, , pYear)
cmdTable.CommandText = "PRINTED_S"
Set obj = cmdTable.Execute
If Not obj.EOF Then
    For I = 0 To obj.Fields.Count - 1
        printed_s = AddFlag(printed_s, obj.Fields(I).Name, obj.Fields(I).Value)
    Next
End If
Set obj = Nothing
End Function
Function RetValChr(ByVal sValue) As String
sValue = Trim(sValue)
Dim str1 As String, str2 As String
For I = 1 To Len(sValue)
    If IsNumeric(Right(sValue, I)) Then
        str1 = Val(Right(sValue, I))
        Exit For
    End If
Next
If str1 <> "" Then
    RetValChr = Left(sValue, Len(str1)) & RetZero(str1, 10)
End If
End Function
Function GetSysDate(Optional ByVal pDate, Optional pChange As Boolean = False) As String
Dim sDate As String
If IsMissing(pDate) Then
    pDate = RetSetting("date", tempPath & "\password.txt")
    If Not IsDate(pDate) Then
        MsgBox " ·« ÌÊÃœ  «—ÌŒ „”Ã· ”Ì „  ”ÃÌ·  «—ÌŒ «·ÌÊ„"
        sysDate = Format(Date, "YYYY-MM-DD")
        Exit Function
    End If
End If

If IsDate(pDate) Then
    If pChange Then
        sDate = InputBox("«œŒ· «· «—ÌŒ : ", "«œŒ«· «· «—ÌŒ", Format(Date, "YYYY-MM-DD"))
        If Not IsDate(sDate) Then
            MsgBox "·ﬁœ ﬁ„  »«œŒ«·  «—ÌŒ Œ«ÿ∆ ! ”Ì „  ÕÊÌ· «· «—ÌŒ «·Ì  «—ÌŒ «·ÌÊ„"
            sDate = Format(Date, "YYYY-MM-DD")
        End If
    Else
        sDate = pDate
    End If
    
    If DateDiff("D", DateValue(Format(Date, "YYYY-MM-DD")), DateValue(sDate)) > 0 Or DateDiff("D", DateValue(Format(Date, "YYYY-MM-DD")), DateValue(sDate)) <= -2 Then
        If MsgBox("«· «—ÌŒ «ﬂ»— «Ê «ﬁ· „‰  «—ÌŒ «·ÌÊ„ »ÌÊ„Ì‰ !! ”Ì „ «· ÕÊÌ· «·Ì  «—ÌŒ «·ÌÊ„", vbOKCancel + vbDefaultButton2) = vbOK Then
            sDate = Format(Date, "YYYY-MM-DD")
        End If
    ElseIf DateDiff("D", DateValue(Format(Date, "YYYY-MM-DD")), DateValue(sDate)) = -1 Then
        MsgBox "«· «—ÌŒ  ⁄œÌ  «—ÌŒ «·ÌÊ„"
        sDate = Format(sDate, "YYYY-MM-DD")
    End If
Else
    MsgBox "«· «—ÌŒ Œ«ÿÌ !! ”Ì „ «· ÕÊÌ· «·Ì  «—ÌÕ «·ÌÊ„"
    sDate = Format(Date, "YYYY-MM-DD")
End If
addSetting "date", sDate, tempPath & "\password.txt"
sysDate = sDate
End Function
Sub SupLookupAll(oForm As Form, oSearch As Form, Optional pName As String = "", Optional pFilter As String)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select code,Desca From FILE4_10" & turn(pFilter) & pFilter
Generalarray(2) = "Order by desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(@@code@@6 or %%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «·„Ê—œÌ‰"
oSearch.Show 1
End Sub
Sub ChargeLookupAll(oForm As Form, oSearch As Form, Optional pFilter As String)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select code,Desca From FILE8_51" & turn(pFilter) & pFilter
Generalarray(2) = "Order by desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(@@code@@3 or %%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «ﬂÊ«œ «·„’«—Ì›"
oSearch.Show 1
End Sub
Sub NotesLookupAll(oForm As Form, oSearch As Form, Optional pFilter As String)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(8, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select bon_move.code,bon_move.Desca,bon_move.bon_count,bon_move.bon_used,bon_move.bon_rest,bon_move.bon,bon_move.bon_last,bon_move.quant,type_gas_codes.desca From " & _
                  " bon_move inner join type_gas_codes on bon_move.type = type_gas_codes.code"
Generalarray(2) = "Order by bon_move.code"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·œ› —"
listarray(0, 1) = "(%%bon_move.DESCA%%)"

listarray(1, 0) = "—ﬁ„ «·»Ê‰"
listarray(1, 1) = "**BON**<= AND **BON_LAST**>="

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "—ﬁ„ «·œ› —"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "⁄œœ «·»Ê‰« "
GrdArray(2, 1) = 1000

GrdArray(3, 0) = "»Ê‰«  „” Œœ„…"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "»Ê‰«  »«ﬁÌ…"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«Ê· »Ê‰"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "√Œ— »Ê‰"
GrdArray(6, 1) = 1500

GrdArray(7, 0) = "«·ﬂ„Ì…"
GrdArray(7, 1) = 1000

GrdArray(8, 0) = "«·‰Ê⁄"
GrdArray(8, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «·⁄„·«¡"
oSearch.Show 1
End Sub
Sub BoxLookupAll(oForm As Form, oSearch As Form, Optional pFilter As String)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE0_50.CODE ,FILE0_50.DESCA,FILE0_50G.DESCA FROM FILE0_50 LEFT JOIN FILE0_50G ON FILE0_50.[GROUP] = FILE0_50G.CODE"
If Trim(pFilter) <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & pFilter
End If

Generalarray(2) = "ORDER BY FILE0_50.[CODE]"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%FILE0_50.DESCA%%)"

listarray(1, 0) = "«·„Ã„Ê⁄…"
listarray(1, 1) = "(cFilter = FILE0_50.[GROUP])"
listarray(1, 2) = "SELECT CODE,DESCA FROM FILE0_50G"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «·Œ“‰"
oSearch.Show 1
End Sub
Sub ItemsLookupAll(oForm As Form, oSearch As Form, Optional pFilter As String = "FILE6_10.OLD = 0", Optional aFlag As Variant)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select FILE6_10.ITEM,FILE6_10.Desca,RELATION_CODES.DESCA,FILE6_10.AGE1,FILE6_10.AGE2 From FILE6_10 left join RELATION_CODES on FILE6_10.RELATION = RELATION_codes.code" & turn(pFilter) & pFilter
Generalarray(2) = "Order by FILE6_10.ITEM"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(**FILE6_10.ITEM** OR %%FILE6_10.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·ﬁ—«»…"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„‰ ”‰"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "Õ Ï ”‰"
GrdArray(4, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ »‰Êœ «·«‘ —«ﬂ« "

oSearch.Show 1
End Sub
Public Sub InstallLookup(oForm As Form, oSearch As Form, Optional pFilter As String = "", Optional pDoc_No As String = "", Optional aFlag As Variant)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(7, 1)
Set Generalarray(0) = oForm
If Not ValidNum(pDoc_No) Then
    Generalarray(1) = "SELECT dbo.f_serial(FILE6_21.CODE,FILE6_21.DATE_DUE), CONVERT(VARCHAR(10),FILE6_21.DATE_DUE,111),FILE6_21.VALUE,FILE6_21.RATE_TAX,dbo.f_install_paid(FILE6_21.ID),FILE6_21.VALUE - dbo.f_install_paid(FILE6_21.ID)," & _
                      "CONVERT(VARCHAR(10),dbo.f_install_paid_date(FILE6_21.ID),111),FILE6_21.ID " & _
                      "  FROM FILE6_21 WHERE dbo.f_install_paid(FILE6_21.ID) < FILE6_21.VALUE"
Else
    Generalarray(1) = "SELECT dbo.f_serial(FILE6_21.CODE,FILE6_21.DATE_DUE), CONVERT(VARCHAR(10),FILE6_21.DATE_DUE,111),FILE6_21.VALUE,FILE6_21.RATE_TAX,dbo.f_install_paid(FILE6_21.ID),FILE6_21.VALUE - dbo.f_install_paid(FILE6_21.ID)," & _
                      "CONVERT(VARCHAR(10),dbo.f_install_paid_date(FILE6_21.ID),111),FILE6_21.ID " & _
                      "  FROM FILE6_21 WHERE dbo.f_install_paid_doc(FILE6_21.ID," & pDoc_No & ") < FILE6_21.VALUE"
End If
If pFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " AND " & pFilter
End If
Generalarray(2) = "ORDER BY FILE6_21.DATE_DUE"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«· «—ÌŒ"
listarray(0, 1) = "(##FILE6_21.DUE_DATE##)"

GrdArray(0, 0) = "—ﬁ„ «·ﬁ”ÿ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·«” Õﬁ«ﬁ"
GrdArray(1, 1) = 1300

GrdArray(2, 0) = "ﬁÌ„… «·ﬁ”ÿ"
GrdArray(2, 1) = 1200

GrdArray(3, 0) = "‰”»… «·÷—Ì»…"
GrdArray(3, 1) = 1200

GrdArray(4, 0) = "„”œœ"
GrdArray(4, 1) = 1200

GrdArray(5, 0) = "€Ì— „”œœ"
GrdArray(5, 1) = 1200

GrdArray(6, 0) = " «—ÌŒ ”œ«œ"
GrdArray(6, 1) = 1350

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„ «ﬁ”«ÿ"
oSearch.Show 1
End Sub
Sub Items_SportLookupAll(oForm As Form, oSearch As Form, Optional pFilter As String, Optional aFlag As Variant)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select FILE1_30.code,FILE1_30.Desca,FILE1_30.VALUE,FILE1_30.VALUE2 From FILE1_30" & turn(pFilter) & pFilter
Generalarray(2) = "Order by FILE1_30.CODE"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(%%FILE1_30.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·ﬁÌ„… ··„Â‰œ”Ì‰"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·ﬁÌ„… ··«⁄÷«¡"
GrdArray(3, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ »‰Êœ «·‰‘«ÿ «·—Ì«÷Ì"
oSearch.Show 1
End Sub
Sub Items_StudentLookupAll(oForm As Form, oSearch As Form, Optional pFilter As String, Optional aFlag As Variant)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select FILE3_20.code,FILE3_20.Desca,FILE3_20.VALUE From FILE3_20" & turn(pFilter) & pFilter
Generalarray(2) = "Order by FILE3_20.CODE"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(%%FILE3_20.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·ﬁÌ„…"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ »‰Êœ «·ÿ·»…"
oSearch.Show 1
End Sub
Sub Items_ServiceLookupAll(oForm As Form, oSearch As Form, Optional pFilter As String, Optional aFlag As Variant)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select FILE7_20.code,FILE7_20.Desca,FILE7_20.VALUE From FILE7_20" & turn(pFilter) & pFilter
Generalarray(2) = "Order by FILE7_20.CODE"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ -«·≈”„"
listarray(0, 1) = "(%%FILE7_20.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·ﬁÌ„…"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ »‰Êœ „—ﬂ“ «·Œœ„« "
oSearch.Show 1
End Sub
Function HandlePassword(pPassWord As String, pCode As String, Optional pCon As ADODB.Connection, Optional pDefault As String) As Boolean
aSec = Empty
If (pPassWord = pDefault And Trim(pDefault) <> "") Or RetSetting("DEFAULT") = "1" Then
    aSec = AddFlag(Empty, "CODE", -1)
    aSec = AddFlag(aSec, "NAME", "Admin")
    aSec = AddFlag(aSec, "SUPER_MODE", True)
    aSec = AddFlag(aSec, "MANAGER", True)
    aSec = AddFlag(aSec, "DAMAGE", True)
    aSec = AddFlag(aSec, "INFORM", False)
    aSec = AddFlag(aSec, "DOOR", False)
    bSupermode = True
    nUsercode = -1
    cUserName = "SUPER_MODE"
    bopt1 = True
    bopt2 = True
    bopt3 = True
    bOpt4 = True
    bOpt5 = True
    HandlePassword = True
    Exit Function
End If
If Not ValidInt(pCode) Then
    MsgBox "«·ﬂÊœ €Ì— „”Ã·"
    Exit Function
End If
Dim loctable As New ADODB.Recordset, cString As String
cString = "select * from users"
If ValidInt(pCode) Then
    cString = cString & turn(cString) & " code = " & pCode
End If
cString = cString & turn(cString) & "password = " & MyParn(UCase(pPassWord))
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.BOF) Then
    aSec = AddFlag(Empty, "CODE", loctable!CODE & "")
    aSec = AddFlag(aSec, "NAME", loctable!Desca & "")
    aSec = AddFlag(aSec, "MANAGER", loctable!Option1)
    aSec = AddFlag(aSec, "CASH", loctable!Option2)
    aSec = AddFlag(aSec, "DOOR", loctable!Option3)
    aSec = AddFlag(aSec, "INFORM", loctable!Option4)
    bopt1 = loctable!Option1
    nUsercode = loctable!CODE
    cUserName = loctable!Desca & ""
    HandlePassword = True
End If
loctable.Close
Set loctable = Nothing
End Function
Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
Dim nLast
For I = IIf(nBegincol = -1, 0, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, IIf(nEndCol > pGrid.Cols - 1, pGrid.Cols - 1, nEndCol))
    If Trim(pGrid.TextMatrix(Row, I)) = "" And pGrid.ColHidden(I) = False Then
        NextEmpty = I
        Exit Function
    End If
Next
NextEmpty = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
Sub FixAddress(pCon As ADODB.Connection)
Dim loctable As New ADODB.Recordset
loctable.Open "select * From Address", pCon, adOpenStatic, adLockReadOnly
If Not (loctable.EOF And loctable.BOF) Then
    cComp_Name = loctable!Desca & ""
    cComp_address = loctable!Address & ""
    cComp_Phone = loctable!Phone & ""
    cComp_Main = loctable!MAIL & ""
End If
loctable.Close
Set loctable = Nothing
aAddress = GetFields("select * from address", GetCon)
End Sub
Public Function addPaidTypes(pCon As ADODB.Connection) As Boolean
aPaidTypes = GetRows("select * from paid_types", pCon)
addPaidTypes = True
End Function
Public Function AgeFieldRel(pTable As String, pDate1, pDate2) As String
If IsDate(pDate1) And IsDate(pDate2) Then
    AgeFieldRel = "dbo.f_age(" & pTable & ".DATE_BIRTH, CASE WHEN " & pTable & ".RELATION = 2 THEN " & addstring(myFormat(pDate1)) & " ELSE " & addstring(myFormat(pDate2)) & " END )"
Else
    AgeFieldRel = "NULL"
End If
End Function
Public Function fixClaim(pDoc_No As String) As String
fixClaim = "UPDATE FILE6_20H SET FILE6_20H.TOTAL_YEAR = dbo.f_inv_total_year(FILE6_20H.DOC_NO)," & _
           " FILE6_20H.TOTAL_YEAR_OTHER = dbo.f_inv_total_year_other(FILE6_20H.DOC_NO)," & _
           " FILE6_20H.TOTAL_LATE = dbo.f_inv_total_late(FILE6_20H.DOC_NO)," & _
           " FILE6_20H.TOTAL_TAX = dbo.f_inv_total_tax(FILE6_20H.DOC_NO)" & _
           "   FROM FILE6_20H WHERE DOC_NO = " & addvalue(pDoc_No)
End Function
Public Function fixYears(pDoc_No As String) As String
fixYears = "UPDATE FILE6_20H SET " & _
          "FILE6_20H.[YEARS] = dbo.fn_get_years_count(" & pDoc_No & ")," & _
          "FILE6_20H.[YEARS_DESCA] = dbo.f_get_years(" & pDoc_No & ") " & _
          "FROM FILE6_20H WHERE DOC_NO = " & pDoc_No
End Function
Public Function fixYears2(pDoc_No As String) As String
fixYears2 = "UPDATE FILE6_60H SET " & _
            "FILE6_60H.[YEARS] = dbo.fn_get_years_count2(" & pDoc_No & ")," & _
            "FILE6_60H.[YEARS_DESCA] = dbo.f_get_years2(" & pDoc_No & ") " & _
            "FROM FILE6_60H WHERE DOC_NO = " & pDoc_No
End Function
Public Function fixMemberPaid(pMember As String) As String
fixMemberPaid = "UPDATE FILE1_10 SET FILE1_10.DOC_NO = [dbo].[f_last_year_doc](FILE1_10.CODE) WHERE FILE1_10.CODE = " & addvalue(pMember)
End Function
Public Function fixClaimOther(pDoc_No As String, pFileHeader As String) As String
fixClaimOther = "UPDATE " & pFileHeader & " SET TOTAL_YEAR = dbo.f_inv_total_year2(DOC_NO)," & _
           " TOTAL_YEAR_OTHER = dbo.f_inv_total_year_other2(DOC_NO)," & _
           " TOTAL_LATE = dbo.f_inv_total_late2(DOC_NO)," & _
           " TOTAL_TAX = dbo.f_inv_total_tax2(DOC_NO)" & _
           " FROM " & pFileHeader & " WHERE DOC_NO = " & addvalue(pDoc_No)
End Function
Sub MemberH_LookupAll(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT FILE3_10.CODE,FILE3_10.NO,FILE3_10.DESCA" & _
                  " FROM FILE3_10"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE3_10.DESCA"
Generalarray(3) = 7000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-—ﬁ„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE3_10.DESCA%%  OR **FILE3_10.NO**)"


GrdArray(0, 0) = "ﬂÊœ «·⁄÷Ê"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "—ﬁ„ «·ﬂ«—‰ÌÂ"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«”„ «·⁄÷Ê"
GrdArray(2, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE2_10.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.sCaption = "≈” ⁄·«„ «·«⁄÷«¡ «·‘—›ÌÌ‰"
oSearch.Show 1
End Sub
Public Function SendCard(Optional pCode As String = "", Optional pCard As String = "", Optional pCon As ADODB.Connection, Optional con2 As ADODB.Connection, Optional pWhere As String = "") As String
Dim loctable As ADODB.Recordset
Dim cInsert As String
cString = "select file1_10.*,YEARS_CODES.DATE2  AS DATE_LAST from file1_10 INNER JOIN YEARS_CODES ON YEARS_CODES.CODE = dbo.f_last_year_CODE(file1_10.code) "
cWhere = "(Not file1_10.card Is Null)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    aInsert = AddFlag(Empty, "[no]", loctable!CODE)
    aInsert = AddFlag(aInsert, "[Name]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[CAT]", "0")
    aInsert = AddFlag(aInsert, "[RELORDER]", "0")
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    aInsert = AddFlag(aInsert, "[END_DATE]", addDate(loctable!DATE_LAST))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(loctable!DATE_LAST))
    
    If Not IsEmpty(GetField("SELECT [ID] FROM [ALL] WHERE [ID] = " & MyParn(loctable!CODE_CARD), con2)) Then
        cInsert = cInsert & addUpdate(aInsert, "[ALL]", "[ID] = " & MyParn(loctable!CODE_CARD)) & ";"
    Else
        aInsert = AddFlag(aInsert, "[ID]", addstring(loctable!CODE_CARD))
        cInsert = cInsert & addInsert(aInsert, "[ALL]") & ";"
    End If
    loctable.MoveNext
Loop

Set loctable = New ADODB.Recordset
cString = "select file1_11.MEMBER,FILE1_11.DESCA,FILE1_11.CODE,FILE1_11.CODE_CARD,FILE1_11.CARD,YEARS_CODES.DATE2  AS DATE_LAST from file1_11 INNER JOIN YEARS_CODES ON YEARS_CODES.CODE = dbo.f_last_year_CODE(file1_11.MEMBER) "
cWhere = "(Not FILE1_11.card Is Null)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.MEMBER = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    aInsert = AddFlag(Empty, "[no]", loctable!CODE)
    aInsert = AddFlag(aInsert, "[Name]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[CAT]", "0")
    aInsert = AddFlag(aInsert, "[RELORDER]", "0")
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    aInsert = AddFlag(aInsert, "[END_DATE]", addDate(loctable!DATE_LAST))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(loctable!DATE_LAST))
    If Not IsEmpty(GetField("SELECT [ID] FROM [ALL] WHERE [ID] = " & MyParn(loctable!CODE_CARD), con2)) Then
        cInsert = cInsert & addUpdate(aInsert, "[ALL]", "[ID] = " & MyParn(loctable!CODE_CARD)) & ";"
    Else
        aInsert = AddFlag(aInsert, "[ID]", addstring(loctable!CODE_CARD))
        cInsert = cInsert & addInsert(aInsert, "[ALL]") & ";"
    End If
    loctable.MoveNext
Loop
SendCard = cInsert
End Function
Public Function SendCardInstall(Optional pCode As String = "", Optional pCard As String = "", Optional pCon As ADODB.Connection, Optional con2 As ADODB.Connection, Optional pWhere As String = "") As String
Dim loctable As ADODB.Recordset
Dim sDateBegin As String, sDateEnd As String
Dim cInsert As String
cString = "select file2_10.* from file2_10"
cWhere = "(Not file2_10.card Is Null) and (NOT DATE_END IS NULL)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.CODE = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    aInsert = AddFlag(Empty, "[no]", loctable!CODE)
    aInsert = AddFlag(aInsert, "[Name]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[CAT]", "1")
    aInsert = AddFlag(aInsert, "[RELORDER]", "0")
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    
    sDateEnd = myFormat(IIf(IsNull(loctable!DATE_PRINT), Date, loctable!DATE_PRINT))
    sDateEnd = myFormat(DateAdd("YYYY", 1, sDateEnd))
    If myFormat(loctable!DATE_END) < sDateEnd Then
        sDateEnd = myFormat(loctable!DATE_END)
    End If
    
    aInsert = AddFlag(aInsert, "[END_DATE]", addDate(sDateEnd))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(sDateEnd))
    
    If Not IsEmpty(GetField("SELECT [ID] FROM [ALL] WHERE [ID] = " & MyParn(loctable!CODE_CARD), con2)) Then
        cInsert = cInsert & addUpdate(aInsert, "[ALL]", "[ID] = " & MyParn(loctable!CODE_CARD)) & ";"
    Else
        aInsert = AddFlag(aInsert, "[ID]", addstring(loctable!CODE_CARD))
        cInsert = cInsert & addInsert(aInsert, "[ALL]") & ";"
    End If
    loctable.MoveNext
Loop

Set loctable = New ADODB.Recordset
cString = "select FILE2_11.MEMBER,FILE2_11.DESCA,FILE2_11.CODE,FILE2_11.CODE_CARD,FILE2_11.CARD,FILE2_10.DATE_PRINT,FILE2_10.DATE_END from FILE2_11 INNER JOIN FILE2_10 ON FILE2_11.MEMBER  = FILE2_10.CODE  "
cWhere = "(Not FILE2_11.card Is Null)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE2_11.MEMBER = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE2_11.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    aInsert = AddFlag(Empty, "[no]", loctable!CODE)
    aInsert = AddFlag(aInsert, "[Name]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[CAT]", "1")
    aInsert = AddFlag(aInsert, "[RELORDER]", "0")
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    
    sDateEnd = myFormat(IIf(IsNull(loctable!DATE_PRINT), Date, loctable!DATE_PRINT))
    sDateEnd = myFormat(DateAdd("YYYY", 1, sDateEnd))
    If myFormat(loctable!DATE_END) < sDateEnd Then
        sDateEnd = myFormat(loctable!DATE_END)
    End If
    
    aInsert = AddFlag(aInsert, "[END_DATE]", addDate(sDateEnd))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(sDateEnd))
    If Not IsEmpty(GetField("SELECT [ID] FROM [ALL] WHERE [ID] = " & MyParn(loctable!CODE_CARD), con2)) Then
        cInsert = cInsert & addUpdate(aInsert, "[ALL]", "[ID] = " & MyParn(loctable!CODE_CARD)) & ";"
    Else
        aInsert = AddFlag(aInsert, "[ID]", addstring(loctable!CODE_CARD))
        cInsert = cInsert & addInsert(aInsert, "[ALL]") & ";"
    End If
    loctable.MoveNext
Loop
SendCardInstall = cInsert
End Function
Public Function SendCardMdb(Optional pCode As String = "", Optional pCard As String = "", Optional pCon As ADODB.Connection, Optional pWhere As String = "", Optional myForm As Form) As Variant
Dim loctable As ADODB.Recordset, aSend()
Dim cInsert As String
cString = "select file1_10.*,YEARS_CODES.DATE2  AS DATE_LAST from file1_10 INNER JOIN YEARS_CODES ON YEARS_CODES.CODE = dbo.f_last_year_CODE(file1_10.code) "
cWhere = "(Not file1_10.card Is Null)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText

nRecordcount = loctable.RecordCount
myForm.prog1.Visible = True
Do Until loctable.EOF
    I = I + 1
    myForm.prog1.Value = IIf(Round(I / (nRecordcount), 2) > 1, 1, Round(I / (nRecordcount), 2)) * 100
    aInsert = AddFlag(Empty, "[no]", loctable!CODE)
    aInsert = AddFlag(aInsert, "[Name]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[CAT]", "0")
    aInsert = AddFlag(aInsert, "[RELORDER]", "0")
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    aInsert = AddFlag(aInsert, "[END_DATE]", addDate(loctable!DATE_LAST))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(loctable!DATE_LAST))
    aInsert = AddFlag(aInsert, "[ID]", addstring(loctable!CODE_CARD))
    SendCardMdb = AddFlag(SendCardMdb, addInsert(aInsert, "[ALL]"))
    loctable.MoveNext
Loop

Set loctable = New ADODB.Recordset
cString = "select file1_11.MEMBER,FILE1_11.DESCA,FILE1_11.CODE,FILE1_11.CODE_CARD,FILE1_11.CARD,YEARS_CODES.DATE2  AS DATE_LAST from file1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE INNER JOIN YEARS_CODES ON YEARS_CODES.CODE = dbo.f_last_year_CODE(file1_11.MEMBER) "
cWhere = "(Not FILE1_11.card Is Null)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.MEMBER = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText

nRecordcount = loctable.RecordCount
I = 0
Do Until loctable.EOF
    I = I + 1
    myForm.prog1.Value = IIf(Round(I / (nRecordcount), 2) > 1, 1, Round(I / (nRecordcount), 2)) * 100
    
    aInsert = AddFlag(Empty, "[no]", loctable!CODE)
    aInsert = AddFlag(aInsert, "[Name]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[CAT]", "0")
    aInsert = AddFlag(aInsert, "[RELORDER]", "0")
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    aInsert = AddFlag(aInsert, "[END_DATE]", addDate(loctable!DATE_LAST))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(loctable!DATE_LAST))
    aInsert = AddFlag(aInsert, "[ID]", addstring(loctable!CODE_CARD))
    SendCardMdb = AddFlag(SendCardMdb, addInsert(aInsert, "[ALL]"))
    loctable.MoveNext
Loop
myForm.prog1.Visible = False
End Function
Public Function getCardMdb(pCon As ADODB.Connection, pCon2 As ADODB.Connection, Optional myForm As Form) As Variant
Dim loctable As ADODB.Recordset, aSend()
Dim cInsert As String
cString = "select * FROM [ALL] "
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon2, adOpenStatic, adLockReadOnly, adCmdText

If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordcount = loctable.RecordCount
    loctable.MoveFirst
End If

myForm.prog1.Visible = True
Do Until loctable.EOF
    I = I + 1
    nRecord = nRecord + 1
    myForm.prog1.Value = IIf(Round(I / (nRecordcount), 2) > 1, 1, Round(I / (nRecordcount), 2)) * 100
    aInsert = AddFlag(Empty, "[no]", loctable!NO)
    aInsert = AddFlag(aInsert, "[Name]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[CAT]", mRound(loctable!cat))
    aInsert = AddFlag(aInsert, "[RELORDER]", mRound(loctable!relorder))
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    aInsert = AddFlag(aInsert, "[END_DATE]", addDate(loctable!END_DATE))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(loctable!MDATE))
    If Not IsEmpty(GetField("SELECT [ID] FROM [ALL] WHERE [ID] = " & MyParn(loctable!ID), pCon)) Then
        cInsert = cInsert & addUpdate(aInsert, "[ALL]", "[ID] = " & MyParn(loctable!ID)) & ";"
    Else
        aInsert = AddFlag(aInsert, "[ID]", addstring(loctable!ID))
        cInsert = cInsert & addInsert(aInsert, "[ALL]") & ";"
    End If
    If nRecord = 10 Then
        getCardMdb = AddFlag(getCardMdb, cInsert)
        nRecord = 0
        cInsert = ""
    End If
    loctable.MoveNext
Loop
If Trim(cInsert) <> "" Then getCardMdb = AddFlag(getCardMdb, cInsert)
myForm.prog1.Visible = False
End Function
Public Function sendCardPhoto(Optional pCode As String = "", Optional pCard As String = "", Optional pCon As ADODB.Connection, Optional pWhere As String = "", Optional myForm As Form) As Variant
Dim loctable As ADODB.Recordset
Dim cInsert As String
cString = "select file1_10.CODE from file1_10 INNER JOIN YEARS_CODES ON YEARS_CODES.CODE = dbo.f_last_year_CODE(file1_10.code) "
cWhere = "(Not file1_10.card Is Null)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
nRecordcount = loctable.RecordCount
myForm.prog1.Visible = True
Do Until loctable.EOF
    I = I + 1
    myForm.prog1.Value = IIf(Round(I / (nRecordcount), 2) > 1, 1, Round(I / (nRecordcount), 2)) * 100
    sendCardPhoto = AddFlag(sendCardPhoto, loctable!CODE)
    loctable.MoveNext
Loop

Set loctable = New ADODB.Recordset
cString = "select file1_11.MEMBER,FILE1_11.CODE from file1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE INNER JOIN YEARS_CODES ON YEARS_CODES.CODE = dbo.f_last_year_CODE(file1_11.MEMBER) "
cWhere = "(Not FILE1_11.card Is Null)"
If pCode <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.MEMBER = " & addvalue(pCode)
If pCard <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.CARD = " & MyParn(pCard)
If pWhere <> "" Then cWhere = cWhere & turn(cWhere, " and ") & pWhere
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set loctable = New ADODB.Recordset
loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
nRecordcount = loctable.RecordCount
myForm.prog1.Visible = True
I = 0
Do Until loctable.EOF
    I = I + 1
    myForm.prog1.Value = IIf(Round(I / (nRecordcount), 2) > 1, 1, Round(I / (nRecordcount), 2)) * 100
    sendCardPhoto = AddFlag(sendCardPhoto, loctable!MEMBER & "-" & loctable!CODE)
    loctable.MoveNext
Loop
myForm.prog1.Visible = False
End Function
Public Function SplitSql(pString As String, pLen As Integer)
Dim aString As Variant, aSplit As Variant, I As Long, i2 As Long
aString = Split(pString, ";")
For I = 0 To UBound(aString)
    If Trim(aString(I)) <> "" Then
        cString = cString & aString(I) & ";"
        i2 = i2 + 1
        If i2 = pLen Then
            SplitSql = AddFlag(SplitSql, cString)
            i2 = 0
            cString = ""
        End If
    End If
Next
End Function
Public Function ValidDateTax(pDate As String) As Boolean
If Not IsDate(pDate) Then Exit Function
If myFormat(pDate) < DATE_TAX1 Then Exit Function
If myFormat(pDate) > DATE_TAX2 Then Exit Function
ValidDateTax = True
End Function
Public Sub ServiceLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional bAddRow As Boolean, Optional sDesca As String)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT vw_last_paid_service.doc_no,convert(varchar(10),[date],111),years_codes.desca" & _
                  " FROM vw_last_paid_service inner join years_codes on vw_last_paid_service.year_code = years_codes.code"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " Where " & cFilter
End If

Generalarray(2) = "Order by vw_last_paid_service.year_code"
Generalarray(3) = 5000
Generalarray(5) = True

listarray(0, 0) = "«· «—ÌŒ"
listarray(0, 1) = "(##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«·„Ê”„"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "Doc_NO")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "ﬂ· «·„ÿ«·»« ")
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

oSearch.sCaption = "≈” ⁄·«„ " & sDesca
oSearch.Show 1
End Sub
Function Document_Files(sCode As String, sId As String) As String
Document_Files = sPath_App & "\Documents\" & sCode & "\" & sId & ".jpg"
End Function
Function Checks_Dir(sCode As String) As String
Checks_Dir = sPath_App & "\Documents\" & sCode
End Function
Public Function paid_once(pType As String, nYear_code As Variant, pCode As String, pCon As ADODB.Connection) As Variant
Dim cmd As New Command
aPrm = AddFlag(aPrm, "TYPE", pType)
aPrm = AddFlag(aPrm, "YEAR_CODE", nYear_code)
aPrm = AddFlag(aPrm, "CODE", pCode)

Set cmd = myCmdEx("[dbo].[sp_paid_once]", pCon, aPrm, 1000)
paid_once = cmd.Parameters("@DOC_NO")
End Function



