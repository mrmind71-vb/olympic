Attribute VB_Name = "toolkit2"
Sub makeMyLoad(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
Print #1, "Private Sub MyLoad()"
With pform.VSFlexGrid1
    .Rows = 0
    For i = 0 To pform.Count - 1
        If (TypeOf pform(i) Is TextBox) Or (TypeOf pform(i) Is DataCombo) Or (TypeOf pform(i) Is CheckBox) Then
            If TypeOf pform(i) Is TextBox Then
                If LCase(pform(i).Tag) = "date" Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
                    .TextMatrix(.Rows - 1, 1) = pform(i).Name & cIndex & ".text = " & "Format(CardTable" & "!" & cCardIndex & Mid(pform(i).Name, 2) & "," & myparn2("dd-mm-yyyy") & ")"
                Else
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
                    .TextMatrix(.Rows - 1, 1) = pform(i).Name & cIndex & ".text = " & " CardTable" & "!" & cCardIndex & Mid(pform(i).Name, 2) & " & " & retPar
                End If
            ElseIf TypeOf pform(i) Is CheckBox Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
                .TextMatrix(.Rows - 1, 1) = pform(i).Name & ".value = " & " iif(CardTable" & "!" & Mid(pform(i).Name, 2) & ",1,0)"
            Else
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
                .TextMatrix(.Rows - 1, 1) = pform(i).Name & ".boundtext = " & " CardTable" & "!" & Mid(pform(i).Name, 2) & " & " & retPar
            End If
        End If
    Next
    .Select 1, 0
    .Sort = flexSortNumericAscending
    For i2 = 0 To .Rows - 1
        Print #1, .TextMatrix(i2, 1)
    Next
End With
Print #1, "End sub"
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyDefine(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
With pform.VSFlexGrid1
    .Rows = 0
    For i = 0 To pform.Count - 1
        If TypeOf pform(i) Is TextBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = pform(i).Name & cIndex & ".text = " & retPar
        ElseIf TypeOf pform(i) Is CheckBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = pform(i).Name & ".value = 0 "
        ElseIf TypeOf pform(i) Is DataCombo Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = pform(i).Name & ".boundtext = " & retPar
        End If
    Next
    .Select 1, 0
    .Sort = flexSortNumericAscending
    For i2 = 0 To .Rows - 1
        Print #1, .TextMatrix(i2, 1)
    Next
End With
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyReplace(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
'Print #1, "Private Sub MyReplace()"

For i = 0 To pform.Count - 1
    If ((TypeOf pform(i) Is TextBox) Or (TypeOf pform(i) Is DataCombo) Or (TypeOf pform(i) Is CheckBox)) Then
        If LCase(pform(i).Tag) <> "ig" Then nCount = nCount + 1
    End If
Next

'Print #1, "Dim aInsert as Variant"
With pform.VSFlexGrid1
.Rows = 0
nOrder = 0
For i = 0 To pform.Count - 1
    If (TypeOf pform(i) Is TextBox) Or (TypeOf pform(i) Is DataCombo) Or (TypeOf pform(i) Is CheckBox) Then
        If TypeOf pform(i) Is TextBox And LCase(pform(i).Tag) <> "ig" Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(i).Name, 2))
            If LCase(pform(i).Tag) = "date" Then
                .TextMatrix(.Rows - 1, 2) = "addDate(" & pform(i).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(i).Tag) = "N" Then
                .TextMatrix(.Rows - 1, 2) = "addValue(" & pform(i).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(i).Tag) = "" Then
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(i).Name & cIndex & ".Text" & ")"
            Else
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(i).Name & cIndex & ".Text" & ")"
            End If
        ElseIf TypeOf pform(i) Is CheckBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(i).Name, 2))
            .TextMatrix(.Rows - 1, 2) = " iif( " & pform(i).Name & ".value = 0 , " & myparn2("FALSE") & "," & myparn2("True") & ")"
        ElseIf TypeOf pform(i) Is DataCombo Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(i).Name, 2))
            .TextMatrix(.Rows - 1, 2) = "addvalue(" & pform(i).Name & ".BoundText" & ")"
        End If
    End If
    nOrder = nOrder + 1
Next

.Select 1, 0
.Sort = flexSortNumericAscending
For i2 = 0 To .Rows - 1
    Print #1, "aInsert = addFlag(aInsert," & .TextMatrix(i2, 1) & "," & .TextMatrix(i2, 2) & ")"
Next
End With
Print #1, "End sub"
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyReplace2(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
Print #1, "Private Sub MyReplace()"

For i = 0 To pform.Count - 1
    If ((TypeOf pform(i) Is TextBox) Or (TypeOf pform(i) Is DataCombo) Or (TypeOf pform(i) Is CheckBox)) Then
        If LCase(pform(i).Tag) <> "ig" Then nCount = nCount + 1
    End If
Next

Print #1, "Dim aInsert(" & nCount - 1 & ",1)"
With pform.VSFlexGrid1
.Rows = 0
nOrder = 0
For i = 0 To pform.Count - 1
    If (TypeOf pform(i) Is TextBox) Or (TypeOf pform(i) Is DataCombo) Or (TypeOf pform(i) Is CheckBox) Then
        If TypeOf pform(i) Is TextBox And LCase(pform(i).Tag) <> "ig" Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(i).Name, 2))
            If LCase(pform(i).Tag) = "date" Then
                .TextMatrix(.Rows - 1, 2) = "addDate(" & pform(i).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(i).Tag) = "N" Then
                .TextMatrix(.Rows - 1, 2) = "addValue(" & pform(i).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(i).Tag) = "" Then
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(i).Name & cIndex & ".Text" & ")"
            Else
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(i).Name & cIndex & ".Text" & ")"
            End If
        ElseIf TypeOf pform(i) Is CheckBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(i).Name, 2))
            .TextMatrix(.Rows - 1, 2) = " iif( " & pform(i).Name & ".value = 0 , " & myparn2("FALSE") & "," & myparn2("True") & ")"
        ElseIf TypeOf pform(i) Is DataCombo Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(i).Name, 2))
            .TextMatrix(.Rows - 1, 2) = "addvalue(" & pform(i).Name & ".BoundText" & ")"
        End If
    End If
    nOrder = nOrder + 1
Next

.Select 1, 0
.Sort = flexSortNumericAscending
For i2 = 0 To .Rows - 1
    Print #1, "aInsert(" & i2 & ",0) = " & .TextMatrix(i2, 1)
    Print #1, "aInsert(" & i2 & ",1) = " & .TextMatrix(i2, 2)
Next
End With
Print #1, "End sub"
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub MFocus(pform As Form)
With pform.VSFlexGrid1
.Rows = 0
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For i = 0 To pform.Count - 1
    If (TypeOf pform(i) Is TextBox Or TypeOf pform(i) Is DataCombo) Then
        Print #1, "Private Sub " & pform(i).Name & "_GotFocus()"
        Print #1, "mygotFocus " & pform(i).Name
        Print #1, "End Sub"
        
        Print #1, "Private Sub " & pform(i).Name & "_LostFocus()"
        Print #1, "myLostFocus " & pform(i).Name
        If UCase(pform(i).Tag) = "D" Then
            Print #1, "myvalidDate " & pform(i).Name
        End If
        If TypeOf pform(i) Is DataCombo Then
            Print #1, "IF NOT " & pform(i).Name & ".matchWithList THEN " & pform(i).Name & pform(i).Name & ".BoundText = " & myparn2("")
        End If
        Print #1, "End Sub"
    End If
Next
Close #1
End With
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub LostFocus(pform As Form)
With pform.VSFlexGrid1
.Rows = 0
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For i = 0 To pform.Count - 1
    If (TypeOf pform(i) Is TextBox Or TypeOf pform(i) Is DataCombo) Then
        If .FindRow(pform(i).Name, , 0) = -1 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).Name
            Print #1, "Private Sub " & pform(i).Name & "_LostFocus()"
            Print #1, "myLostFocus " & pform(i).Name
            Print #1, "End Sub"
        End If
    End If
Next
Close #1
End With
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub MakeKeyDown(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For i = 0 To pform.Count - 1
    If (TypeOf pform(i) Is DataCombo) Then
        Print #1, "Private Sub " & pform(i).Name & "_KeyDown(KeyCode As Integer, Shift As Integer)"
        Print #1, "if KeyCode = 40 Then "
        Print #1, "    KeyCode = 0"
        Print #1, "    SendKeys " & myparn2("{TAB}")
        Print #1, "elseif KeyCode = 38 Then"
        Print #1, "    KeyCode = 0"
        Print #1, "    SendKeys " & myparn2("+{TAB}")
        Print #1, "End If"
        Print #1, "End Sub"
    ElseIf (TypeOf pform(i) Is TextBox) Then
        Print #1, "Private Sub " & pform(i).Name & "_KeyDown(" & IIf(pform(i).Index <> "", "Index As Integer", "") & "KeyCode As Integer, Shift As Integer)"
        Print #1, "if KeyCode = 40 Then  SendKeys " & myparn2("{TAB}")
        Print #1, "if KeyCode = 38 Then  SendKeys " & myparn2("+{TAB}")
        Print #1, "End Sub"
    End If
Next
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyMenu(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For i = 0 To pform.Count - 1
    Print #1, pform(i).Name
Next
Close #1
End Sub
Function retPar()
retPar = String(2, Chr(34))
End Function
Function GetText(Optional pFileName) As String
Dim TextLine
On Error GoTo myerror
Open pFileName For Input As #1    ' Open file.
Do While Not EOF(1)   ' Loop until end of file.
   Line Input #1, TextLine   ' Read line into variable.
   GetText = GetText & IIf(GetText = "", "", vbCrLf) & TextLine
Loop
Close #1   ' Close file.
Exit Function
myerror:
'MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
GetText = ""
End Function
Function RetAdd(sTable As String, bcon As ADODB.Connection)
cFile = App.Path & "\temp.txt"
Dim loctable As New ADODB.Recordset
Dim nLast As Integer
nLast = 20
Open cFile For Output As #1   ' Open file for output.
Print #1, "' --- ADD SUB --------------------------"

loctable.Open sTable, bcon, adOpenStatic, adLockReadOnly, adCmdTable
cString1 = "insert into " & sTable & "( "
For i = 0 To loctable.Fields.Count - 1
    cString1 = cString1 & turnFound2(cString1, loctable.Fields(0).Name, ",", "") & loctable.Fields(i).Name
Next
cString1 = Chr(34) & cString1 & ")" & Chr(34) & " & _"
cString1 = "cString  = " & cString1

Print #1, cString1

Print #1, myparn2("Values (") & " & _"
For i = 0 To loctable.Fields.Count - 1
    If loctable.Fields(i).Type = 3 Then
       cField = "addValue(x" & loctable.Fields(i).Name & ".text)"
    ElseIf loctable.Fields(i).Type = 11 Then
       cField = "iif(x" & loctable.Fields(i).Name & ".value = 0," & Chr(34) & "FALSE" & Chr(34) & "," & Chr(34) & "TRUE" & Chr(34) & " )"
    ElseIf loctable.Fields(i).Type = 7 Then
        cField = "adddate(x" & loctable.Fields(i).Name & ".text)"
    Else
       cField = "addstring(x" & loctable.Fields(i).Name & ".text)"
    End If
    
    If i = nLast Then
        Print #1, ""
        cField = "cString2 = " & cField
    End If
    
    If i = nLast - 1 And i <> loctable.Fields.Count - 1 Then
        cField = cField & " & " & myparn2(",")
    Else
        cField = cField & IIf(i = loctable.Fields.Count - 1, "& _", " & " & myparn2(",") & " & _")
    End If
    Print #1, cField
Next
Print #1, myparn2(")")
If i > nLast Then Print #1, "cString = cString & cString2"


Print #1, "' --- EDIT SUB  --------------------------"
Print #1, "cString  = " & myparn2("update " & sTable & " Set") & "& _"

For i = 0 To loctable.Fields.Count - 1
    If loctable.Fields(i).Type = 3 Then
       cField = myparn2(loctable.Fields(i).Name & " = ") & " &  addValue(x" & loctable.Fields(i).Name & ".Text )"
    ElseIf loctable.Fields(i).Type = 11 Then
       cField = myparn2(loctable.Fields(i).Name & " = ") & " & iif(x" & loctable.Fields(i).Name & ".value = 0," & myparn2("FALSE") & "," & myparn2("TRUE") & " )"
    ElseIf loctable.Fields(i).Type = 7 Then
       cField = myparn2(loctable.Fields(i).Name & " = ") & " &  addDate(x" & loctable.Fields(i).Name & ".Text )"
    Else
       cField = myparn2(loctable.Fields(i).Name & " = ") & " & addstring(x" & loctable.Fields(i).Name & ".text)"
    End If

    If i = nLast Then
        Print #1, ""
        cField = "cString2 = " & cField
    End If
    
    If i = nLast - 1 And i <> loctable.Fields.Count - 1 Then
        cField = cField & " & " & myparn2(",")
    Else
        cField = cField & IIf(i = loctable.Fields.Count, "", " & " & myparn2(",") & " & _")
    End If
    Print #1, cField
Next
Print #1, myparn2(")")
If i > nLast Then Print #1, "cString = cString & cString2"

Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Function
Private Function myparn2(cString)
myparn2 = Chr(34) & cString & Chr(34)
End Function
Sub makeMyValidate(pform As Form, Optional withfocus As Boolean = False, Optional withValidate As Boolean = True, Optional withLost As Boolean = False)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
Print #1, String(50, "-")
If withValidate Then
    For i = 0 To pform.Count - 1
        If (TypeOf pform(i) Is DataCombo) Then
            Print #1, "Private Sub " & pform(i).Name & "_Validate(Cancel As Boolean)"
            Print #1, "if Not " & pform(i).Name & ".MatchedWithList Then " & pform(i).Name & ".BoundText = " & myparn2("")
            Print #1, "End Sub"
        ElseIf (TypeOf pform(i) Is TextBox) Then
            If LCase(pform(i).Tag) = "date" Then
                Print #1, "Private Sub " & pform(i).Name & "_Validate(Cancel As Boolean)"
                Print #1, "With " & pform(i).Name
                Print #1, "If (Not IsDate(.Text)) And Trim(.Text) <> " & myparn2("") & " Then "; ".text = " & myparn2("")
                'Print #1, vbTab & "Cancel = True"
               ' Print #1,
               ' Print #1, "End If"
                Print #1, ".Text = Format(.Text," & myparn2("dd-mm-yyyy") & ")"
                Print #1, "End With"
                Print #1, "End Sub"
            End If
        End If
    Next
End If

If withfocus Then
    Print #1, String(50, "-")
    For i = 0 To pform.Count - 1
        'If (TypeOf pform(I) Is DataCombo) Or TypeOf pform(I) Is TextBox Then
         If TypeOf pform(i) Is TextBox Then
            If pform(i).TabIndex > 2 Then
                Print #1, "Private Sub " & pform(i).Name & "_GotFocus()"
                'Print #1, "StatusBar1.Panels(2).Text = RetCap(Me.Name, ActiveControl.Name)"
                If TypeOf pform(i) Is TextBox Then
                    Print #1, pform(i).Name & ".SelStart = 0"
                    Print #1, pform(i).Name & ".SelLength = Len(" & pform(i).Name & ".text)"
                End If
                Print #1, "End Sub"
            End If
        End If
    Next
End If
Close #1

cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub

Private Sub wNo_GotFocus()
wNo.SelStart = 0
wNo.SelLength = Len(wNo.Text)
End Sub
Private Sub xdate_GotFocus()
xdate.SelStart = 0
xdate.SelLength = Len(xdate.Text)
End Sub
Sub mRetNumber(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
With pform.VSFlexGrid1
.Rows = 0
For i = 0 To pform.Count - 1
    If .FindRow(pform(i).Name, , 0) = -1 Then
        If LCase(Left(pform(i).Name, 2)) = "m0" Or LCase(Left(pform(i).Name, 2)) = "m1" Or LCase(Left(pform(i).Name, 2)) = "m2" Then
            If pform(i).Index <> "" Then
                cIndex = "(" & pform(i).Index & ")"
            End If
        End If
        If (TypeOf pform(i) Is TextBox And pform(i).Tag <> "date") Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(i).Name
            If cIndex <> "" Then
                Print #1, "Private Sub " & pform(i).Name & "_KeyPress(Index As Integer,KeyAscii As Integer)"
                Print #1, "if " & pform(i).Name & "(index).tag <> " & myparn2("dec") & " and lcase(" & pform(i).Name & "(index).tag) <>  " & myparn2("date") & " then  "
                Print #1, "if " & pform(i).Name & "(index).tag  = " & myparn2("dec"); " then " & "KeyAscii = RetNumber(KeyAscii,TRUE) else KeyAscii = RetNumber(KeyAscii)"
                Print #1, "end if"
            Else
                Print #1, "Private Sub " & pform(i).Name & "_KeyPress(KeyAscii As Integer)"
                Print #1, "KeyAscii = RetNumber(KeyAscii)"
            End If
            Print #1, "End Sub"
        End If
    End If
Next
Close #1
End With
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
