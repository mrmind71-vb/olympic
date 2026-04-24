Attribute VB_Name = "DataProc"
Public Enum cmType
adStoredProc = adCmdStoredProc
adText = adCmdText
adTable = adCmdTable
End Enum
Public Function myCmd(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 200) As ADODB.Recordset
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.Command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set myCmd = cmd.Execute
End Function
Public Function myCommand(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100, Optional ByRef nRecords As Long) As ADODB.Command
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.Command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString

If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
cmd.Execute nRecords
Set myCommand = cmd
End Function
Public Function myCmdEx(pString As String, Optional con As ADODB.Connection, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100) As ADODB.Command
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.Command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = pString

Dim i As Long
For i = 0 To UBound(aParam) Step 2
    cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
Next
cmd.Execute
Set myCmdEx = cmd
End Function
Public Function myCmdProc(pString As String, Optional con As ADODB.Connection, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim i As Long
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = pString
Dim loctable As New ADODB.Recordset
If Not IsEmpty(aParam) Then
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set myCmdProc = cmd.Execute
End Function
Public Function myCmdText(pString As String, Optional con As ADODB.Connection, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim i As Long
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
cmd.CommandText = pString
Dim loctable As New ADODB.Recordset
If Not IsEmpty(aParam) Then
    If Not IsEmpty(aParam) Then
        For i = 0 To UBound(aParam) Step 2
            cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
        Next
    End If
End If
Set myCmdText = cmd.Execute
End Function
Public Function myFunc(pFunction As String, Optional pParam1 As String = "", Optional pParam2 As String = "", Optional pParam3 As String = "", Optional pParam4 As String = "", Optional pParam5 As String = "", Optional pParam6 As String = "", Optional pParam7 As String = "", Optional pParam8 As String = "", Optional pParam9 As String = "", Optional pParam10 As String = "") As String
If pParam1 <> "" Then myFunc = pParam1
If pParam2 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam2
If pParam3 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam3
If pParam4 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam4
If pParam5 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam5
If pParam6 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam6
If pParam7 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam7
If pParam8 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam8
If pParam9 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam9
If pParam10 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam10
myFunc = pFunction & "(" & myFunc & ")"
End Function
Public Function MyFuncValue(pFunction As String, pCon As ADODB.Connection, Optional pParam1 As String = "", Optional pParam2 As String = "", Optional pParam3 As String = "", Optional pParam4 As String = "", Optional pParam5 As String = "", Optional pParam6 As String = "", Optional pParam7 As String = "", Optional pParam8 As String = "", Optional pParam9 As String = "", Optional pParam10 As String = "") As Variant
Dim cString As String
cString = myFunc(pFunction, pParam1, pParam2, pParam3, pParam4, pParam5, pParam6, pParam7, pParam8, pParam9, pParam10)
MyFuncValue = GetField("Select " & cString, pCon)
End Function
Public Function myField(pString As String, pField As String, Optional con As ADODB.Connection, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
Dim cmd As New ADODB.Command
Dim loctable As ADODB.Recordset
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set loctable = cmd.Execute
If Not loctable.EOF Then
    myField = loctable.Fields(pField)
ElseIf Not IsEmpty(pDef) Then
    myField = pDef
End If
End Function
Public Function cmd(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 1000) As ADODB.Command
Set cmd = New ADODB.Command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
End Function
