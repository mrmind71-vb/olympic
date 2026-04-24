Attribute VB_Name = "NewMdl"
Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
                            As Long, ByVal dwMilliseconds As Long) As Long

Public Function GetExecutableForFile(strFileName As String) As String
   Dim lngRetval As Long
   Dim strExecName As String * 255
   lngRetval = FindExecutable(strFileName, vbNullString, strExecName)
   GetExecutableForFile = Left$(strExecName, InStr(strExecName, Chr$(0)) - 1)
End Function
Sub RunIt(strNewFullPath As String, Optional pType As Double = vbNormalFocus)
   Dim exeName As String

   exeName = GetExecutableForFile(strNewFullPath)
   Shell exeName & " " & Chr(34) & strNewFullPath & Chr(34), pType
End Sub
Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
End Function
Function ValidNumChr(ByVal pCode As Variant, Optional bZero As Boolean = False) As Boolean
Dim sNumber As String
If Len(pCode & "") > 18 Then Exit Function
pCode = Trim(pCode & "")
If Mid(pCode, 1, 1) = "0" And (Len(pCode) > 1) Then Exit Function
For I = 1 To Len(pCode)
    If Not IsNumeric(Mid(pCode, I, 1)) Then
        Exit Function
    End If
Next
ValidNumChr = True
End Function
Function myFormatShort(sDate As Variant) As String
myFormatShort = Format(sDate, "YYYYMM")
End Function


