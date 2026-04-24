Attribute VB_Name = "DirsMdl"
Public contemp As New ADODB.Connection
Public Function MakeLocal(ByRef sError As String) As String
Dim aDrive As Variant
Dim i As Long
aDrive = aLastDrive(2)
nCount = retFlag(aDrive, "COUNT")
'If nCount > 1 Then
'    aDrive = aLastDrive(2, 2)
'    sDrive = retFlag(aDrive, "LETTER")
'    tempPath = sDrive & ":\TempMrshd" & RetSetting("BRANCH", App.Path & "\conf.txt")
'    sError = createLocal
'    If sError = "" Then Exit Function
'End If
Dim sDrive As String
For i = IIf(nCount > 1, 2, 1) To 1 Step -1
    aDrive = aLastDrive(2, i)
    sDrive = retFlag(aDrive, "LETTER")
    tempPath = sDrive & ":\TempMrshd" & RetSetting("BRANCH", App.Path & "\conf.txt")
    sError = createLocal
    If sError = "" Then Exit Function
Next

For i = 1 To nCount
    aDrive = aLastDrive(2, i)
    sDrive = retFlag(aDrive, "LETTER")
    tempPath = sDrive & ":\TempMrshd" & RetSetting("BRANCH", App.Path & "\conf.txt")
    sError = createLocal
    If sError = "" Then Exit Function
Next
MakeLocal = sError
End Function
Public Function createLocal() As String
On Error GoTo myerror
Dim fs As New FileSystemObject
MyCreateFolder tempPath
tempFile = tempPath & "\temp.mdb"
fs.CopyFile App.Path & "\temp.mdb", tempFile
contemp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & tempFile
Exit Function
myerror:
createLocal = Err.Description
Err.Clear
End Function
