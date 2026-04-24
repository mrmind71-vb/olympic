Attribute VB_Name = "compess_files"
Public Function ZipFolder(zip As ChilkatZip, pSource As String, pTarget As String, Optional bNew As Boolean = False, Optional ByRef pError As String, Optional pform As Form) As Boolean
Dim success As Long
Dim fs As New FileSystemObject

On Error GoTo myerror

If Not fs.FolderExists(pSource) Then
    pError = "„”«— «·„’œ— €Ì— ’ÕÌÕ"
    Exit Function
End If

If bNew Or (Not fs.FileExists(pTarget)) Then
    success = zip.NewZip(pTarget)
End If

If (success <> 1) Then
    pError = zip.LastErrorText
    Exit Function
End If


Dim fso As New FileSystemObject
Dim fld As Folder
Dim File As File
Dim nCount As Long
Dim sCaption As String


Dim nRecordcount As Long
Set fld = fso.GetFolder(pSource)
nRecordcount = fld.Files.Count

If Not pform Is Nothing Then
    pform.Prog1.Value = 0
    pform.Prog1.Visible = True
    sCaption = pform.Caption
End If

Dim nRecord As Long
For Each File In fld.Files
    If Not pform Is Nothing Then
        success = zip.AppendOneFileOrDir(File.Path, 0)
    End If
    If success = 0 Then
        pError = zip.LastErrorText
        GoTo myExit
    End If
    
    nRecord = nRecord + 1
    If Not pform Is Nothing Then
        pform.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordcount
        pform.Prog1.Value = mRound(nRecord / nRecordcount, 2) * 100
    End If
Next

If Not pform Is Nothing Then
    pform.Caption = sCaption & " " & "Ì „ ﬂ «»… «·„·› «·„÷€Êÿ"
End If


success = zip.WriteZipAndClose()
If (success <> 1) Then
    pError = zip.LastErrorText
    GoTo myExit
End If
ZipFolder = True

myExit:
If Not pform Is Nothing Then
    pform.Prog1.Value = 0
    pform.Prog1.Visible = False
    pform.Caption = sCaption
End If
Exit Function
myerror:
    pError = Err.Description
    Err.Clear
End Function
Public Function UnZipFolder(myZip As ChilkatZip, pSource As String, pTarget As String, ByRef pError As String, pform As Form) As Boolean
Dim success As Long
Dim fs As New FileSystemObject

On Error GoTo myerror

If Not fs.FileExists(pSource) Then
    pError = "«·„·› «·„÷€Êÿ €Ì— „ÊÃÊœ"
    Exit Function
End If

If Not fs.FolderExists(pTarget) Then
    pError = "„”«— «·„·› €Ì— „ÊÃÊœ"
    Exit Function
End If

success = myZip.OpenZip(pSource)

If (success <> 1) Then
    pError = myZip.LastErrorText
    Exit Function
End If

Dim sCaption As String
If Not pform Is Nothing Then
    sCaption = pform.Caption
    pform.Caption = sCaption & " " & "Ì „ ›ﬂ «·„·› «·„÷€Êÿ"
End If

Dim nCount As Long
myZip.PercentDoneScale = 100

unZipCount = myZip.Unzip(pTarget)

If (success <> 1) Then
    pError = myZip.LastErrorText
    GoTo lastFunc
End If
UnZipFolder = True
lastFunc:
If Not pform Is Nothing Then
    pform.Caption = sCaption
End If
Exit Function
myerror:
pError = Err.Description
Err.Clear
GoTo lastFunc
End Function

