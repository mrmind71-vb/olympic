Attribute VB_Name = "data"
Public strCon As String
Public GetCon As New ADODB.Connection
Public cExpress As String, sPath_App As String
Public aSetting As Variant
Public sCatalog As String, sMdfName As String
Function openCon(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "", Optional nTimeOut As Long = 5) As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = strCon Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.CommandTimeout = nTimeOut
pCon.Open cString
openCon = "ok"
Exit Function
myerror:
openCon = Err.Description
Err.Clear
End Function
Function openConMdb(ByRef pCon As ADODB.Connection, ByVal pString As String)
pCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & pString
End Function
Function closeCon(ByRef pCon As ADODB.Connection) As Boolean
On Error GoTo myerror
If pCon.State = adStateOpen Then pCon.Close
Set pCon = Nothing
closeCon = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Function ReadFile(cFile) As String
Dim TextLine
On Error GoTo myerror
Open cFile For Input As #1   ' Open file.
Do While Not EOF(1)
    Line Input #1, TextLine  ' Read line into variable.
    ReadFile = ReadFile & turn(ReadFile, " ") & TextLine
Loop
Close #1   ' Close file.
Exit Function
myerror:
Err.Clear
ReadFile = ""
End Function
Function createCommand(pString As String, pCon As ADODB.Connection) As Boolean
Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = pCon
FS1.CommandText = pString
FS1.Execute
Set FS1 = Nothing
End Function
Sub fixSetting()
aSetting = GetFields("select TOP 1 * from address ORDER BY ID DESC", GetCon)
End Sub
Public Function LoadConString(Optional aServer As Variant = Empty, Optional pCatalog As String = "")
Dim cServerName As String, cUserId As String, cPassword As String
Dim myCatalog As String
myCatalog = IIf(pCatalog = "", sCatalog, pCatalog)
If IsEmpty(aServer) Then
    cServerName = RetSetting("server", App.Path & "\conf.txt")
    cUserId = decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
    cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")
Else
    cServerName = retFlag(aServer, "server")
    cUserId = retFlag(aServer, "userId")
    cPassword = retFlag(aServer, "password")
End If
If cServerName = "" Then cServerName = "." & turn(cExpress, "\" & cExpress)
If cUserId = "" Or cPassword = "" Then
    LoadConString = "provider=SQLOLEDB;data source= " & cServerName & " ;initial " _
            & "catalog=" & myCatalog & ";Trusted_Connection=yes;Connection Timeout = 2"
Else
    LoadConString = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
            & "catalog=" & myCatalog & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Connection Timeout = 2"
End If
End Function
Public Function testData() As String
Dim cString As String, cError As String
cString = LoadConString
cError = openCon(GetCon, cString)
If cError = "ok" Then
    testData = "ok"
    Exit Function
End If

cString = "The user is not associated with a trusted SQL Server connection."
If LCase(Right(cError, Len(cString))) = LCase(cString) Then
    cError = CreateRemote
    If cError = "ok" Then
        Inform " „ «÷«›… ’·«ÕÌ«  «·»Ì«‰«  »‰Ã«Õ «·—Ã«¡ «⁄«œ…  ‘€Ì· «·ÃÂ«“ »⁄œ «·«‰ Â«¡"
        End
    Else
        MsgBox cError
        testData = cError
    End If
End If

cString = "Login failed for user"
If LCase(Mid(cError, 1, 21)) = LCase(cString) Then
    cError = createLogin
    If cError = "ok" Then
        Inform " „ «÷«›… „” Œœ„ »‰Ã«Õ"
        cError = openCon(GetCon)
    End If
End If

If cError <> "ok" Then
    cString = "Cannot open database"
    If Left(LCase(cError), 20) = LCase(cString) Then
        cError = AttachData
        If cError = "ok" Then Inform " „ —»ÿ «·»Ì«‰«  »‰Ã«Õ"
        cError = openCon(GetCon)
    End If
End If

If cError <> "ok" Then
    cString = "Cannot open database"
    If Left(LCase(cError), 20) = LCase(cString) Then
        cError = bringOnLine
        If cError = "ok" Then
            Inform " „ › Õ «·„·› »‰Ã«Õ"
            cError = openCon(GetCon)
         End If
    End If
End If

If cError <> "ok" Then
    MsgBox cError
    confFrm.Show 1
End If
testData = "ok"
End Function
Public Function CreateRemote() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
conMaster.Open LoadConString(, "master")

cString = "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'LoginMode', REG_DWORD, 2"
createCommand cString, conMaster

closeCon conMaster
CreateRemote = "ok"
Exit Function
myerror:
   CreateRemote = Err.Description
   Err.Clear
End Function
Public Function createLogin() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection, cString As String

cString = LoadConString(, "master")

cString = "CREATE LOGIN [elmorshed] WITH PASSWORD=N'2010', DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF"
cString = cString & turn(cString, vbCrLf) & "EXEC sys.sp_addsrvrolemember @loginame = N'elmorshed', @rolename = N'sysadmin'"
createCommand cString, conMaster
closeCon conMaster
createLogin = "ok"
Exit Function
myerror:
   createLogin = Err.Description
   Err.Clear
End Function
Public Function AttachData() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection, cString As String, aData As Variant
Dim cFile As String

cString = LoadConString(aData, "master")
conMaster.Open cString

cFile = App.Path & "\mdf\" & sMdfName
cString = "CREATE DATABASE [" & sCatalog & "] ON (FILENAME = N'" & cFile & ".mdf" & "' )," & _
      "(FILENAME = N'" & cFile & "_LOG.ldf" & "' )" & _
      " FOR ATTACH"

createCommand cString, conMaster

closeCon conMaster
AttachData = "ok"
Exit Function
11 myerror:
   AttachData = Err.Description
   Err.Clear
End Function
Public Function bringOnLine() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String

cString = LoadConString(aData, "master")
conMaster.Open cString

Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = conMaster
cString = "alter database [" & sCatalog & "]"
cString = cString & turn(cString, vbCrLf) & "set online"
FS1.CommandText = cString
FS1.Execute
bringOnLine = "ok"
Exit Function
myerror:
bringOnLine = Err.Description
Err.Clear
End Function
Public Function checkCopy(Optional bNetWork As Boolean = True, Optional bCopy As Boolean = False) As Boolean
'Dim aDrive As Variant, fs As New FileSystemObject, nCount As Long, i As Long
Dim fs As New FileSystemObject
'aDrive = aLastDrive()

'nCount = retFlag(aDrive, "COUNT")

'For i = nCount To 1 Step -1
 '   On Error Resume Next
    'aDrive = aLastDrive(, i)
    'cDir = retFlag(aDrive, "LETTER") & ":\DataBackup"
    
'    fs.CopyFile App.Path & "\temp.mdb", cDir & "\TEMP_CHECK.MDB", True
 '   If Err.Number = 0 Then Exit For Else Err.Clear
'Next

'Dim cDir As String
'cDir = RetSetting("PATH", App.Path & "\conf.txt")
'If Not fs.FolderExists(cDir) Then
'    MsgBox "„”«— «·„’«œ— €Ì— „ÊÃÊœ"
'    Exit Function
'End If

Dim cDir As String
cDir = sPath_App & "\BACKUP"
If Not MyCreateFolder(cDir) Then
    MsgBox "„”«— «·‰”Œ €Ì— ’«·Õ"
    Exit Function
End If

cFileName = cDir & "\" & sCatalog & "_" & Format(Date, "yyyymmdd") & ".bak"
If (Not fs.FileExists(cFileName)) Or bCopy Then
    MsgBox "”Ì „ ⁄„· ‰”Œ… «Õ Ì«ÿÌ…"
    If DoBackUp(cDir) Then
        Inform " „ ⁄„· ‰”Œ… »‰Ã«Õ"
        checkCopy = True
    End If
End If
End Function
Function DoBackUp(Optional cDir, Optional bRecopy As Boolean = True, Optional bNetWork As Boolean = True) As Boolean
On Error GoTo myerror
'cDir = sDrive & ":\DataBackup"
Dim sMsg As String
'MyCreateFolder (cDir)
FixFiles (cDir)
'FixFiles App.Path
cFileName = cDir & "\" & sCatalog & "_" & Format(Date, "yyyymmdd") & ".bak"
If createBackUp(cFileName) Then DoBackUp = True
DoEvents
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Function createBackUp(pFileName, Optional bRecopy As Boolean = True, Optional bNetWork As Boolean = True) As Boolean
Dim con As New ADODB.Connection, FS1 As New ADODB.Command, fs As New FileSystemObject, cFile As String
'If bNetWork Then
'    cFile = IIf(bNetWork, App.Path & "\" & sCatalog & "_" & Format(Date, "yyyymmdd") & ".bak", pFileName)
'Else
''    cFile = pFileName
'End If
On Error GoTo myerror
If (Not fs.FileExists(pFileName)) Then
    openCon con, LoadConString
    FS1.CommandType = adCmdText
    Set FS1.ActiveConnection = con
    cString = "BACKUP DATABASE " & sCatalog & " TO  DISK = N'" & pFileName & "' WITH  RETAINDAYS = 1, NOFORMAT, INIT,  NAME = N'over-Full Database Backup', SKIP,  NOREWIND, NOUNLOAD, STATS = 10"
    FS1.CommandText = cString
    FS1.Execute
    Set FS1 = Nothing
    closeCon con
End If
'If bNetWork Then fs.CopyFile cFile, pFileName
createBackUp = True
Err.Clear
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Function aLastDrive(Optional ntype As Integer = -1, Optional nCount As Long = -1) As Variant
Dim fs, d, DC, letter, i As Long
Set fs = CreateObject("Scripting.FileSystemObject")
Set DC = fs.Drives
For Each d In DC
    If (d.DriveType = 1 Or d.DriveType = 2) And (ntype = -1 Or d.DriveType = ntype) Then
        i = i + 1
        On Error Resume Next
        aLastDrive = AddFlag(Empty, "LETTER", d.DriveLetter)
        aLastDrive = AddFlag(aLastDrive, "SERIAL", d.SerialNumber)
        aLastDrive = AddFlag(aLastDrive, "TYPE", d.DriveType)
        aLastDrive = AddFlag(aLastDrive, "COUNT", i)
        If i = nCount Then Exit For
    End If
Next
End Function
Function FixFiles(pDir As String, Optional nMaxFiles As Integer = 10) As Boolean
Dim fs As New FileSystemObject
Dim aRet As Variant, nDelete As Long
On Error Resume Next
aRet = retFArray(pDir, "bak")
nDelete = (UBound(aRet) + 1) - nMaxFiles
For i = 0 To (nDelete)
    fs.DeleteFile pDir & "\" & aRet(i)
Next
Err.Clear
End Function
Function retFArray(pFolder As String, sExt As String) As Variant
Dim fso As New FileSystemObject, FileCount As Long
Dim fNames()
ReDim fNames(0)
If Not fso.FolderExists(pFolder) Then
    retFArray = fNames
    Exit Function
End If
Set fold = fso.GetFolder(pFolder)
For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, Len(sCatalog) + 2, 8)) Then FileCount = FileCount + 1
    End If
Next


ReDim fNames(FileCount)
cFcount = 0

For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, Len(sCatalog) + 2, 8)) Then
            cFcount = cFcount + 1
            fNames(cFcount) = LCase(File.Name)
        End If
    End If
Next

For tName = 1 To FileCount
    For nName = (tName + 1) To FileCount
        If StrComp(fNames(tName), fNames(nName), 0) = 1 Then
            buffer = fNames(nName)
            fNames(nName) = fNames(tName)
            fNames(tName) = buffer
        End If
    Next
Next
retFArray = fNames
End Function
Function retAllArray(pFolder As String, sExt As String) As Variant
Dim fso As New FileSystemObject, FileCount As Long
Dim fNames()
ReDim fNames(0)
If Not fso.FolderExists(pFolder) Then
    retAllArray = fNames
    Exit Function
End If
Set fold = fso.GetFolder(pFolder)
For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
         FileCount = FileCount + 1
    End If
Next


ReDim fNames(FileCount)
cFcount = 0

For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        cFcount = cFcount + 1
        fNames(cFcount) = LCase(File.Name)
    End If
Next

For tName = 1 To FileCount
    For nName = (tName + 1) To FileCount
        If StrComp(fNames(tName), fNames(nName), 0) = 1 Then
            buffer = fNames(nName)
            fNames(nName) = fNames(tName)
            fNames(tName) = buffer
        End If
    Next
Next
retAllArray = fNames
End Function
Public Function CreateConStr2() As String
Dim aServer As Variant
aServer = AddFlag(Empty, "SERVER", RetSetting("SERVER2", App.Path & "\conf.txt"))
aServer = AddFlag(aServer, "USERID", RetSetting("USERID2", App.Path & "\conf.txt"))
aServer = AddFlag(aServer, "PASSWORD", RetSetting("PASSWORD2", App.Path & "\conf.txt"))
CreateConStr2 = LoadConString(aServer, "client")
End Function
Public Function CreateConStr3() As String
Dim aServer As Variant
aServer = AddFlag(Empty, "SERVER", RetSetting("SERVER", App.Path & "\conf.txt"))
aServer = AddFlag(aServer, "USERID", RetSetting("USERID", App.Path & "\conf.txt"))
aServer = AddFlag(aServer, "PASSWORD", RetSetting("PASSWORD", App.Path & "\conf.txt"))
CreateConStr3 = LoadConString(aServer, "client")
End Function
Function MyParnAndNo(cSearch, cField) As String
Dim aString, cString2
aString = Split(Trim(cSearch), " ")
For i2 = 0 To UBound(aString)
    If Trim(aString(i2)) <> "" Then cString2 = cString2 & IIf(cString2 = "", "", " and ") & " ( Not " & cField & " Like " & " '%" & aString(i2) & "%')"
Next
MyParnAndNo = cString2
End Function






