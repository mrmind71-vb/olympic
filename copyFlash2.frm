VERSION 5.00
Begin VB.Form copyflashfrm 
   Caption         =   "⁄„· ‰”Œ… «Õ Ì«ÿÌ…"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRestoreFrom 
      Caption         =   "«” —Ã«⁄ «·‰”Œ…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      TabIndex        =   2
      Top             =   720
      Width           =   4560
   End
   Begin VB.CommandButton cmdExit 
      Height          =   645
      Left            =   45
      Picture         =   "copyFlash2.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1395
      Width           =   4560
   End
   Begin VB.CommandButton cmdCopyTo 
      Caption         =   "⁄„· ‰”Œ… «Õ Ì«ÿÌ… "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4560
   End
End
Attribute VB_Name = "copyflashfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nMaxFiles As Long
Private Sub cmdCopyto_Click()
On Error GoTo myerror
Me.MousePointer = 11
If copyToFlash Then MsgBox " „ «·‰”Œ »‰Ã«Õ", , "Copy result"
Me.MousePointer = 1
Exit Sub
myerror:
MsgBox Err.Description, , "Error description"
Err.Clear
Me.MousePointer = 1
End Sub
Private Function copyToFlash() As Boolean
On erorr GoTo myerror
Dim fs As FileSystemObject, cDir As String, cFileName As String, cDrive As String
Set fs = CreateObject("Scripting.FileSystemObject")
'cDrive = LastDrive(True)
'cDrive = "C"
'cDir = cDrive & ":\DataBackup"
'MyCreateFolder cDir
cDir = App.Path
FixFiles (cDir)
cFileName = cDir & "\" & Format(Date, "yyyymmdd") & ".bak"
If createBackUp(cFileName) Then copyToFlash = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdRestoreFrom_Click()
Dim cString As String
cString = Trim(LCase(InputBox("«œŒ· ﬂ·„… «·”—")))
If cString = "20132014" Then
    If RestoreFromFlash Then
        MsgBox " „ «” —Ã«⁄ «·»Ì«‰«  »‰Ã«Õ!!«·—Ã«¡ › Õ «·»—‰«„Ã „—… «Œ—Ì"
        End
    End If
ElseIf cString <> "" Then
    MsgBox "ﬂ·„… ”— €Ì— ’ÕÌÕ"
End If
End Sub

Private Sub Form_Load()
nMaxFiles = 9
Me.cmdRestoreFrom.Enabled = cExpress <> ""
End Sub
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
        If IsNumeric(Mid(File.Name, 1, Len(File.Name) - 4)) Then FileCount = FileCount + 1
    End If
Next


ReDim fNames(FileCount)
cFcount = 0

For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, 1, Len(File.Name) - 4)) Then
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
Private Function FixFiles(pDir As String) As Boolean
Dim fs As New FileSystemObject
Dim aret As Variant, nDelete As Long
On Error Resume Next
aret = retFArray(pDir, "bak")
nDelete = (UBound(aret) + 1) - nMaxFiles
For I = 0 To (nDelete)
    fs.DeleteFile pDir & "\" & aret(I)
Next
Err.Clear
End Function
Private Function LastFile(pDir As String) As String
Dim fs As New FileSystemObject
Dim aret As Variant
On Error Resume Next
aret = retFArray(pDir, "bak")
If UBound(aret) > 0 Then
    LastFile = aret(UBound(aret)) & ""
End If
Err.Clear
End Function
Private Function createBackUp(pFileName) As Boolean
Dim cFile As String
Dim con As New ADODB.Connection
openCon con

Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText

Set FS1.ActiveConnection = con
cString = "BACKUP DATABASE [" & sCatalog & "] TO  DISK = N'" & pFileName & "' WITH NOFORMAT, INIT,  NAME = N'TABLES-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
FS1.CommandText = cString

FS1.Execute
Set FS1 = Nothing
closeCon con
createBackUp = True
End Function
Private Function RestoreBackUp(pFileName) As Boolean
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString

Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = conMaster
cString = "alter database  [" & sCatalog & "] set offline with rollback immediate"
cString = cString & turn(cString, vbCrLf) & "alter database [" & sCatalog & "]"
cString = cString & turn(cString, vbCrLf) & "set online"

FS1.CommandText = cString
FS1.Execute

cString = "RESTORE DATABASE [" & sCatalog & "] FROM  DISK = N'" & pFileName & "' WITH  FILE = 1,  NOUNLOAD,  REPLACE,  STATS = 10"
FS1.CommandText = cString
FS1.Execute
Set FS1 = Nothing
closeCon conMaster
RestoreBackUp = True
End Function
Private Function RestoreFromFlash() As Boolean
On erorr GoTo myerror
Dim fs As FileSystemObject, cDir As String, cFileName As String, cDrive As String, cLastFile As String
Set fs = CreateObject("Scripting.FileSystemObject")
'cDir = LastDrive(True)
'cDir = "C"
'cDir = cDir & turn(cDir, ":\") & "DataBackup"
cDir = App.Path
cFileName = LastFile(cDir)
If cFileName = "" Then
    MsgBox "·« ÌÊÃœ „·› ·«” —Ã«⁄ «·»Ì«‰«  „‰Â"
    Exit Function
End If
cFileName = cDir & turn(cDir, "\") & cFileName
If RestoreBackUp(cFileName) Then RestoreFromFlash = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function



