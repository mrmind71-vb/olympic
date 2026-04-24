VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form restore_datafrm 
   Caption         =   "Úãá äÓÎÉ ÇÍÊíÇØíÉ"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Height          =   690
      Left            =   90
      Picture         =   "restore_data.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   855
      Width           =   4515
   End
   Begin Threed.SSCommand cmdCopyTo 
      Height          =   735
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1296
      _Version        =   196610
      ForeColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "restore_data.frx":246C
      Caption         =   "ÇÓÊÑÌÇÚ äÓÎÉ ÇÍÊíÇØíÉ"
      ButtonStyle     =   1
      PictureAlignment=   1
      BevelWidth      =   10
      ShapeSize       =   1
   End
End
Attribute VB_Name = "restore_datafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nMaxFiles As Long
Private Sub cmdCopyto_Click()
On Error GoTo myerror
Me.MousePointer = 11
If Restore_data Then MsgBox "Êã ÇÓÊÑÌÇÚ äÓÎÉ ÇÍÊíÇØíÉ ÈäÌÇÍ", , "ÇÓÊÑÌÇÚ äÓÎÉ ÇÍÊíÇØíÉ"
lastsub:
Me.MousePointer = 1
Exit Sub
myerror:
MsgBox Err.Description, , "Error description"
Err.Clear
Me.MousePointer = 1
GoTo lastsub
End Sub
Private Function copyToFlash() As Boolean
On erorr GoTo myerror
Dim fs As FileSystemObject, cDir As String, cFileName As String, cDrive As String
Set fs = CreateObject("Scripting.FileSystemObject")
cDrive = LastDrive(True)
cDir = cDrive & ":\DataBackup"
If RestoreBackUp(cFileName) Then copyToFlash = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
nMaxFiles = 9
End Sub
Function retFArray(pFolder As String)
Dim fso As New FileSystemObject, FileCount As Long
Dim fNames()
ReDim fNames(0)

If Not fso.FolderExists(pFolder) Then
    retFArray = fNames
    Exit Function
End If

Set fold = fso.GetFolder(pFolder)
For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = ".bak" And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, 1, Len(File.Name) - 4)) Then FileCount = FileCount + 1
    End If
Next


ReDim fNames(FileCount)
cFcount = 0

For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = ".bak" And Len(File.Name) > 4 Then
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
aret = retFArray(pDir)
nDelete = (UBound(aret) + 1) - nMaxFiles
For i = 0 To (nDelete)
    fs.DeleteFile pDir & "\" & aret(i)
Next
Err.Clear
End Function
Private Function RestoreBackUp(pFileName As String) As Boolean
Dim aret As Variant
aret = retFArray(pFileName)
If UBound(aret) > 0 Then
    pFileName = pFileName & turn(pFileName, "\") & aret(UBound(aret))
End If

Dim cFile As String
Dim con As New ADODB.Connection
openCon con

Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText

Set FS1.ActiveConnection = con
cString = "BACKUP DATABASE " & sCatalog & " TO  DISK = N'" & pFileName & "' WITH  RETAINDAYS = 1, NOFORMAT, INIT,  NAME = N'over-Full Database Backup', SKIP,  NOREWIND, NOUNLOAD, STATS = 10"
FS1.CommandText = cString

FS1.Execute
Set FS1 = Nothing
closeCon con
createBackUp = True
End Function
