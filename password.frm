VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form password 
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "password.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   5280
      Begin VB.TextBox xPass 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xUser 
         Height          =   360
         Left            =   135
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "≈”„ «·„” Œœ„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„… «·”—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   675
         Width           =   1005
      End
   End
   Begin VB.CheckBox xEditLogin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "»Ì«‰«  «·œŒÊ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1590
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   180
      MaskColor       =   &H00FFFFFF&
      Picture         =   "password.frx":12632
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   " ‘€Ì·"
      Top             =   1215
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   1905
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   555
      Left            =   1620
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1215
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   979
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "password.frx":14A9E
      Caption         =   "œŒÊ·"
      PictureAlignment=   10
   End
End
Attribute VB_Name = "PassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTimes As Integer, nTime, userTable As Recordset
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
If (Not HandlePassword(xPass.text, xuser.BoundText, GetCon, myFormatShort(Date))) Then
'If (Not HandlePassword(xPass.text, xuser.BoundText, GetCon, "1")) Then
    nTime = nTime + 1
    MsgBox ArbString("ﬂ·„… ”— €Ì— ’ÕÌÕ… " & vbCrLf & "„Õ«Ê·… " & nTime & " „‰  3")
    If nTime >= 3 Then
        End
    Else
        Exit Sub
    End If
End If

Dim fs As New FileSystemObject
sPath_App = RetSetting("PATH", App.Path & "\conf.txt")
If Not fs.FolderExists(sPath_App) Then sPath_App = App.Path

If Trim(LCase(RetSetting("BACKUP", App.Path & "\conf.txt"))) <> "no" And Not DefUser Then
    checkCopy False
End If

If xEditLogin.Value = 1 And bSupermode Then
    confFrm.Show 1
    End
End If
SaveSetting
Unload Me
Main.Show
Exit Sub
LOCALERROR:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
On Error GoTo myerror

ValidDate
sCatalog = "OLYMPIC"
sMdfName = "OLYMPIC"


'myLoadDirs
Dim sError As String
If MakeLocal(sError) <> "" Then
    MsgBox "„‘ﬂ·… «À‰«¡ ⁄„· «·„·› «·„ƒﬁ " & vbCrLf & sError
End If

strCon = LoadConString

Dim cError As String
cError = testData
If cError <> "ok" Then
    MsgBox cError
    GoTo myerror
End If

Set data1.Recordset = myRecordSet("SELECT * FROM USERS", GetCon)
Set xuser.RowSource = data1

xuser.ListField = "Desca"
xuser.BoundColumn = "Code"
xuser.BoundText = RetSetting("user", tempPath & "\password.txt")
Exit Sub
myerror:
    If Err.Number <> 0 Then MsgBox Err.Description
    confFrm.Show 1
    Err.Clear
    End
End Sub
'Private Sub MakeLocal()
'On Error GoTo myerror
'Dim fs As New FileSystemObject
'MyCreateFolder tempPath
'fs.CopyFile App.Path & "\temp.mdb", tempFile
'contemp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & tempFile
'Exit Sub
'myerror:
''MsgBox "„‘ﬂ·… ›Ï ‰”Œ «·„·› «·„ƒﬁ " & vbCrLf & Err.Number & vbCrLf & Err.Description
'MsgBox Err.Description
'Err.Clear
'End Sub

Private Sub xDate_Validate(Cancel As Boolean)
With xDate
If IsDate(.text) Then
    .text = Format(.text, "YYYY-MM-DD")
Else
    .text = ""
End If
End With
End Sub

Private Sub xPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdApply_Click
End Sub

Private Sub xUser_Click(Area As Integer)
'If Not xUser.MatchedWithList Then xUser.BoundText = ""
'cmdApply.Enabled = xUser.BoundText <> ""
End Sub
Private Sub xUser_LostFocus()
If Not xuser.MatchedWithList Then xuser.BoundText = ""
CmdApply.Enabled = xuser.BoundText <> ""
End Sub
Private Sub myLoadDirs()
'Dim sDrive As String
'If RetSetting("DRIVE_TEMP") <> "" Then
'LocalPath = App.Path
'tempPath = "d:\TempMrshd"
'tempFile = tempPath & "\temp.mdb"
'If MakeLocal <> "" Then
'    MsgBox "„‘ﬂ·… «À‰«¡ ⁄„· «·„·› «·„ƒﬁ "
'End If
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
            & "catalog=" & myCatalog & ";Trusted_Connection=yes"
Else
    LoadConString = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
            & "catalog=" & myCatalog & ";user id = " & cUserId & ";" & "password = " & cPassword
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set PassWord = Nothing
End Sub
Private Sub SaveSetting()
addSetting "editLogin", xEditLogin.Value, tempPath & "\password.txt"
addSetting "user", xuser.BoundText, tempPath & "\password.txt"
End Sub
Private Function PrepData(nErr As Long, Optional sError As String)
End Function
Private Sub CreateUserMaster()
On Error GoTo myerror

Dim cString As String, cServerName As String
cServerName = RetSetting("server", App.Path & "\conf.txt")
If cServerName = "" Then cServerName = MyParn(".\SQLEXPRESS")
cString = "provider=SQLOLEDB;data source= " & cServerName & " ;initial " _
        & "catalog=master;Trusted_Connection=yes"
Dim con As New ADODB.Connection
con.Open cString

cString = "IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name = N'elmorshed') " & _
           "CREATE LOGIN [elmorshed] WITH PASSWORD=N'2010', DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF"
createCommand cString, con

cString = "EXEC sys.sp_addsrvrolemember @loginame = N'elmorshed', @rolename = N'sysadmin'"
createCommand cString, con
LastSub:
closeCon con
Exit Sub
myerror:
   MsgBox Err.Description
   Err.Clear
   GoTo LastSub
End Sub

Private Function createCommand(pString As String, pCon As ADODB.Connection) As Boolean
Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = pCon
FS1.CommandText = pString
FS1.Execute
Set FS1 = Nothing
End Function
Private Function CheckId() As Boolean
Dim cString As String, cId As String
cId = CpuId
'cString = RetSetting("ID", App.Path & "\controls.ocx")
If cString = "" Then Exit Function
If Trim(crypt(cId, "MRVB1971")) = Trim(cString) Then CheckId = True
End Function
Private Function testData() As String
Dim cString As String, cError As String
cError = openCon(GetCon)
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
    testData = cError
    Exit Function
    'confFrm.Show 1
End If
testData = "ok"
End Function
Private Function CreateRemote() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString

cString = "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'LoginMode', REG_DWORD, 2"
createCommand cString, conMaster

closeCon conMaster
CreateRemote = "ok"
Exit Function
myerror:
   CreateRemote = Err.Description
   Err.Clear
End Function
Private Function createLogin() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cServerName As String, cString As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString
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
Private Function AttachData() As String
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
Private Function bringOnLine() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString

Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = conMaster
cString = "alter database " & sCatalg
cString = cString & turn(cString, vbCrLf) & "set online"
FS1.CommandText = cString
FS1.Execute
bringOnLine = "ok"
Exit Function
myerror:
bringOnLine = Err.Description
Err.Clear
End Function
Private Function IsDir(strPath As String) As Boolean
  If Len(Dir$(strPath, vbNormal)) = 0 Then
    IsDir = True
  Else
    IsDir = False
  End If
End Function

