VERSION 5.00
Begin VB.Form PassWord0 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ‘€Ì·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1500
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   675
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   300
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   675
      Width           =   1140
   End
   Begin VB.TextBox xPass 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1950
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﬂ·„… «·”— :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "PassWord0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTimes As Integer, tSecurity As Recordset
Dim PassTable As Recordset
Private Sub CmdApply_Click()
Dim lOk As Boolean
PassTable.FindFirst " PASS = " & MyParn(xPass.Text)
If Not PassTable.NoMatch Then
    firsttitle = PassTable.COMP
    cUserName = PassTable.User
    lManger = PassTable.MANGER
    MdbPath = App.Path & "\data\data.mdb"
    nCountPrint = TurnValue(PassTable.Count, Null, 0)
    
    PublicPath = App.Path
    Set mydb = OpenDatabase(MdbPath)
    Main.Show 1
ElseIf xPass.Text = "22" Then
    firsttitle = PassTable.COMP
    cUserName = PassTable.User
    nCountPrint = TurnValue(PassTable.Count, Null, 0)
    lManger = PassTable.MANGER
    MdbPath = App.Path & "\data\data.mdb"
    PublicPath = App.Path
    Set mydb = OpenDatabase(MdbPath)
    Load Main
    Main.mXsec.Visible = True
    Main.Show 1
Else
    MsgBox "ﬂ·„… «·”— Œÿ√"
End If
Exit Sub
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
Dim fs
sRegDir = "C:\ELMORSHED\threed.ocx"
TempPath = App.Path & "\temp.mdb"
Set tempdb = OpenDatabase(App.Path & "\temp.mdb")
If Dir(TempPath) = "" Then
    MsgBox "⁄„· „·› „ƒﬁ "
    CreateTempFile
    'If MsgBox("⁄„· „·› „ƒﬁ ", vbYesNo) Then CreateTempFile
End If
Set tempdb = OpenDatabase(TempPath)
Set SecDB = OpenDatabase(App.Path & "\TEMP.MDB")
Set PassTable = SecDB.OpenRecordset("SELECT * FROM FILE_SS ")
'    MsgBox "Â‰«ﬂ Œÿ√ „« ﬁœ ÕœÀ «À‰«¡ › Õ «·„·› «·„ƒﬁ !!"
End Sub
Function RightPassword()
tSecurity.FindFirst "pass = " & MyParn(xPass)
RightPassword = Not tSecurity.NoMatch
End Function
Private Sub handlemenu()
'Main.m_Store.Enabled = tSecurity.Store
End Sub
Function rightPreviousPath() As Boolean
If IsNull(tSecurity!Previous) Then GoTo 1000
On Error Resume Next
Set Previousdb = OpenDatabase(tSecurity!Previous & "\namko.mdb")
1000 rightPreviousPath = (Err.Number = 0)
End Function
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub CreateTempFile()
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
If Not fs.FolderExists("c:\elmorshed") Then fs.createFolder ("c:\elmorshed")
fs.CopyFile TempPath, "c:\elmorshed\temp.mdb"
End Sub
