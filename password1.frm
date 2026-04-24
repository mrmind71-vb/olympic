VERSION 5.00
Begin VB.Form PassWord1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   1545
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
   ScaleHeight     =   1545
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      BackColor       =   &H00CAD29F&
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
      Left            =   1350
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   825
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00CAD29F&
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
      Left            =   150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   825
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
Attribute VB_Name = "PassWord1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTimes As Integer, tSecurity As Recordset
Dim PassTable As Recordset
Dim cComp As String
Private Sub CmdApply_Click()
Dim lOk As Boolean
PassTable.FindFirst " PASS = " & MyParn(xPass.Text)
If Not PassTable.NoMatch Then
    cComp = TurnValue(PassTable.comp, Null, "«·„—‘‹‹‹‹‹‹‹‹‹‹œ")
    firsttitle = TurnValue(PassTable.comp, Null, "«·„—‘‹‹‹‹‹‹‹‹‹‹œ")
    lManger = PassTable.MANGER
    MdbPath = App.Path & "\data\data.mdb"
    PublicPath = App.Path
    Set mydb = OpenDatabase(MdbPath)
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
PublicPath = App.Path
TempPath = App.Path & "\temp.mdb"
Set tempdb = OpenDatabase(TempPath)
Set SecDB = OpenDatabase(App.Path & "\SEC.MDB")
Set PassTable = SecDB.OpenRecordset("SELECT * FROM FILE_SS")
Exit Sub
MyError:
    End
    MsgBox "Â‰«ﬂ Œÿ√ „« ﬁœ ÕœÀ «À‰«¡ › Õ «·„·› «·„ƒﬁ !!"
End Sub
Function RightPassword()
tSecurity.FindFirst "pass = " & MyParn(xPass)
RightPassword = Not tSecurity.NoMatch
End Function
Function rightPreviousPath() As Boolean
If IsNull(tSecurity!Previous) Then GoTo 1000
On Error Resume Next
Set Previousdb = OpenDatabase(tSecurity!Previous & "\namko.mdb")
1000 rightPreviousPath = (Err.Number = 0)
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tSecurity.Close
Set tSecurity = Nothing
End Sub

