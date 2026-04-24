VERSION 5.00
Begin VB.Form SettingFrm 
   Caption         =   "÷»ÿ „”«— „·› «·„Õ·"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1500
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1035
      Width           =   1500
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   " ⁄œÌ·"
      Height          =   375
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1035
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      Begin VB.TextBox xPath 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   7890
      End
      Begin VB.TextBox xcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   6315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„”«— «·„·›«  :"
         Height          =   240
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂÊœ «·⁄„Ì· :"
         Height          =   330
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1005
      End
   End
End
Attribute VB_Name = "SettingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdUpdate_Click()
Dim aInsert(1, 1)
aInsert(0, 0) = "code"
aInsert(0, 1) = addstring(xcode.Text)

aInsert(1, 0) = "Path"
aInsert(1, 1) = addstring(xPath.Text)
On Error GoTo MYERROR
con.BeginTrans
If GetDesca("Select count(*) from path") = "0" Then
   con.Execute CreateInsert(aInsert, "path")
Else
   con.Execute CreateUpdate(aInsert, "path", "")
End If
con.CommitTrans
Inform " „ «· ⁄œÌ· »‰Ã«Õ"
Exit Sub
MYERROR:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
OpenCon con
myload
xPath.Text = "\\Al-salam\elmorshed\DATA"
End Sub
Private Sub myload()
Dim aret As Variant
aret = aGetDesca("Select code,Path from path")
xcode.Text = aret(1) & ""
xPath.Text = aret(2) & ""
xCode_LostFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode_LostFocus()
xcode.BackColor = &H80000005
xCodeDesca.Caption = ""
If xcode.Text = "" Then Exit Sub
xcode.Text = RetZero(xcode.Text, 6)
xCodeDesca.Caption = GetDesca("select desca from file3_10 where code = " & MyParn(xcode.Text)) & ""
End Sub
