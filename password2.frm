VERSION 5.00
Begin VB.Form PassWord2 
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
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
   ScaleHeight     =   1065
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ‘€Ì·"
      Height          =   420
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   585
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   1140
   End
   Begin VB.TextBox xPass 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2610
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·”— :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   210
      Width           =   870
   End
End
Attribute VB_Name = "PassWord2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm
Dim nTry As Integer
Private Sub cmdApply_Click()
'If LCase(Trim(xPass.Text)) = "mor2008" Then
If LCase(Trim(xPass.Text)) = "1" Then
    myForm.bRetvalue = True
    Unload Me
Else
    MsgBox "ﬂ·„… «·”— €Ì— ’ÕÌÕ…"
    nTry = nTry + 1
    If nTry = 3 Then
        myForm.bRetvalue = False
        Unload Me
    End If
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set PassWord2 = Nothing
End Sub
