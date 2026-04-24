VERSION 5.00
Begin VB.Form InformfrmOK 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   840
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4215
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   -15
   End
   Begin VB.Label lbl_inform 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   225
      TabIndex        =   0
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "InformfrmOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Unload Me
End Sub

Private Sub Form_Load()
'Timer1.Interval = 900

End Sub
Private Sub Timer1_Timer()
'Timer1.Enabled = False
'Unload Me
End Sub
