VERSION 5.00
Begin VB.Form PaySal 
   BackColor       =   &H00E0E0E0&
   Caption         =   "”œ«œ «·»Ê‰"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox xVisa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C53C9&
      Height          =   350
      Left            =   1950
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Õ›‹‹‹‹‹‹‹Ÿ"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1125
      Width           =   1740
   End
   Begin VB.TextBox xCash 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C53C9&
      Height          =   350
      Left            =   1950
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   1500
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "”œ«œ ›Ì‹‹‹“«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3525
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "”œ«œ ‰ﬁ‹‹‹œÏ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   180
      Width           =   870
   End
End
Attribute VB_Name = "PaySal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    If Val(xCash.Text) + Val(xVisa.Text) = Val(Vs_Inv.xTotItem.Text) - Val(Vs_Inv.xDisc.Text) Then
        Vs_Inv.xCash.Text = xCash.Text
        Vs_Inv.xVisa.Text = xVisa.Text
        Me.Hide
    Else
        MsgBox "„—«Ã⁄…  ﬁÌ„… «·”œ«œ ··»Ê‰"
        xCash.SetFocus
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub xCash_GotFocus()
    xCash.Text = Format(Val(Vs_Inv.xTotItem.Text) - Val(Vs_Inv.xDisc.Text), "#0.00")
    xCash.SelStart = 0
    xCash.SelLength = Len(Trim(xCash.Text))
End Sub
Private Sub xCash_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If xCash.Text = Val(Vs_Inv.xTotItem.Text) - Val(Vs_Inv.xDisc.Text) Then
            xVisa.Enabled = False
            cmdSave_Click
        Else
            xVisa.Enabled = True
        End If
    End If
End Sub
Private Sub xVisa_GotFocus()
    xVisa.Text = Format(Val(Vs_Inv.xTotItem.Text) - Val(Vs_Inv.xDisc.Text) - Val(xCash.Text), "#0.00")
    xVisa.SelStart = 0
    xVisa.SelLength = Len(xVisa.Text)
End Sub
