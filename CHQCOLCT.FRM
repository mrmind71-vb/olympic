VERSION 5.00
Begin VB.Form ChqCollect 
   BackColor       =   &H00E0E0E0&
   Caption         =   " Õ’Ì· «·‘Ìﬂ«  "
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.5
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xBox 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   525
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00C0FFFF&
      Caption         =   " —«Ã⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   975
      Width           =   1140
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1275
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   975
      Width           =   1590
   End
   Begin VB.TextBox xMemo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   525
      Width           =   2565
   End
   Begin VB.TextBox xDate_3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   150
      Width           =   2565
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Œ“‰…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ«   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4875
      TabIndex        =   1
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4875
      TabIndex        =   0
      Top             =   225
      Width           =   465
   End
End
Attribute VB_Name = "ChqCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
If MsgBox(CmdOk.Caption & " Â· «‰  „Ê«›ﬁ", vbOKCancel, "‰Ÿ«„ «·„Œ«“‰") = vbCancel Then Exit Sub
If Not MYVALID Then Exit Sub
If lColect Then
    chq.xDate_3.Text = xDate_3.Text
    chq.xMemo.Text = xMemo.Text
    chq.xBox.Text = xBox.Text
    chq.Collect_Chq
Else
    chq.xDate_3.Text = xDate_3.Text
    chq.xMemo.Text = xMemo.Text
    chq.xBox.Text = xBox.Text
    chq.Collect_Chq
End If
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
If lColect Then
    CmdOk.Caption = " Õ’Ì· ‘Ìﬂ"
Else
    CmdOk.Caption = "—›÷ ‘Ìﬂ"
End If
xDate_3.Text = ""
xMemo.Text = ""
End Sub
Function MYVALID() As Boolean
If Not IsDate(xDate_3.Text) Then
    MsgBox "ÌÃ»  ”ÃÌ· «· «—ÌŒ »ÿ—Ìﬁ… ”·Ì„…"
    Exit Function
End If
MYVALID = True
End Function
