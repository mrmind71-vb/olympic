VERSION 5.00
Begin VB.Form RepBank 
   Caption         =   " ﬁ«—Ì— «·»‰Êﬂ"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "√Ê—«ﬁ ﬁ»÷ „Õ’·… Œ·«· › —… "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   5
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2240
      Width           =   2445
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "√Ê—«ﬁ œ›⁄ „”œœ… Œ·«· › —…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   6
      Left            =   4485
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   2325
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "»Ì«‰ »√Ê—«ﬁ œ›⁄  ·»‰ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   4
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1840
      Width           =   1725
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "»Ì«‰ »√Ê—«ﬁ ﬁ»÷  Õ  «· Õ’Ì· ·»‰ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   3
      Left            =   4005
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   2805
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ï Õ—ﬂ… «·»‰Êﬂ ( —’Ìœ «Ê· - Ã. ≈Ìœ«⁄«  - Ã. „”ÕÊ»«  - Ã. „’«—Ì› »‰ﬂÌ… - Ã.‘Ìﬂ« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   2
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1040
      Width           =   6600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Œ‹‹‹—ÊÃ"
      Height          =   375
      Left            =   2475
      TabIndex        =   3
      Top             =   4290
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈” Ã«»…"
      Height          =   375
      Left            =   5175
      TabIndex        =   2
      Top             =   4275
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂ‘› Õ”«» »‰ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   1
      Left            =   5205
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   640
      Width           =   1605
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "√—’œ… «·»‰Êﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   0
      Left            =   5565
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "RepBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Byte
Private Sub Command1_Click()
Select Case nOption
Case 0
    Load R_BANK1
    R_BANK1.Show 1
Case 1
    Load R_Bank2
    R_Bank2.Show 1
Case 2
    publicFlag = 1
    Load R_Bank3
    R_Bank3.Show 1
Case 3
    publicFlag = 1
    R_Bank4.Show 1
Case 4
    publicFlag = 2
    R_Bank4.Show 1
Case 5
    publicFlag = 1
    Load R_Bank4
    R_Bank4.Show 1
Case 6
    publicFlag = 2
    Load R_Bank4
    R_Bank4.Show 1
Case 7
    publicFlag = 3
    Load R_Bank4
    R_Bank4.Show 1
End Select
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Opt_Click(Index As Integer)
If Opt(Index).Value = True Then nOption = Index
End Sub

