VERSION 5.00
Begin VB.Form RepCharge 
   Caption         =   " ﬁ«—Ì— «·„’«—Ì›"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ï „”ÕÊ»«  «·‘—ﬂ«¡ ·› —…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   2235
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1485
      Width           =   3525
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   " ›’Ì·Ï „”ÕÊ»«  «·‘—ﬂ«¡ ·› —…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   3525
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "„ «»⁄… ÌÊ„Ì… ··„’«—Ì›"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   2235
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1050
      Width           =   3525
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2310
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   " ›’Ì· „’—Ê› Œ·«· › —…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2235
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   615
      Width           =   3525
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Œ—ÊÃ"
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
      Left            =   255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2310
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ï «·„’«—Ì› ﬂ„Ã„Ê⁄« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2235
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   3525
   End
End
Attribute VB_Name = "RepCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Integer
Private Sub CmdApply_Click()
0 publicFlag = nOption
Select Case nOption
Case 1
    TProft.Show 1
Case 2
    Load RCharg1
    RCharg1.Show 1
Case 3
    Load RCharg1
    RCharg1.xCharge.Visible = True
    RCharg1.LblCharge.Visible = True
    RCharg1.Show 1
Case 4
    Load RepIncome1
    RepIncome1.Show 1
Case 5
    Load RCharg1
    RCharg1.Show 1
Case 6
    Load RCharg1
    RCharg1.xCharge.Visible = True
    RCharg1.LblCharge.Visible = True
    RCharg1.Show 1
Case 7
    R_Chr0.Show 1
Case 8
    RCharg8.Show 1

Case 11
    R_Chr11.Show 1
Case 12
    R_Chr12.Show 1

End Select
End Sub
Private Sub CmdOk_Click()
Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(Index).Value = True Then nOption = Index
End Sub
