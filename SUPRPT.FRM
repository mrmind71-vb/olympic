VERSION 5.00
Begin VB.Form SupRpt 
   BackColor       =   &H8000000B&
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
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
   ScaleHeight     =   6885
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6660
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   75
      Width           =   8265
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "»Ì«‰ »√’‰«› ( ⁄œœ -  ﬁÌ„ )"
         Height          =   315
         Index           =   16
         Left            =   4575
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   5805
         Width           =   3390
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„” Õﬁ«  „Ê—œ"
         Height          =   315
         Index           =   15
         Left            =   4575
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   5490
         Width           =   3390
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ï ›Ê« Ì— „—œÊœ „‘ —Ì«  Œ·«· › —…"
         Height          =   315
         Index           =   14
         Left            =   4575
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   5167
         Width           =   3390
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   225
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6075
         Width           =   1290
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "«” Ã«»…"
         Height          =   390
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   6075
         Width           =   1365
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "√—’œ… „Ê—œÌ‰"
         Height          =   315
         Index           =   0
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1290
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„ «»⁄… «Ã„«·Ï ›Ê« Ì— „Ê—œ Œ·«· › —…"
         Height          =   315
         Index           =   3
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1284
         Width           =   2790
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ï Õ—ﬂ… «·„Ê—œÌ‰ Œ·«· › —…"
         Height          =   315
         Index           =   1
         Left            =   5025
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   578
         Width           =   2940
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„ «»⁄…  ›’Ì·Ï ›Ê« Ì— „‘ —Ì«  ·„Ê—œ Œ·«· › —… "
         Height          =   315
         Index           =   7
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2696
         Width           =   4140
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„ «»⁄… «Ã„«·Ï „—œÊœ „‘ —Ì«  ·„Ê—œ Œ·«· › —… "
         Height          =   315
         Index           =   5
         Left            =   4425
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1980
         Width           =   3540
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„ «»⁄…  ›’Ì·Ï ›Ê« Ì— „—œÊœ „‘ —Ì«  ·„Ê—œ Œ·«· › —…"
         Height          =   315
         Index           =   8
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3049
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "’«›Ï „‘ —Ì«  «·√’‰«› „‰ „Ê—œ Œ·«· › —…"
         Height          =   315
         Index           =   9
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3402
         Width           =   3840
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ—ﬂ… „Ê—œ Œ·«· › —… "
         Height          =   315
         Index           =   12
         Left            =   6150
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   4461
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„ «»⁄…  „œ›Ê⁄«  ≈·Ï „Ê—œ Œ·«· › —…"
         Height          =   315
         Index           =   11
         Left            =   5100
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   4108
         Width           =   2865
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„ «»⁄… «Ã„«·Ï ›Ê« Ì— „Ã„Ê⁄…  Œ·«· › —…"
         Height          =   315
         Index           =   4
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1637
         Width           =   3165
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ï Õ—ﬂ… «·„Ã„Ê⁄«  Œ·«· › —… "
         Height          =   315
         Index           =   2
         Left            =   4875
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   931
         Width           =   3090
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "„ «»⁄… «Ã„«·Ï „—œÊœ „‘ —Ì«  ·„Ã„Ê⁄…  „Ê—œÌ‰ Œ·«· › —… "
         Height          =   315
         Index           =   6
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   2343
         Width           =   5040
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "’«›Ï „‘ —Ì«  «·√’‰«›  ·„Ã„Ê⁄«  «·√’‰«›"
         Height          =   315
         Index           =   10
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   3755
         Width           =   3915
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ï ›Ê« Ì—  „‘ —Ì«   Œ·«· › —…"
         Height          =   315
         Index           =   13
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   4814
         Width           =   3465
      End
   End
End
Attribute VB_Name = "SupRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Byte
Private Sub CmdApply_Click()
publicFlag = nOption
Select Case nOption
Case 0
    SupRpt1.Show 1
Case 1
    Load SupRpt1
    SupRpt1.Show 1
Case 2
    Load SupRpt1
    SupRpt1.Show 1
Case 3, 5
    Load SupRpt2
    SupRpt2.Show 1
Case 4
    Load SupRpt1
    SupRpt1.Show 1
Case 6
    Load SupRpt1
    SupRpt1.Show 1
Case 7, 8
    Load SupRpt2
    SupRpt2.Show 1
Case 9
    Load SupRpt2
    SupRpt2.Show 1
Case 10
    Load SupRpt1
    SupRpt1.Show 1
Case 11
    Load SupRpt2
    SupRpt2.Show 1
Case 12
    Load SupRpt2
    SupRpt2.Show 1
Case 13
    Load SupRpt2
    SupRpt2.Show 1
Case 14
    Load SupRpt2
    SupRpt2.Show 1
Case 15
'    Load Supp_Inv
'  Supp_Inv.Show 1
    SupRpt15.Show 1

Case 16
    SupRpt16.Show 1
End Select
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub
Private Sub Opt_Click(Index As Integer)
If Opt(Index).Value = True Then nOption = Index
End Sub
Private Sub Opt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
CmdApply_Click
End If
End Sub
