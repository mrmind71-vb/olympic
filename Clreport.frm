VERSION 5.00
Begin VB.Form ClientReports 
   BackColor       =   &H8000000B&
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   6150
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
   ScaleHeight     =   6150
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4065
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   8490
      Begin VB.CommandButton Command1 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   3525
         Width           =   1365
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "‘Ìﬂ«  „Õ’·… Œ·«· › —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3600
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "‘Ìﬂ«  ⁄„·«¡ „” Õﬁ… Œ·«· › —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2799
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "‘Ìﬂ«  „Õ——… Œ·«· › —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3198
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "’«›Ï „»Ì⁄«  «·«’‰«› ·⁄„Ì· Œ·«· › —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "«Ã„«·Ï „— Ã⁄«  Œ·«· › —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2001
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "„»Ì⁄«  ›Ì“« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1575
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "„ «»⁄… „ﬁ»Ê÷«  „‰ ⁄„Ì· Œ·«· › —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1203
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "„ «»⁄…  ›’Ì·Ï „—œÊœ „»Ì⁄«  ⁄„Ì· Œ·«· › —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   804
         Width           =   4665
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "„ «»⁄…  ›’Ì·Ï ›Ê« Ì— ⁄„Ì· Œ·«· › —… "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   405
         Width           =   4665
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   150
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   5475
         Width           =   1290
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "«” Ã«»…"
         Height          =   390
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   3525
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ClientReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Byte
Private Sub CmdApply_Click()
Select Case nOption
Case 0
    Rep_502.Show 1
Case 1
    Rep_501.Show 1
Case 2
    Rep_503.Show 1
Case 7
    publicFlag = 3
    Load ClientReport3
    ClientReport3.Show 1
Case 8
    publicFlag = 4
    Load ClientReport3
    ClientReport3.Show 1
Case 9
    publicFlag = 5
    Load ClientReport3
    ClientReport3.Show 1
Case 11
    publicFlag = 6
    Load ClientReport3
    ClientReport3.Show 1
Case 13
    publicFlag = 3
    Load ItemsRep1
    ItemsRep1.Show 1
Case 14
    publicFlag = 4
    Load ItemsRep1
    ItemsRep1.Show 1
End Select
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
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
