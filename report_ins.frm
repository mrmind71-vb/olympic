VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form report_insfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ«—Ì— «·⁄÷ÊÌ… «·„ﬁ”ÿ…"
   ClientHeight    =   6720
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   12840
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "report_ins.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   12705
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   1
         Left            =   8955
         TabIndex        =   0
         Top             =   135
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«⁄œ«œ «·«⁄÷«¡"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   2
         Left            =   8955
         TabIndex        =   3
         Top             =   585
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»Ì«‰«  «·«⁄÷«¡"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   3
         Left            =   8955
         TabIndex        =   5
         Top             =   1035
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»Ì«‰«  «·”œ«œ Õ”»  «—ÌŒ »œ«Ì… «·⁄÷ÊÌ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   4
         Left            =   8955
         TabIndex        =   6
         Top             =   1485
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»Ì«‰«  «·”œ«œ Õ”»  «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   5
         Left            =   8955
         TabIndex        =   4
         Top             =   1935
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "√ﬁ”«ÿ „ √Œ—… ⁄‰  «—ÌŒ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   6
         Left            =   8955
         TabIndex        =   7
         Top             =   2385
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«ﬁ”«ÿ „”œœ… Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   7
         Left            =   8955
         TabIndex        =   8
         Top             =   2835
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "√⁄÷«¡ «ﬁ”«ÿÂ«  Œ ·› ⁄‰ ﬁÌ„… «·⁄ﬁœ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   9
         Left            =   8955
         TabIndex        =   9
         Top             =   3735
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ì «Ì’«·«  «ﬁ”«ÿ Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   8
         Left            =   8955
         TabIndex        =   10
         Top             =   3285
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»Ì«‰«  «·⁄÷ÊÌ… Õ”» ‰Ê⁄ «·”œ«œ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   10
         Left            =   8955
         TabIndex        =   11
         Top             =   4185
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "„ «»⁄… „ﬁœ„ «ﬁ”«ÿ «·«⁄÷«¡"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   11
         Left            =   8955
         TabIndex        =   12
         Top             =   4635
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«»‰«¡ Ê’·Ê« «·Ì ”‰ „⁄Ì‰"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "report_insfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim nOption As Integer
Private Sub cmdApply_Click()
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1
     report_insfrm1.Show 1
Case 2
     report_insfrm2.Show 1
Case 3
     report_insfrm3.Show 1
Case 4
     report_insfrm4.Show 1
Case 5
     report_insfrm5.Show 1
Case 6
     report_insfrm6.Show 1
Case 7
     report_insfrm7.Show 1
Case 8
     report_insfrm8.Show 1
Case 9
     grdpaid_Installfrm1.Show
Case 10
     report_insfrm10.Show 1
Case 11
     report_insfrm11.Show 1
End Select
End Sub
Private Sub cmdgo_MouseEnter(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdGo(Index).ForeColor = &HC00000
End Sub
Private Sub cmdgo_MouseExit(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdGo(Index).ForeColor = &H80000008
End Sub
Private Sub CmdOk_Click()
Unload Me
End Sub

