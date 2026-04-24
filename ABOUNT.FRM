VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "⁄‰ «·„Œ«“‰ "
   ClientHeight    =   3750
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2588.317
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "Abount.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   3375
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "  : 5825347 - 680 ÿ—Ìﬁ «·Õ—Ì… ·Ê—«‰"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1650
      TabIndex        =   9
      Top             =   2250
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   2895
      TabIndex        =   8
      Top             =   2550
      Width           =   60
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   " Õ–Ì—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1425
      Width           =   540
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "⁄„—Ê⁄»œ«·Õ„Ìœ„Õ„œ - ⁄»œ«·Õ„Ìœ „Õ„œ «·”„«‰"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Top             =   3075
      Width           =   3000
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "«· ’„Ì„ Ê «·»—„Ã… "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   1
      Left            =   2250
      TabIndex        =   5
      Top             =   2625
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ﬂ· «·ÕﬁÊﬁ „Õ›ÊŸ… ··„ƒ””… «·œÊ·Ì… ·· ﬂ‰Ê·ÊÃÌ« - «·‰⁄Ì‰⁄ ≈ŒÊ«‰ "
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   795
      TabIndex        =   4
      Top             =   1950
      Width           =   4350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   70.429
      X2              =   5295.313
      Y1              =   828.261
      Y2              =   828.261
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "«·„—‘œ ·Õ”«»«  «·„Œ«“‰ - «·⁄„·«¡ - «·„Ê—œÌ‰ - «·»‰Êﬂ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   225
      Width           =   4125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   70.429
      X2              =   5281.227
      Y1              =   880.028
      Y2              =   880.028
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "«·«’œ«— —ﬁ„    3.5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   1500
      TabIndex        =   3
      Top             =   750
      Width           =   2910
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
  Unload Me
End Sub

