VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmClose 
   Caption         =   "≈€·«ﬁ «·„” ‰œ« "
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   675
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2625
      Width           =   1365
   End
   Begin VB.TextBox xDate1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   675
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   1365
   End
   Begin Threed.SSCommand CMD_CLOSE 
      Height          =   540
      Left            =   300
      TabIndex        =   6
      Top             =   1125
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   953
      _Version        =   196610
      ForeColor       =   4210752
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " —ÕÌ·"
      ButtonStyle     =   2
   End
   Begin Threed.SSCheck xClos 
      Height          =   315
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   225
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   196610
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«·„»Ì⁄« "
      Alignment       =   1
      Value           =   1
   End
   Begin Threed.SSCheck xClos 
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   628
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   196610
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«·„‘ —Ì« "
      Alignment       =   1
      Value           =   1
   End
   Begin Threed.SSCheck xClos 
      Height          =   315
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   1031
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   196610
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "„—œÊœ «·„‘ —Ì« "
      Alignment       =   1
      Value           =   1
   End
   Begin Threed.SSCheck xClos 
      Height          =   315
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   1434
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   196610
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«· ÕÊÌ·« "
      Alignment       =   1
      Value           =   1
   End
   Begin Threed.SSCheck xClos 
      Height          =   315
      Index           =   4
      Left            =   3000
      TabIndex        =   4
      Top             =   1837
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   196610
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«·Ã—œ"
      Alignment       =   1
      Value           =   1
   End
   Begin Threed.SSCheck xClos 
      Height          =   315
      Index           =   5
      Left            =   3000
      TabIndex        =   5
      Top             =   2240
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   196610
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "‰ﬁœÌ…"
      Alignment       =   1
      Value           =   1
   End
   Begin Threed.SSCommand Cmd_Open 
      Height          =   540
      Left            =   300
      TabIndex        =   9
      Top             =   3150
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   953
      _Version        =   196610
      ForeColor       =   65535
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "≈⁄«œ… › Õ «·„” ‰œ« "
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand CMD_EXIT 
      Height          =   540
      Left            =   300
      TabIndex        =   12
      Top             =   4050
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   953
      _Version        =   196610
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      ButtonStyle     =   2
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "› Õ «·„” ‰œ«  ·· ⁄œÌ· »⁄œ  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   165
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2175
      Width           =   2475
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   1800
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   2100
      Width           =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " —ÕÌ· «·„” ‰œ«  Õ Ï  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   330
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   1800
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   2640
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   3900
      Left            =   2775
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   2715
   End
End
Attribute VB_Name = "FrmClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_close_Click()
If Not IsDate(xdate1.Text) Then Exit Sub
' «·„»Ì⁄« 
If xClos(0).Value Then
    cString = " UPDATE FILE6_20 SET FILE6_20.POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
End If

' «·„‘ —Ì« 
If xClos(1).Value Then
    cString = " UPDATE FILE7_20 SET FILE7_20.POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
End If

' „— Ã⁄« 
If xClos(2).Value Then
    cString = " UPDATE FILE6_11 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
End If

'  ÕÊÌ·« 
If xClos(3).Value Then
    cString = " UPDATE FILE1_60 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
End If

' Ã—œ
If xClos(4).Value Then
    cString = " UPDATE FILE0_10 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
End If

' „ﬁ»Ê÷« 
If xClos(5).Value Then
    cString = " UPDATE FILE8_20 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString

    cString = " UPDATE FILE8_40 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
    
    cString = " UPDATE FILE8_10 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
    
    cString = " UPDATE FILE8_30 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString

    cString = " UPDATE FILE8_50 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString

    cString = " UPDATE FILE8_90 SET .POSTED = TRUE  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString

End If


MsgBox "  „ ≈€·«ﬁ «·„” ‰œ«    0000   ·Ì‰ Ì „ «Ï  ⁄œÌ· ›ÌÂ« ≈·« »⁄œ › ÕÂ« „—… «Œ—Ï ", , ""

End Sub

Private Sub CMD_EXIT_Click()
Unload Me
End Sub
Private Sub Form_Load()
For I = 0 To 5
    xClos(I).Value = True
Next I
xdate1.Text = Format(DateAdd("D", Date, -1), "YYYY-MM-DD")
XDATE2.Text = ""
End Sub
Private Sub CMD_OPEN_Click()
If Not IsDate(XDATE2.Text) Then Exit Sub
' «·„»Ì⁄« 
If xClos(0).Value Then
    cString = " UPDATE FILE6_20 SET FILE6_20.POSTED = FALSE  Where Date >= DateValue(" & MyParn(XDATE2.Text) & ")"
    mydb.Execute cString
End If

' «·„‘ —Ì« 
If xClos(1).Value Then
    cString = " UPDATE FILE7_20 SET FILE7_20.POSTED = FALSE  Where Date >= DateValue(" & MyParn(XDATE2.Text) & ")"
    mydb.Execute cString
End If

' „— Ã⁄« 
If xClos(2).Value Then
    cString = " UPDATE FILE6_11 SET .POSTED = FALSE  Where Date >= DateValue(" & MyParn(XDATE2.Text) & ")"
    mydb.Execute cString
End If

'  ÕÊÌ·« 
If xClos(3).Value Then
    cString = " UPDATE FILE1_60 SET .POSTED = FALSE  Where Date >= DateValue(" & MyParn(XDATE2.Text) & ")"
    mydb.Execute cString
End If

' Ã—œ
If xClos(4).Value Then
    cString = " UPDATE FILE0_10 SET .POSTED = FALSE  Where Date >= DateValue(" & MyParn(XDATE2.Text) & ")"
    mydb.Execute cString
End If

' „ﬁ»Ê÷« 
If xClos(5).Value Then
    cString = " UPDATE FILE8_20 SET .POSTED = false  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString

    cString = " UPDATE FILE8_40 SET .POSTED = false  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
    
    cString = " UPDATE FILE8_10 SET .POSTED = false  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString
    
    cString = " UPDATE FILE8_30 SET .POSTED = false  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString

    cString = " UPDATE FILE8_50 SET .POSTED = false  Where Date <= DateValue(" & MyParn(xdate1.Text) & ")"
    mydb.Execute cString


End If

MsgBox "  „ › Õ «·„” ‰œ«    0000   Ì„ﬂ‰ «· ⁄œÌ· Ê «·Õ–›  ", , ""

End Sub

