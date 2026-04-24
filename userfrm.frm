VERSION 5.00
Begin VB.Form userfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " €ÌÌ— ﬂ·„… «·”—"
   ClientHeight    =   2520
   ClientLeft      =   4050
   ClientTop       =   3825
   ClientWidth     =   5955
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   5955
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   225
      TabIndex        =   6
      Top             =   1620
      Width           =   4515
      Begin VB.CommandButton cmdExit 
         Height          =   600
         Left            =   45
         Picture         =   "userfrm.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2250
         MaskColor       =   &H00FFFFFF&
         Picture         =   "userfrm.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Save"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   2220
      End
   End
   Begin VB.TextBox xPass2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   2130
   End
   Begin VB.CheckBox xShow 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "≈ŸÂ«—"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   495
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   765
      Width           =   870
   End
   Begin VB.TextBox xPass 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2130
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   5955
      TabIndex        =   3
      Top             =   0
      Width           =   5955
      Begin VB.Label xDesca 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   90
         Width           =   5835
      End
   End
   Begin VB.Label Label2 
      Caption         =   " √ﬂÌœ ﬂ·„… «·”—"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3870
      TabIndex        =   9
      Top             =   1170
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "ﬂ·„… «·”—"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3870
      TabIndex        =   5
      Top             =   810
      Width           =   1590
   End
End
Attribute VB_Name = "userfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public nTop As Long, nLeft As Long
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Trim(xPass.text) = "" Then
    MsgBox "ﬂ·„… ”— €Ì— „”Ã·"
    Exit Sub
End If

If Len(Trim(xPass.text)) < 3 And Trim(xPass.text) <> "" Then
    MsgBox "«ﬁ· ⁄œœ ··Õ—Ê› 3"
    Exit Sub
End If

If xShow.Value = 0 Then
    If Trim(xPass.text) <> Trim(xPass2.text) Then
        MsgBox "ﬂ·„… «·”—  Õ ·› ⁄‰ ﬂ·„… «· √ﬂÌœ"
        Exit Sub
    End If
End If

con.BeginTrans
On Error GoTo myError
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "password", addstring(xPass.text))
con.Execute addUpdate(aInsert, "USERS", "code = " & nUsercode)
con.CommitTrans
Inform " „  €ÌÌ— ﬂ·„… «·”— »‰Ã«Õ"
End
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Form_Up(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    aUser = Empty
    Unload Me
End If
End Sub
Private Sub Form_Load()
If nTop > 0 Then Me.Top = Me.Top + nTop
If nLeft > 0 Then Me.Left = Me.Left + nLeft
xDesca.Caption = cUserName
openCon con
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set userfrm = Nothing
End Sub
Private Sub xPass_Change()
cmdSave.Enabled = Trim(xPass.text) <> ""
End Sub

Private Sub xPass_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
ElseIf KeyCode = 13 And xPass.text <> "" Then
    KeyCode = 0
    cmdSave_Click
End If
End Sub
Private Sub xShow_Click()
xPass.PasswordChar = IIf(xShow.Value = 1, "", "*")
xPass2.text = ""
xPass2.Enabled = xShow.Value = 0
End Sub

Private Sub xPass_GotFocus()
myGotFocus xPass
End Sub
Private Sub xPass_LostFocus()
myLostFocus xPass
End Sub
Private Sub xrecords_GotFocus()
myGotFocus xRecords
End Sub
Private Sub xrecords_LostFocus()
myLostFocus xRecords
End Sub
