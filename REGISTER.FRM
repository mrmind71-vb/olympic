VERSION 5.00
Begin VB.Form RegisterFrm 
   Caption         =   "√›Êﬂ« Ê"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1635
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
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
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1125
      Width           =   1140
   End
   Begin VB.CommandButton CmdRegister 
      Caption         =   "«· —ŒÌ’"
      Enabled         =   0   'False
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
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1125
      Width           =   1215
   End
   Begin VB.TextBox xReg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   675
      Width           =   3240
   End
   Begin VB.Label Label2 
      Caption         =   "—ﬁ„ «· —ŒÌ’ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3525
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   750
      Width           =   1065
   End
   Begin VB.Label xKey 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "—ﬁ„ «·ﬂÊœ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3525
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   825
   End
End
Attribute VB_Name = "RegisterFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myRegister As Boolean
Dim CaseTable As Recordset, ClientTable As Recordset
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdRegister_Click()
On Error GoTo MyError
If Not CheckSerial(Val(xReg.Text)) Then
    MsgBox "«·—ﬁ„ «·„”Ã· €Ì— ’ÕÌÕ", , systemName
    Unload Me
Else
    'If Dir(App.Path & "\ss.txt") <> "" Then Kill App.Path & "\ss.txt"
     Open App.Path & "\Serial.txt" For Output As #1
     Write #1, CodeSerial(Serial("C:"), NCODE2)
     ' Close before reopening in another mode.
     Close #1
     MsgBox " „ «· —ŒÌ’ »‰Ã«Õ", , systemName
     myRegister = True
     Unload Me
End If
Exit Sub
MyError:
    Unload Me
End Sub
 Private Sub Form_Load()
'Label3.Caption = Serial("C:")
xKey.Caption = CodeSerial(Serial("C"), NCODE1)
Clipboard.SetText xKey
End Sub
Private Function CheckSerial(pSerial) As Boolean
SERIALNUMBER = UnCodeSerial(pSerial, NCODE2)
CheckSerial = (SERIALNUMBER = Serial("C"))
End Function
Private Sub Form_Unload(Cancel As Integer)
If Not myRegister Then End
End Sub
Private Sub xReg_Change()
On Error GoTo MyError:
CmdRegister.Enabled = IsNumeric(xReg.Text)
CmdRegister.Enabled = CheckSerial(xReg.Text)
Exit Sub
MyError:
CmdRegister.Enabled = False
End Sub
