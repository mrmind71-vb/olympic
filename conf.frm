VERSION 5.00
Begin VB.Form confFrm 
   Caption         =   "Connection Configuration"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1530
      Width           =   2850
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1485
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1545
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   7035
      Begin VB.TextBox xServerName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1845
         TabIndex        =   0
         Top             =   225
         Width           =   4200
      End
      Begin VB.TextBox xUseridNew 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1845
         TabIndex        =   1
         Top             =   630
         Width           =   4200
      End
      Begin VB.TextBox xpasswordNew 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1845
         TabIndex        =   2
         Top             =   1035
         Width           =   4200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Server Name"
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
         Left            =   180
         TabIndex        =   9
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name"
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
         Left            =   135
         TabIndex        =   8
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Password"
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
         Left            =   135
         TabIndex        =   7
         Top             =   1125
         Width           =   1545
      End
      Begin VB.Label xuserid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label xpassword 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Visible         =   0   'False
         Width           =   1950
      End
   End
End
Attribute VB_Name = "confFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conFileName As String
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub

addSetting "Server", xServerName.Text, conFileName
If Trim(xUseridNew.Text) <> "" Then addSetting "userid", crypt(xUseridNew.Text, "dr"), conFileName
If Trim(xpasswordNew.Text) <> "" Then addSetting "password", crypt(xpasswordNew.Text, "dr"), conFileName

Inform " „ «·Õ›Ÿ »‰Ã«Õ"
Unload Me
End Sub
Private Sub CmdExit_Click()
Unload Me
End
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
conFileName = App.Path & "\conf.txt"
myload
End Sub
Private Sub myload()
xServerName.Text = RetSetting("server", conFileName)
xuserid.Caption = RetSetting("userid", conFileName)
xpassword.Caption = RetSetting("password", conFileName)
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Function MYVALID() As Boolean
Dim conMaster As New ADODB.Connection
On Error GoTo myerror
Dim aData As Variant
aData = AddFlag(Empty, "server", xServerName.Text)
aData = AddFlag(aData, "userid", xUseridNew.Text)
aData = AddFlag(aData, "password", xpasswordNew.Text)
cString = LoadConString(aData, "master")
conMaster.Open cString
MYVALID = True
Exit Function
myerror:
    MsgBox Err.Description
    Err.Clear
End Function
Private Sub xServerName_GotFocus()
myGotFocus xServerName
End Sub
Private Sub xServerName_LostFocus()
myLostFocus xServerName
End Sub
Private Sub xUseridNew_GotFocus()
myGotFocus xUseridNew
End Sub
Private Sub xUseridNew_LostFocus()
myLostFocus xUseridNew
End Sub
Private Sub xpasswordNew_GotFocus()
myGotFocus xpasswordNew
End Sub
Private Sub xpasswordNew_LostFocus()
myLostFocus xpasswordNew
End Sub
