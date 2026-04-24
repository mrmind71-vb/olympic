VERSION 5.00
Begin VB.Form SettingFrm 
   Caption         =   "ÈíÇäÇÊ ÇáÔÑßÉ"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9675
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   2220
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Height          =   465
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Setting.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1710
      UseMaskColor    =   -1  'True
      Width           =   1365
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
      Height          =   465
      Left            =   1530
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Setting.frx":241E
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "ÍÝÙ"
      Top             =   1710
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   9465
      Begin VB.TextBox xMail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1260
         Width           =   7890
      End
      Begin VB.TextBox xPhone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   7890
      End
      Begin VB.TextBox xAddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   7890
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   7890
      End
      Begin VB.Label Label4 
         Caption         =   "ÈÑíÏ ÇáíßÊÑæäí :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1305
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "ÇáÊáíÝæä :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "ÇáÚäæä :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "ÇáÈíÇä :"
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1005
      End
   End
End
Attribute VB_Name = "SettingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Private Sub cmdSave_Click()
Dim aInsert(3, 1)

aInsert(0, 0) = "Desca"
aInsert(0, 1) = addstring(xDesca.Text)

aInsert(1, 0) = "Address"
aInsert(1, 1) = addstring(xAddress.Text)

aInsert(2, 0) = "Phone"
aInsert(2, 1) = addstring(xPhone.Text)

aInsert(3, 0) = "Mail"
aInsert(3, 1) = addstring(xMail.Text)

On Error GoTo myerror
con.BeginTrans
If Val(GetDesca("Select count(*) from address")) = 0 Then
   con.Execute CreateInsert(aInsert, "Address")
Else
   con.Execute CreateUpdate(aInsert, "Address", "", Array(-1))
End If
con.CommitTrans
Inform "Êã ÇáÊÚÏíá ÈäÌÇÍ"
FixAddress GetCon
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
openCon con
myload
End Sub
Private Sub myload()
Dim aret As Variant
aret = aGetDesca("Select desca,address,Phone,Mail from Address")
If UBound(aret) > 0 Then
    xDesca.Text = aret(1) & ""
    xAddress.Text = aret(2) & ""
    xPhone.Text = aret(3) & ""
    xMail.Text = aret(4) & ""
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
