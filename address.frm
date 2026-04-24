VERSION 5.00
Begin VB.Form addressfrm 
   Caption         =   "»Ì«‰«  «·‘—ﬂ…"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10065
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
   ScaleHeight     =   2250
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Height          =   465
      Left            =   180
      MaskColor       =   &H00FFFFFF&
      Picture         =   "address.frx":0000
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
      Left            =   1575
      MaskColor       =   &H00FFFFFF&
      Picture         =   "address.frx":241E
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Õ›Ÿ"
      Top             =   1710
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   9780
      Begin VB.TextBox xMail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1260
         Width           =   7845
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
         Width           =   7845
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
         Width           =   7845
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
         Width           =   7845
      End
      Begin VB.Label Label4 
         Caption         =   "»—Ìœ «·Ìﬂ —Ê‰Ì :"
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
         TabIndex        =   9
         Top             =   1305
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "«· ·Ì›Ê‰ :"
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
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "«·⁄‰Ê‰ :"
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
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "«·»Ì«‰ :"
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
         TabIndex        =   6
         Top             =   225
         Width           =   1005
      End
   End
End
Attribute VB_Name = "addressfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdSave_Click()
Dim aInsert As Variant
aInsert = AddFlag(Empty, "desca", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "Address", addstring(xAddress.Text))
aInsert = AddFlag(aInsert, "Phone", addstring(xPhone.Text))
aInsert = AddFlag(aInsert, "Mail", addstring(xMail.Text))
aInsert = AddFlag(aInsert, "RATE1", Val(xrate1.Text))
aInsert = AddFlag(aInsert, "RATE2", Val(xrate2.Text))

On Error GoTo myerror
con.BeginTrans
If IsEmpty(GetField("SELECT ID FROM ADDRESS")) Then
   con.Execute addInsert(aInsert, "Address")
Else
   con.Execute addUpdate(aInsert, "Address", "")
End If
con.CommitTrans
Inform " „ «· ⁄œÌ· »‰Ã«Õ"
FixAddress GetCon
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
myload
End Sub
Private Sub myload()
Dim aRet As Variant
aRet = GetFields("Select * from Address ORDER BY ID DESC")
If Not IsEmpty(aRet) Then
    xDesca.Text = retFlag(aRet, "DESCA") & ""
    xAddress.Text = retFlag(aRet, "ADDRESS") & ""
    xPhone.Text = retFlag(aRet, "PHONE") & ""
    xMail.Text = retFlag(aRet, "MAIL") & ""
    xrate1.Text = Myvalue(retFlag(aRet, "RATE1"))
    xrate2.Text = Myvalue(retFlag(aRet, "RATE2"))
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xrate2_GotFocus()
myGotFocus xrate2
End Sub
Private Sub xrate2_LostFocus()
myLostFocus xrate2
End Sub
Private Sub xrate1_GotFocus()
myGotFocus xrate1
End Sub
Private Sub xrate1_LostFocus()
myLostFocus xrate1
End Sub
Private Sub xMail_GotFocus()
myGotFocus xMail
End Sub
Private Sub xMail_LostFocus()
myLostFocus xMail
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub
Private Sub xPhone_LostFocus()
myLostFocus xPhone
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xDesca_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
