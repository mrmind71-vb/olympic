VERSION 5.00
Begin VB.Form listitemsfrm2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   6450
      LargeChange     =   2
      Left            =   11745
      Max             =   2
      Min             =   1
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   180
      Value           =   1
      Width           =   375
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   29
      Left            =   270
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5580
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   28
      Left            =   2565
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5580
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   27
      Left            =   4860
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5580
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   26
      Left            =   7155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5580
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   25
      Left            =   9450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5580
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   24
      Left            =   270
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4500
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   23
      Left            =   2565
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4500
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   22
      Left            =   4860
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4500
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   21
      Left            =   7155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4500
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   20
      Left            =   9450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4500
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   19
      Left            =   270
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3420
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   18
      Left            =   2565
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3420
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   17
      Left            =   4860
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3420
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   16
      Left            =   7155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3420
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   15
      Left            =   9450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3420
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   14
      Left            =   270
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2340
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   13
      Left            =   2565
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2340
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   12
      Left            =   4860
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2340
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   11
      Left            =   7155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2340
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   10
      Left            =   9450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2340
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   9
      Left            =   270
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1260
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   8
      Left            =   2565
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1260
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   7
      Left            =   4860
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1260
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   6
      Left            =   7155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1260
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   5
      Left            =   9450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1260
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   4
      Left            =   270
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   3
      Left            =   2565
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   2
      Left            =   4860
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   1
      Left            =   7155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   2265
   End
   Begin VB.CommandButton cmdItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cmdItem"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   0
      Left            =   9450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   6630
      Left            =   180
      Top             =   90
      Width           =   12030
   End
End
Attribute VB_Name = "listitemsfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Dim loctable As New adodb.Recordset
Public nGroup As String, cGroupname As String
Private Sub cmdItem_Click(Index As Integer)
salesfrm.myproc3 cmdItem(Index).Tag, cmdItem(Index).Caption
'Me.Hide
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
openCon con
Me.Caption = "ÞÇÆãÉ ÇáÇÕäÇÝ" & turn(cGroupname, " - ") & cGroupname
Dim cString As String
cString = "select * from file1_10"
cString = cString & turn(cString) & "file1_10.[GROUP] = " & MyParn(nGroup)
cString = cString & turn(cString) & "(file1_10.SHOW = 1)"

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
loctable.PageSize = 30
VScroll1.Max = loctable.PageCount
VScroll1.Visible = VScroll1.Max > 1
myload 0
End Sub
Private Sub myload(nPage)
Dim nCol As Integer, nRow As Integer
For i = 0 To 29
    cmdItem(i).Visible = False
Next
If nPage > 0 And nPage <= loctable.PageCount Then
    loctable.AbsolutePage = nPage
End If
i = 0
Do Until (loctable.EOF Or i > 29)
    cmdItem(i).Visible = True
    cmdItem(i).Tag = loctable!Item & ""
    cmdItem(i).Caption = loctable!desca & "(" & loctable!price & ")"
    loctable.MoveNext
    i = i + 1
Loop
End Sub
Private Sub VScroll1_Change()
myload VScroll1.Value
End Sub
