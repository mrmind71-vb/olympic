VERSION 5.00
Begin VB.Form Cust 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·⁄„·«¡"
   ClientHeight    =   3975
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   9585
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
   ScaleHeight     =   3975
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1425
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   750
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Data Data1 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   750
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      MaxLength       =   200
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   7335
   End
   Begin VB.TextBox xADDress 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   300
      MaxLength       =   250
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2715
      Width           =   7305
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9585
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3435
      Width           =   9585
      Begin VB.CommandButton CmdPrevious 
         BackColor       =   &H00C0FFFF&
         Caption         =   "”«»ﬁ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8175
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   75
         Width           =   1290
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«Õﬁ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6935
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   75
         Width           =   1215
      End
      Begin VB.CommandButton CmdFirst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√Ê·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5770
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   75
         Width           =   1140
      End
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   75
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   9585
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ–›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4788
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdUndo 
         BackColor       =   &H00C0FFFF&
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5736
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8580
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«÷«›…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7632
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6675
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   915
      End
   End
   Begin VB.TextBox xClass 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   270
      MaxLength       =   200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2170
      Width           =   7335
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6090
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   1515
   End
   Begin VB.TextBox xScool 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   270
      MaxLength       =   200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1725
      Width           =   7335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·›’· «·œ—«”Ï"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2250
      Width           =   1380
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·⁄‰Ê«‰ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2775
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·«”„ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " ·Ì›Ê‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   825
      Width           =   555
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·„œ—”…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   4
      X1              =   150
      X2              =   9510
      Y1              =   1575
      Y2              =   1575
   End
End
Attribute VB_Name = "Cust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cardtable As Recordset
Dim formMode As Byte
Dim movetable As Recordset
Sub myDefine()
xDescA.Text = ""
xADDress.Text = ""
xScool.Text = ""
xClass.Text = ""
End Sub
Sub myProc()
    cardtable.FindFirst "Code = " & MyParn(GrdText(Search.Grid1, 0))
    If Not cardtable.NoMatch Then MyLoad
End Sub
Sub MyLoad()
If cardtable.RecordCount = 0 Then Exit Sub
xCode.Text = TurnValue(cardtable.CODE, Null, "")
xDescA.Text = TurnValue(cardtable.DESCA, Null, "")
xADDress.Text = TurnValue(cardtable.ADDRESS, Null, "")
xScool.Text = TurnValue(cardtable!SCOOL, Null, "")
xClass.Text = TurnValue(cardtable!Class, Null, "")
End Sub
Sub MyReplace()
cardtable.FindFirst "Code = " & MyParn(xCode.Text)
If cardtable.NoMatch Then
    cardtable.AddNew
Else
    cardtable.Edit
End If
cardtable.SCOOL = TurnValue(xScool.Text, "", Null)
cardtable!Class = TurnValue(xClass.Text, "", Null)
cardtable.DESCA = TurnValue(xDescA.Text, "", Null)
cardtable.ADDRESS = TurnValue(xADDress.Text, "", Null)
cardtable.CODE = TurnValue(xCode.Text, "", Null)
cardtable.Update
End Sub
Function MYVALID()
MYVALID = True
If xCode.Text = "" Then
    MsgBox "«· ·Ì›Ê‰ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    MYVALID = False
    Exit Function
End If
If xDescA.Text = "" Then
    MsgBox "«·≈”„  ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    MYVALID = False
    Exit Function
End If

'If cardtable.RecordCount > 0 And formMode = addmode Then
'     cardtable.FindFirst "Code = " & MyParn(xCode.Text)
'     If Not cardtable.NoMatch Then
'        MYVALID = False
'        Exit Function
'    End If
'End If
MYVALID = True
End Function
Sub UnEmptyTable()
   CmdFirst.Enabled = True
   CmdLast.Enabled = True
   CmdNext.Enabled = True
   CmdPrevious.Enabled = True
   CmdDel.Enabled = True
   Cmdgo.Enabled = True
   xDescAfilter.Enabled = True
End Sub
Private Sub CmdAdd_Click()
    AddProc
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    cardtable.Delete
    cardtable.Requery
    If cardtable.RecordCount > 0 Then
        cardtable.MoveLast
        MyLoad
    Else
       EmptyProc
    End If
End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
    cardtable.MoveFirst
    MyLoad
End Sub
Private Sub CmdInform_Click()
    Dim Generalarray(3)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «· ·Ì›Ê‰,DescA As «·«”„  From File3_20"
    Generalarray(3) = "Where DescA Like '%cFilter%'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Private Sub CmdLast_Click()
cardtable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
cardtable.MoveNext
If cardtable.EOF Then
    cardtable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
cardtable.MovePrevious
If cardtable.BOF Then
    cardtable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub cmdSave_Click()
msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub
'If Not MsgBox(msgBoxStr, 4) = 6 Then
'    CmdUndo_Click
'    Exit Sub
'End If
MyReplace
cCust = xCode.Text
Unload Me
End Sub
Private Sub CmdUndo_Click()
Select Case formMode
Case EmptyMode
    myDefine
Case addmode
    cardtable.MoveLast
    editProc
    MyLoad
Case Editmode
    MyLoad
End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If cardtable.RecordCount = 0 Then Exit Sub
If formMode <> Editmode Then Exit Sub
If KeyCode = 34 Then CmdPrevious_Click
If KeyCode = 33 Then CmdNext_Click
End Sub
Private Sub Form_Load()
    Set cardtable = mydb.OpenRecordset("select * from file3_20 order by code", dbOpenDynaset)
    If cCust = "" Then
        CmdAdd_Click
    Else
        cardtable.FindFirst " CODE = " & MyParn(cCust)
        If cardtable.NoMatch Then
            myDefine
            CmdAdd_Click
        Else
            MyLoad
        End If
    End If
End Sub
Sub Handlecontrols()
Select Case formMode
Case Editmode
'     CmdAdd.Enabled = True
     CmdDel.Enabled = True
     CmdInform.Enabled = True
     CmdExit.Enabled = True
     CmdSave.Enabled = True
     CmdUndo.Enabled = True
     CmdPrevious.Enabled = True
     CmdNext.Enabled = True
     CmdLast.Enabled = True
     CmdFirst.Enabled = True

Case addmode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
'    CmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False

Case EmptyMode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
'    CmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False

End Select
End Sub
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols
myDefine
xCode.Text = ""
End Sub
Sub editProc()
formMode = Editmode
Handlecontrols
End Sub
Sub AddProc()
formMode = addmode
Handlecontrols
myDefine
xCode.Text = ""
End Sub
Private Sub xCode_LostFocus()
cardtable.FindFirst " CODE = " & MyParn(xCode.Text)
If cardtable.NoMatch Then
    myDefine
Else
    MyLoad
End If
End Sub
Private Sub xDescA_LostFocus()
cardtable.FindFirst " DESCA = " & MyParn(xDescA.Text)
If Not cardtable.NoMatch Then
    If cardtable!CODE <> xCode.Text Then
        MsgBox "«·≈”„ „”Ã· „‰ ﬁ»· "
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
If KeyAscii = 19 Then cmdSave_Click
End Sub

