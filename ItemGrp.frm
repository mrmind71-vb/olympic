VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form ItemsGrp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "„Ã„Ê⁄«  «·«’‰«›"
   ClientHeight    =   2640
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   5925
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin MSDBCtls.DBCombo xM_Group 
      Bindings        =   "ItemGrp.frx":0000
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   1575
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5925
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2085
      Width           =   5925
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         Height          =   390
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton CmdFirst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√Ê·"
         Height          =   390
         Left            =   3090
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«Õﬁ"
         Height          =   390
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton CmdPrevious 
         BackColor       =   &H00C0FFFF&
         Caption         =   "”«»ﬁ"
         Height          =   390
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.Label xRecordNumber 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   225
         TabIndex        =   19
         Top             =   75
         Width           =   1890
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3750
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   690
   End
   Begin VB.TextBox xF_date 
      Enabled         =   0   'False
      Height          =   285
      Left            =   -240
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2835
      Width           =   285
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      Height          =   315
      Left            =   135
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   4305
   End
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   5925
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   5925
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«÷«›…"
         Height          =   390
         Left            =   3885
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«” ⁄·«„"
         Height          =   390
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton CmdUndo 
         BackColor       =   &H00C0FFFF&
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   2970
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ–›"
         Height          =   390
         Left            =   1140
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ›Ÿ"
         Height          =   390
         Left            =   2055
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   225
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1650
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰ «·„Ã„Ê⁄… :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·„Ã„Ê⁄… :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   705
      Width           =   1155
   End
End
Attribute VB_Name = "ItemsGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim cardtable As Recordset
Dim nRecordNumber As Integer
Sub AddProc()
If Not formMode = addmode Then Handlecontrols addmode
formMode = addmode
myDefine
xCode.Text = IncRec(myLastField(cardtable, "Code"))
End Sub
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols EmptyMode
myDefine
xCode.Text = IncRec(myLastField(cardtable, "Code"))
End Sub
Sub Handlecontrols(nMode)
Select Case nMode
Case Editmode
     cmdAdd.Enabled = True
     CmdDel.Enabled = True
     CmdInform.Enabled = True
     CmdExit.Enabled = True
     CmdSave.Enabled = True
     CmdUndo.Enabled = True
     CmdPrevious.Enabled = True
     CmdNext.Enabled = True
     CmdLast.Enabled = True
     CmdFirst.Enabled = True
     xCode.Enabled = False
Case addmode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    cmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xCode.Enabled = True
Case EmptyMode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    cmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xCode.Enabled = True
End Select
End Sub
Sub CardLookup()
Dim Generalarray(4)
Dim GrdArray(2)
Set Generalarray(1) = Me
Generalarray(2) = "Select Code as «·ﬂÊœ,DescA as [»Ì«‰ «·„Ã„Ê⁄…]From FILE1_50"
Generalarray(3) = " Where DescA Like '*cFilter*'"
Generalarray(4) = " ORDER BY VAL(CODE) "
GrdArray(1) = 1200
GrdArray(2) = 4000
    
Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Sub editProc()
formMode = Editmode
Handlecontrols Editmode
End Sub
Sub myDefine()
xDescA.Text = ""
xM_Group.BoundText = ""
xRecordNumber = ""
End Sub
Sub myProc()
cardtable.FindFirst "Code = " & MyParn(GrdText(Search.Grid1, 0))
MyLoad
End Sub
Sub MyLoad()
xCode.Text = cardtable.Code
xDescA.Text = cardtable.DESCA
xM_Group.BoundText = TurnValue(cardtable.m_Group, Null, "")
xRecordNumber = "”Ã· " & cardtable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
End Sub
Sub MyReplace()
cardtable.FindFirst "Code = " & MyParn(xCode.Text)
If cardtable.NoMatch Then
    cardtable.AddNew
    formMode = addmode
Else
    cardtable.Edit
    formMode = Editmode
End If
cardtable.Code = xCode.Text
cardtable.DESCA = xDescA.Text
cardtable.m_Group = TurnValue(xM_Group.BoundText, "", Null)
cardtable.Update
End Sub
Function MYVALID()
If xCode.Text = "" Then
    MsgBox "ﬂÊœ «·„Ã„Ê⁄… ·« ÌÕ ÊÏ ⁄·Ï »Ì«‰« "
    Exit Function
End If

If formMode <> Editmode Then
    cardtable.FindFirst "Code = '" & xCode.Text & "'"
    If Not cardtable.NoMatch Then Exit Function
End If

If xDescA.Text = "" Then
    MsgBox "»Ì«‰ «·„Ã„Ê⁄… ·« ÌÕ ÊÏ ⁄·Ï »Ì«‰« "
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
    AddProc
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    cardtable.Delete
    cardtable.Requery
    If cardtable.RecordCount > 0 Then
        cardtable.MoveLast
        nRecordNumber = cardtable.RecordCount
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
CardLookup
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

If Not MsgBox(msgBoxStr, 4) = 6 Then
    CmdUndo_Click
    Exit Sub
End If
MyReplace
Select Case formMode
Case addmode, EmptyMode
    cardtable.Requery
    cardtable.MoveLast
    nRecordNumber = cardtable.RecordCount
    AddProc
Case Editmode
    editProc
End Select
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
Private Sub Form_Load()
Set cardtable = mydb.OpenRecordset("file1_50", dbOpenDynaset)
'Me.Picture = LoadPicture(App.Path & "\graph\02-02.jpg")
Data1.DatabaseName = MdbPath
Data1.RecordSource = "select * from file1_70 where flag = 2 order by desca "
xM_Group.ListField = "Desca"
xM_Group.BoundColumn = "CODE"
If cardtable.RecordCount > 0 Then
     cardtable.MoveLast
     nRecordNumber = cardtable.RecordCount
     MyLoad
     editProc
 Else
     EmptyProc
 End If
End Sub
