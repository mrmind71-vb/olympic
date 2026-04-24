VERSION 5.00
Begin VB.Form SecUser 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ﬂ·„… ”— «·‰Ÿ«„"
   ClientHeight    =   4125
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3075
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2415
      Width           =   690
   End
   Begin VB.TextBox xComp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      Height          =   315
      Left            =   900
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1968
      Width           =   2865
   End
   Begin VB.CheckBox xManger 
      BackColor       =   &H00FFFFFF&
      Caption         =   "’·«ÕÌ«  "
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
      Height          =   315
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3075
      Value           =   1  'Checked
      Width           =   4140
   End
   Begin VB.TextBox xPass 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      Height          =   315
      Left            =   900
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1537
      Width           =   2865
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3570
      Width           =   5925
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         Height          =   390
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.Label xRecordNumber 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   225
         TabIndex        =   18
         Top             =   75
         Width           =   1890
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3075
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
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2835
      Width           =   285
   End
   Begin VB.TextBox xUser 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      Height          =   315
      Left            =   900
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1106
      Width           =   2865
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   5925
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«÷«›…"
         Height          =   390
         Left            =   3885
         Style           =   1  'Graphical
         TabIndex        =   11
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Visible         =   0   'False
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   2
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÿ»«⁄… »Ê‰ «·»Ì⁄"
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
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·‘—ﬂ…"
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
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1950
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﬂ·„… «·”—"
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
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1551
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "≈”„ «·„” Œœ„"
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
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1128
      Width           =   1200
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "—ﬁ„ «·„” Œœ„"
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
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   705
      Width           =   1125
   End
End
Attribute VB_Name = "SecUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim CardTable As Recordset
Dim nRecordNumber As Integer
Sub AddProc()
If Not formMode = addmode Then Handlecontrols addmode
formMode = addmode
myDefine
xCode.Text = IncRec(myLastField(CardTable, "Code"))
End Sub
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols EmptyMode
myDefine
xCode.Text = IncRec(myLastField(CardTable, "Code"))
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
Sub editProc()
formMode = Editmode
Handlecontrols Editmode
End Sub
Sub myDefine()
xPass.Text = ""
xUser.Text = ""
xManger.Value = 0
xComp.Text = ""
End Sub
Sub MyLoad()
xCode.Text = CardTable.CODE
xUser.Text = CardTable.User
xComp.Text = CardTable.COMP
xCount.Text = Format(CardTable.Count, "#0")

xPass.Text = TurnValue(CardTable.PASS, Null, "")
If CardTable!MANGER Then
    xManger.Value = 1
Else
    xManger.Value = 0
End If
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
End Sub
Sub MyReplace()
CardTable.FindFirst "Code = " & MyParn(xCode.Text)
If CardTable.NoMatch Then
    CardTable.AddNew
    formMode = addmode
Else
    CardTable.Edit
    formMode = Editmode
End If
CardTable.CODE = xCode.Text
CardTable.COMP = xComp.Text
CardTable.Count = Val(xCount.Text)
CardTable.PASS = TurnValue(xPass.Text, "", Null)
CardTable.User = TurnValue(xUser.Text, "", Null)
CardTable.MANGER = xManger.Value
CardTable.Update
End Sub
Function MYVALID()
If formMode <> Editmode Then
    CardTable.FindFirst "Code = '" & xCode.Text & "'"
    If Not CardTable.NoMatch Then Exit Function
End If

MYVALID = True
End Function
Private Sub CmdAdd_Click()
    AddProc
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    CardTable.Delete
    CardTable.Requery
    If CardTable.RecordCount > 0 Then
        CardTable.MoveLast
        nRecordNumber = CardTable.RecordCount
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
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
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
    CardTable.Requery
    CardTable.MoveLast
    nRecordNumber = CardTable.RecordCount
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
    CardTable.MoveLast
    editProc
    MyLoad
Case Editmode
    MyLoad
End Select
End Sub
Private Sub Form_Load()
Set CardTable = sec.OpenRecordset("FILE_SS", dbOpenDynaset)
If CardTable.RecordCount > 0 Then
     CardTable.MoveLast
     nRecordNumber = CardTable.RecordCount
     MyLoad
     editProc
 Else
     EmptyProc
 End If
End Sub

