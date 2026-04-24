VERSION 5.00
Begin VB.Form MonyBox 
   Caption         =   "√ﬂÊ«œ Œ“‰"
   ClientHeight    =   2895
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
   ScaleHeight     =   2895
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xF_Date 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1475
      Width           =   1515
   End
   Begin VB.TextBox xF_Bal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1875
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   5925
      TabIndex        =   12
      Top             =   2415
      Width           =   5925
      Begin VB.CommandButton CmdPrevious 
         BackColor       =   &H00C0FFFF&
         Caption         =   "”«»ﬁ"
         Height          =   390
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   75
         Width           =   1290
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«Õﬁ"
         Height          =   390
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   75
         Width           =   1215
      End
      Begin VB.CommandButton CmdFirst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√Ê·"
         Height          =   390
         Left            =   4425
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   75
         Width           =   1140
      End
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         Height          =   390
         Left            =   750
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   75
         Width           =   1065
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3675
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   690
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1075
      Width           =   2865
   End
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
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
      TabIndex        =   2
      Top             =   0
      Width           =   5925
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000018&
         Caption         =   "«÷«›…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   10
         Top             =   90
         Width           =   870
      End
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H80000018&
         Caption         =   "«” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4800
         TabIndex        =   7
         Top             =   90
         Width           =   870
      End
      Begin VB.CommandButton CmdUndo 
         BackColor       =   &H00C0C0C0&
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2910
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Õ–›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1065
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1995
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ìœ «Ê·"
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
      Top             =   1950
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «Ê·"
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
      TabIndex        =   11
      Top             =   1535
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰"
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
      Top             =   1120
      Width           =   315
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·Œ“‰…"
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
      TabIndex        =   8
      Top             =   705
      Width           =   810
   End
End
Attribute VB_Name = "MonyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim CardTable As Recordset
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
Sub CardLookup()
Dim Generalarray(3)
Dim GrdArray(2)
Set Generalarray(1) = Me
Generalarray(2) = "Select Code as «·ﬂÊœ,DescA as [»Ì«‰ ]From FILE0_50"
Generalarray(3) = " Where DescA Like '*cFilter*'"
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
xdesca.Text = ""
xF_date.Text = ""
xF_Bal.Text = ""
End Sub
Sub myProc()
CardTable.FindFirst "Code = " & MyParn(GrdText(Search.Grid1, 0))
MyLoad
End Sub
Sub MyLoad()
xCode.Text = CardTable.CODE
xdesca.Text = TurnValue(CardTable.DESCA, Null, "")
xF_date.Text = TurnValue(Format(CardTable.F_DATE, "DD-MM-YYYY"), Null, "")
xF_Bal.Text = TurnValue(CardTable.F_BAL, Null, "")
End Sub
Sub MyReplace()
CardTable.FindFirst "Code = " & MyParn(xCode.Text)
If CardTable.NoMatch Then
    CardTable.AddNew
Else
    CardTable.Edit
End If
CardTable.CODE = xCode.Text
CardTable.DESCA = xdesca.Text
CardTable.F_BAL = Val(xF_Bal.Text)
CardTable.F_DATE = xF_date.Text
CardTable.Update
End Sub
Function MYVALID()
If xCode.Text = "" Then
    MsgBox " ”ÃÌ· ﬂÊœ «·„Ã„Ê⁄… "
    Exit Function
End If

If xdesca.Text = "" Then
    MsgBox " ”ÃÌ· ≈”„ «·Œ“‰… "
    Exit Function
End If

If xF_date.Text = "" Then
    MsgBox " ”ÃÌ·  «—ÌŒ "
    Exit Function
End If

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
Private Sub CmdInform_Click()
CardLookup
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
CardTable.FindFirst "Code = " & MyParn(xCode.Text)
If Not CardTable.NoMatch Then
    MyReplace
Else
    MyReplace
    AddProc
End If
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
Set CardTable = mydb.OpenRecordset("SELECT * FROM file0_50 ORDER BY CODE ", dbOpenDynaset)
If CardTable.RecordCount > 0 Then
     MyLoad
     editProc
 Else
     EmptyProc
 End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
End Sub
