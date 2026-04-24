VERSION 5.00
Begin VB.Form TransBox 
   BackColor       =   &H00E0E0E0&
   Caption         =   " ÕÊÌ·«  ‰ﬁœÌ… „‰ Œ“‰… ≈·Ï  √Œ—Ï"
   ClientHeight    =   3150
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
   ScaleHeight     =   3150
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   3225
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1515
   End
   Begin VB.TextBox xNo2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   600
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1425
      Width           =   690
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   3225
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1020
      Width           =   1515
   End
   Begin VB.TextBox xValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   3225
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2250
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   5925
      TabIndex        =   16
      Top             =   2670
      Width           =   5925
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         Height          =   390
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   70
         Width           =   1065
      End
      Begin VB.CommandButton CmdFirst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√Ê·"
         Height          =   390
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   70
         Width           =   1140
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«Õﬁ"
         Height          =   390
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   70
         Width           =   1215
      End
      Begin VB.CommandButton CmdPrevious 
         BackColor       =   &H00C0FFFF&
         Caption         =   "”«»ﬁ"
         Height          =   390
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   70
         Width           =   1290
      End
   End
   Begin VB.TextBox xNo1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   4125
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1425
      Width           =   615
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   1875
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1845
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
      TabIndex        =   7
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
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   6
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”·"
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
      Left            =   5025
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   675
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï Œ“‰…"
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
      Left            =   1500
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁÌ„…"
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
      Left            =   4965
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2325
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
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
      Left            =   4965
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1080
      Width           =   390
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
      Left            =   4965
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1905
      Width           =   315
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ Œ“‰…"
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
      Left            =   4965
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1500
      Width           =   660
   End
End
Attribute VB_Name = "TransBox"
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
'xCode.Text = IncRec(myLastField(cardtable, "Code"))
xCode.Text = "000001"
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
'    CmdPrevious.Enabled = True
'    CmdNext.Enabled = True
'    CmdLast.Enabled = True
'    CmdFirst.Enabled = True
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
Dim GrdArray(3)
Set Generalarray(1) = Me
Generalarray(2) = "Select Code as «·ﬂÊœ, [DATE] as [ «—ÌŒ] ,DescA as [»Ì«‰ ]From FILE0_60"
Generalarray(3) = " Where DescA Like '*cFilter*'"
Generalarray(4) = " ORDER BY DATE "

GrdArray(1) = 1200
GrdArray(2) = 1500
GrdArray(3) = 2500
    
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
xDate.Text = ""
xNo1.Text = ""
xNo2.Text = ""
xValue.Text = ""
End Sub
Sub myProc()
CardTable.FindFirst "Code = " & MyParn(GrdText(Search.Grid1, 0))
MyLoad
End Sub
Sub MyLoad()
xCode.Text = CardTable.CODE
xdesca.Text = TurnValue(CardTable.DESCA, Null, "")
xNo1.Text = TurnValue(CardTable.no1, Null, "")
xNo2.Text = TurnValue(CardTable.no2, Null, "")
xDate.Text = TurnValue(Format(CardTable!Date, "DD-MM-YYYY"), Null, "")
xValue.Text = TurnValue(CardTable!Value, Null, "")
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
CardTable.Value = Val(xValue.Text)
CardTable.Date = xDate.Text
CardTable.no1 = xNo1.Text
CardTable.no2 = xNo2.Text
CardTable.Update
End Sub
Function MYVALID()
If xCode.Text = "" Then
    MsgBox " ”ÃÌ· „”·”· "
    Exit Function
End If

If xDate.Text = "" Then
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
Set CardTable = mydb.OpenRecordset("SELECT * FROM file0_60 ORDER BY CODE ", dbOpenDynaset)
If CardTable.RecordCount > 0 Then
     MyLoad
     editProc
Else
     EmptyProc
End If
End Sub
