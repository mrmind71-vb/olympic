VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form ChargeSub 
   Caption         =   "«ﬂÊ«œ «·„’«—Ì›"
   ClientHeight    =   2925
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   6390
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
   ScaleHeight     =   2925
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdLast 
      BackColor       =   &H00C0FFFF&
      Caption         =   "√ŒÌ—"
      Height          =   390
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2325
      Width           =   1065
   End
   Begin VB.CommandButton CmdFirst 
      BackColor       =   &H00C0FFFF&
      Caption         =   "√Ê·"
      Height          =   390
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2325
      Width           =   1140
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00C0FFFF&
      Caption         =   "·«Õﬁ"
      Height          =   390
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2325
      Width           =   1215
   End
   Begin VB.CommandButton CmdPrevious 
      BackColor       =   &H00C0FFFF&
      Caption         =   "”«»ﬁ"
      Height          =   390
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2325
      Width           =   1290
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   525
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   3825
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   690
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   1650
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   2865
   End
   Begin VB.PictureBox SSPanel2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   525
      ScaleWidth      =   6390
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   6390
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
         Height          =   390
         Left            =   150
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Height          =   390
         Left            =   1980
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Height          =   390
         Left            =   4725
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
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
         Height          =   390
         Left            =   3810
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
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
         Height          =   390
         Left            =   2895
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Height          =   390
         Left            =   1065
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
   Begin MSDBCtls.DBCombo xMainGroup 
      Bindings        =   "Chrgsub.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   1575
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…:"
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
      Index           =   0
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1650
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "«·»Ì‹‹‹«‰"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   570
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "ﬂÊœ"
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
      TabIndex        =   5
      Top             =   750
      Width           =   270
   End
End
Attribute VB_Name = "ChargeSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim cardtable As Recordset
Dim sFileName As String
Sub AddProc()
If Not formMode = addmode Then
    formMode = addmode
    Handlecontrols addmode
End If
myDefine
xCode.Text = IncRec(myLastField(cardtable, "Code"))
End Sub
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols EmptyMode
myDefine
xCode.Text = "01"
End Sub
Sub Handlecontrols(pMode)
Select Case pMode
Case Editmode
     CmdAdd.Enabled = True
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
    CmdAdd.Enabled = False
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
    CmdAdd.Enabled = False
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
    Generalarray(2) = "Select Code as «·ﬂÊœ,DescA as [»Ì«‰ «·„Ã„Ê⁄…] From " & sFileName
    Generalarray(3) = " Where DescA Like '%cFilter%'"
           
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
xDesca.Text = ""
xMainGroup.BoundText = ""
End Sub
Sub myProc()
cardtable.FindFirst "Code = " & MyParn(GrdText(Search.Grid1, 0))
MyLoad
End Sub
Sub MyLoad()
xCode.Text = cardtable.CODE
xDesca.Text = cardtable.DESCA
xMainGroup.BoundText = TurnValue(cardtable.MainGroup, Null, "")

End Sub
Sub MyReplace()
cardtable.FindFirst "Code = " & MyParn(xCode.Text)
If cardtable.NoMatch Then
    cardtable.AddNew
Else
    cardtable.Edit
End If
cardtable.CODE = xCode.Text
cardtable.DESCA = xDesca.Text
cardtable.MainGroup = TurnValue(xMainGroup.BoundText, "", Null)

cardtable.Update
cardtable.Requery
End Sub
Function MYVALID()
If xCode.Text = "" Then
    MsgBox "ﬂÊœ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If formMode <> Editmode Then
    cardtable.FindFirst "Code = '" & xCode.Text & "'"
    If Not cardtable.NoMatch Then Exit Function
End If

If xDesca.Text = "" Then
    MsgBox "»Ì«‰ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ ›«—€«"
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
        MyLoad
    Else
        EmptyProc
    End If
End If
End Sub
Private Sub cmdExit_Click()
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
SendKeys "{tab}"
Select Case formMode
Case addmode
    MyReplace
    AddProc
    
Case Editmode
    MyReplace

Case EmptyMode
    MyReplace
    AddProc
End Select
End Sub
Private Sub CmdSave_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then SendKeys "{TAB}"
End Sub

Private Sub CmdUndo_Click()
Select Case formMode
Case EmptyMode
    myDefine
Case addmode
    cardtable.Requery
    cardtable.MoveLast
    editProc
    MyLoad
Case Editmode
    MyLoad
End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If cardtable.RecordCount = 0 Then Exit Sub
If formMode <> Editmode Then Exit Sub
If KeyCode = 34 Then CmdPrevious_Click
If KeyCode = 33 Then CmdNext_Click
End Sub
Private Sub Form_Load()
formMode = Editmode
sFileName = IIf(publicFlag = 0, "File8_70", "File8_71")
Me.Caption = IIf(publicFlag = 0, "«·„’«—Ì›", "«·«Ì—«œ« ")
Set cardtable = mydb.OpenRecordset(sFileName, dbOpenDynaset)
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select * From file1_70 where flag = " & IIf(publicFlag = 0, 7, 13)
xMainGroup.ListField = "Desca"
xMainGroup.BoundColumn = "Code"


If cardtable.RecordCount > 0 Then
     MyLoad
     editProc
 Else
     EmptyProc
 End If
End Sub
Private Sub xMainGroup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xMainGroup.BoundText = ""
End Sub
