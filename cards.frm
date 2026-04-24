VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form cardsfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "‘Õ‰ ﬂ«—‰ÌÂ«  «·«⁄÷«¡"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19155
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   19155
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9765
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   979
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "cards.frx":0000
      Caption         =   "Exit"
      ButtonStyle     =   2
      PictureAlignment=   9
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9645
      Left            =   0
      ScaleHeight     =   9615
      ScaleWidth      =   18930
      TabIndex        =   1
      Top             =   90
      Width           =   18960
      Begin VB.TextBox xCard 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   17055
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   2205
         Visible         =   0   'False
         Width           =   1770
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "cards.frx":2323
         Height          =   2130
         Index           =   0
         Left            =   17055
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "-1"
         Top             =   45
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   3757
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "íß“Ωª∫≠Ω≥´±“™ºØ´¥æÆØUBOR-FEOEONZI-EPCP6gI"
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Index           =   0
         Left            =   17055
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2610
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cards.frx":2765
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "cards.frx":515A
      End
      Begin Threed.SSCommand cmdBarcode 
         Height          =   465
         Index           =   0
         Left            =   17955
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2610
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cards.frx":79F3
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "cards.frx":9E6E
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   375
         Index           =   0
         Left            =   17055
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2205
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cards.frx":C15B
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "cards.frx":E2BB
      End
      Begin VB.Label cardLbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   17055
         TabIndex        =   10
         Top             =   4365
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label xType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   17055
         TabIndex        =   8
         Top             =   3141
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label xName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   17055
         TabIndex        =   7
         Top             =   3510
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label xCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   17055
         TabIndex        =   6
         Top             =   3990
         Visible         =   0   'False
         Width           =   1770
      End
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "cardsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nImageIndex As Integer, nType As Integer
Attribute nImageIndex.VB_VarHelpID = -1
Dim cFile As String, cFileRel As String, cProcName As String
Dim fs As New FileSystemObject, bAct As Boolean
Public sCode As String
Dim aPhoto() As String
Dim con As New ADODB.Connection

Private Sub cmdBarcode_Click(Index As Integer)
Dim sPath As String
Clipboard.Clear
sPath = sPath_App & "\CardReader\CardReader.exe"
'ShellExecute Me.hwnd, sPath, sPath, vbNullString, "C:\", SW_SHOWNORMAL
'RunIt sPath, vbNormalFocus
'CardTimer.Enabled = True
xCard(Index).SetFocus
'myGotFocus xCard(Index)

ShellExWait sPath, vbNullString, Me
sCard = Clipboard.GetText


If sCard <> "" Then
    xCard(Index).Text = sCard
    If cmdSave(Index).Enabled Then
        cmdSave_Click (Index)
    End If
Else
    MsgBox "·„ Ì „ ﬁ—«¡… «·ﬂ«—‰ÌÂ"
End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function MyReplace(aData, Index) As Boolean
On Error GoTo myerror
con.BeginTrans
On Error GoTo myerror
If ValidNum(retFlag(aData, "ID")) Then
    con.Execute "update " & cFileRel & " set card = " & addstring(xCard(Index).Text) & " where id = " & retFlag(aData, "id")
Else
    con.Execute "update " & cFile & "  set card = " & addstring(xCard(Index).Text) & " where code = " & addvalue(retFlag(aData, "member"))
End If
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function

Private Sub CmdUndo2_Click()
myload
End Sub
Private Sub cmdSave_Click(Index As Integer)
If Trim(xCard(Index).Text) <> "" Then
    If Not MYVALID(Index, cmdSave(Index).TagVariant) Then Exit Sub
End If
If MyReplace(cmdSave(Index).TagVariant, Index) Then
    Inform " „ «·Õ›Ÿ »‰Ã«Õ"
    xCard(Index).Tag = xCard(Index).Text
    cmdSave(Index).Enabled = False
    cardLbl(Index).Caption = xCard(Index).Text
End If
End Sub
Private Function MYVALID(Index, aData) As Boolean
MYVALID = True
If (Not ValidNumAny(xCard(Index).Text)) Then
    MsgBox "—ﬁ„ €Ì— ’ÕÌÕ"
    MYVALID = False
Else
    If Trim(xCard(Index).Text) <> "" Then
        cWhere = "card = " & MyParn(xCard(Index).Text)
        If ValidNum(retFlag(aData, "ID")) Then
            cWhere1 = cWhere & " and " & cFileRel & ".id <> " & retFlag(aData, "ID")
            cWhere2 = cWhere
        Else
            cWhere1 = cWhere
            cWhere2 = cWhere & " and " & cFile & ".code <> " & retFlag(aData, "member")
        End If
            
        aRet = GetFields("Select member,code,ID from " & cFileRel & " where " & cWhere1, con)
        If Not IsEmpty(aRet) Then
            If MsgBox("ﬂ«—‰ÌÂ »‰›” «·—ﬁ„ ··⁄÷ÊÌ… " & retFlag(aRet, "member") & "  «»⁄ —ﬁ„ " & retFlag(aRet, "code") & vbCrLf & "Õ–› —ﬁ„ «·ﬂ«—‰ÌÂ", vbOKCancel + vbDefaultButton2) = vbOK Then
                On Error GoTo myerror
                con.Execute "update " & cFileRel & " set card = null  WHERE ID = " & retFlag(aRet, "ID")
                Inform " „ «·Õ–› »‰Ã«Õ"
            Else
                MYVALID = False
            End If
        End If
        
        aRet = GetFields("Select code from " & cFile & " where " & cWhere2, con)
        If Not IsEmpty(aRet) Then
            If MsgBox("ﬂ«—‰ÌÂ »‰›” «·—ﬁ„ ··⁄÷ÊÌ… " & retFlag(aRet, "code") & vbCrLf & "Õ–› —ﬁ„ «·ﬂ«—‰ÌÂ", vbOKCancel + vbDefaultButton2) = vbOK Then
                On Error GoTo myerror
                con.Execute "update " & cFile & "  set card = null WHERE CODE = " & retFlag(aRet, "CODE")
                Inform " „ «·Õ–› »‰Ã«Õ"
            Else
                MYVALID = False
            End If
        End If
    End If
End If
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
MYVALID = False
End Function
Private Sub CmdUndo_Click(Index As Integer)
xCard(Index).Text = Member_Load(retFlag(cmdSave(Index).TagVariant, "member"), "card", con) & ""
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    xCard(1).SetFocus
    myGotFocus xCard(1)
End If
End Sub
Private Sub Form_Load()
If nType = 0 Then
    Me.Caption = " ‰‘Ìÿ ﬂ«—‰ÌÂ«  «·«⁄÷«¡ «·⁄«„·Ì‰"
    cFile = "FILE1_10"
    cFileRel = "FILE1_11"
    cProcName = "MEMBER_LOAD"
ElseIf nType = 1 Then
    Me.Caption = " ‰‘Ìÿ ﬂ«—‰ÌÂ«  «·⁄÷ÊÌ«  «·„ﬁ”ÿ…"
    cFile = "FILE2_10"
    cFileRel = "FILE2_11"
    cProcName = "MEMBER_LOAD_INSTALL"
End If
openCon con
myload
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set twain = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set cardsfrm = Nothing
End Sub
Sub myload()
Dim RelationTable As New ADODB.Recordset
cString = "Select  member," & cFileRel & ".code," & cFileRel & ".desca," & cFileRel & ".title,relation_codes.desca as rel_Desca," & cFileRel & ".card," & cFileRel & ".id  From " & cFileRel & " left join relation_codes on " & cFileRel & ".relation = relation_codes.code where Member = " & sCode & " order by " & cFileRel & ".code"
RelationTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (RelationTable.EOF And RelationTable.BOF) Then
    ReDim aPhoto(RelationTable.RecordCount, 5)
Else
    ReDim aPhoto(0, 5)
End If
aPhoto(0, 0) = "⁄÷Ê «”«”Ì"
aPhoto(0, 1) = Member_Load(sCode, "desca", con, cProcName)
aPhoto(0, 2) = sCode
aPhoto(0, 3) = sCode
aPhoto(0, 4) = ""
aPhoto(0, 5) = Member_Load(sCode, "card", con, cProcName) & ""

i = 0
With RelationTable
    Do Until RelationTable.EOF
        i = i + 1
        aPhoto(i, 0) = !REL_DESCA & ""
        aPhoto(i, 1) = !desca & ""
        aPhoto(i, 2) = !MEMBER & "-" & !CODE
        aPhoto(i, 3) = !MEMBER
        aPhoto(i, 4) = !ID
        aPhoto(i, 5) = !card & ""
       .MoveNext
    Loop
myloadcontrols
End With
End Sub
Private Sub myloadcontrols()
On Error GoTo myerror
nCols = 10
NROWS = 2
nWidth = Photo1(0).Width + 100
nHeight = 4700

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load Photo1(Photo1.Count)
    Photo1(Photo1.Count - 1).Top = Photo1(0).Top + ((nRow - 1) * nHeight)
    Photo1(Photo1.Count - 1).Left = Photo1(0).Left - (nWidth * (nCol - 1))
    If nType = 0 Then
        If validPhoto(RetPhoto(aPhoto(i - 1, 2))) Then
            Photo1(Photo1.Count - 1).Import.FromFile RetPhoto(aPhoto(i - 1, 2))
        Else
            If fs.FileExists(RetPhoto(aPhoto(i - 1, 2))) Then MsgBox "Œÿ√ ›Ï «·’Ê—… —ﬁ„ " & aPhoto(i - 1, 2)
        End If
    Else
        If validPhoto(RetPhoto_I(aPhoto(i - 1, 2))) Then
            Photo1(Photo1.Count - 1).Import.FromFile RetPhoto_I(aPhoto(i - 1, 2))
        Else
            If fs.FileExists(RetPhoto_I(aPhoto(i - 1, 2))) Then MsgBox "Œÿ√ ›Ï «·’Ê—… —ﬁ„ " & aPhoto(i - 1, 2)
        End If
    End If
    Photo1(Photo1.Count - 1).Visible = True
    Photo1(Photo1.Count - 1).Tag = aPhoto(i - 1, 2)
Next

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmdSave(cmdSave.Count)
    cmdSave(cmdSave.Count - 1).Top = cmdSave(0).Top + ((nRow - 1) * nHeight)
    cmdSave(cmdSave.Count - 1).Left = cmdSave(0).Left - (nWidth * (nCol - 1))
    cmdSave(cmdSave.Count - 1).Visible = True
    
    cmdSave(cmdSave.Count - 1).TagVariant = AddFlag(Empty, "member", aPhoto(i - 1, 3))
    cmdSave(cmdSave.Count - 1).TagVariant = AddFlag(cmdSave(cmdSave.Count - 1).TagVariant, "id", aPhoto(i - 1, 4))
Next

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load xCard(xCard.Count)
    xCard(xCard.Count - 1).Top = xCard(0).Top + ((nRow - 1) * nHeight)
    xCard(xCard.Count - 1).Left = xCard(0).Left - (nWidth * (nCol - 1))
    xCard(xCard.Count - 1).Visible = True
    xCard(xCard.Count - 1).Tag = aPhoto(i - 1, 5)
    xCard(xCard.Count - 1).Tag = aPhoto(i - 1, 5)
    xCard(xCard.Count - 1).TabIndex = xCard.Count - 1
    xCard(i).TabStop = True
Next

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load xCode(xCode.Count)
    xCode(xCode.Count - 1).Top = xCode(0).Top + ((nRow - 1) * nHeight)
    xCode(xCode.Count - 1).Left = xCode(0).Left - (nWidth * (nCol - 1))
    xCode(xCode.Count - 1).Visible = True
    xCode(xCode.Count - 1).Caption = aPhoto(i - 1, 2)
Next

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load xName(xName.Count)
    xName(xName.Count - 1).Top = xName(0).Top + ((nRow - 1) * nHeight)
    xName(xName.Count - 1).Left = xName(0).Left - (nWidth * (nCol - 1))
    xName(xName.Count - 1).Visible = True
    xName(xName.Count - 1).Caption = aPhoto(i - 1, 1)
Next

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load xType(xType.Count)
    xType(xType.Count - 1).Top = xType(0).Top + ((nRow - 1) * nHeight)
    xType(xType.Count - 1).Left = xType(0).Left - (nWidth * (nCol - 1))
    xType(xType.Count - 1).Visible = True
    xType(xType.Count - 1).Caption = aPhoto(i - 1, 0)
Next

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmdBarcode(cmdBarcode.Count)
    cmdBarcode(cmdBarcode.Count - 1).Top = cmdBarcode(0).Top + ((nRow - 1) * nHeight)
    cmdBarcode(cmdBarcode.Count - 1).Left = cmdBarcode(0).Left - (nWidth * (nCol - 1))
    cmdBarcode(cmdBarcode.Count - 1).Visible = True
Next


'nRow = 0
'nCol = 0
'For i = 1 To UBound(aPhoto) + 1
'    nCol = IIf(nCol = nCols, 1, nCol + 1)
'    nRow = IIf(nCol = 1, nRow + 1, nRow)
'    Load cmdUndo(cmdUndo.Count)
'    cmdUndo(cmdUndo.Count - 1).Top = cmdUndo(0).Top + ((nRow - 1) * nHeight)
'    cmdUndo(cmdUndo.Count - 1).Left = cmdUndo(0).Left - (nWidth * (nCol - 1))
'    cmdUndo(cmdUndo.Count - 1).Visible = True
'Next

nRow = 0
nCol = 0
For i = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cardLbl(cardLbl.Count)
    cardLbl(cardLbl.Count - 1).Top = cardLbl(0).Top + ((nRow - 1) * nHeight)
    cardLbl(cardLbl.Count - 1).Left = cardLbl(0).Left - (nWidth * (nCol - 1))
    cardLbl(cardLbl.Count - 1).Visible = True
    cardLbl(cardLbl.Count - 1).Caption = aPhoto(i - 1, 5)
Next

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function WrongFile(cFileName As String) As Boolean
If fs.FileExists(cFileName) Then
    If Not validPhoto(cFileName) Then
        WrongFile = True
        Exit Function
    End If
End If
End Function


Private Sub xCard_Change(Index As Integer)
cmdSave(Index).Enabled = xCard(Index).Text <> xCard(Index).Tag
End Sub

Private Sub xCard_GotFocus(Index As Integer)
myGotFocus xCard(Index)
End Sub

Private Sub xCard_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave_Click Index
End If
End Sub
Private Sub xCard_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    xCard(Index).Text = ""
End If
End Sub
Private Sub xCard_LostFocus(Index As Integer)
myLostFocus xCard(Index)
End Sub
