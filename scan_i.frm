VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form Scan_ifrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„”Õ ÷Ê∆Ì ·’Ê— «·«⁄÷«¡ Ê«· «»⁄Ì‰"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand cmdExit 
      Height          =   780
      Left            =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9090
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1376
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
      Picture         =   "scan_i.frx":0000
      Caption         =   "Exit"
      Alignment       =   8
      ButtonStyle     =   2
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9060
      Left            =   45
      ScaleHeight     =   9030
      ScaleWidth      =   15015
      TabIndex        =   0
      Top             =   0
      Width           =   15045
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "scan_i.frx":2323
         DragMode        =   1  'Automatic
         Height          =   2130
         Index           =   0
         Left            =   13140
         TabIndex        =   1
         Tag             =   "-1"
         Top             =   90
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   3757
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "íß“Ωª∫≠Ω≥´±“™ºØ´¥æÆØUBOR-FEOEONZI-EPCP6gI"
      End
      Begin Threed.SSCommand cmdScan 
         Height          =   465
         Index           =   0
         Left            =   14040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2250
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
         Picture         =   "scan_i.frx":2765
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "scan_i.frx":4C6D
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   330
         Index           =   0
         Left            =   14040
         TabIndex        =   6
         Top             =   2745
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ›Ÿ"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   12
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   330
         Index           =   0
         Left            =   13140
         TabIndex        =   7
         Top             =   2745
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —«Ã⁄"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   12
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   330
         Index           =   0
         Left            =   14040
         TabIndex        =   8
         Top             =   3105
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ–›"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   12
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   330
         Index           =   0
         Left            =   13140
         TabIndex        =   9
         Top             =   3105
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "‰ﬁ·"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   12
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdFile 
         Height          =   465
         Index           =   0
         Left            =   13140
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2250
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
         Picture         =   "scan_i.frx":6EEE
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "scan_i.frx":9133
      End
      Begin VB.Label xCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   0
         Left            =   13140
         TabIndex        =   5
         Top             =   4095
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label xName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   0
         Left            =   13140
         TabIndex        =   4
         Top             =   3780
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label xType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   0
         Left            =   13140
         TabIndex        =   3
         Top             =   3465
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
Attribute VB_Name = "Scan_ifrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents twain As ImgXTwain, nImageIndex As Integer
Attribute twain.VB_VarHelpID = -1
Dim fs As New FileSystemObject
Public sCode As String
Dim aPhoto() As String
Dim con As New ADODB.Connection
Private Sub CmdDel_Click(Index As Integer)
If Not validPhoto(RetPhoto_I(aPhoto(Index - 1, 2))) Then
    MsgBox "·«  ÊÃœ ’Ê—…"
    Exit Sub
End If
If MsgBox("”Ì „ «·«‰ Õ–› ’Ê—… «·⁄÷Ê ", vbOKCancel + vbDefaultButton2, "Õ–› «·’Ê—…") = vbCancel Then Exit Sub
Photo1(Index).Images.Clear
If fs.FileExists(RetPhoto_I(aPhoto(Index - 1, 2))) Then fs.DeleteFile RetPhoto_I(aPhoto(Index - 1, 2))
Inform " „ Õ–› «·’Ê—… »‰Ã«Õ"
Handlecontrols
End Sub

Private Sub cmdFile_Click(Index As Integer)
Dim cFile As String, cNewFile As String
On Error GoTo myerror
Common1.FileName = ""
Common1.InitDir = App.Path & "\PICT"
Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
Common1.ShowOpen
If Common1.FileTitle <> "" Then
    cFile = Common1.FileName
    If cFile <> "" Then
        'fs.CopyFile cFile, RetPhoto_i(xCode.Text)
        Photo1(Index).Images.Clear
        Photo1(Index).Import.FromFile cFile
    End If
    'LoadPhoto xCode.Text
End If
'MyReplace Index
Handlecontrols
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmdMove_Click(Index As Integer)
Dim cStr1, cStr2
On Error GoTo myerror
cStr1 = InputBox("√œŒ· —ﬁ„ «·⁄÷Ê «·„—«œ ‰ﬁ· «·’Ê—… «·ÌÂ")

If Trim(cStr1) = "" Then Exit Sub

cPos = InStr(1, cStr1, "-")
If cPos = 0 Then cStr2 = cStr1 Else cStr2 = Mid(cStr1, 1, cPos - 1)

If Trim(cStr2) = "" Then Exit Sub

If GetDesca("select code from file2_10 where code = " & cStr2) = "" Then
    MsgBox "·« ÌÊÃœ ⁄÷Ê »Â–« «·—ﬁ„"
    Exit Sub
End If

If fs.FileExists(RetPhoto_I(cStr1)) Then If MsgBox(" ÊÃœ ’Ê—… »‰›” «·≈”„ !! «” »œ«· «·’Ê—… «·Õ«·Ì… „ﬂ«‰Â«", vbOKCancel + vbDefaultButton2) = vbCancel Then Exit Sub

If MsgBox("”Ì „ «·«‰ ‰ﬁ· «·’Ê—… «·Ì «·⁄÷Ê " & cStr1, vbOKCancel + vbDefaultButton1, "‰ﬁ· «·’Ê—…") = vbCancel Then Exit Sub
Photo1(Index).Export.ToFile RetPhoto_I(cStr1), ixfsJPG
If fs.FileExists(RetPhoto_I(aPhoto(Index - 1, 2))) Then fs.DeleteFile RetPhoto_I(aPhoto(Index - 1, 2))
Photo1(Index).Images.Clear
Exit Sub
myerror:
MsgBox Err.Description
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub MyReplace(Index)
On Error GoTo myerror
If Not Photo1(Index).IsEmpty Then
    Photo1(Index).Export.ToFile RetPhoto_I(aPhoto(Index - 1, 2)), ixfsJPG
    Inform " „ Õ›Ÿ « ·„·›"
End If
Exit Sub
myerror:
Err.Clear
End Sub
Private Sub cmdSave_Click(Index As Integer)
MyReplace Index
Handlecontrols
End Sub
Private Sub cmdScan_Click(Index As Integer)
On Error GoTo myerror
nImageIndex = Index
Set twain = New ImgXTwain
twain.OpenTwain Me.hwnd
'If twain.QuerySupport(ixtcResolution) Then
'     twain.Resolution = 100
'End If
If twain.Sources.Count > 1 Then twain.SelectSource
twain.Acquire False, Me.hwnd
If cmdSave(Index).Enabled Then cmdSave(Index).SetFocus
'ImgX.Filters.AutoBrightness
Exit Sub
myerror:
    MsgBox Err.Description
End Sub
Private Sub Command3_Click()
For I = 0 To Photo1.Count - 1
    Photo1(I).Images.Clear
Next
End Sub

Private Sub CmdUndo_Click(Index As Integer)
Photo1(Index).Images.Clear
If validPhoto(RetPhoto_I(aPhoto(Index - 1, 2))) Then
    Photo1(Index).Import.FromFile RetPhoto_I(aPhoto(Index - 1, 2))
End If
End Sub

Private Sub CmdUndo2_Click()
myload
End Sub
Private Sub Form_Load()
openCon con
myload
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set twain = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set Scanfrm = Nothing
End Sub

Private Sub Photo1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
On Error GoTo myerror
If WrongFile(RetPhoto_I(aPhoto(Source.Index - 1, 2))) Then Exit Sub
If WrongFile(RetPhoto_I(aPhoto(Index - 1, 2))) Then Exit Sub
Photo1(0).Images.Clear
If Photo1(Index).IsEmpty Then
    Photo1(Index).Images.Replace Source.Image
    Photo1(Source.Index).Images.Clear
    MyReplace Index
    MyReplace Source.Index
    Handlecontrols
    If fs.FileExists(RetPhoto_I(aPhoto(Source.Index - 1, 2))) Then fs.DeleteFile RetPhoto_I(aPhoto(Source.Index - 1, 2))
ElseIf Photo1(Source.Index).IsEmpty Then
    Photo1(Source.Index).Images.Replace Photo1(Index).Image
    Photo1(Index).Images.Clear
    If fs.FileExists(RetPhoto_I(aPhoto(Index - 1, 2))) Then fs.DeleteFile RetPhoto_I(aPhoto(Index - 1, 2))
    MyReplace Source.Index
    Handlecontrols
Else
    Photo1(0).Images.Replace Photo1(Index).Image
    Photo1(Index).Images.Replace Source.Image
    Source.Images.Replace Photo1(0).Image
    MyReplace Index
    MyReplace Source.Index
    Handlecontrols
End If
'CmdSave(Source.Index).Enabled = True
'CmdUndo(Source.Index).Enabled = True
'CmdSave(Index).Enabled = True
'CmdUndo(Index).Enabled = True
Exit Sub
myerror:
MsgBox Err.Description
End Sub

Private Sub Twain_ImageAcquired(Image As ImgX_Image)
Photo1(nImageIndex).Visible = True
Photo1(nImageIndex).Images.Replace Image
Handlecontrols
End Sub
Private Sub Twain_TwainError(ByVal erNum As Long, ByVal erSource As String, ByVal Description As String)
MsgBox "Error Number:  " & erNum & vbCrLf & vbCrLf & Description, vbInformation, erSource
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub
Sub myload()
Dim RelationTable As New ADODB.Recordset
cString = "Select file2_11.member, file2_11.code,file2_11.desca,file2_11.title,relation_codes.desca as rel_Desca  From file2_11 left join relation_codes on file2_11.relation = relation_codes.code where file2_11.Member = " & sCode & " order by code"
RelationTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (RelationTable.EOF And RelationTable.BOF) Then
    ReDim aPhoto(RelationTable.RecordCount, 2)
Else
    ReDim aPhoto(0, 2)
End If

aPhoto(0, 0) = "⁄÷Ê «”«”Ì"
aPhoto(0, 1) = GetField("select desca from file2_10 where code = " & sCode)
aPhoto(0, 2) = sCode

I = 0
With RelationTable
    Do Until RelationTable.EOF
        I = I + 1
        aPhoto(I, 0) = !Rel_desca
        aPhoto(I, 1) = !Title & turn(!Title & "", "/") & !Desca
        aPhoto(I, 2) = !Member & "-" & !CODE
       .MoveNext
    Loop
myloadcontrols
End With
End Sub
Private Sub myloadcontrols()
On Error GoTo myerror
nCols = 7
NROWS = 2
nWidth = Photo1(0).Width + 100
nHeight = 4500

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load Photo1(Photo1.Count)
    Photo1(Photo1.Count - 1).Top = Photo1(0).Top + ((nRow - 1) * nHeight)
    Photo1(Photo1.Count - 1).Left = Photo1(0).Left - (nWidth * (nCol - 1))
    If validPhoto(RetPhoto_I(aPhoto(I - 1, 2))) Then
        Photo1(Photo1.Count - 1).Import.FromFile RetPhoto_I(aPhoto(I - 1, 2))
    Else
        If fs.FileExists(RetPhoto_I(aPhoto(I - 1, 2))) Then MsgBox "Œÿ√ ›Ï «·’Ê—… —ﬁ„ " & aPhoto(I - 1, 2)
    End If
    Photo1(Photo1.Count - 1).Visible = True
    Photo1(Photo1.Count - 1).Tag = aPhoto(I - 1, 2)
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load xCode(xCode.Count)
    xCode(xCode.Count - 1).Top = xCode(0).Top + ((nRow - 1) * nHeight)
    xCode(xCode.Count - 1).Left = xCode(0).Left - (nWidth * (nCol - 1))
    xCode(xCode.Count - 1).Visible = True
    xCode(xCode.Count - 1).Caption = "—ﬁ„ : " & aPhoto(I - 1, 2)
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load xName(xName.Count)
    xName(xName.Count - 1).Top = xName(0).Top + ((nRow - 1) * nHeight)
    xName(xName.Count - 1).Left = xName(0).Left - (nWidth * (nCol - 1))
    xName(xName.Count - 1).Visible = True
    xName(xName.Count - 1).Caption = aPhoto(I - 1, 1)
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load xType(xType.Count)
    xType(xType.Count - 1).Top = xType(0).Top + ((nRow - 1) * nHeight)
    xType(xType.Count - 1).Left = xType(0).Left - (nWidth * (nCol - 1))
    xType(xType.Count - 1).Visible = True
    xType(xType.Count - 1).Caption = aPhoto(I - 1, 0)
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmdScan(cmdScan.Count)
    cmdScan(cmdScan.Count - 1).Top = cmdScan(0).Top + ((nRow - 1) * nHeight)
    cmdScan(cmdScan.Count - 1).Left = cmdScan(0).Left - (nWidth * (nCol - 1))
    cmdScan(cmdScan.Count - 1).Visible = True
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmdFile(cmdFile.Count)
    cmdFile(cmdFile.Count - 1).Top = cmdFile(0).Top + ((nRow - 1) * nHeight)
    cmdFile(cmdFile.Count - 1).Left = cmdFile(0).Left - (nWidth * (nCol - 1))
    cmdFile(cmdFile.Count - 1).Visible = True
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmdSave(cmdSave.Count)
    cmdSave(cmdSave.Count - 1).Top = cmdSave(0).Top + ((nRow - 1) * nHeight)
    cmdSave(cmdSave.Count - 1).Left = cmdSave(0).Left - (nWidth * (nCol - 1))
    cmdSave(cmdSave.Count - 1).Visible = True
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmdUndo(cmdUndo.Count)
    cmdUndo(cmdUndo.Count - 1).Top = cmdUndo(0).Top + ((nRow - 1) * nHeight)
    cmdUndo(cmdUndo.Count - 1).Left = cmdUndo(0).Left - (nWidth * (nCol - 1))
    cmdUndo(cmdUndo.Count - 1).Visible = True
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmddel(cmddel.Count)
    cmddel(cmddel.Count - 1).Top = cmddel(0).Top + ((nRow - 1) * nHeight)
    cmddel(cmddel.Count - 1).Left = cmddel(0).Left - (nWidth * (nCol - 1))
    cmddel(cmddel.Count - 1).Visible = True
Next

nRow = 0
nCol = 0
For I = 1 To UBound(aPhoto) + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    Load cmdMove(cmdMove.Count)
    cmdMove(cmdMove.Count - 1).Top = cmdMove(0).Top + ((nRow - 1) * nHeight)
    cmdMove(cmdMove.Count - 1).Left = cmdMove(0).Left - (nWidth * (nCol - 1))
    cmdMove(cmdMove.Count - 1).Visible = True
Next
Handlecontrols
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Handlecontrols(Optional bChanged As Boolean = True)
On Error Resume Next
For i2 = 1 To cmdMove.Count - 1
'    cmdSave(i2).Enabled = (Photo1(i2).IsChanged) And Not Photo1(i2).IsEmpty
'    cmdUndo(i2).Enabled = Photo1(i2).IsChanged And Not Photo1(i2).IsEmpty
    
    cmdSave(i2).Enabled = Not Photo1(i2).IsEmpty
    cmdUndo(i2).Enabled = Not Photo1(i2).IsEmpty
    
    cmddel(i2).Enabled = fs.FileExists(RetPhoto_I(aPhoto(i2 - 1, 2)))
    
    'cmdFile(i2).Enabled = Photo1(i2).IsEmpty
    'cmdScan(i2).Enabled = Photo1(i2).IsEmpty
    
    If fs.FileExists(RetPhoto_I(aPhoto(i2 - 1, 2))) And Not validPhoto(RetPhoto_I(aPhoto(i2 - 1, 2))) Then
        If Photo1(i2).BackColor <> vbRed Then Photo1(i2).BackColor = vbRed
    Else
        If Photo1(i2).BackColor <> &H8000000F Then Photo1(i2).BackColor = &H8000000F
    End If
    cmdMove(i2).Enabled = validPhoto(RetPhoto_I(aPhoto(i2 - 1, 2)))
Next
End Sub
Private Function WrongFile(cFileName As String) As Boolean
If fs.FileExists(cFileName) Then
    If Not validPhoto(cFileName) Then
        WrongFile = True
        Exit Function
    End If
End If
End Function

