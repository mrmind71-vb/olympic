VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form t 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”Õ» «·»Ì«‰«  „‰ «·—∆Ì”Ì"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Drive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2340
      Width           =   1050
      Begin VB.TextBox xDrive 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÿ—Ìﬁ… «·”Õ»"
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2295
      Width           =   1590
      Begin VB.CheckBox xAuto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ·ﬁ«∆Ì"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   315
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      TabIndex        =   3
      Top             =   945
      Width           =   5730
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "”Õ» »Ì«‰«  «·«’‰«›"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5730
   End
   Begin VB.Frame Frame11 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1665
      Width           =   5730
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "t"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection, cDataFolder As String, cDataFile As String
Public nFlag As Integer
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub CmdGo_Click()
Me.Caption = "”Õ» «·»Ì«‰«  „‰ «·›—⁄ «·—∆Ì”Ì"
If Not getData Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ ”Õ» «·»Ì«‰«  „‰ «·›—⁄ «·—∆Ì”Ì"
    Exit Sub
Else
    MsgBox " „ﬂ‰ «·‰Ÿ«„ „‰ ”Õ» «·»Ì«‰«  „‰ «·›—⁄ «·—∆Ì”Ì"
End If
End Sub
Private Sub Form_Load()
xDrive.Text = RetSetting("drive", TempSave(Me))
cDataFolder = App.Path & "\mdf"
cDataFile = "data"
openCon con
End Sub
Private Function getData() As Boolean
Dim bExit As Boolean
    con.BeginTrans
    
    nRecordCount = GetGroup
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ „Ã„Ê⁄«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getItems
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getCode("FILE1_10SC", True)
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«ﬁ”«„", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getCode("FILE1_50G")
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „Ã„Ê⁄«  «·«’‰«› «·—∆Ì”Ì…", "»‰Ã«Õ" Else GoTo myerror
        
    con.CommitTrans
    getData = True
Exit Function
myerror:
con.RollbackTrans
End Function
Private Function GetGroup() As Long
Dim conmdb As New adodb.Connection, loctable As New adodb.Recordset, cFile As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

con.Execute "DELETE FROM FILE1_50"

cFile = "FILE1_50"
cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert(6, 1)
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If


Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addvalue(loctable!CODE)
    
    aInsert(1, 0) = "desca"
    aInsert(1, 1) = addstring(loctable!desca)
       
    aInsert(2, 0) = "[GROUP]"
    aInsert(2, 1) = addvalue(loctable!Group)
    
    con.Execute CreateInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    GetGroup = GetGroup + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetGroup = -1
GoTo lastsub
End Function
Private Function getItems() As Long
Dim conmdb As New adodb.Connection, loctable As New adodb.Recordset
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
Dim cFile As String

cFile = "FILE1_10"
cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert(25, 1)
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If

con.Execute "DELETE FROM FILE1_10"
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert(0, 0) = "item"
    aInsert(0, 1) = addstring(loctable!Item)

    aInsert(1, 0) = "desca"
    aInsert(1, 1) = addstring(loctable!desca & "")

    aInsert(2, 0) = "[Group]"
    aInsert(2, 1) = addvalue(loctable!Group & "")

    aInsert(3, 0) = "[SECTION]"
    aInsert(3, 1) = addvalue(loctable!Section)

    aInsert(4, 0) = "[COLOR]"
    aInsert(4, 1) = addstring(loctable!Color)

    aInsert(5, 0) = "[COST]"
    aInsert(5, 1) = Val(loctable!cost & "")

    aInsert(6, 0) = "[PRICE]"
    aInsert(6, 1) = Val(loctable!price & "")

    aInsert(7, 0) = "[part_no]"
    aInsert(7, 1) = addstring(loctable!Part_no & "")
       
    aInsert(8, 0) = "[COLOR_NO]"
    aInsert(8, 1) = addstring(loctable!Color)
       
    con.Execute CreateInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    getItems = getItems + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getItems = -1
GoTo lastsub
End Function
Private Function getCode(cFile As String, Optional isNumber As Boolean = False) As Long
Dim conmdb As New adodb.Connection, loctable As New adodb.Recordset

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cString = "SELECT * FROM " & cFile

loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert(2, 1)
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
con.Execute "DELETE FROM " & cFile
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    If isNumber Then cString = addvalue(loctable!CODE) Else cString = addstring(loctable!CODE)
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = cString
    
    aInsert(1, 0) = "desca"
    aInsert(1, 1) = addstring(loctable!desca)
           
    con.Execute CreateInsert(aInsert, cFile), nAffect
    getCode = getCode + nAffect
    loctable.MoveNext
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getCode = -1
GoTo lastsub
End Function
Private Function CopyData() As Boolean
Dim fs As New FileSystemObject
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & "blk.mdb") Then
    fs.CopyFile cDataFolder & "\" & cDataFile & "_" & "blk.mdb", cDataFolder & "\" & cDataFile & ".mdb"
End If
CopyData = True
End Function
Private Function CopyToBranch() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
If fs.FileExists(cDataFolder & "\" & cDataFile & ".mdb") Then
    If Trim(xDrive.Text) <> "" Then sLastDrive = xDrive.Text Else sLastDrive = LastDrive(True)
    If sLastDrive <> "" Then
        noReadOnly cDataFolder & "\" & cDataFile & ".mdb"
        noReadOnly sLastDrive & ":\elmorshed\mdb" & "\" & cDataFile & ".mdb"
        MyCreateFolder (sLastDrive & ":\elmorshed\mdb")
        fs.CopyFile cDataFolder & "\" & cDataFile & ".mdb", sLastDrive & ":\elmorshed\mdb" & "\" & cDataFile & ".mdb"
    End If
End If
CopyToBranch = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function CopyFromMain() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
If Trim(xDrive.Text) <> "" Then sLastDrive = xDrive.Text Else sLastDrive = LastDrive(True)
MyCreateFolder cDataFolder
If sLastDrive <> "" Then
    noReadOnly sLastDrive & ":\elmorshed\mdb" & "\" & cDataFile & ".mdb"
    noReadOnly cDataFolder & "\" & cDataFile & ".mdb"
    If fs.FileExists(sLastDrive & ":\elmorshed\mdb" & "\" & cDataFile & ".mdb") Then
        fs.CopyFile sLastDrive & ":\elmorshed\mdb" & "\" & cDataFile & ".mdb", cDataFolder & "\" & cDataFile & ".mdb"
    End If
End If
CopyFromMain = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function CopyToMain() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & sBranchCode & ".mdb") Then
    If Trim(xDrive.Text) <> "" Then sLastDrive = xDrive.Text Else sLastDrive = LastDrive(True)
    If sLastDrive <> "" Then
        MyCreateFolder (sLastDrive & ":\elmorshed\mdb")
        fs.CopyFile cDataFolder & "\" & cDataFile & "_" & sBranchCode & ".mdb", sLastDrive & ":\elmorshed\mdb" & "\" & cDataFile & "_" & sBranchCode & ".mdb", True
    End If
End If
CopyToMain = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function copyData2() As Boolean
Dim fs As New FileSystemObject
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & "blk.mdb") Then
    fs.CopyFile cDataFolder & "\" & cDataFile & "_" & "blk.mdb", cDataFolder & "\" & cDataFile & "_" & sBranchCode & ".mdb"
End If
copyData2 = True
End Function
Private Function validData() As Boolean
Dim fs As New FileSystemObject
If Not fs.FileExists(cDataFolder & "\" & cDataFile & ".mdb") Then Exit Function
validData = True
End Function
Private Sub DeleteValid(sFile As String, sField As String, Optional bNum As Boolean, Optional cFilter As String)
Dim loctable As New adodb.Recordset, cString As String
cString = "SELECT " & sField & " FROM " & sFile
If cFilter <> "" Then cString = cString & turn(cString) & cFilter

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
On Error Resume Next
prog1.Value = 0
prog1.Visible = True
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    con.Execute "DELETE FROM " & sFile & " WHERE " & sField & " = " & IIf(bNum, loctable(sField & ""), MyParn(loctable(sField & "")))
    If Err.Number <> 0 Then Err.Clear
    loctable.MoveNext
Loop
prog1.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
addSetting "drive", xDrive.Text, TempSave(Me)
Set transDatafrm = Nothing
End Sub
