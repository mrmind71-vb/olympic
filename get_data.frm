VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form getDataFrm 
   Caption         =   "äÞá ÇáÈíÇäÇÊ"
   ClientHeight    =   1185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   2490
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "get_data.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   2310
      End
   End
   Begin VB.Frame Frame5 
      Height          =   825
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      Begin VB.TextBox xDrive 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   810
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton cmdGetData 
         Caption         =   "ÓÍÈ ÇáÈíÇäÇÊ"
         Height          =   450
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhotoNew 
         Caption         =   "ÓÍÈ ÇáÕæÑÉ ÇáÍÏíËÉ"
         Height          =   450
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhoto 
         Caption         =   "ÓÍÈ ßá ÇáÕæÑ"
         Height          =   450
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Drive "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   555
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   945
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "getDataFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdGetData_Click()
Dim fs As New FileSystemObject, conMdb As New ADODB.Connection, aInsert As Variant
Dim sSource As String
If Trim(xDrive.Text) = "" Then
    MsgBox "ÇáÞÑÕ ÛíÑ ãÓÌá"
    Exit Sub
End If


Dim sTarget As String, sòSource As String
sSource = xDrive.Text & ":\olymbic_door_sql\data_trans.mdb"
sTarget = App.Path & "\mdb\data_trans.mdb"

fs.CopyFile sSource, sTarget
openConMdb conMdb, sTarget

aSql = getCardMdb(con, conMdb, Me)

If Not IsEmpty(aSql) Then
    Prog1.Visible = True
    Prog1.Value = 0
    
    For i = 0 To UBound(aSql)
        Prog1.Value = mRound((i / (UBound(aSql))) * 100, 2)
        con.Execute aSql(i)
    Next
    
    Prog1.Visible = False
    Prog1.Value = 0
    
    closeCon con

    
    
    Inform "Êã ÇÑÓÇá ÇáÈíÇäÇÊ ÈäÌÇÍ"
Else
    MsgBox "áÇ ÊæÌÏ Çí ÈíÇäÇÊ áÇÑÓÇáåÇ"
End If
sSource = xDrive.Text & ":\olymbic_door_sql\photo1"
sTarget = "D:\PHOTO"

If fs.FolderExists(sSource) Then
    If GetPhotos(sSource, sTarget) Then Inform "Êã ÇÑÓÇá ÇáÕæÑ ÈäÌÇÍ"
    MsgBox "Êã ÇÑÓÇá ÇáÈíÇäÇÊ ÈäÌÇÍ"
Else
    MsgBox "ãáÝ ÇáÈíÇäÇÊ ÛíÑ ãæÌæÏ"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Prog1.Visible = False
Prog1.Value = 0
End Sub
Private Sub Form_Load()
xDrive.Text = RetSetting(xDrive.Name, TempSave(Me))
sError = openCon(con, LoadConString(, "client"))
End Sub
Private Function GetPhotos(pFolder As String, pFolderTarget As String) As Boolean
Dim fs As New FileSystemObject, sSource As String, nRecordcount As Double, i As Long

aString = retAllArray(pFolder, "jpg")
'On Error GoTo myerror
Prog1.Visible = True
Prog1.Value = 0
For i = 0 To UBound(aString)
    Prog1.Value = mRound(i / UBound(aString) * 100, 2)
    sSource = pFolder & "\" & aString(i)
    sTarget = pFolderTarget & "\" & aString(i)
    If fs.FileExists(sSource) Then
        bCopy = True
        If fs.FileExists(sTarget) Then
           If myFormat(fs.GetFile(sTarget).DateLastModified) >= myFormat(fs.GetFile(sSource).DateLastModified) Then
               bCopy = False
           End If
        End If
        If bCopy Then fs.CopyFile sSource, sTarget
    End If
Next
Prog1.Visible = False
Prog1.Value = 0
GetPhotos = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
Prog1.Visible = False
Prog1.Value = 0
End Function
Private Sub Form_Unload(Cancel As Integer)
addSetting xDrive.Name, xDrive.Text, TempSave(Me)
Set SendDataFrm = Nothing
Unload Me
End Sub

Private Sub xDrive_Change()
xDrive.Text = UCase(xDrive.Text)
End Sub
