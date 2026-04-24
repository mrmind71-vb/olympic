VERSION 5.00
Begin VB.Form SendDataFrm 
   Caption         =   "‰ﬁ· «·»Ì«‰« "
   ClientHeight    =   1215
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
   ScaleHeight     =   1215
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
         Picture         =   "send_data_old.frx":0000
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
         Caption         =   "‰ﬁ· «·»Ì«‰« "
         Height          =   450
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhotoNew 
         Caption         =   "”Õ» «·’Ê—… «·ÕœÌÀ…"
         Height          =   450
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhoto 
         Caption         =   "”Õ» ﬂ· «·’Ê—"
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
End
Attribute VB_Name = "SendDataFrm"
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
If Trim(xDrive.text) = "" Then
    MsgBox "«·ﬁ—’ €Ì— „”Ã·"
    Exit Sub
End If

If Not MyCreateFolder(xDrive.text & ":\olymbic_door_sql") Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·»—‰«„Ã"
    Exit Sub
End If

If Not MyCreateFolder(xDrive.text & ":\photo_i") Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·»—‰«„Ã"
    Exit Sub
End If

If Not MyCreateFolder(xDrive.text & ":\olymbic_door_sql\photo1") Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·’Ê—"
    Exit Sub
End If

If Not MyCreateFolder(xDrive.text & ":\olymbic_door_sql\photo2") Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·’Ê—"
    Exit Sub
End If

If Not MyCreateFolder(xDrive.text & ":\olymbic_door_sql\photo3") Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·’Ê—"
    Exit Sub
End If

If Not MyCreateFolder(xDrive.text & ":\olymbic_door_sql\photo4") Then
    MsgBox "„‘ﬂ·… ›Ï «‰‘«¡ „”«— «·’Ê—"
    Exit Sub
End If


Dim sTarget As String, sÚSource As String
sTarget = xDrive.text & ":\olymbic_door_sql\data_trans.mdb"
sSource = App.Path & "\mdb\data_trans.mdb"

fs.CopyFile App.Path & "\mdb\data_empty.mdb", sSource

openConMdb conMdb, sSource

aSql = SendCardMdb("", "", con, "dbo.f_last_year_CODE(file1_10.code) = " & sSeason, Me)
'If sString <> "" Then
'    aSql = SplitSql(sString, 1)
'End If

If Not IsEmpty(aSql) Then
    prog1.Visible = True
    prog1.Value = 0
    
    For I = 0 To UBound(aSql)
        prog1.Value = mRound((I / (UBound(aSql))) * 100, 2)
        conMdb.Execute aSql(I)
    Next
    
    prog1.Visible = False
    prog1.Value = 0
    
    closeCon conMdb
    
    fs.CopyFile sSource, sTarget
    
    
    Inform " „ «—”«· «·»Ì«‰«  »‰Ã«Õ"
Else
    MsgBox "·«  ÊÃœ «Ì »Ì«‰«  ·«—”«·Â«"
End If

aString = Empty
aString = sendCardPhoto("", "", con, "dbo.f_last_year_CODE(file1_10.code) = " & sSeason, Me)

If fs.FileExists(sSource) Then
    If sendPhotos(aString) Then Inform " „ «—”«· «·’Ê— »‰Ã«Õ"
    MsgBox " „ «—”«· «·»Ì«‰«  »‰Ã«Õ"
Else
    MsgBox "„·› «·»Ì«‰«  €Ì— „ÊÃÊœ"
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
prog1.Visible = False
prog1.Value = 0
End Sub
Private Sub Form_Load()
xDrive.text = RetSetting(xDrive.Name, TempSave(Me))
openCon con
End Sub
Private Function sendPhotos(aString) As Boolean
Dim fs As New FileSystemObject, sSource As String, nRecordcount As Double, I As Long
If IsEmpty(aString) Then Exit Function

On Error GoTo myError
prog1.Visible = True
prog1.Value = 0
For I = 0 To UBound(aString)
    I = I + 1
    prog1.Value = mRound(I / UBound(aString) * 100, 2)
    sSource = RetPhoto(aString(I))
    sTarget = myPhoto_Path(aString(I), xDrive & ":\olymbic_door_sql\photo1")
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
prog1.Visible = False
prog1.Value = 0
sendPhotos = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
prog1.Visible = False
prog1.Value = 0
End Function
Private Sub Form_Unload(Cancel As Integer)
addSetting xDrive.Name, xDrive.text, TempSave(Me)
Set SendDataFrm = Nothing
Unload Me
End Sub

Private Sub xDrive_Change()
xDrive.text = UCase(xDrive.text)
End Sub
