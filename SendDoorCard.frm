VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form DoorCardSend 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«—”«· «·»Ì«‰«  ··»Ê«»…"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1185
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   90
      Width           =   3390
      Begin VB.TextBox xCode1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   1860
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„‰ —ﬁ„ "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2535
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈·Ì —ﬁ„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2550
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   555
      End
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   555
      Left            =   5175
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
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
      Caption         =   "«—”«· «·»Ì«‰«  ··»Ê«»…"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   2370
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   45
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   714
         _Version        =   196610
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
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   2175
      Visible         =   0   'False
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmddel 
      Height          =   555
      Left            =   3915
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   979
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "SendDoorCard.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "SendDoorCard.frx":279C
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   2430
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   979
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "SendDoorCard.frx":4C30
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
End
Attribute VB_Name = "DoorCardSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ntype As Integer
Public sFile As String
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
SendCards
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
sFile = IIf(ntype = 0, "FILE1_10", "FILE2_10")
LoadText Me, , ntype & ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set DoorSendCard = Nothing
End Sub

Private Sub xClosed_Click()
cmdApply.Caption = IIf(xClosed.Value = 0, "› Õ", "«€·«ﬁ")
End Sub
Private Sub SendCards()
Me.MousePointer = 11
Dim loctable As New ADODB.Recordset, nRecordcount As Long
Dim con2 As New ADODB.Connection, cInsert As String
sMsg = openCon(con2, CreateConStr2)
If sMsg <> "ok" Then
    MsgBox sMsg
    Exit Sub
End If

Dim cString As String, cWhere As String
cString = "SELECT * FROM " & sFile & " WHERE (NOT CARD IS NULL)"

If IsNumeric(xCode1.Text) Then
    cString = cString & turn(cString) & " CODE  " & IIf(IsNumeric(xCode2.Text), " >= ", " = ") & xCode1.Text
End If

If IsNumeric(xCode2.Text) Then
    cString = cString & turn(cString) & " CODE <= " & xCode2.Text
End If

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

nRecordcount = loctable.RecordCount


Dim i As Long
If loctable.EOF And loctable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰« "
    Exit Sub
End If


prog1.Visible = True
prog1.Value = 0
con2.BeginTrans
'On Error GoTo myerror
Do Until loctable.EOF
    i = i + 1
    prog1.Value = Round(i / (nRecordcount), 2) * 100
    If ntype = 0 Then
        cInsert = SendCard(loctable!CODE, , con, con2)
    Else
        cInsert = SendCardInstall(loctable!CODE, , con, con2)
    End If
    If cInsert <> "" Then
        con2.Execute cInsert
    End If
    loctable.MoveNext
Loop
con2.CommitTrans
closeCon con2
Me.MousePointer = 0
prog1.Visible = False
MsgBox " „ ‰ﬁ· «·»Ì«‰«  ··»Ê«»… »‰Ã«Õ"
panel1(0).Caption = " „ ‰ﬁ· " & i & " ”Ã·"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
prog1.Visible = False
End Sub
