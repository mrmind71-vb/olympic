VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FixMemberPaidfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "÷»ÿ ”œ«œ «·«⁄÷«¡"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1515
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdApply 
      Height          =   555
      Left            =   5175
      TabIndex        =   0
      Top             =   225
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
      Caption         =   "÷»ÿ «·”œ«œ"
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
      Top             =   1050
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
      Top             =   855
      Visible         =   0   'False
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   3690
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   225
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
      Picture         =   "FixMemberPaidt.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
End
Attribute VB_Name = "FixMemberPaidfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sFile As String, sFieldClose As String, sFieldDate As String, pFilter As String, sCaption As String
Public bTrans As Boolean
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
CreateFawry
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set FixMemberPaidfrm = Nothing
End Sub
Private Sub CreateFawry(Optional bBegin As Boolean = False)
Me.MousePointer = 11
Dim loctable As New ADODB.Recordset, nRecordcount As Long
loctable.Open "Select code from file1_10", con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Visible = True
prog1.Value = 0
con.BeginTrans
Dim I As Long
nRecordcount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    I = I + 1
    Me.Caption = sCaption & "-" & I & " „‰ " & nRecordcount
    prog1.Value = Round((I) / (nRecordcount), 2) * 100
    con.Execute fixMemberPaid(loctable!CODE)
    loctable.MoveNext
Loop
Me.Caption = sCaption
con.CommitTrans
Me.MousePointer = 0
MsgBox " „ ÷»ÿ ”œ«œ " & nRecords & " ⁄÷Ê »‰Ã«Õ"
prog1.Visible = False
panel1(0).Caption = "⁄„· " & nRercords & "”œ«œ"
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
prog1.Visible = False
Me.Caption = sCaption
End Sub
Public Function fixMemberPaid(pMember As String) As String
fixMemberPaid = "UPDATE FILE1_10 SET FILE1_10.DOC_NO = [dbo].[f_last_year_doc](FILE1_10.CODE) WHERE FILE1_10.CODE = " & addvalue(pMember)
End Function

