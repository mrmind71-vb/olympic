VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form EMPFRM 
   Caption         =   "»Ì«‰«  «·⁄«„·Ì‰"
   ClientHeight    =   4770
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9435
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
   ScaleHeight     =   4770
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   675
      Width           =   9240
      Begin VB.TextBox xFirst_balance 
         Alignment       =   1  'Right Justify
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
         Left            =   5490
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Tag             =   "D"
         Top             =   2970
         Width           =   1905
      End
      Begin VB.TextBox xDate_end 
         Alignment       =   1  'Right Justify
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
         Left            =   5490
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "D"
         Top             =   2565
         Width           =   1905
      End
      Begin VB.TextBox xDate_Begin 
         Alignment       =   1  'Right Justify
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
         Left            =   5490
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Tag             =   "D"
         Top             =   2160
         Width           =   1905
      End
      Begin VB.TextBox xPhone3 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   315
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   2445
      End
      Begin VB.TextBox xPhone1 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4995
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   2400
      End
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   2700
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   4695
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6075
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox xPhone2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2790
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox xAddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   315
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1260
         Width           =   7080
      End
      Begin VB.CommandButton cmdJob 
         Caption         =   "..."
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   900
         Width           =   420
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   1125
         Top             =   675
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo xJob 
         Height          =   315
         Left            =   4185
         TabIndex        =   2
         Top             =   900
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄Âœ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3060
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «‰ Â«¡ «·⁄„·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2610
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ »œ«Ì… «·⁄„·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2205
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«· ·Ì›Ê‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1845
         Width           =   600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÊŸÌ›…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   570
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   -45
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "EMP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EMP.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EMP.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EMP.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EMP.frx":9A85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdsave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EMP.frx":BFFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4095
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   25
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "EMP.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "EMP.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   26
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "EMP.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "EMP.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   27
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "EMP.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "EMP.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   28
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "EMP.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "EMP.frx":1CE87
      End
   End
End
Attribute VB_Name = "EMPFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim formMode As Byte
Dim CardTable As ADODB.Recordset
Dim oSearch As New Search3
Const LoadMode = 1, DefineMode = 2
Private Sub cmdGroup_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String
sCode = xgroup.BoundText
oFlagfrm.sTable = "JOB_CODES"
oFlagfrm.sCaption = "«‰Ê⁄ «·—Œ’…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
data1.Refresh
xJob.BoundText = sCode
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub

Private Sub EMPfrm_Click()

End Sub
Private Sub cmdJOB_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String
sCode = xJob.BoundText
oFlagfrm.sTable = "JOB_CODES"
oFlagfrm.sCaption = "«·ÊŸ«∆›"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
data1.Refresh
xJob.BoundText = sCode
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM JOB_CODES"
Set xJob.RowSource = data1
xJob.ListField = "Desca"
xJob.BoundColumn = "Code"

openCardTable
myUndo
End Sub
Private Sub CmdAdd_Click()
mydefine
xDescA.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From EMP  Where code = " & MyParn(xCode.Text)
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
    EMPLookupAll Me, oSearch
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = ""
xDescA.Text = ""
xJob.BoundText = ""
xAddress.Text = ""
xPhone1.Text = ""
xPhone2.Text = ""
xPhone3.Text = ""
xDate_Begin.Text = ""
xDate_end.Text = ""
Handlecontrols DefineMode
End Sub
Sub myload()
xCode.Text = CardTable!CODE & ""
xDescA.Text = CardTable!Desca
xAddress.Text = CardTable!Address & ""
xPhone1.Text = CardTable!phone1 & ""
xPhone2.Text = CardTable!phone2 & ""
xPhone3.Text = CardTable!phone3 & ""
xDate_Begin.Text = Format(CardTable!Date_Begin, "dd-mm-yyyy")
xDate_end.Text = Format(CardTable!Date_End, "dd-mm-yyyy")
xJob.BoundText = CardTable!JOB & ""
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDescA.Text))
aInsert = AddFlag(aInsert, "PHONE1", addstring(xPhone1.Text))
aInsert = AddFlag(aInsert, "PHONE2", addstring(xPhone2.Text))
aInsert = AddFlag(aInsert, "PHONE3", addstring(xPhone3.Text))
aInsert = AddFlag(aInsert, "date_Begin", addDate(xDate_Begin.Text))
aInsert = AddFlag(aInsert, "date_End", addDate(xDate_end.Text))
aInsert = AddFlag(aInsert, "[JOB]", addvalue(xJob.BoundText))
On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    xCode.Text = RetZero(Val(Newflag("EMP", "code")), 3)
    aInsert = AddFlag(aInsert, "CODE", addstring(xCode.Text))
    con.Execute addInsert(aInsert, "EMP")
Else
    con.Execute addUpdate(aInsert, "EMP", "code = " & addstring(xCode.Text))
End If
addBox
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myproc()
 xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
oSearch.Hide
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
On Error Resume Next
Unload oSearch
Set oSearch = Nothing
Err.Clear
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 3)
CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xCode.Tag = LoadMode Then
    mydefine
End If
End Sub
Function MYVALID() As Boolean

If Not xJob.MatchedWithList Then
    MsgBox " ”ÃÌ· ‰Ê⁄ «·—Œ’…"
    Exit Function
End If

If Trim(xDescA.Text) = "" Then
    MsgBox "«·«”„ €Ì— „”Ã·"
    Exit Function
End If

If Not IsDate(xDate_Begin.Text) Then
    MsgBox " «—ÌŒ »œ«Ì… «·⁄„· €Ì— „”Ã·"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xCode.Text) <> "" Then
        CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT EMP.* FROM EMP"
cString = cString & " ORDER BY EMP.[CODE]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub xDate_End_GotFocus()
myGotFocus xDate_end
End Sub
Private Sub xDate_End_LostFocus()
myLostFocus xDate_end
myValidDate xDate_end
End Sub
Private Sub xDate_Begin_GotFocus()
myGotFocus xDate_Begin
End Sub
Private Sub xDate_Begin_LostFocus()
myLostFocus xDate_Begin
myValidDate xDate_Begin
End Sub

Private Sub xPhone3_GotFocus()
myGotFocus xPhone3
End Sub
Private Sub xPhone3_LostFocus()
myLostFocus xPhone3
End Sub
Private Sub xPhone1_GotFocus()
myGotFocus xPhone1
End Sub
Private Sub xPhone1_LostFocus()
myLostFocus xPhone1
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDescA
End Sub
Private Sub xDescA_LostFocus()
myLostFocus xDescA
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xPhone2_GotFocus()
myGotFocus xPhone2
End Sub
Private Sub xPhone2_LostFocus()
myLostFocus xPhone2
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xJOB_GotFocus()
myGotFocus xJob
End Sub
Private Sub xJOB_LostFocus()
myLostFocus xJob
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub
Private Function addBox()
Dim aRet As Variant, aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDescA.Text))
aInsert = AddFlag(aInsert, "f_date", addDate(xDate_Begin.Text))
aInsert = AddFlag(aInsert, "[f_Bal]", Val(xFirst_balance.Text))
aRet = GetField("select code from file0_50 where code = " & MyParn(xCode.Text))
If IsEmpty(aRet) Then
    aInsert = AddFlag(aInsert, "CODE", addstring(xCode.Text))
    con.Execute addInsert(aInsert, "FILE0_50")
Else
    con.Execute addUpdate(aInsert, "FILE0_50", "code = " & addstring(xCode.Text))
End If
End Function
Private Sub xFirst_Balance_GotFocus()
myGotFocus xFirst_balance
End Sub
Private Sub xFirst_Balance_LostFocus()
myLostFocus xFirst_balance
End Sub

