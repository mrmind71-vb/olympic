VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form clients 
   Caption         =   "»Ì«‰«  «·⁄„·«¡"
   ClientHeight    =   4575
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
   ScaleHeight     =   4575
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   630
      Width           =   9240
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
         Left            =   2700
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   1545
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
         Left            =   5850
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   1545
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
         MaxLength       =   50
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
         MaxLength       =   6
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
         Left            =   4275
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   1545
      End
      Begin VB.TextBox xAddress 
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
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   330
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   1035
         Top             =   135
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
      Begin MSDataListLib.DataCombo xgroup 
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«· ·Ì›Ê‰ :"
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
         TabIndex        =   27
         Top             =   1845
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ :"
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
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„ :"
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
         TabIndex        =   25
         Top             =   615
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰ :"
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
         TabIndex        =   24
         Top             =   1260
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Top             =   945
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   -45
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "Clients.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "Clients.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "Clients.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "Clients.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "Clients.frx":9A85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdSave 
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
         Picture         =   "Clients.frx":BFFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "—’Ìœ «›  «ÕÌ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2880
      Width           =   9240
      Begin VB.TextBox xMinus 
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
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1545
      End
      Begin VB.TextBox xPlus 
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
         Left            =   5850
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   1545
      End
      Begin VB.CheckBox xcash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄„Ì· ‰ﬁœÌ"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   630
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox xf_Date 
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
         Left            =   5850
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Tag             =   "date"
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„œÌ‰ :"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "œ«∆‰ :"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ √Ê · «·„œ… :"
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
         Left            =   7485
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   615
         Width           =   1320
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3870
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   29
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
         Picture         =   "Clients.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   30
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
         Picture         =   "Clients.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   31
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
         Picture         =   "Clients.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   32
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
         Picture         =   "Clients.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":1CE87
      End
   End
End
Attribute VB_Name = "clients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Dim con As New ADODB.Connection
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim sGroup As String
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdGroup_Click()
Dim oFlagfrm As New flag_mainfrm, cValue As String
cValue = xgroup.BoundText
oFlagfrm.sTable = "FILE3_50"
oFlagfrm.sCaption = "„Ã„Ê⁄… «·⁄„·«¡"
oFlagfrm.nZero = 2
oFlagfrm.bedit = True
oFlagfrm.Show 1
data1.Refresh
xgroup.BoundText = cValue
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
sGroup = myDef("file3_50", "CODE")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
sGroup = myDef("file3_50", "CODE")
cTableName = "FILE3_10"
cGroupname = "FILE3_50"
Me.Caption = "»Ì«‰«  «·⁄„·«¡"

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE3_50"
Set xgroup.RowSource = data1
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

openCardTable
myUndo
End Sub
Private Sub CmdAdd_Click()
mydefine
xDescA.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From FILE3_10  Where code = " & MyParn(xCode.Text)
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
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·⁄„Ì· »‰Ã«Õ"
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
    CardLookup
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
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE3_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "≈” ⁄·«„ «·⁄„·«¡"
Search3.Show 1
End Sub
Sub mydefine()
xCode.Text = RetZero(Newflag(cTableName, "code"), 6)
xDescA.Text = ""
xgroup.BoundText = sGroup
xAddress.Text = ""
xPhone1.Text = ""
xPhone2.Text = ""
xPhone3.Text = ""
xf_Date.Text = ""
xPlus.Text = ""
xMinus.Text = ""
xcash.Value = 1
Handlecontrols DefineMode
End Sub
Sub myload()
xCode.Text = CardTable!code & ""
xDescA.Text = CardTable!desca
xAddress.Text = CardTable!Address & ""
xPhone1.Text = CardTable!phone1 & ""
xPhone2.Text = CardTable!phone2 & ""
xPhone3.Text = CardTable!phone3 & ""
xf_Date.Text = Format(CardTable!F_DATE, "YYYY-MM-DD")
xPlus.Text = IIf(Val(CardTable!f_Bal1 & "") > 0, Val(CardTable!f_Bal1 & ""), "")
xMinus.Text = IIf(Val(CardTable!f_Bal1 & "") < 0, Abs(Val(CardTable!f_Bal1 & "")), "")
xcash.Value = IIf(CardTable!cash, 1, 0)
xgroup.BoundText = CardTable!Group & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDescA.Text))
aInsert = AddFlag(aInsert, "ADDRESS", addstring(xAddress.Text))
aInsert = AddFlag(aInsert, "PHONE1", addstring(xPhone1.Text))
aInsert = AddFlag(aInsert, "PHONE2", addstring(xPhone2.Text))
aInsert = AddFlag(aInsert, "PHONE3", addstring(xPhone3.Text))
aInsert = AddFlag(aInsert, "F_DATE", addDate(xf_Date.Text))
aInsert = AddFlag(aInsert, "F_BAL1", IIf(Val(xPlus.Text) > 0, Val(xPlus.Text), -1 * Val(xMinus.Text)))
aInsert = AddFlag(aInsert, "[GROUP]", addstring(xgroup.BoundText))
aInsert = AddFlag(aInsert, "[CASH]", xcash.Value)
On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    xCode.Text = RetZero(Val(Newflag("FILE3_10", "code")))
    aInsert = AddFlag(aInsert, "CODE", addstring(xCode.Text))
    con.Execute addInsert(aInsert, "FILE3_10")
Else
    con.Execute addUpdate(aInsert, "FILE3_10", "code = " & addstring(xCode.Text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
 xCode.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Search3.Hide
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
If MsgBox(cMsgExit, vbOKCancel + vbDefaultButton1) = vbOK Then
    cmdSave_Click
End If
CardTable.Close
Set CardTable = Nothing
closeCon con
On Error Resume Next
Unload Search3
Set Search3 = Nothing
Err.Clear
End Sub

Private Sub xCode_LostFocus()
xCode.BackColor = &H80000005
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If xCode.Text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xgroup.BoundText = "" Then
    MsgBox " ”ÃÌ· «·„Ã„Ê⁄… "
    Exit Function
End If

If xDescA.Text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If Not IsDate(xf_Date.Text) Then
    MsgBox " «—ÌŒ «Ê· «·„œ… ÷—Ê—Ì"
    Exit Function
End If

Dim aRet As Variant
aRet = GetField("Select code from file3_10 where desca = " & MyParn(xDescA.Text) & " and code <> " & MyParn(xCode.Text))
If Not IsEmpty(aRet) Then
    MsgBox "«·«”„ „ÊÃÊœ „‰ ﬁ»· ›Ï «·ﬂÊœ " & aRet
    Exit Function
End If
MYVALID = True
End Function
Private Sub xDisc2_LostFocus()
xDisc2.BackColor = &H80000005
End Sub
Private Sub xDisc1_LostFocus()
xDisc1.BackColor = &H80000005
End Sub
Private Sub xBalancePlus_LostFocus()
xBalancePlus.BackColor = &H80000005
End Sub
Private Sub xF_DATE_LostFocus()
xf_Date.BackColor = &H80000005
End Sub
Private Sub xBalanceMinus_LostFocus()
xBalanceMinus.BackColor = &H80000005
End Sub
Private Sub xAddress_LostFocus()
xAddress.BackColor = &H80000005
End Sub
Private Sub xManager_LostFocus()
xManager.BackColor = &H80000005
End Sub
Private Sub xFAx_LostFocus()
xFAx.BackColor = &H80000005
End Sub

Private Sub xPhone2_LostFocus()
xPhone2.BackColor = &H80000005
End Sub
Private Sub xfileNo_LostFocus()
xFileNo.BackColor = &H80000005
End Sub
Private Sub xDesca_LostFocus()
xDescA.BackColor = &H80000005
End Sub
Private Sub xPhone1_LostFocus()
xPhone1.BackColor = &H80000005
End Sub
Private Sub xPhone3_LostFocus()
xPhone3.BackColor = &H80000005
End Sub
Private Sub xGroup_LostFocus()
xgroup.BackColor = &H80000005
End Sub
Private Sub xDisc2_GotFocus()
xDisc2.SelStart = 0
xDisc2.SelLength = Len(xDisc2.Text)
xDisc2.BackColor = &HC0FFFF
End Sub
Private Sub xDisc1_GotFocus()
xDisc1.SelStart = 0
xDisc1.SelLength = Len(xDisc1.Text)
xDisc1.BackColor = &HC0FFFF
End Sub
Private Sub xBalancePlus_GotFocus()
xBalancePlus.SelStart = 0
xBalancePlus.SelLength = Len(xBalancePlus.Text)
xBalancePlus.BackColor = &HC0FFFF
End Sub
Private Sub xF_DATE_GotFocus()
xf_Date.SelStart = 0
xf_Date.SelLength = Len(xf_Date.Text)
xf_Date.BackColor = &HC0FFFF
End Sub
Private Sub xBalanceMinus_GotFocus()
xBalanceMinus.SelStart = 0
xBalanceMinus.SelLength = Len(xBalanceMinus.Text)
xBalanceMinus.BackColor = &HC0FFFF
End Sub
Private Sub xAddress_GotFocus()
xAddress.SelStart = 0
xAddress.SelLength = Len(xAddress.Text)
xAddress.BackColor = &HC0FFFF
End Sub
Private Sub xManager_GotFocus()
xManager.SelStart = 0
xManager.SelLength = Len(xManager.Text)
xManager.BackColor = &HC0FFFF
End Sub
Private Sub xFAx_GotFocus()
xFAx.SelStart = 0
xFAx.SelLength = Len(xFAx.Text)
xFAx.BackColor = &HC0FFFF
End Sub
Private Sub xPhone2_GotFocus()
xPhone2.SelStart = 0
xPhone2.SelLength = Len(xPhone2.Text)
xPhone2.BackColor = &HC0FFFF
End Sub
Private Sub xfileNo_GotFocus()
xFileNo.SelStart = 0
xFileNo.SelLength = Len(xFileNo.Text)
xFileNo.BackColor = &HC0FFFF
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.Text)
xCode.BackColor = &HC0FFFF
End Sub
Private Sub xDescA_GotFocus()
xDescA.SelStart = 0
xDescA.SelLength = Len(xDescA.Text)
xDescA.BackColor = &HC0FFFF
End Sub
Private Sub xPhone1_GotFocus()
xPhone1.SelStart = 0
xPhone1.SelLength = Len(xPhone1.Text)
xPhone1.BackColor = &HC0FFFF
End Sub
Private Sub xPhone3_GotFocus()
xPhone3.SelStart = 0
xPhone3.SelLength = Len(xPhone3.Text)
xPhone3.BackColor = &HC0FFFF
End Sub
Private Sub xGroup_GotFocus()
xgroup.BackColor = &HC0FFFF
End Sub
Private Sub xf_Date_Validate(Cancel As Boolean)
With xf_Date
If (Not IsDate(.Text)) And Trim(.Text) <> "" Then .Text = ""
.Text = Format(.Text, "YYYY-MM-DD")
End With
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub

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
cString = "SELECT FILE3_10.* FROM FILE3_10"
cString = cString & " ORDER BY FILE3_10.[CODE]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub xPlus_LostFocus()
If IsNumeric(xPlus.Text) And Val(xPlus.Text) <> 0 Then xMinus.Text = ""
End Sub
Private Sub xMinus_LostFocus()
If IsNumeric(xMinus.Text) And Val(xMinus.Text) <> 0 Then xPlus.Text = ""
End Sub

