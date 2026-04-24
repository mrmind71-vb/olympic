VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form itemsfrm_org 
   Caption         =   "»‰Êœ «·‰‘«ÿ"
   ClientHeight    =   3615
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
   ScaleHeight     =   3615
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   675
      Width           =   9285
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         Height          =   330
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   990
         Width           =   330
      End
      Begin VB.TextBox xValue2 
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
         TabIndex        =   4
         Tag             =   "S"
         Top             =   1755
         Width           =   1905
      End
      Begin VB.TextBox xValue 
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
         TabIndex        =   3
         Tag             =   "S"
         Top             =   1350
         Width           =   1905
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
         Height          =   360
         Left            =   270
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   7125
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
         Height          =   375
         Left            =   6075
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   225
         Top             =   270
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
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   330
         Left            =   2835
         TabIndex        =   2
         Top             =   990
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄…"
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
         TabIndex        =   24
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«»‰«¡ €Ì— «·„Â‰œ”Ì‰"
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
         TabIndex        =   22
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«»‰«¡ „Â‰œ”Ì‰"
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
         TabIndex        =   21
         Top             =   1395
         Width           =   1050
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰"
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
         TabIndex        =   13
         Top             =   570
         Width           =   420
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   -45
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "items2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "items2.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "items2.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "items2.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "items2.frx":9A85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "items2.frx":BFFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2880
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   16
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
         Picture         =   "items2.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items2.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   17
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
         Picture         =   "items2.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items2.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   18
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
         Picture         =   "items2.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items2.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   19
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
         Picture         =   "items2.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items2.frx":1CE87
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
   Begin VB.Label xRecordNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2970
      Width           =   6045
   End
End
Attribute VB_Name = "itemsfrm_org"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim formMode As Byte
Dim cFilter As String
Dim CardTable As ADODB.Recordset
Dim bEdit As Boolean, bEditRecord As Boolean
Dim oSearch As New Search3
Const LoadMode = 1, DefineMode = 2
Private Sub cmdGroup_Click()
Dim oFlagfrm As New flag_mainfrm, sBoundText As String
sBoundText = xgroup.BoundText
oFlagfrm.sTable = "FILE1_30G"
oFlagfrm.sCaption = "«·„Ã„Ê⁄…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
Set Data2.Recordset = myRecordSet("select * from FILE1_30G", con)
xgroup.BoundText = sBoundText
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con

'data1.ConnectionString = strCon
'data2.RecordSource = "SELECT * FROM FILE1_30G ORDER BY CODE"
Set Data2.Recordset = myRecordSet("SELECT * FROM FILE1_30G ORDER BY CODE", con)
Set xgroup.RowSource = Data2
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

bEdit = True

openCardTable
myUndo
End Sub
Private Sub CmdAdd_Click()
myDefine
xDescA.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myError
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From FILE1_30  Where code = " & MyParn(xCode.Text)
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        myDefine
    End If
End If
Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
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
Items_SportLookupAll Me, oSearch, cFilter
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
bEditRecord = bEdit
cmdAdd.Enabled = (nMode = LoadMode) And bEditRecord
CmdDel.Enabled = (nMode = LoadMode) And bEditRecord
cmdsave.Enabled = bEditRecord
CmdInform.Enabled = (nMode = LoadMode)
CmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
CmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
CmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
CmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub myDefine()
xCode.Text = ""
xDescA.Text = ""
xgroup.BoundText = ""
xValue.Text = ""
xValue2.Text = ""
xRecordNo.Caption = "«÷«›… ”Ã· ÃœÌœ " & "(" & CardTable.RecordCount & ")"
Handlecontrols DefineMode
End Sub
Sub myload()
xCode.Text = CardTable!code & ""
xDescA.Text = CardTable!Desca
xgroup.BoundText = CardTable!Group & ""
xValue.Text = Myvalue(CardTable!Value)
xValue2.Text = Myvalue(CardTable!Value2)
xRecordNo.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
Handlecontrols LoadMode
End Sub
Private Function MyReplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDescA.Text))
aInsert = AddFlag(aInsert, "VALUE", addvalue(xValue.Text))
aInsert = AddFlag(aInsert, "VALUE2", Val(xValue2.Text))
aInsert = AddFlag(aInsert, "[GROUP]", addvalue(xgroup.BoundText))
On Error GoTo myError
con.BeginTrans
If xCode.Enabled Then
    xCode.Text = Val(Newflag("FILE1_30", "code"))
    aInsert = AddFlag(aInsert, "CODE", addstring(xCode.Text))
    con.Execute addInsert(aInsert, "FILE1_30")
Else
    con.Execute addUpdate(aInsert, "FILE1_30", "code = " & addstring(xCode.Text))
End If
con.CommitTrans
MyReplace = True
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
 xCode.Text = oSearch.Grid1.TextMatrix(oSearch.Grid1.Row, 0)
oSearch.Hide
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
If MsgBox(ArbString("»«·Œ—ÊÃ ”Ì „ ” ›ﬁœ ﬂ· «· ⁄œÌ·«  ⁄·Ì «·”Ã· ?"), vbOKCancel + vbDefaultButton2) <> vbOK Then
    Cancel = True
    Exit Sub
End If
CardTable.Close
Set CardTable = Nothing
closeCon con
Set FILE1_30frm = Nothing
On Error Resume Next
Unload oSearch
Set oSearch = Nothing
Err.Clear
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidInt(xCode.Text) Then Exit Sub
CardTable.Find "CODE = " & xCode.Text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xCode.Tag = LoadMode Then
    myDefine
End If
End Sub
Function MYVALID() As Boolean

If Not xgroup.MatchedWithList Then
    MsgBox " ”ÃÌ· «·„Ã„Ê⁄…"
    Exit Function
End If

If Trim(xDescA.Text) = "" Then
    MsgBox "«·»Ì«‰ €Ì— „”Ã·"
    Exit Function
End If

If Val(xValue.Text) < 0 Then
    MsgBox "«·ﬁÌ„… €Ì— ’ÕÌÕ…"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    myDefine
Else
    If ValidInt(xCode.Text) Then
        CardTable.Find "CODE = " & xCode.Text, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cFilter = ""
cString = "SELECT FILE1_30.* FROM FILE1_30"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY FILE1_30.[CODE]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub

Private Sub xAge2_GotFocus()
myGotFocus xAge2
End Sub
Private Sub xAge2_LostFocus()
myLostFocus xAge2
End Sub
Private Sub xAge1_GotFocus()
myGotFocus xAge1
End Sub
Private Sub xAge1_LostFocus()
myLostFocus xAge1
End Sub
Private Sub xValue_GotFocus()
myGotFocus xValue
End Sub
Private Sub xValue_LostFocus()
myLostFocus xValue
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDescA
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDescA
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xRelation_GotFocus()
myGotFocus xRelation
End Sub
Private Sub xRelation_LostFocus()
myLostFocus xRelation
If Not xRelation.MatchedWithList Then xRelation.BoundText = ""
End Sub
