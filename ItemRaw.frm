VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form itemRawfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «·„Ê«œ «·Œ«„"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   7710
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3960
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2745
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6435
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "Œ«„… ÃœÌœ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5175
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton CmdDelInv 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› «·Œ«„…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   8745
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4275
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2580
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         MaxLength       =   150
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   6765
      End
      Begin MSDataListLib.DataCombo XGROUP 
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Top             =   900
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7020
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "„”·”· «·’‰› :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7035
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·’‰› :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7020
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   1020
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4110
      Left            =   135
      TabIndex        =   7
      Top             =   1935
      Width           =   10410
      _cx             =   18362
      _cy             =   7250
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   135
      TabIndex        =   8
      Top             =   6030
      Width           =   1995
      Begin VB.CommandButton cmdFirst 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "«·«Ê·"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "«·«ŒÌ—"
         Top             =   180
         Width           =   465
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   45
      Top             =   405
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
End
Attribute VB_Name = "itemrawfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim oSearchColor As New Search3
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(3, 1)
aInsert(0, 0) = "code"
aInsert(0, 1) = addstring(xCode.Text)

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(xdesca.Text)

aInsert(2, 0) = "[GROUP]"
aInsert(2, 1) = addstring(xGroup.BoundText)

aInsert(3, 0) = "[isRaw]"
aInsert(3, 1) = "1"

'On Error GoTo myerror
con.BeginTrans
If xCode.Tag = DefineMode Then
'    xCode.Text = Val(Newflag("FILE1_50", "CODE"))
'    aInsert(0, 1) = addstring(xCode.Text)
    con.Execute CreateInsert(aInsert, "FILE1_50")
Else
    con.Execute CreateUpdate(aInsert, "FILE1_50", " where CODE = " & MyParn(xCode.Text))
End If
'If xCode.Tag = DefineMode Then myreplaceGrd
myReplacegrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "code = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    If Not CardTable.EOF Then myload
    Unload Search3
End If
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearchColor.grid1.TextMatrix(oSearchColor.grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    grid1.EditText = oSearchColor.grid1.TextMatrix(oSearchColor.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 0) = oSearchColor.grid1.TextMatrix(oSearchColor.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 1) = oSearchColor.grid1.TextMatrix(oSearchColor.grid1.Row, 1)
    
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = ""
        grid1.AddItem ""
        grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.TextMatrix(grid1.Row, 3) = ""
        grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Select grid1.Row + 1, 1
    End If
'    CalcTotals
End If

End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE1_10 where [Group] = " & MyParn(xCode.Text)
    con.Execute "Delete From FILE1_50 where code = " & MyParn(xCode.Text)
    con.CommitTrans
    CardTable.Requery
    On Error GoTo myerror2
    If CardTable.BOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "code < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
myerror2:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE1_50.Code, File1_50.Desca " & _
                  " FROM FILE1_50 where israw = 1"
Generalarray(2) = " Order by FILE1_50.Code"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·„Ã„Ê⁄…"
listarray(0, 1) = "%%Desca%%"

GrdArray(0, 0) = "ﬂÊœ «·„Ã„Ê⁄…"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ã„Ê⁄…"
GrdArray(1, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "„Ã„Ê⁄«  «·Œ«„« "
Search3.Show 1
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
Private Sub CmdNewInv_Click()
'xCode.Text = Newflag("FILE1_50", "CODE")
mydefine
End Sub
Private Sub cmdSave_Click()
mysave
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
    Exit Sub
End If
CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT CODE,DESCA,[GROUP] FROM FILE1_50 WHERE ISRAW = 1 ORDER BY CODE", con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE1_50G WHERE ISRAW = 1"

Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
    Fixgrd
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
SetKbLayout Lang_AR
closeCon con
On Error Resume Next
Unload oSearchColor
Err.Clear
Set itemrawfrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    If Trim(grid1.TextMatrix(Row, Col)) <> "" Then
        grid1.TextMatrix(Row, 1) = GetDesca("SELECT DESCA FROM COLOR WHERE CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 0)))
    End If
End If

If xCode.Tag = DefineMode Then
    mysave
    Exit Sub
End If
If Not myValidRow(Row) Then
    myloadgrd
    Exit Sub
End If
With grid1
Dim aInsert(7, 1)

aInsert(0, 0) = "SubItem"
aInsert(0, 1) = addstring(grid1.TextMatrix(Row, 0))
        
aInsert(1, 0) = "item"
aInsert(1, 1) = addstring(xCode.Text & "-" & grid1.TextMatrix(Row, 0))

aInsert(2, 0) = "DESCASUB"
aInsert(2, 1) = addstring(grid1.TextMatrix(Row, 1))
        
aInsert(3, 0) = "[GROUP]"
aInsert(3, 1) = addstring(xCode.Text)
                
aInsert(4, 0) = "Cost"
aInsert(4, 1) = Val(.TextMatrix(Row, 2))
                
aInsert(5, 0) = "Fact_No"
aInsert(5, 1) = addstring(.TextMatrix(Row, 3))
                
aInsert(6, 0) = "isRaw"
aInsert(6, 1) = "1"
                
aInsert(7, 0) = "Desca"
aInsert(7, 1) = addstring(xdesca.Text & " " & grid1.TextMatrix(Row, 1))
                
On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(Row, .Cols - 1) = "" Then
    con.Execute CreateInsert(aInsert, "FILE1_10")
    grid1.TextMatrix(Row, grid1.Cols - 1) = grid1.TextMatrix(Row, 0)
Else
    con.Execute CreateUpdate(aInsert, "FILE1_10", " where ITEM = " & MyParn(xCode.Text & "-" & grid1.TextMatrix(Row, 0)))
End If
con.CommitTrans
FixLast
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myloadgrd
End Sub
Private Sub grid1_EnterCell()
With grid1
'If xCode.Enabled Then
'    .Editable = flexEDNone
If .Col = 0 And .TextMatrix(.Row, .Cols - 1) <> "" Then
    .Editable = flexEDNone
ElseIf .Col <> 0 And .TextMatrix(.Row, .Cols - 1) = "" Then
    .Editable = flexEDKbd
Else
    .Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 2 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from FILE1_10 where ITEM = " & MyParn(xCode.Text & "-" & grid1.TextMatrix(grid1.Row, 0)), nDelete
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
        grid1_EnterCell
        'FixLast
    End If
End If
If KeyCode = 112 Then
    If grid1.Col = 0 And grid1.Row <> 0 Then colorLookup
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    FixLast
End If
End Sub
Private Function MYVALID() As Boolean
If Trim(xCode.Text) = "" Then
    MsgBox "ﬂÊœ «·„Ã„Ê⁄… ·„ Ì”Ã·"
    Exit Function
End If
If xCode.Tag = DefineMode Then
    For I = 0 To grid1.Rows - 2
        If Not myValidRow(I) Then Exit Function
    Next
End If
MYVALID = True
End Function
Private Function myValidRow(ByVal nRow) As Boolean
With grid1
    If Trim(.TextMatrix(nRow, 0)) = "" Then
        MsgBox "ﬂÊœ «·Œ«„… ·„ Ì”Ã·"
        Exit Function
    End If

    If .TextMatrix(nRow, 1) = "" Then
       '.Select i, 0, i, grid1.Cols - 1
        'MsgBox "«”„ «·Œ«„… ·„ Ì”Ã·"
'        Exit Function
    End If
End With
myValidRow = True
End Function
Private Sub myload()
xCode.Text = CardTable!code
xdesca.Text = CardTable!desca
xGroup.BoundText = CardTable!Group & ""
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE1_10.subItem,file1_10.Descasub,FILE1_10.COST,FACT_NO,FILE1_10.subItem as subitem2" & _
               " FROM FILE1_10 " & _
               " where [group] = " & MyParn(xCode.Text) & " Order by SubItem"

    DATA10.RecordSource = cString
    DATA10.Refresh
    grid1.AddItem ""
    'grid1.TextMatrix(grid1.Rows - 1, 0) = Val(grid1.TextMatrix(grid1.Rows - 2, 0)) + 1
End With
Fixgrd
End Sub
Private Sub mydefine()
xdesca.Text = ""
xGroup.BoundText = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bedit)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xCode.Enabled = (nMode = DefineMode)
xCode.Tag = nMode
cmdSave.Enabled = bedit
CmdDelInv.Enabled = bedit
End Sub
Private Sub Fixgrd()
With grid1
    .Cols = 5
     .FormatString = "ﬂÊœ «·’‰›|" & "≈”„ «·’‰›|" & "«· ﬂ·›…|" & "—ﬁ„ «·„’‰⁄|" & "ﬂÊœ «·’‰›"
    .ColWidth(0) = 1200
    .ColWidth(1) = 5000
    .ColWidth(2) = 1000
    .ColWidth(3) = 2800
    .ColHidden(.Cols - 1) = True
    For I = 1 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If .EditText = "" And Col = 0 Then
        MsgBox "»Ì«‰«  „ÿ·Ê»…"
        Cancel = True
   ElseIf .EditText <> "" And Col = 0 Then
        For I = 1 To grid1.Rows - 1
            If I <> Row Then
                If Trim(.TextMatrix(I, 0)) = Trim(.EditText) Then
                    MsgBox "«·ﬂÊœ „ÊÃÊœ „‰ ﬁ»·"
                    Cancel = True
                End If
            End If
        Next
    End If
End With
End Sub
Private Sub FixLast()
'With grid1
'    If grid1.Rows = 2 Then
'        .TextMatrix(.Rows - 1, 0) = "1"
'    Else
'        .TextMatrix(.Rows - 1, 0) = Val(.TextMatrix(.Rows - 2, 0)) + 1
'    End If
'End With
End Sub
Private Sub myReplacegrd()
Dim aInsert(7, 1)
With grid1
For I = 1 To grid1.Rows - 2
    aInsert(0, 0) = "SubItem"
    aInsert(0, 1) = addstring(grid1.TextMatrix(I, 0))
            
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(xCode.Text & "-" & grid1.TextMatrix(I, 0))
    
    aInsert(2, 0) = "Descasub"
    aInsert(2, 1) = addstring(grid1.TextMatrix(I, 1))
            
    aInsert(3, 0) = "[GROUP]"
    aInsert(3, 1) = addstring(xCode.Text)
                    
    aInsert(4, 0) = "Cost"
    aInsert(4, 1) = Val(.TextMatrix(I, 2))
                        
    aInsert(5, 0) = "Fact_No"
    aInsert(5, 1) = addstring(.TextMatrix(I, 3))
                        
    aInsert(6, 0) = "isRaw"
    aInsert(6, 1) = "1"
                        
    aInsert(7, 0) = "Desca"
    aInsert(7, 1) = addstring(xdesca.Text & " " & grid1.TextMatrix(I, 1))
                        
    If grid1.TextMatrix(I, .Cols - 1) = "" Then
        con.Execute CreateInsert(aInsert, "FILE1_10")
        grid1.TextMatrix(Row, grid1.Cols - 1) = grid1.TextMatrix(I, 0)
    Else
        con.Execute CreateUpdate(aInsert, "FILE1_10", " where ITEM = " & MyParn(xCode.Text & "-" & grid1.TextMatrix(I, 0)))
    End If
Next
FixLast
End With
End Sub
Private Function myvalidGrd() As Boolean
End Function
Private Sub mysave(Optional igMsg As Boolean = False)
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Requery
If Not igMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Find "Code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
If Not CardTable.EOF Then myload
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.Text)
xCode.BackColor = &HC0FFFF
SetKbLayout Lang_EN
End Sub
Private Sub xCode_LostFocus()
SetKbLayout Lang_AR
xCode.Text = UCase(xCode.Text)
xCode.BackColor = &H80000005
If Trim(xCode.Text) = "" Then Exit Sub
xCode.Text = xCode.Text
CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub xDescA_GotFocus()
xdesca.SelStart = 0
xdesca.SelLength = Len(xdesca.Text)
xdesca.BackColor = &HC0FFFF
End Sub
Private Sub xGroup_GotFocus()
xGroup.BackColor = &HC0FFFF
End Sub
Private Sub xDescA_LostFocus()
xdesca.BackColor = &H80000005
End Sub
Private Sub xGroup_LostFocus()
xGroup.BackColor = &H80000005
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub colorLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From color"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ-«·»Ì«‰"
listarray(0, 1) = "(%%CODE%% OR %%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«··Ê‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchColor.Caption = "≈” ⁄·«„ «·«·Ê«‰"
oSearchColor.Show 1
End Sub
