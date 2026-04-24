VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form ReDem 
   BackColor       =   &H00E0E0E0&
   Caption         =   "≈” ⁄Ê«÷ „‰ „Ê—œÌ‰"
   ClientHeight    =   6480
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   9270
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   675
      Width           =   1500
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   975
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   525
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00DEE7D3&
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   3150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   75
      Width           =   915
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00DEE7D3&
      Caption         =   "ÿ»«⁄…"
      Height          =   390
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      Width           =   915
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00DEE7D3&
      Caption         =   " —«Ã⁄"
      Height          =   390
      Left            =   1200
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00DEE7D3&
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75
      Width           =   1035
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   150
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   6495
      Left            =   225
      TabIndex        =   5
      Top             =   1425
      Width           =   11490
      _cx             =   20267
      _cy             =   11456
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "ReDem.frx":0000
      DataSource      =   "Data2"
      Height          =   315
      Left            =   8310
      TabIndex        =   8
      Top             =   270
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label xclientDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   675
      Width           =   3780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ‹‹‹‹‹‹‹‹Êœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10995
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   1125
      Left            =   4275
      Top             =   75
      Width           =   7515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„ÃÕÊ⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10890
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   315
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   6600
      Left            =   150
      Top             =   1335
      Width           =   11640
   End
End
Attribute VB_Name = "ReDem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datatable As Recordset
Dim BalTable As Recordset
Dim SuppInvTable As Recordset
Dim ItemTable As Recordset
Dim GrTable As Recordset
Dim ClinTable As Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "ÿ·» »÷«⁄… „‰ «·„Ê—œ " & xclientDescA.Caption
    VsItem.ColHidden(4) = True
    VsItem.ColHidden(5) = True
    VsItem.ColHidden(6) = True
    Load PrintGrd
    PrintGrd.Doprint Me.VsItem, 1, -2, cHead1, cHead2, , False, False, 10
    PrintGrd.Show 1
    VsItem.ColHidden(4) = False
    VsItem.ColHidden(5) = False
    VsItem.ColHidden(6) = False
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdOk_Click()
If xCode.Text = "" Then Exit Sub
cString = "SELECT FILE1_10.ITEM ,FILE1_10.FACTCODE , FILE1_10.DESCA , FILE1_10.PRICE ,FILE1_10.GROUP ,  FILE1_10.R1 , FILE1_10.R2  , SUM(FILE1_11.[IN]) AS TIN , SUM(FILE1_11.OUT) AS TOUT" & _
          " FROM FILE1_11 LEFT JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM   " & _
          " Where FILE1_10.ITEM IS NOT NULL  "
If xGroup.Text <> "" Then cString = cString & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
cString = cString & " GROUP BY FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.PACK , FILE1_10.COST1 , FILE1_10.COST2 , FILE1_10.PRICE , FILE1_10.COST4 , FILE1_10.GROUP  , R1 , R2 ,FILE1_10.FACTCODE"
Set datatable = mydb.OpenRecordset(cString)

Me.MousePointer = 11
If datatable.RecordCount > 0 Then
    Fillgrd
Else
    VsItem.Rows = 1
End If
Me.MousePointer = 0
End Sub
Sub Fillgrd()
Dim nPack As Double
Dim nBal As Double
Dim cCode As String
With VsItem
.FixedRows = 1
.FrozenCols = 2
.ExplorerBar = flexExSortShow
.Rows = 1
datatable.MoveFirst
.SubtotalPosition = flexSTBelow
Do While Not datatable.EOF
    cCode = ""
    If TurnValue(datatable.R1, Null, 0) > 0 Then
        nBal = TurnValue(datatable.TIN, Null, 0) - TurnValue(datatable.TOUT, Null, 0)
        If nBal <= datatable.R2 Then
            cStr1 = " SELECT FILE4_10.CODE , FILE4_10.DESCA, FILE7_20.DATE FROM FILE4_10 RIGHT JOIN FILE7_20 ON FILE4_10.CODE = FILE7_20.CODE WHERE ITEM = " & MyParn(datatable.Item) & " ORDER BY DATE "
            Set SuppInvTable = mydb.OpenRecordset(cStr1)
            If SuppInvTable.RecordCount > 0 Then
                SuppInvTable.MoveLast
                cCode = SuppInvTable.CODE
            End If
            If xCode.Text = "" Or xCode.Text = cCode Then
                .AddItem ""
                If SuppInvTable.RecordCount > 0 Then
                    SuppInvTable.MoveLast
'                   .TextMatrix(.Rows - 1, 0) = SuppInvTable.CODE
'                   .TextMatrix(.Rows - 1, 1) = SuppInvTable.DESCA
                End If
                .TextMatrix(.Rows - 1, 0) = datatable.Item
                .TextMatrix(.Rows - 1, 1) = datatable.DESCA
                .TextMatrix(.Rows - 1, 2) = datatable.FACTCODE
                .TextMatrix(.Rows - 1, 3) = datatable.price
                .TextMatrix(.Rows - 1, 4) = TurnValue(datatable.R1, Null, 0)
                .TextMatrix(.Rows - 1, 5) = TurnValue(datatable.R2, Null, 0)
                .TextMatrix(.Rows - 1, 6) = Format(TurnValue(datatable.TIN, Null, 0) - TurnValue(datatable.TOUT, Null, 0), "##0")
                .TextMatrix(.Rows - 1, 7) = Format(Val(.TextMatrix(.Rows - 1, 4)) - Val(.TextMatrix(.Rows - 1, 6)), "#0")
            End If
        End If
    End If
    datatable.MoveNext
Loop
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 6, "#0", , vbRed, , " "
End With
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Set ItemTable = mydb.OpenRecordset("select * from FILE1_10")
Set GrTable = mydb.OpenRecordset("select * from FILE1_50")
Set ClinTable = mydb.OpenRecordset("SELECT * FROM FILE4_10")
Data2.DatabaseName = MdbPath
Data2.RecordSource = "FILE1_50"
Data2.Refresh
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"


With VsItem
.Cols = 8
.Rows = 1
.RowHeight(0) = 1000
.WordWrap = True

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·’‰›"
.TextMatrix(0, 2) = "»«— ﬂÊœ"
.TextMatrix(0, 3) = "”⁄— «·„” Â·ﬂ"
.TextMatrix(0, 4) = "«·—›"
.TextMatrix(0, 5) = "≈⁄«œ… «·ÿ·»"
.TextMatrix(0, 6) = "«·—’Ìœ"

.TextMatrix(0, 7) = "ﬂ„Ì… ≈” ⁄Ê«÷"

.ColWidth(0) = 1300
.ColWidth(1) = 5000
.ColWidth(2) = 2000

.ColWidth(3) = 800
.ColWidth(4) = 800
.ColWidth(5) = 800
.ColWidth(6) = 800
.ColWidth(7) = 800

.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
End With
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As ﬂÊœ , DescA As ≈”„ From File4_10"
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Private Sub xCode_LostFocus()
    ClinTable.FindFirst " CODE = " & MyParn(xCode.Text)
    If Not ClinTable.NoMatch Then Me.xclientDescA.Caption = ClinTable.DESCA
End Sub
