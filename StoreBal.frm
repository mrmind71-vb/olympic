VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form StoreBal 
   Caption         =   "√—’œ… «·«’‰«› „Ê“⁄… ⁄·Ï «·„Œ«“‰"
   ClientHeight    =   10950
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   90
      Width           =   10410
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6495
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   4875
         TabIndex        =   7
         Top             =   1320
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   4875
         TabIndex        =   8
         Top             =   960
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   4875
         TabIndex        =   9
         Top             =   600
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3675
         TabIndex        =   14
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8430
         TabIndex        =   13
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„"
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
         Index           =   3
         Left            =   8430
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8430
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1005
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄…"
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
         Index           =   1
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1350
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "⁄—÷"
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
      Left            =   2535
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1545
      Width           =   1275
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1305
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1545
      Width           =   1230
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1545
      Width           =   1275
   End
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   8535
      Left            =   90
      TabIndex        =   0
      Top             =   1980
      Width           =   15015
      _cx             =   26485
      _cy             =   15055
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -225
      Top             =   375
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Left            =   -90
      Top             =   525
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -90
      Top             =   675
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -90
      Top             =   75
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
Attribute VB_Name = "StoreBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub Cmd_Print_Click()
Dim cHead1 As String
Dim cHead2 As String
cHead1 = "»Ì«‰ «’‰«› ·Â« «” ⁄Ê«÷ „‰ «·„Ê—œÌ‰ "
Load PrintGrd
PrintGrd.doprint Me.VsItem, 1, -2, cHead1, cHead2, , False, False, 9
PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Sub myload()
Dim nPack As Double
Dim nBal As Double
Dim cCode As String

Dim datatable As New ADODB.Recordset
Dim StoreBalTable As New ADODB.Recordset

cString = "SELECT FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.PRICE ,FILE1_10.[GROUP] ,  FILE1_10.R1 , FILE1_10.R2  , SUM(FILE1_11.[IN]) AS TIN , SUM(FILE1_11.OUT) AS TOUT" & _
          " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM   "

If xGroup.Text <> "" Then
    cwhere = cwhere & " file1_10.[GROUP] = " & MyParn(xGroup.BoundText)
End If

cString = cString & TurnWhere(cwhere) & cwhere & " GROUP BY FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.PACK , FILE1_10.COST1 , FILE1_10.COST2 , FILE1_10.PRICE , FILE1_10.COST4 , FILE1_10.[GROUP]  , R1 , R2 "

datatable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

cString = "SELECT FILE1_11.ITEM , FILE1_11.STORE , SUM(FILE1_11.[IN]) AS TIN , SUM(FILE1_11.OUT) AS TOUT" & _
          " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM   "
cString = cString & " GROUP BY FILE1_11.ITEM , FILE1_11.STORE "
StoreBalTable.Open cString, con, adOpenStatic, adLockReadOnly

With VsItem
.FixedRows = 1
.FrozenCols = 4
.ExplorerBar = flexExSortShow
.Rows = 2
.SubtotalPosition = flexSTBelow
Do Until datatable.EOF
    nBal = Val(datatable!TIN & "") - Val(datatable!TOUT & "")
    If nBal <> 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = datatable!Item
        .TextMatrix(.Rows - 1, 1) = datatable!Desca & ""
        .TextMatrix(.Rows - 1, 2) = datatable!price & ""
        .TextMatrix(.Rows - 1, 3) = Val(datatable!R1 & "")
        .TextMatrix(.Rows - 1, 4) = Val(datatable!R2 & "")
        .TextMatrix(.Rows - 1, 5) = Format(Val(datatable!TIN & "") - Val(datatable!TOUT & ""), "##0")
        For nCol = 6 To .Cols - 1
            StoreBalTable.Filter = " ITEM = " & MyParn(datatable!Item) & " AND STORE = " & MyParn(.TextMatrix(1, nCol))
            If Not (StoreBalTable.EOF And StoreBalTable.BOF) Then .TextMatrix(.Rows - 1, nCol) = Format(Val(StoreBalTable!TIN & "") - Val(StoreBalTable!TOUT & ""), "##0")
        Next
    End If
    datatable.MoveNext
Loop
datatable.Close
StoreBalTable.Close
End With
End Sub

Private Sub CmdOk_Click()
myload
End Sub

Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Dim storeTable As New ADODB.Recordset
storeTable.Open "Stores", con, adOpenKeyset, adLockReadOnly, adCmdTable

DATA1.ConnectionString = strCon
DATA1.RecordSource = "SELECT * FROM FILE1_50"
Set xGroup.RowSource = DATA1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

With VsItem
.Cols = 6
.Rows = 2
.RowHeight(0) = 1000
.RowHeight(1) = 0
.WordWrap = True

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·’‰›"
.TextMatrix(0, 2) = "”⁄— «·„” Â·ﬂ"
.TextMatrix(0, 3) = "«·—›"
.TextMatrix(0, 4) = "≈⁄«œ… «·ÿ·»"
.TextMatrix(0, 5) = "«·—’Ìœ"


.ColWidth(0) = 800
.ColWidth(1) = 3000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000

.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble

Do Until storeTable.EOF
    .Cols = .Cols + 1
    .ColWidth(.Cols - 1) = 1000
    .ColDataType(.Cols - 1) = flexDTDouble
    .TextMatrix(0, .Cols - 1) = storeTable!Desca
    .TextMatrix(1, .Cols - 1) = storeTable!Code
    storeTable.MoveNext
Loop
storeTable.Close
Set storeTable = Nothing
End With
End Sub
