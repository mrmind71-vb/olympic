VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form VsStore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ãÊÇÈÚÉ ÇáãæÏíáÇÊ ( ÇáÅÖÇÝÉ - ÇáãÈíÚÇÊ - äÓÈÉ ÇáãÈíÚÇÊ ) "
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E3C7AB&
      Caption         =   "ÎÑæÌ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1050
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00E3C7AB&
      Caption         =   "ÊÑÇÌÚ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1350
      Width           =   915
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00E3C7AB&
      Caption         =   "ØÈÇÚÉ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1350
      Width           =   915
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "ÚÑÖ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5265
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   6660
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   8280
      Width           =   8565
      Begin VSFlex7LCtl.VSFlexGrid grid2 
         Height          =   1275
         Left            =   90
         TabIndex        =   17
         Top             =   225
         Width           =   8295
         _cx             =   14631
         _cy             =   2249
         _ConvInfo       =   1
         Appearance      =   1
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
         BackColor       =   16777215
         ForeColor       =   12089119
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   12089119
         BackColorBkg    =   13822956
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
         SelectionMode   =   0
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
         BackColorFrozen =   13822956
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   8385
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
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1815
      End
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xItem 
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
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   1815
      End
      Begin VB.TextBox xDesca 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   3030
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   4320
         TabIndex        =   7
         Top             =   540
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   4320
         TabIndex        =   14
         Top             =   1260
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ãÌãæÚÉ"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ãä ÊÇÑíÎ"
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
         Left            =   7455
         TabIndex        =   13
         Top             =   195
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Åáì ÊÇÑíÎ"
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
         Left            =   3195
         TabIndex        =   12
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ãÌãæÚÉ"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   615
         Width           =   630
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÕäÝ"
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
         Left            =   7455
         TabIndex        =   10
         Top             =   1050
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "æÕÝ ÇáÕäÝ"
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
         Left            =   3195
         TabIndex        =   9
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label xItemDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   5415
      End
   End
   Begin VB.TextBox xModel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   225
      Visible         =   0   'False
      Width           =   690
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   6495
      Left            =   90
      TabIndex        =   0
      Top             =   1755
      Width           =   15045
      _cx             =   26538
      _cy             =   11456
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
      BackColor       =   13822956
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   13822956
      GridColor       =   0
      GridColorFixed  =   4210752
      TreeColor       =   -2147483632
      FloodColor      =   0
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   945
      Top             =   810
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Left            =   540
      Top             =   630
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
Attribute VB_Name = "VsStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataSaleTable As Recordset
Public nRemRow As Double
Dim BalModelTable As Recordset, PModelTable As Recordset
Dim SaleDateTable As Recordset
Dim datatable As Recordset
Dim SuppTable As Recordset
Dim Bal1Table As Recordset
Dim Bal2Table As Recordset

Dim temptable As Recordset
Dim storeTable As Recordset
Dim FlagTable As Recordset
Dim TPurchTable As Recordset
Dim SuppModel As Recordset
Dim FDateTable As Recordset
Dim nCrow As Double
Dim cString As String
Dim cMosm As String
Private Sub Cmd_Print_Click()
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -2, "ÈíÇä ÊÝÕíáì ãæÞÝ ãÍá" & Me.xStore.Text, , , , , 8
    PrintGrd.Show 1
End Sub
Private Sub CMDEXIT_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
MyLoad
End Sub

Private Sub CmdOk1_Click()
MyLoad
End Sub

Private Sub Form_Load()
data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "SELECT * FROM FILE1_50"
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

data2.ConnectionString = CON.ConnectionString
data2.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data2
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

'data2.DatabaseName = MdbPath
'data2.RecordSource = "FILE1_50"
'data2.Refresh
'xgroup.ListField = "Desca"
'xgroup.BoundColumn = "Code"

'Data3.DatabaseName = MdbPath
'Data3.RecordSource = "Stores"
'xStore.ListField = "desca"
'xStore.BoundColumn = "code"


With grid1
.Cols = 18
.Rows = 2
.RowHeight(0) = 800
.WordWrap = True
.ColHidden(0) = True
'.Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 12
'.Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = "Traditional Arabic"
.Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
.RowHidden(1) = True

.TextMatrix(0, 0) = "ßæÏ"
.TextMatrix(0, 1) = "ãÌãæÚÉ"
.TextMatrix(0, 2) = "ÕäÝ "
.TextMatrix(0, 3) = "ÈíÇä ÇáÕäÝ"

.TextMatrix(0, 4) = "ÑÕíÏ Ãæá"
.TextMatrix(0, 5) = "ÊÓæíÉ ÌÑÏ"
.TextMatrix(0, 6) = "ÕÇÝì ãÔÊÑíÇÊ"
.TextMatrix(0, 7) = "ÕÇÝì ÊÍæíáÇÊ"
.TextMatrix(0, 8) = "åÇáß"
.TextMatrix(0, 9) = "ÊÕäíÚ"
.TextMatrix(0, 10) = "Ì. æÇÑÏ"
.TextMatrix(0, 11) = "Ì. ãÈíÚÇÊ"
.TextMatrix(0, 12) = "ÞíãÉ ãÈíÚÇÊ"
.TextMatrix(0, 13) = "ÑÈÍ ãÈíÚÇÊ"
.TextMatrix(0, 14) = "äÓÈÉ ãÈíÚÇÊ"
.TextMatrix(0, 15) = " ÓÚÑ ÈíÚ"
.TextMatrix(0, 16) = "ÑÕíÏ ÃÎÑ"
.TextMatrix(0, 17) = " ÞíãÉ ÑÕíÏ"
.ColHidden(15) = Not bopt2
.ColWidth(0) = 300
.ColWidth(1) = 1000
.ColWidth(2) = 800
.ColWidth(3) = 2000
.ColWidth(4) = 800

.ColWidth(5) = 800
.ColWidth(6) = 800
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(9) = 800
.ColWidth(10) = 800
.ColWidth(11) = 800
.ColWidth(12) = 800
.ColWidth(13) = 800
.ColWidth(14) = 800
.ColWidth(15) = 800
.ExplorerBar = flexExSort
.ColDataType(0) = flexDTString
.ColDataType(1) = flexDTString
.ColDataType(2) = flexDTString
.ColDataType(3) = flexDTString

.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble
.ColDataType(10) = flexDTDouble
.ColDataType(11) = flexDTDouble
.ColDataType(12) = flexDTDouble
.ColDataType(13) = flexDTDouble
.ColDataType(14) = flexDTDouble
.ColDataType(15) = flexDTDouble

.Editable = flexEDNone

'VsStore.Cols = 1
'VsStore.Rows = 6
'VsStore.RowHidden(0) = True
'storeTable.MoveFirst
'VsStore.WordWrap = True
End With
With Grid2
Grid2.Cols = 1
Grid2.ColWidth(0) = 2000
Grid2.Rows = 6
.TextMatrix(1, 0) = "ÇáãÍá"
.TextMatrix(2, 0) = "ãÔÊÑíÇÊ"
.TextMatrix(3, 0) = "ãÈíÚÇÊ"
.TextMatrix(4, 0) = "ÊÍæíáÇÊ"
.TextMatrix(5, 0) = "ÑÕíÏ"
.FixedRows = 2
.RowHidden(0) = True

data2.Recordset.MoveFirst
Do Until data2.Recordset.EOF
    Grid2.Cols = Grid2.Cols + 1
    Grid2.TextMatrix(0, Grid2.Cols - 1) = data2.Recordset!CODE
    Grid2.TextMatrix(1, Grid2.Cols - 1) = data2.Recordset!desca
    data2.Recordset.MoveNext
Loop
For I = 1 To Grid2.Cols - 1
    Grid2.ColWidth(I) = (Val(Grid2.Width) - (Grid2.ColWidth(0) + 400)) / (Grid2.Cols - 1)
Next
End With
'nColW = (VsStore.Width - 1700) / storeTable.RecordCount


'Do While Not storeTable.EOF
'
'    VsStore.Cols = VsStore.Cols + 1
'    VsStore.ColWidth(VsStore.Cols - 1) = nColW
'    VsStore.TextMatrix(1, VsStore.Cols - 1) = storeTable.desca
'    VsStore.TextMatrix(0, VsStore.Cols - 1) = storeTable.CODE
'
'    storeTable.MoveNext
'Loop

'VsStore.Cell(flexcpAlignment, 0, 0, VsStore.Rows - 1, VsStore.Cols - 1) = 4
'.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
'End With

End Sub
Private Sub grid1_DblClick()
With grid1
    If .Row > 1 Then
        Frame2.Caption = grid1.TextMatrix(grid1.Row, 3)
        MYLOAD2 grid1.TextMatrix(grid1.Row, 2)
        
'        xModel.Text = .TextMatrix(.row, 2)
'        For i = 1 To VsStore.Cols - 1
            'AllMoveModelStore .TextMatrix(.Row, 2), VsStore.TextMatrix(0, i), i
'        Next i
    End If
End With
End Sub

Private Sub xfact_Change()
    xCode.Text = ""
    xDesca.Text = ""
End Sub
Private Sub MyLoad()
Dim GRDTABLE As New ADODB.Recordset
cWhere = ""
If IsDate(xdate1.Text) Then cWhere = " date < " & DateSq(xdate1.Text)
cField1 = myiif(cWhere, "VAL([IN] & '')  - VAL(OUT & '')") & " as BalanceFirst"


cWhere = ""
If IsDate(xdate1.Text) Then cWhere = " date >= " & DateSq(xdate1.Text)
If IsDate(xDate2.Text) Then cWhere = cWhere & TurnAnd(cWhere) & " date <= " & DateSq(xDate2.Text)

cWhere = cWhere & TurnAnd(cWhere)

cField2 = myiif(cWhere & "FILE1_11.TYPE = 'Z' ", "[IN]") & " as StockDiffer "
cField3 = myiif(cWhere & "(FILE1_11.TYPE = '2' OR FILE1_11.TYPE = '7') ", "VAL([IN] & '') - VAL(OUT & '')") & " as Purchase "
cField4 = myiif(cWhere & "(FILE1_11.TYPE = 'T' OR FILE1_11.TYPE = 'F') ", "VAL([IN] & '') - VAL(OUT & '')") & " as Trans "
cField5 = myiif(cWhere & "FILE1_11.TYPE = '9' ", "[OUT]") & " as Damage "
cField6 = myiif(cWhere & "(FILE1_11.TYPE = '10' OR FILE1_11.TYPE = '11') ", "VAL([IN] & '') - VAL(OUT & '')") & " as Product "
cField7 = myiif(cWhere & "(FILE1_11.TYPE = '6' OR FILE1_11.TYPE = '3') ", "VAL(OUT & '') - VAL([IN] & '')") & " as Sales"
cField8 = myiif(cWhere & "(FILE1_11.TYPE = '6' OR FILE1_11.TYPE = '3') ", " IIF(FILE1_11.TYPE = '6',1,-1) *  VAL(TOTAL & '')") & " as SalesValue "
cField9 = myiif(cWhere & "(FILE1_11.TYPE = '6' OR FILE1_11.TYPE = '3') ", " IIF(FILE1_11.TYPE = '6',1,-1) *  VAL(FILE1_11.COST & '')") & " as SalesCost "

cWhere = ""
If IsDate(xDate2.Text) Then cWhere = TurnAnd(cWhere) & " date <= " & DateSq(xDate2.Text)

cField10 = myiif(cWhere & TurnAnd(cWhere) & " Not(FILE1_11.TYPE =  '6' or FILE1_11.TYPE = '3') ", "VAL([IN] & '') - VAL(OUT & '')") & " as InAll"
cField11 = myiif(cWhere, "VAL([IN] & '') - VAL(OUT & '' )") & "  as BalanceLast "
cWhere = ""
cString = "SELECT FILE1_10.ITEM , FILE1_10.DESCA, FILE1_50.DESCA as GroupDesca ,FILE1_10.PRICE, FILE1_10.COST1 , FILE1_10.GROUP " & _
           "," & cField1 & _
           "," & cField2 & _
           "," & cField3 & _
           "," & cField4 & _
           "," & cField5 & _
           "," & cField6 & _
           "," & cField7 & _
           "," & cField8 & _
           "," & cField9 & _
           "," & cField10 & _
           "," & cField11

cString = cString & " FROM (FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE "
If xGroup.Text <> "" Then cWhere = cWhere & TurnAnd(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
If xDesca.Text <> "" Then cWhere = cWhere & TurnAnd(cWhere) & MyParnAnd(Trim(xDesca.Text), "file1_10.Desca")
If XITEM.Text <> "" Then cWhere = cWhere & TurnAnd(cWhere) & "file1_11.item = " & MyParn(XITEM.Text)
If xStore.BoundText <> "" Then cWhere = cWhere & TurnAnd(cWhere) & "file1_11.store = " & MyParn(xStore.BoundText)

cString = cString & TurnWhere(cWhere) & cWhere & " GROUP BY FILE1_10.ITEM , FILE1_10.DESCA,FILE1_50.DESCA ,FILE1_10.PRICE,FILE1_10.COST1, FILE1_10.GROUP  "
GRDTABLE.Open cString, CON, adOpenKeyset, adLockReadOnly, adCmdText
With grid1
.Rows = 1
Me.MousePointer = 11
Do Until GRDTABLE.EOF
    .AddItem ""
   .TextMatrix(.Rows - 1, 0) = GRDTABLE!Group & ""
   .TextMatrix(.Rows - 1, 1) = GRDTABLE!GroupDesca & ""
   .TextMatrix(.Rows - 1, 2) = GRDTABLE!Item & ""
   .TextMatrix(.Rows - 1, 3) = GRDTABLE!desca & "'"
   .TextMatrix(.Rows - 1, 4) = Val(GRDTABLE!BalanceFirst & "")
   .TextMatrix(.Rows - 1, 5) = Val(GRDTABLE!StockDiffer & "")
   .TextMatrix(.Rows - 1, 6) = Val(GRDTABLE!purchase & "")
   .TextMatrix(.Rows - 1, 7) = Val(GRDTABLE!TRANS & "")
   .TextMatrix(.Rows - 1, 8) = Val(GRDTABLE!Damage & "")
   .TextMatrix(.Rows - 1, 9) = Val(GRDTABLE!Product & "")
   .TextMatrix(.Rows - 1, 10) = Val(GRDTABLE!InAll & "")
   
   .TextMatrix(.Rows - 1, 11) = Val(GRDTABLE!Sales & "")
   .TextMatrix(.Rows - 1, 12) = Val(GRDTABLE!salesvalue & "")
   .TextMatrix(.Rows - 1, 13) = Val(GRDTABLE!salesvalue & "") - Val(GRDTABLE!SalesCost & "")
    If Val(GRDTABLE!InAll & "") <> 0 Then
        .TextMatrix(.Rows - 1, 14) = Round(Val(GRDTABLE!Sales & "") / GRDTABLE!InAll * 100, 2) & "%"
    End If
   .TextMatrix(.Rows - 1, 15) = retitem(GRDTABLE!Item & "", "price") & ""
   .TextMatrix(.Rows - 1, 16) = Val(GRDTABLE!BalanceLast & "")
   .TextMatrix(.Rows - 1, 17) = Val(GRDTABLE!BalanceLast & "") * Val(retitem(GRDTABLE!Item & "", "price") & "")
    GRDTABLE.MoveNext
Loop
.Subtotal flexSTClear
.Subtotal flexSTSum, -1, 4, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 5, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 6, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 7, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 8, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 9, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 11, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 14, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 15, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 16, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"
.Subtotal flexSTSum, -1, 17, "##0", , RGB(255, 0, 0), True, "ÅÌãÇáì"

Me.MousePointer = 0
End With
End Sub
Private Sub MYLOAD2(pItem)
Dim GRDTABLE As New ADODB.Recordset
cWhere = ""
If IsDate(xdate1.Text) Then cWhere = " date >= " & DateSq(xdate1.Text)
If IsDate(xDate2.Text) Then cWhere = cWhere & TurnAnd(cWhere) & " date <= " & DateSq(xDate2.Text)
cWhere = cWhere & TurnAnd(cWhere)
cField1 = myiif(cWhere & TurnAnd(cWhere) & TurnAnd(cWhere) & "type = '2' or type = '7' ", "VAL([IN] & '') - VAL(OUT & '') ") & " AS PURCHASE"
cField2 = myiif(cWhere & TurnAnd(cWhere) & TurnAnd(cWhere) & "type = '6' or type = '6' ", "VAL(OUT & '') - VAL([IN] & '') ") & " AS SALES"
cField3 = myiif(cWhere & TurnAnd(cWhere) & TurnAnd(cWhere) & "type = 'T' or type = 'F' ", "VAL([IN] & '') - VAL(OUT & '') ") & " AS TRANS"
cWhere = ""
If IsDate(xDate2.Text) Then cWhere = cWhere & TurnAnd(cWhere) & " date <= " & DateSq(xDate2.Text)
cField4 = myiif(cWhere, "VAL([IN] & '') - VAL(OUT & '') ") & " AS BALANCE"

cString = "Select " & _
          cField1 & _
          "," & cField2 & _
          "," & cField3 & _
          "," & cField4 & _
          " FROM FILE1_11 " & _
          " WHERE FILE1_11.ITEM = " & MyParn(pItem)

For I = 1 To Grid2.Cols - 1
    cWhere = ""
    cString1 = cString & " and  store = " & MyParn(Grid2.TextMatrix(0, I)) & _
               " GROUP BY FILE1_11.ITEM "
    GRDTABLE.Open cString1, CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not (GRDTABLE.EOF And GRDTABLE.BOF) Then
        Grid2.TextMatrix(2, I) = GRDTABLE!purchase & ""
        Grid2.TextMatrix(3, I) = GRDTABLE!Sales & ""
        Grid2.TextMatrix(4, I) = GRDTABLE!TRANS & ""
        Grid2.TextMatrix(5, I) = GRDTABLE!balance & ""
    End If
    GRDTABLE.Close
Next
Set GRDTABLE = Nothing
End Sub
Private Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select Item ,DescA , FILE1_10.PRICE From file1_10 "
Generalarray(2) = "Order by item"
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "ÅÓã ÇáÕäÝ"
listarray(0, 1) = "(FILE1_10.item LIKE '%cFilter%' or %%DESCA%%)"



GrdArray(0, 0) = "ßæÏ ÇáÕäÝ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ÅÓã ÇáÕäÝ"
GrdArray(1, 1) = 6000

GrdArray(2, 0) = "ÓÚÑ ÇáÕäÝ"
GrdArray(2, 1) = 1200


searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "ÅÓÊÚáÇã ÇáÇÕäÇÝ"
Search3.Show 1
End Sub

Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookup
End Sub

Private Sub xitem_LostFocus()
xitemDesca.Caption = ""
If Trim(XITEM.Text) <> "" Then
    xitemDesca.Caption = retitem(XITEM.Text, "desca") & ""
End If
End Sub

Sub myProc()
XITEM.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
xitemDesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
Unload Search3
End Sub

