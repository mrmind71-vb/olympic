VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form TDaySal 
   ClientHeight    =   8385
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15300
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   15300
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCloseday 
      Caption         =   "«€·«ﬁ «·ÌÊ„"
      Height          =   510
      Left            =   7695
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8685
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   11340
      MaskColor       =   &H00FFFFFF&
      Picture         =   "TDaySal.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   7740
      UseMaskColor    =   -1  'True
      Width           =   3885
   End
   Begin VB.Frame Frame2 
      Height          =   3030
      Left            =   11385
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4680
      Width           =   3795
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   225
         Width           =   2760
         _ExtentX        =   4868
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
      Begin MSACAL.Calendar xDate 
         Height          =   2430
         Left            =   180
         TabIndex        =   15
         Top             =   540
         Width           =   3495
         _Version        =   524288
         _ExtentX        =   6165
         _ExtentY        =   4286
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2006
         Month           =   5
         Day             =   21
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483635
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂ«‘Ì— :"
         Height          =   270
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "«·„»Ì⁄« "
      Height          =   2265
      Left            =   11385
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   135
      Width           =   3795
      Begin VB.Label xSalesNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì «·„»Ì⁄«  :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1845
         Width           =   1950
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì ﬁÌ„… «·Œ’„ :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   1950
      End
      Begin VB.Label xSalesDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Label xSalesCount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì ⁄œœ «·»Ê‰«  :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1905
      End
      Begin VB.Label xSalesValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   1410
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·„»Ì⁄«  :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   1905
      End
      Begin VB.Label xSalesValueRet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   990
         Width           =   1410
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì „—œÊœ «·„»Ì⁄«  :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1035
         Width           =   1905
      End
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   1065
      Left            =   11565
      TabIndex        =   1
      Top             =   8190
      Visible         =   0   'False
      Width           =   960
      _cx             =   1693
      _cy             =   1879
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Traditional Arabic"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   0   'False
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   100
      MarginTop       =   100
      MarginRight     =   100
      MarginBottom    =   100
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   200
      ZoomMode        =   0
      ZoomMax         =   400
      ZoomMin         =   200
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   1
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483626
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›"
      TabPicture(0)   =   "TDaySal.frx":246C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPrint2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " ›’Ì·Ì »Ê‰«  «·»Ì⁄"
      TabPicture(1)   =   "TDaySal.frx":2488
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grid1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "xDesca"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "xitem"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox xitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   7650
         Width           =   3570
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   7650
         Width           =   4020
      End
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   7095
         Left            =   135
         TabIndex        =   2
         Top             =   450
         Width           =   11040
         _cx             =   19473
         _cy             =   12515
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   7100
         Left            =   -74865
         TabIndex        =   3
         Top             =   450
         Width           =   11040
         _cx             =   19473
         _cy             =   12524
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
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
      Begin Threed.SSCommand cmdPrint2 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   -74865
         TabIndex        =   21
         Top             =   7605
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
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
         Picture         =   "TDaySal.frx":24A4
      End
      Begin VB.Label Label9 
         Caption         =   "ﬂÊœ «·’‰› :"
         Height          =   285
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   7695
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "«·’‰› :"
         Height          =   330
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   7695
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
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
   Begin VB.Frame Frame3 
      Caption         =   "«·‰ﬁœÌ…"
      Height          =   2265
      Left            =   11340
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2430
      Width           =   3795
      Begin VB.Label xVisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1755
         Width           =   1410
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "›Ì“« :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1800
         Width           =   1950
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’«—Ì› :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   720
         Width           =   1950
      End
      Begin VB.Label XCHARGE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   675
         Width           =   1410
      End
      Begin VB.Label xCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·‰ﬁœÌ… :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   315
         Width           =   1950
      End
      Begin VB.Label xbalcash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ï ‰ﬁœÌ… «·ÌÊ„ :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1125
         Width           =   1950
      End
   End
End
Attribute VB_Name = "TDaySal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bShowPrf As Boolean
Dim con As New ADODB.Connection
Dim TSalTable As Recordset
Dim CountTable As Recordset
Dim SalTable As Recordset
Dim Sal2Table As Recordset
Dim temptable As Recordset
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmd_close_Click()
'    cString = " UPDATE FILE6_20 SET FILE6_20.CLOSED = TRUE  Where Date = " & DATESQ(xDate.Value)
'    mydb.Execute cString
End Sub
Private Sub myload()
myload1
myload2
MyLoadTotal
End Sub
Private Sub myload1()
Dim bNew As Boolean
Dim nTotalDiscount As Double, nTotal As Double, nTotalSalesDis As Double, nTotalcash As Double, nTotalVisa As Double
If IsDate(xDate.Value) Then cwhere = turn(cwhere, " and ") & " DATE = " & DateSq(Format(xDate.Value, "YYYY-MM-DD"))
If xBox.BoundText <> "" Then cwhere = cwhere & turn(cwhere, " and ") & " BOX = " & MyParn(xBox.BoundText)
cwhere = cwhere & turn(cwhere, " AND ") & " PRINTED = 1"
If Trim(xDesca.Text) <> "" Then
    cwhere = cwhere & turn(cwhere, " AND ") & " DOC_NO IN (SELECT DOC_NO FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE " & MyParnAnd(xDesca, "FILE1_10.DESCA") & ")"
End If

If Trim(xItem.Text) <> "" Then
    cwhere = cwhere & turn(cwhere, " AND ") & " DOC_NO IN (SELECT DOC_NO FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE file6_20.item = " & MyParn(xItem.Text) & " )"
End If

cString = "SELECT SALESDTL.* " & _
          " FROM SALESDTL"
cString = cString & turn(cwhere, " WHERE ") & cwhere
cString = cString & " ORDER BY DOC_NO,FLAG"

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With grid1
grid1.Rows = 1
Do Until loctable.EOF
    .AddItem ""
    If loctable!flag = 0 Then
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no & ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Time, "HH:NN")
        .TextMatrix(.Rows - 1, 2) = loctable!Item & ""
        .TextMatrix(.Rows - 1, 3) = loctable!Desca & ""
        .TextMatrix(.Rows - 1, 4) = loctable!Quant
        .TextMatrix(.Rows - 1, 5) = loctable!price & ""
        .TextMatrix(.Rows - 1, 6) = Format(Val(loctable!TOTAL & ""), "Fixed")
        nTotal = nTotal + Val(loctable!TOTAL & "")
    ElseIf loctable!flag = 1 Then
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no
        For I = 0 To 5
            .TextMatrix(.Rows - 1, I) = "«·Œ’„"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 6) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalDiscount = nTotalDiscount + Val(.TextMatrix(.Rows - 1, 6))
    ElseIf loctable!flag = 3 Then
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no
        For I = 0 To 5
            .TextMatrix(.Rows - 1, I) = "«·«Ã„«·Ì"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 6) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    ElseIf loctable!flag = 4 Then
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no
        For I = 0 To 5
            .TextMatrix(.Rows - 1, I) = "‰ﬁœÌ…"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 6) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalcash = nTotalcash + Val(loctable!TOTAL & "")
    ElseIf loctable!flag = 5 Then
        '.TextMatrix(.Rows - 1, 0) = LOCTABLE!doc_no
        For I = 0 To 5
            .TextMatrix(.Rows - 1, I) = "›Ì“«"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 6) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalVisa = nTotalVisa + Val(loctable!TOTAL & "")
    End If
    loctable.MoveNext
Loop

If nTotal <> 0 Then
    grid1.AddItem ""
    For I = 0 To 5
        .TextMatrix(.Rows - 1, I) = "≈Ã„«·Ì «·ÌÊ„"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 6) = Round(nTotal, 2) - Round(nTotalDiscount, 2)
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If nTotalcash <> 0 Then
    .AddItem ""
    For I = 0 To 5
        .TextMatrix(.Rows - 1, I) = "≈Ã„«·Ì «·‰ﬁœÌ…"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 6) = nTotalcash
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If nTotalVisa <> 0 Then
    .AddItem ""
    For I = 0 To 5
        .TextMatrix(.Rows - 1, I) = "≈Ã„«·Ì «·›Ì“«"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 6) = nTotalVisa
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If (nTotal - nTotalDiscount) - (nTotalcash + nTotalVisa) <> 0 Then
    .AddItem ""
    For I = 0 To 5
        .TextMatrix(.Rows - 1, I) = "≈Ã„«·Ì «·¬Ã·"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 6) = Round((nTotal - nTotalDiscount) - (nTotalcash + nTotalVisa), 2)
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If
End With
fixgrd1
End Sub
Private Sub myload2()
Dim bNew As Boolean, DDate1 As String, DDate2 As String
If IsDate(xDate.Value) Then cwhere = cwhere & turn(cwhere, " AND ") & "FILE6_20H.DATE = " & DateSq(Format(xDate.Value, "YYYY-MM-DD"))
If xBox.BoundText <> "" Then cwhere = cwhere & turn(cwhere, " AND ") & " FILE6_20H.BOX = " & MyParn(xBox.BoundText)
cwhere = cwhere & turn(cwhere, " AND ") & "PRINTED = 1"

cField1 = "Select SUM(FILE6_20H.DISCOUNT) FROM FILE6_20H"

If cwhere <> "" Then cField1 = cField1 & turn(cField1) & cwhere
cField1 = "(" & cField1 & ")" & " as SumofDiscount"


cField2 = myiif("FILE6_20.Quant > 0", "QUANT * FILE6_20.PRICE") & " AS SalesValue"
cField3 = myiif("FILE6_20.Quant < 0", "-1 * QUANT * FILE6_20.PRICE") & " AS SalesValueRet"




cString = "Select " & cField1 & "," & cField2 & "," & cField3 & _
          ", file6_20.price , FILE6_20.ITEM,FILE1_10.DESCA,SUM(FILE6_20.QUANT) AS salesQuantNet,SUM(FILE6_20.QUANT * FILE6_20.PRICE) AS SalesValueNet FROM (FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO) INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"

If cwhere <> "" Then
    cString = cString & turn(cString) & cwhere
End If
cString = cString & " group by FILE6_20.price,FILE6_20.ITEM,FILE1_10.DESCA"

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

Dim nTotalValue As Double
With grid2
.Rows = 1
Do Until sourcetable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = sourcetable!Item
    .TextMatrix(.Rows - 1, 1) = sourcetable!Desca & ""
    
    .TextMatrix(.Rows - 1, 2) = Format(sourcetable!SalesQuantNet, "Fixed")
    .TextMatrix(.Rows - 1, 3) = Format(sourcetable!price, "Fixed")
    .TextMatrix(.Rows - 1, 4) = Format(sourcetable!SalesValue, "fixed")
    .TextMatrix(.Rows - 1, 5) = Format(sourcetable!SalesValueRet, "fixed")
    .TextMatrix(.Rows - 1, 6) = Format(sourcetable!SalesValueNet, "fixed")
    nTotalValue = Val(sourcetable!SalesValueNet & "") + nTotalValue
    nDiscount = sourcetable!sumOfDiscount
    sourcetable.MoveNext
Loop

If nDiscount <> 0 Then
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = ""
    .TextMatrix(.Rows - 1, 1) = "«·Œ’„"
    .TextMatrix(.Rows - 1, 2) = ""
    .TextMatrix(.Rows - 1, 6) = -1 * nDiscount
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFFF
    nTotalValue = nTotalValue - nDiscount
End If

End With
fixGrd2
End Sub
Private Sub cmdCloseDay_Click()
If Not xBox.MatchedWithList Then
    MsgBox " ÕœÌœ «·Œ“‰…"
    Exit Sub
Else
    If Not CheckOpen Then Exit Sub
    Dim dLastdate As String
    dLastdate = Format(GetDesca("Select Max(Date) from FILE6_20H"), "YYYY-MM-DD")
    If Not IsDate(dLastdate) Then
        MsgBox "·«  ÊÃœ ›« Ê—… „»Ì⁄«   „ «œŒ«·Â«"
    End If
    If Format(dLastdate, "YYYY-MM-DD") = Format(dSalesDate, "YYYY-MM-DD") Then
        If IsDate(dSalesDate) Then
            Dim sMsg As String, cString As String
            sMsg = " €ÌÌ— «· «—ÌŒ «·Ì ÌÊ„ ÃœÌœ" & Format(dDate, "YYYY-MM-DD")
            If MsgBox(sMsg, vbOKCancel + vbDefaultButton2) = vbOK Then
                cString = "UPDATE datesales SET DATE = " & addDate(Format(DateAdd("d", 1, dSalesDate), "YYYY-MM-DD"))
                con.BeginTrans
                On Error GoTo myerror
                con.Execute cString
                con.CommitTrans
                dSalesDate = Format(DateAdd("d", 1, dSalesDate), "YYYY-MM-DD")
                Firsttitle = Secondtitle & Format(dSalesDate, "YYYY-MM-DD")
                main.Caption = Firsttitle
                salesfrm.Caption = Format(dSalesDate, "YYYY-MM-DD")
                salesfrm.mydefine
                Inform " „  €ÌÌ— «·ÌÊ„ »‰Ã«Õ"
            End If
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans

End Sub
Private Sub Command1_Click()
doprint2
End Sub
Private Sub Command2_Click()
doprint2
End Sub

Private Sub CmdPrint2_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "≈Ã„«·Ì „»Ì⁄«  «·«’‰«› ›Ï " & Format(xDate.Value, "YYYY-MM-DD")
    PrintGrd.doprint grid2, 0.9, -1, cHead1, , , False, False, 8, , Array(grid2.Rows - 1)
    PrintGrd.Show 1

End Sub

Private Sub Command3_Click()
Dim loctable As New ADODB.Recordset
On Error GoTo myerror
Me.MousePointer = 11
loctable.Open "select file6_20.doc_no,file6_20.item,file6_20h.date,file6_20.row from file6_20 inner join file6_20h on file6_20.doc_no = file6_20h.doc_no", con, adOpenStatic, adLockReadOnly
con.BeginTrans
Do
    con.Execute "update file6_20 set file6_20.cost = " & Round(itemCost(loctable!Item, Format(loctable!Date, "yyyy-mm-dd"), con), 2) & _
                 " where file6_20.doc_no = " & MyParn(loctable!doc_no) & " and " & _
                 " file6_20.row = " & Val(loctable!Row & "")
    loctable.MoveNext
Loop Until loctable.EOF
con.CommitTrans
Me.MousePointer = 0
MsgBox "done ..."
Exit Sub
Me.MousePointer = 0
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub Command4_Click(Index As Integer)
Unload Me
End Sub
Private Sub Form_Load()
openCon con
 xDate.Visible = True
data1.ConnectionString = strCon

'VsSales.RowHeight(0) = 400
'VsSales.Sort = flexSortNone
'VsSales.ExplorerBar = flexExSortShow
xDate.Value = DateValue(Format(salesfrm.xDate.Text, "YYYY-MM-DD"))
xDate.Visible = bopt1
xBox.Enabled = bopt1
MakeDataString
Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "CODE"
Dim cdefman As String
'xBox.BoundText = salesfrm.xBox.BoundText
'cmdCloseday.Enabled = cdefman <> ""

myload
End Sub
Private Sub Cmd_Print_Click()
    Load REP1_27
    REP1_27.Date1.Text = Format(xDate.Value, "YYYY-MM-DD")
    REP1_27.date2.Text = Format(xDate.Value, "YYYY-MM-DD")
    REP1_27.Show 1
End Sub
Private Sub SSCommand3_Click()
    Load Rep_107
    Rep_107.xDate1.Text = xDate.Value
    Rep_107.xDate2.Text = xDate.Value
    Rep_107.xStore.BoundText = xStore.BoundText
    Rep_107.Show 1
End Sub
Private Sub VsSales_DBLClick()
    ViewSale.Show 1
    myload
End Sub

Private Sub SSCommand7_Click()

End Sub

Private Sub X5_Click()
End Sub

Private Sub xDate_BeforeUpdate(Cancel As Integer)
   ' MyLoad
   ' MyLoadTotal
End Sub
Private Sub xDate_Click()
    myload
    MyLoadTotal

End Sub

Private Sub xStore_Click(Area As Integer)
    myload
End Sub
Private Sub MyLoadTotal()
cField = "Sum(1) as SalesCount"

cField = cField & turn(cField, ",") & _
    "Sum(Cash) as SumofCash"

cField = cField & turn(cField, ",") & _
       "Sum(Discount) as SumofDiscount"

cField = cField & turn(cField, ",") & _
       "Sum(Visa) as SumofVisa"

If IsDate(xDate.Value) Then cwhere = cwhere & turn(cwhere, " and ") & " FILE6_20H.DATE = " & DateSq(Format(xDate.Value, "YYYY-MM-DD"))
If xBox.BoundText <> "" Then cwhere = cwhere & turn(cwhere, " and ") & "FILE6_20H.BOX = " & MyParn(xBox.BoundText)
cwhere = cwhere & turn(cwhere, " and ") & " printed = 1"

cString = "SELECT " & _
           cField & _
          " FROM FILE6_20H"

If cwhere <> "" Then cString = cString & turn(cwhere, " where ") & cwhere
                                                      
                                                      
Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenKeyset, adCmdText
With sourcetable
If Not (sourcetable.EOF And sourcetable.BOF) Then
    xSalesCount.Caption = Val(!SalesCount & "")
    xcash.Caption = Val(!sumofCash & "")
    xSalesDiscount.Caption = Val(!sumOfDiscount & "")
    xVisa.Caption = Val(!sumofVisa & "")

    Dim aRet As Variant
    cField1 = myiif("quant > 0", "file6_20.total")
    cField2 = myiif("quant < 0", "file6_20.total")
    cString = "Select " & cField1 & "," & cField2 & _
              " from file6_20 inner join file6_20h on file6_20.doc_no = file6_20h.doc_no" & _
              turn(cwhere, " where ") & cwhere
    aRet = aGetDesca(cString)
    If UBound(aRet) <> 0 Then
        xSalesValue.Caption = Format(Val(aRet(1) & ""), "fixed")
        xSalesValueRet.Caption = Format(Abs(Val(aRet(2) & "")), "fixed")
        xSalesNet.Caption = Val(xSalesValue.Caption) - Val(xSalesValueRet.Caption) - Val(xSalesDiscount.Caption)
    End If
End If
End With
                           
cString = "SELECT SUM(VALUE) FROM FILE8_50 INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO"
cString = cString & turn(cString) & "FILE8_50H.DATE = " & DateSq(xDate.Value)
If xBox.BoundText <> "" Then
    cString = cString & turn(cString) & "BOX = " & MyParn(xBox.BoundText)
End If
XCHARGE.Caption = Val(GetDesca(cString) & "")
xbalcash.Caption = Format(Val(xcash.Caption & "") - Val(XCHARGE.Caption & ""), "#0.00")

End Sub
Private Sub doprint2()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To grid2.Rows - 2
    temptable.AddNew
    temptable!str21 = "≈Ã„«·Ì „»Ì⁄«  «’‰«› ÌÊ„ " & NameOfDay(xDate.Value) & " «·„Ê«›ﬁ : " & Format(xDate.Value, "dd-mm-yy")
    If xBox.MatchedWithList Then temptable!str22 = "«·ﬂ«‘Ì— : " & xBox.Text
    temptable!str1 = TurnValue(grid2.TextMatrix(I, 0))
    temptable!str2 = TurnValue(grid2.TextMatrix(I, 1))
    temptable!val1 = Val(grid2.TextMatrix(I, 2))
    temptable!val2 = Val(grid2.TextMatrix(I, 3))
    temptable!Val3 = Val(grid2.TextMatrix(I, 4))
    temptable!val4 = Val(grid2.TextMatrix(I, 5))
    temptable!Val5 = Val(grid2.TextMatrix(I, 6))
'    temptable!Val6 = Val(grid2.TextMatrix(i, 7))
    temptable.Update
Next

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\salesday1.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub xBox_Change()
If xBox.MatchedWithList Or Trim(xBox.Text) = "" Then
    myload
    MyLoadTotal
End If
End Sub
Private Sub MakeDataString()
data1.RecordSource = "SELECT code,DESCA FROM FILE0_50 "
End Sub
Private Sub xbox_LostFocus()
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub
Private Sub fixGrd2()
With grid2
.Cols = 7
.TextMatrix(0, 0) = "ﬂÊœ «·’‰›"
.TextMatrix(0, 1) = "«·«”„"
.TextMatrix(0, 2) = "„Ì⁄« "
.TextMatrix(0, 3) = "”⁄— «·»Ì⁄"
.TextMatrix(0, 4) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 5) = "ﬁÌ„… „— Ã⁄"
.TextMatrix(0, 6) = "’«›Ì ﬁÌ„… «·»Ì⁄"

.ColWidth(0) = 2000
.ColWidth(1) = 3500
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000

.MergeCells = flexMergeFree
.MergeCol(0) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 2, "#0", vbYellow, , True, ""
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, ""
.Subtotal flexSTSum, -1, 5, "#0", vbYellow, , True, ""
.Subtotal flexSTSum, -1, 6, "#0", vbYellow, , True, ""
If .Rows > 1 Then
    .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
    .TextMatrix(.Rows - 1, 1) = "«·«Ã„«·Ì"
    .MergeCells = flexMergeNever
    .MergeRow(.Rows - 1) = True
End If
End With
End Sub
Private Sub fixgrd1()
With grid1
.Cols = 7
.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«·Êﬁ "
.TextMatrix(0, 2) = "ﬂÊœ"
.TextMatrix(0, 3) = "«·’‰›"
.TextMatrix(0, 4) = "„»Ì⁄« "
.TextMatrix(0, 5) = "«·”⁄—"
.TextMatrix(0, 6) = "«·≈Ã„«·Ï"
.MergeCells = flexMergeFree
.MergeCol(0) = True

.ColWidth(0) = 1000
.ColWidth(1) = 800
.ColWidth(2) = 2000
.ColWidth(3) = 4000
.ColWidth(4) = 700
.ColWidth(5) = 700
.ColWidth(6) = 1200
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
End With
End Sub
Sub PrintTSales(dDate, cBox)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(1)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "⁄œœ «·»Ê‰« "
    temptable!Val3 = Val(GetDesca("SELECT COUNT(DOC_NO) FROM FILE6_20H WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update


    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "⁄œœ «·«’‰«›"
    temptable!Val3 = Val(GetDesca("SELECT SUM(T_QUANT) FROM T_SALESDOC WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "ﬁÌ„… „»Ì⁄« "
    temptable!Val3 = Val(GetDesca("SELECT SUM(T_TOTAL) FROM T_SALESDOC WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "Œ’„ „»Ì⁄« "
    temptable!Val3 = Val(GetDesca("SELECT SUM(DISCOUNT) FROM T_SALESDOC WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "’«›Ï „»Ì⁄« "
    temptable!Val3 = Val(GetDesca("SELECT SUM(T_TOTAL -DISCOUNT) FROM T_SALESDOC WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "”œ«œ ‰ﬁœÏ"
    NTCASH = Val(GetDesca("SELECT SUM(CASH) FROM T_SALESDOC WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable!Val3 = NTCASH
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "”œ«œ ›Ì“«"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VISA) FROM T_SALESDOC WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "1 - ≈Ã„«·Ï «·„»Ì⁄« "
    temptable!str2 = xBox.Text
    temptable!str5 = "”œ«œ √Ã·"
    temptable!Val3 = Val(GetDesca("SELECT SUM(T_TOTAL - DISCOUNT - CASH - VISA) FROM T_SALESDOC WHERE DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update

' ************************

'    temptable.AddNew
'    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
'    temptable!str1 = "12 - ≈Ã„«·Ï «·‰ﬁœÌ… "
'    temptable!str2 = xBox.Text
'    temptable!str5 = "—’Ìœ √Ê· «·ÌÊ„"
'    temptable!val3 = Val(GetDesca("SELECT SUM(PLUS - MINUS) FROM BOXMOVE WHERE DATE < " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
'    temptable.Update

    ntday = 0
    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "„»Ì⁄«  ‰ﬁœÏ"
    temptable!Val3 = NTCASH
    temptable.Update
    ntday = ntday + NTCASH
    
    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "≈Ã„«·Ï „’«—Ì›"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VALUE) FROM FILE8_50 INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO WHERE FILE8_50H.DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    ntday = ntday - Val(temptable!Val3)
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "”œ«œ „Ê—œÌ‰"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VALUE) FROM FILE8_20 INNER JOIN FILE8_20H ON FILE8_20.DOC_NO = FILE8_20H.DOC_NO WHERE FILE8_20H.DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    ntday = ntday - Val(temptable!Val3)
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "”œ«œ ⁄„·«¡"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VALUE) FROM FILE8_10 INNER JOIN FILE8_10H ON FILE8_10.DOC_NO = FILE8_10H.DOC_NO WHERE FILE8_10H.DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    ntday = ntday + Val(temptable!Val3)
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "„”ÕÊ»«  ‘—ﬂ«¡"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VALUE_M) FROM FILE8_70 INNER JOIN FILE8_70H ON FILE8_70.DOC_NO = FILE8_70H.DOC_NO WHERE FILE8_70H.DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    ntday = ntday - Val(temptable!Val3)
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "≈Ìœ«⁄«  ‘—ﬂ«¡"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VALUE_P) FROM FILE8_70 INNER JOIN FILE8_70H ON FILE8_70.DOC_NO = FILE8_70H.DOC_NO WHERE FILE8_70H.DATE = " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    ntday = ntday + Val(temptable!Val3)
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "’«œ— ‰ﬁœÌ…"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VALUE) FROM FILE0_51 WHERE DATE = " & DateSq(dDate) & " AND NO1 = " & MyParn(cBox)) & "")
    ntday = ntday - Val(temptable!Val3)
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "2 - ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "Ê«—œ ‰ﬁœÌ…"
    temptable!Val3 = Val(GetDesca("SELECT SUM(VALUE) FROM FILE0_51 WHERE DATE = " & DateSq(dDate) & " AND NO2 = " & MyParn(cBox)) & "")
    ntday = ntday + Val(temptable!Val3)
    temptable.Update




    temptable.AddNew
    temptable!str3 = " ≈Ã„«·Ï ÌÊ„ " & dDate
    temptable!str1 = "4- ≈Ã„«·Ï «·‰ﬁœÌ… "
    temptable!str2 = xBox.Text
    temptable!str5 = "—’Ìœ √Œ— «·ÌÊ„"
    temptable!Val3 = ntday
'    temptable!val3 = Val(GetDesca("SELECT SUM(PLUS - MINUS) FROM BOXMOVE WHERE DATE < " & DateSq(dDate) & " AND BOX = " & MyParn(cBox)) & "")
    temptable.Update


contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\TDAY.RPT"
main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.REPORT1.Action = 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xitem_GotFocus()
myGotFocus xItem
End Sub
Private Sub xDesca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then myload1
End Sub

Private Sub xITEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then myload1
End Sub

Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Function CheckOpen() As Boolean
Dim cString As String, nCount As Long
cString = "Select count(*) from file6_20h"
cString = cString & turn(cString) & "File6_20h.printed = 0"
cString = cString & turn(cString) & " [DATE] = " & DateSq(dSalesDate)
cString = cString & turn(cString) & "File6_20h.BOX = " & MyParn(xBox.BoundText)
nCount = Val(GetDesca(cString))
If nCount > 0 Then
    MsgBox "Â‰«ﬂ ⁄œœ " & nCount & " »Ê‰«  »Ì⁄ „› ÊÕ…!!«·—Ã«¡ «·Õ–› «Ê «· ”ÃÌ·", vbCritical
    CheckOpen = nCount
    Exit Function
End If
CheckOpen = True
End Function

