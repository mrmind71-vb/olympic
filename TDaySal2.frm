VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TDaySal2 
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox xDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8475
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   4425
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   1470
      Left            =   150
      TabIndex        =   19
      Top             =   1050
      Visible         =   0   'False
      Width           =   1455
      _cx             =   2566
      _cy             =   2593
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
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "≈Ã„«·Ï „»Ì⁄«  «·„ÊœÌ·«  ·ÌÊ„"
      TabPicture(0)   =   "TDaySal2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "VsSales"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " ›’Ì·Ï »Ê‰«  »Ì⁄ ·ÌÊ„"
      TabPicture(1)   =   "TDaySal2.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "VsSales2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid VsSales 
         Height          =   7260
         Left            =   -74925
         TabIndex        =   1
         Top             =   375
         Width           =   7695
         _cx             =   13573
         _cy             =   12806
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
      Begin VSFlex7LCtl.VSFlexGrid VsSales2 
         Height          =   7290
         Left            =   75
         TabIndex        =   2
         Top             =   375
         Width           =   7845
         _cx             =   13838
         _cy             =   12859
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
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2340
      Left            =   8040
      TabIndex        =   3
      Top             =   120
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   4128
      _Version        =   196610
      BackColor       =   13822956
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand Day_Sale 
         Height          =   390
         Left            =   2025
         TabIndex        =   4
         Top             =   50
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   192
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ï ⁄œœ »Ê‰«  «·»Ì⁄"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   390
         Left            =   2025
         TabIndex        =   5
         Top             =   930
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   192
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ï ﬁÌ„… «·„»Ì⁄« "
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   390
         Left            =   2025
         TabIndex        =   6
         Top             =   510
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   192
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ï ﬂ. „»Ì⁄« "
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand X5 
         Height          =   390
         Left            =   795
         TabIndex        =   7
         Top             =   1800
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   12089119
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   390
         Left            =   2025
         TabIndex        =   8
         Top             =   1365
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   192
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ﬁÌ„… «·Œ’‹‹‹„"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   390
         Left            =   2025
         TabIndex        =   9
         Top             =   1800
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   192
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "’«›Ï ﬁÌ„… «·„»Ì⁄« "
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand X2 
         Height          =   390
         Left            =   795
         TabIndex        =   10
         Top             =   510
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   12089119
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand X3 
         Height          =   390
         Left            =   795
         TabIndex        =   11
         Top             =   930
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   12089119
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand X4 
         Height          =   390
         Left            =   795
         TabIndex        =   12
         Top             =   1365
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   12089119
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand X1 
         Height          =   390
         Left            =   795
         TabIndex        =   13
         Top             =   50
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   12089119
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand X40 
         Height          =   390
         Left            =   90
         TabIndex        =   14
         Top             =   1365
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   12089119
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         ButtonStyle     =   2
      End
   End
   Begin MSDBCtls.DBCombo xStore 
      Bindings        =   "TDaySal2.frx":0038
      DataSource      =   "Data2"
      Height          =   330
      Left            =   8040
      TabIndex        =   15
      Top             =   3690
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   13822956
      ForeColor       =   12089119
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
   Begin Threed.SSCommand CMD_END 
      Height          =   540
      Left            =   8025
      TabIndex        =   18
      Top             =   3100
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   953
      _Version        =   196610
      Font3D          =   1
      ForeColor       =   192
      BackColor       =   13822956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÿ»«⁄… „»Ì⁄«  «·ÌÊ„ & ≈€·«ﬁ «·ÌÊ„Ì…"
   End
   Begin Threed.SSCommand CMD_Print 
      Height          =   540
      Left            =   8025
      TabIndex        =   17
      Top             =   2510
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   953
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
      BackColor       =   12440240
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Traditional Arabic"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Print"
      ButtonStyle     =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Œ“‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B8771F&
      Height          =   195
      Left            =   10935
      TabIndex        =   16
      Top             =   3765
      Width           =   540
   End
End
Attribute VB_Name = "TDaySal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TSalTable As Recordset
Dim CountTable As Recordset
Dim SalTable As Recordset
Dim Sal2Table As Recordset
Dim TempTable As Recordset
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub xmGroup_Change()
End Sub
Private Sub cmd_close_Click()
'    cString = " UPDATE FILE6_20 SET FILE6_20.CLOSED = TRUE  Where Date = " & DateSql(XDATE.TEXT)
'    mydb.Execute cString
End Sub
Private Sub CmdOk_Click()
VsSales.Rows = 1
VsSales2.Rows = 1
Dim nX1 As Double, nX2 As Double, nX3 As Double, nX4 As Double, nX40 As Double, nX5 As Double
cStr1 = " SELECT FILE6_20.DOC_NO FROM FILE6_20 WHERE FILE6_20.Date = " & DateSql(xDate.Text)
If xStore.BoundText <> "" Then cStr1 = cStr1 & " and FILE6_20.STORE2 = " & MyParn(xStore.BoundText)
cStr1 = cStr1 & " GROUP BY FILE6_20.DOC_NO "

Set CountTable = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT Sum(FILE6_20.QUANT) AS TQUANT, Sum(FILE6_20.TOTAL) AS TTOTAL, FILE6_20.DOC_NO, FILE6_20.STORE , SUM([QUANT] * [PRICE]) AS TPRICE FROM FILE6_20 " & _
        " WHERE FILE6_20.Date = " & DateSql(xDate.Text)
If xStore.BoundText <> "" Then cStr1 = cStr1 & " and FILE6_20.STORE2 = " & MyParn(xStore.BoundText)
cStr1 = cStr1 & " GROUP BY FILE6_20.DOC_NO, FILE6_20.STORE "
Set TSalTable = mydb.OpenRecordset(cStr1, dbOpenDynaset)

cStr1 = " SELECT FIRST(FILE1_10.DESCA) AS DESCA , FILE1_10.ITEM, Sum(FILE6_20.[QUANT]) AS T_QUANT, Sum(FILE6_20.TOTAL) AS T_TOTAL  " & _
        " FROM FILE1_10 RIGHT JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM " & _
        " WHERE file1_10.ITEM IS NOT NULL   " & _
        " AND FILE6_20.Date = " & DateSql(xDate.Text)
If xStore.BoundText <> "" Then cStr1 = cStr1 & " and FILE6_20.STORE2 = " & MyParn(xStore.BoundText)
cStr1 = cStr1 & " GROUP BY FILE1_10.ITEM  "

cStr2 = " SELECT FILE6_20.* , FILE1_10.DESCA  " & _
        " FROM FILE1_10 RIGHT JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM " & _
        " WHERE FILE6_20.Date = " & DateSql(xDate.Text)

cStr2 = cStr2 & " ORDER BY DOC_NO  "
With TSalTable
    If .RecordCount > 0 Then
        Do While Not .EOF
            If .Store <> "zz" Then
                nX2 = nX2 + TurnValue(.TQUANT, Null, 0)
                nX3 = nX3 + TurnValue(.TPRICE, Null, 0)
            Else
                nX4 = nX4 + TurnValue(.TTOTAL, Null, 0)
            End If
            nX5 = nX5 + TurnValue(.TTOTAL, Null, 0)
            .MoveNext
        Loop
    End If
End With

If CountTable.RecordCount > 0 Then
    CountTable.MoveFirst
    X1.Caption = Format(CountTable.RecordCount, "#0")
    X2.Caption = Format(nX2, "#0")
    X3.Caption = Format(nX3, "#0.00")
    X4.Caption = Format(nX4, "#0.00")
    If X4 > 0 Then X40.Caption = "%" & Format(nX3 / nX4 * 100, "#0.00")
    X5.Caption = Format(nX5, "#0.00")
    Set SalTable = mydb.OpenRecordset(cStr1, dbOpenDynaset)
    VsSales.Rows = 1
    If SalTable.RecordCount = 0 Then Exit Sub
    With VsSales
        .Rows = 1
        .Sort = flexSortNone
        SalTable.MoveFirst
        Do While True
            .AddItem ""
            
            .TextMatrix(.Rows - 1, 0) = DelZero(SalTable.Item)
            .TextMatrix(.Rows - 1, 1) = TurnValue(SalTable.DESCA, Null, "")
            .TextMatrix(.Rows - 1, 5) = Format(SalTable.T_QUANT, "#0")
            .TextMatrix(.Rows - 1, 6) = Format(SalTable.T_TOTAL, "#0.00")
            SalTable.MoveNext
            If SalTable.EOF Then Exit Do
        Loop
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 5, "#0", , vbBlue
        .Subtotal flexSTSum, -1, 6, "#0.00", , vbBlue
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End If

Set Sal2Table = mydb.OpenRecordset(cStr2, dbOpenDynaset)
If Sal2Table.RecordCount > 0 Then
    VsSales2.Rows = 1
    With VsSales2
        .Rows = 1
        .Sort = flexSortNone
        Sal2Table.MoveFirst
        Do While True
            .AddItem ""
            .TextMatrix(.Rows - 1, 1) = Sal2Table.DOC_NO
'           .TextMatrix(.Rows - 1, 2) = TurnValue(Sal2Table.ManName, Null, "")
            If Sal2Table.Store = "zz" Then
                .TextMatrix(.Rows - 1, 4) = "Œ’„ «·»Ê‰"
                .Cell(flexcpBackColor, .Rows - 1, 4, .Rows - 1, .Cols - 1) = &HD3BD78
            Else
                .TextMatrix(.Rows - 1, 3) = TurnValue(Sal2Table.Item, Null, "")
                .TextMatrix(.Rows - 1, 4) = TurnValue(Sal2Table!DESCA, Null, "")
                If Sal2Table.Quant > 0 Then
                    .TextMatrix(.Rows - 1, 5) = Format(Sal2Table.Quant, "#0")
                Else
                    .TextMatrix(.Rows - 1, 6) = Format(Sal2Table.Quant, "#0")
                End If
                .TextMatrix(.Rows - 1, 7) = Format(Sal2Table.price, "#0.00")
            End If
            .TextMatrix(.Rows - 1, 8) = Format(Sal2Table.total, "#0.00")
            Sal2Table.MoveNext
            If Sal2Table.EOF Then Exit Do
        Loop
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 5, "#0", , vbBlue, , "Ã. «·ÌÊ„"
        .Subtotal flexSTSum, -1, 6, "#0", , vbBlue, , "Ã. «·ÌÊ„"
        .Subtotal flexSTSum, -1, 8, "#0.00", , vbBlue, , "Ã. «·ÌÊ„"
        
        .Subtotal flexSTSum, 1, 5, "#0", &HE0E0E0, vbRed, ""
        .Subtotal flexSTSum, 1, 6, "#0", &HE0E0E0, vbRed, ""
        .Subtotal flexSTSum, 1, 8, "#0.00", &HE0E0E0, vbRed, , ""
                
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
            
                        
    End With
End If
End Sub
Private Sub Form_Load()
xDate.Visible = bopt1
CMD_Print.Visible = bopt1
tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("TEMP", dbOpenDynaset)

xDate.Text = Date
VsSales.Rows = 1
VsSales.Cols = 10
VsSales.FixedRows = 1
VsSales.FixedCols = 0
VsSales.ColWidth(0) = 1000
VsSales.ColWidth(1) = 3000
VsSales.ColWidth(2) = 0
VsSales.ColWidth(3) = 0
VsSales.ColWidth(4) = 0
VsSales.ColWidth(5) = 1200
VsSales.ColWidth(6) = 1200
VsSales.ColWidth(7) = 0
VsSales.ColWidth(8) = 0
VsSales.ColWidth(9) = 0

With VsSales2
.Cols = 9
.FixedRows = 1
.OutlineBar = 1
.ColWidth(0) = 300
.ColWidth(1) = 1700
.ColWidth(2) = 0
.ColWidth(3) = 800
.ColWidth(4) = 1500
.ColWidth(5) = 600
.ColWidth(6) = 600
.ColWidth(7) = 700
.ColWidth(8) = 900
.TextMatrix(0, 1) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 3) = "ﬂÊœ"
.TextMatrix(0, 4) = "«·’‰›"
.TextMatrix(0, 5) = "„»Ì⁄« "
.TextMatrix(0, 6) = "„— Ã⁄"
.TextMatrix(0, 7) = "«·”⁄—"
.TextMatrix(0, 8) = "«·≈Ã„«·Ï"
.MergeCells = flexMergeFree
.MergeCol(1) = True
.MergeCol(2) = True
End With

VsSales.TextMatrix(0, 0) = "ﬂÊœ"
VsSales.TextMatrix(0, 1) = " «·≈”„"
VsSales.TextMatrix(0, 5) = "⁄œœ „»Ì⁄« "
VsSales.TextMatrix(0, 6) = "ﬁÌ„…"

VsSales.RowHeight(0) = 400
VsSales.Sort = flexSortNone
VsSales.ExplorerBar = flexExSortShow

xDate.Text = DateValue(Vs_Inv.xDate.Text)

CmdOk_Click
End Sub
Private Sub Cmd_Print_Click()
    Load REP1_27
    REP1_27.Date1.Text = Format(xDate.Text, "dd-mm-yyyy")
    REP1_27.date2.Text = Format(xDate.Text, "dd-mm-yyyy")
    REP1_27.Show 1
End Sub
Private Sub SSCommand3_Click()
    Load Rep_107
    Rep_107.xDate1.Text = xDate.Text
    Rep_107.xDate2.Text = xDate.Text
    Rep_107.xStore.BoundText = xStore.BoundText
    Rep_107.Show 1
End Sub
Private Sub VsSales_DBLClick()
    ViewSale.Show 1
    CmdOk_Click
End Sub
Private Sub xDate_BeforeUpdate(Cancel As Integer)
    CmdOk_Click
End Sub
Private Sub xDate_Click()
    CmdOk_Click
End Sub
Private Sub xStore_Click(Area As Integer)
    CmdOk_Click
End Sub
Private Sub CMD_END_Click()
'On Error GoTo myerror
Dim cStr1 As String
Dim SalTable As Recordset
Dim nTQunat As Double
Dim nTtotal As Double
Dim nTdisc As Double
If MsgBox("ÿ»«⁄… , ≈€·«ﬁ ÌÊ„Ì… «·»Ì⁄ ", vbOKCancel) = vbCancel Then Exit Sub
nTtotal = 0
nTdisc = 0
nTQunat = 0
 
 cField4 = myiif(" STORE = 'zz' " _
           , "val(Format([TOTAL ]))") & _
           " As DISC "

cStr1 = " SELECT FILE6_20.DOC_NO, Min(FILE6_20.STORE) AS STORE, FILE6_22.CASH, FILE6_22.VISA, FILE6_22.total FROM FILE6_20 LEFT JOIN FILE6_22 ON FILE6_20.DOC_NO = FILE6_22.DOC_NO WHERE FILE6_20.DATE = " & DateSql(xDate.Text) & "  GROUP BY FILE6_20.DOC_NO, FILE6_22.CASH, FILE6_22.VISA, FILE6_22.total ORDER BY FILE6_20.DOC_NO "
Set SalTable = mydb.OpenRecordset(cStr1)
With vp
     vp = " "
'   .Device = "GP-80160II"
     .Device = pDevice

    .StartDoc
    .TextAlign = taRightMiddle
    .TextColor = vbBlack
    
    .Width = 3700
    .Height = 3000
    .MarginLeft = 100
    .MarginRight = 100
    .MarginTop = 10
    .MarginFooter = 10
    .MarginBottom = 100
    
    .FontSize = 14
    .FontName = "Traditional Arabic"
    .FontBold = True
    .TextAlign = taCenterMiddle
    .Paragraph = "**  " & firsttitle & "  **"
    .FontSize = 10
    .TextAlign = taRightMiddle
    .Paragraph = "„·Œ’ »Ê‰«  »Ì⁄ ÌÊ„ " & NameOfDay(xDate.Text) & " «·„Ê«›ﬁ : " & Format(xDate.Text, "dd-mm-yy")
    
    .FontSize = 9
    .FontBold = False
    .StartTable
    .TableCell(tcRows) = 1
    .TableCell(tcCols) = 4
    .TableCell(tcColWidth, 1, 4) = 800
    .TableCell(tcColWidth, 1, 3) = 700
    .TableCell(tcColWidth, 1, 2) = 900
    .TableCell(tcColWidth, 1, 1) = 1200
    .BorderStyle = bsNone
    
    .TableBorder = tbTopBottom
    .TableCell(tcRowBorderAbove, 1, 1, 1, 5) = 4
    .TableCell(tcText, 1, 4) = "»Ê‰ —ﬁ„"
    .TableCell(tcText, 1, 3) = "„Õ·"
    .TableCell(tcText, 1, 2) = "≈Ã„«·Ï"
    .TableCell(tcText, 1, 1) = "ﬂ«‘"
    If SalTable.RecordCount = 0 Then Exit Sub
    SalTable.MoveFirst
    Do While Not SalTable.EOF
        .TableCell(tcRows) = .TableCell(tcRows) + 1
        .TableCell(tcText, .TableCell(tcRows), 4) = SalTable.DOC_NO
        .TableCell(tcText, .TableCell(tcRows), 3) = SalTable.Store
        .TableCell(tcText, .TableCell(tcRows), 2) = Format(SalTable.total, "#0.00")
        .TableCell(tcText, .TableCell(tcRows), 1) = Format(SalTable.CASH, "#0.00")
        nTtotal = nTtotal + TurnValue(SalTable.CASH, Null, 0)
'       nTdisc = nTDisv + TurnValue(SalTable.DISC, Null, 0)
'       nTQunat = nTQunat + TurnValue(SalTable.TQUANT, Null, 0)
        SalTable.MoveNext
    Loop
    .TableCell(tcRowBorderBelow, 1, 1, 1, 4) = 4

    .TableCell(tcRows) = .TableCell(tcRows) + 1
    .TableCell(tcText, .TableCell(tcRows), 4) = "≈Ã„«·Ï"
'    .TableCell(tcText, .TableCell(tcRows), 3) = Format(nTQunat, "#0")
'    .TableCell(tcText, .TableCell(tcRows), 2) = Format(nTdisc, "#0.00")
    .TableCell(tcText, .TableCell(tcRows), 1) = Format(nTtotal, "#0.00")
    .TableCell(tcRowBorder, .TableCell(tcRows), 1, .TableCell(tcRows), 4) = 4
    .EndTable
    
    
    cStr1 = " SELECT Sum(MOVEBOX.boxin) AS TSAL , MOVEBOX.box, FILE0_50.DESCA FROM MOVEBOX LEFT JOIN FILE0_50 ON MOVEBOX.box = FILE0_50.CODE WHERE MOVEBOX.TBOX = 4 AND DATE = " & DateSql(xDate.Text)
    cStr1 = cStr1 & " GROUP BY MOVEBOX.box, FILE0_50.DESCA "
    Set SalTable = mydb.OpenRecordset(cStr1)
    
    
    .StartTable
    .TableCell(tcRows) = 1
    .TableCell(tcCols) = 2
    .TableCell(tcColWidth, 1, 2) = 2000
    .TableCell(tcColWidth, 1, 1) = 1500
    .BorderStyle = bsNone
    
    .TableBorder = tbTopBottom
    .TableCell(tcRowBorderAbove, 1, 1, 1, 2) = 4
    .TableCell(tcText, 1, 2) = "„Œ“‰"
    .TableCell(tcText, 1, 1) = "≈Ã„«·Ï "
    If SalTable.RecordCount = 0 Then Exit Sub
    SalTable.MoveFirst
    nTtotal = 0
    Do While Not SalTable.EOF
        .TableCell(tcRows) = .TableCell(tcRows) + 1
        .TableCell(tcText, .TableCell(tcRows), 2) = TurnValue(SalTable.DESCA, Null, "")
        .TableCell(tcText, .TableCell(tcRows), 1) = Format(SalTable.TSAL, "#0.00")
        nTtotal = nTtotal + TurnValue(SalTable.TSAL, Null, 0)
        SalTable.MoveNext
    Loop
    .TableCell(tcRowBorderBelow, 1, 1, 1, 2) = 4

    .TableCell(tcRows) = .TableCell(tcRows) + 1
    .TableCell(tcText, .TableCell(tcRows), 2) = "≈Ã„«·Ï"
    .TableCell(tcText, .TableCell(tcRows), 1) = Format(nTtotal, "#0.00")
    .TableCell(tcRowBorder, .TableCell(tcRows), 1, .TableCell(tcRows), 4) = 4
    .EndTable
    
    .EndDoc
End With
cString = " UPDATE FILE6_20 SET FILE6_20.[POSTED] = TRUE  Where Date = " & DateSql(xDate.Text)
mydb.Execute cString

Exit Sub
myerror:
MsgBox Err.Description
MsgBox "Try Again This Order "
Err.Clear
Unload Me
End Sub



