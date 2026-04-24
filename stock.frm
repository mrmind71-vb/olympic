VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StockFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã—œ „Œ«“‰"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14835
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   14835
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExel 
      Height          =   600
      Left            =   90
      Picture         =   "stock.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "⁄—÷"
      Top             =   1035
      Width           =   1545
   End
   Begin VB.Frame Frame7 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   8370
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   35
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
         Picture         =   "stock.frx":27EB
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "stock.frx":49BB
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   36
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
         Picture         =   "stock.frx":6B03
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "stock.frx":8CCB
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   37
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
         Picture         =   "stock.frx":AE1A
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "stock.frx":CFFA
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   38
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
         Picture         =   "stock.frx":F155
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "stock.frx":11311
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   9945
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   8370
      Width           =   4740
      Begin VB.CommandButton cmdPost 
         BackColor       =   &H00E0E0E0&
         Caption         =   " —ÕÌ·"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdunPost 
         BackColor       =   &H00E0E0E0&
         Caption         =   "≈·€«¡  —ÕÌ·"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdFix 
         BackColor       =   &H00E0E0E0&
         Caption         =   "«⁄«œ… ÷»ÿ «·Ã—œ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1185
         RightToLeft     =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton CmdAdditem 
         Caption         =   "«÷«›… ’‰›"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   1095
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6630
      Left            =   90
      TabIndex        =   23
      Top             =   1665
      Width           =   14640
      _cx             =   25823
      _cy             =   11695
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
      Cols            =   9
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
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   9270
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   0
      Width           =   5460
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":13460
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":1587E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":18118
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "√÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4095
         Picture         =   "stock.frx":1A6C4
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame frmProg1 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5535
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8325
      Width           =   2670
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   465
         Left            =   90
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   820
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "»ÕÀ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   8280
      Width           =   2085
      Begin VB.TextBox xfilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   270
         Width           =   1905
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1140
      Left            =   5175
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   495
      Width           =   1365
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":1CE97
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":1F1FA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "«·ﬁ«∆„…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   450
      Width           =   3480
      Begin VB.CommandButton cmdaddList 
         Caption         =   "«÷«›…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   3300
      End
      Begin MSDataListLib.DataCombo xList 
         Height          =   345
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   2670
      Begin VB.CommandButton cmdAdd 
         Caption         =   "«÷«›… «’‰«› ·Â« —’Ìœ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   135
         Width           =   2580
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   630
      Width           =   8160
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
         Left            =   3870
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   3165
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   585
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin VB.TextBox xDoc_No 
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
         Left            =   5670
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xDate 
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
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   2040
      End
      Begin VB.Label Label4 
         Caption         =   "»Ì«‰ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Œ“‰ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   570
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   2325
      Top             =   1725
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   5130
      Top             =   2835
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   1125
      Top             =   3060
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
   Begin VB.Frame frmTotals 
      Caption         =   "«·«Ã„«·Ì"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8325
      Width           =   1725
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   270
         Width           =   1545
      End
   End
End
Attribute VB_Name = "StockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As ADODB.Recordset, oSearchDoc As New Search3, osearchitem As New Search3
Dim clist1 As String
Dim itemBalTable As ADODB.Recordset
Dim formMode
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional nRow As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[STORE]", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xDesca.Text))
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE0_10H", "doc_no")))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, "FILE0_10h")
Else
    con.Execute addUpdate(aInsert, "FILE0_10h", "doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd nRow, bNewOnly
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 0) = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        Unload osearchitem
        CellPos 13, grid1.Row, grid1.Col
    Else
        grid1.Row = grid1.Rows - 1
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    oSearchDoc.Hide
ElseIf ActiveControl.Name = xdocImp.Name Then
    xdocImp.Text = searchImp.grid1.TextMatrix(searchImp.grid1.Row, 0)
    searchImp.Hide
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmd_item_Click()
Load frmItem
frmItem.Show 1
End Sub
Private Sub CmdAdd_Click()
If MsgBox("Â·  Êœ «÷«›… «’‰«› ·Â« —’Ìœ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
Dim loctable As New ADODB.Recordset, nRecordCount As Integer
cString = "SELECT Sum(FILE1_11.[IN] - FILE1_11.[out]) AS Balance,FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_10.COST" & _
          " FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
          " WHERE FILE1_11.DATE <= " & DateSq(xDate.Text)
cString = cString & turn(cString) & " FILE1_11.STORE = " & MyParn(xStore.BoundText)
cString = cString & " GROUP BY FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_10.COST"
cString = cString & turn(cString, " HAVING ", " AND ") & " Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
With grid1
    prog1.Visible = True
    prog1.Value = 0
    Do Until loctable.EOF
        If grid1.FindRow(loctable!Item, , 0) = -1 Then
            I = I + 1
            prog1.Value = Round(I / (nRecordCount), 2) * 100
            grid1.TextMatrix(.Rows - 1, 0) = loctable!Item
            grid1.TextMatrix(.Rows - 1, 1) = loctable!Desca & ""
            grid1.TextMatrix(.Rows - 1, 3) = Val(loctable!BALANCE & "")
            grid1.TextMatrix(.Rows - 1, 4) = 0
            grid1.TextMatrix(.Rows - 1, 5) = -1 * Val(loctable!BALANCE & "")
            grid1.TextMatrix(.Rows - 1, 6) = Val(loctable!cost & "")
            grid1.TextMatrix(.Rows - 1, 7) = -1 * loctable!BALANCE * Val(loctable!cost & "")
            grid1.AddItem ""
        End If
        loctable.MoveNext
    Loop
    prog1.Visible = False
End With
If Not MYVALID Then Exit Sub
If Not myreplace(, True) Then Exit Sub
openCardTable
myUndo
End Sub
Private Sub CmdAddList_Click()
If Not xList.MatchedWithList Then Exit Sub
If MsgBox("Â·  Êœ «÷«›… ﬁ«∆„… «·Ã—œ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
Dim loctable As New ADODB.Recordset, nRecordCount As Integer
Dim cString As String, nCost As Double
cString = "SELECT FILE1_10.ITEM,FILE1_10.DESCA" & _
          " FROM FILE1_10 INNER JOIN ST_LIST ON FILE1_10.ITEM = ST_LIST.ITEM"
If xList.MatchedWithList Then
    cString = cString & turn(cString) & " ST_LIST.DOC_NO = " & MyParn(xList.BoundText)
End If
cString = cString & " GROUP BY FILE1_10.ITEM,FILE1_10.DESCA"

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
With grid1
    prog1.Visible = True
    prog1.Value = 0
    Do Until loctable.EOF
        If grid1.FindRow(loctable!Item, , 0) = -1 Then
            I = I + 1
            prog1.Value = Round(I / (nRecordCount), 2) * 100
            grid1.TextMatrix(.Rows - 1, 0) = loctable!Item
            GrdDesc grid1.Rows - 1, 0
            grid1.AddItem ""
        End If
        loctable.MoveNext
    Loop
    prog1.Visible = False
End With
If Not MYVALID Then Exit Sub
If Not myreplace(, True) Then Exit Sub
openCardTable
myUndo
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE0_10 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From FILE0_10H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    If (CardTable.EOF And CardTable.BOF) Then
        mydefine
    Else
        CardTable.Find "doc_no < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub

Private Sub cmdExel_Click()
ToFileExel grid1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdFix_Click()
If Not MYVALIDfix Then Exit Sub
With grid1
'Time = Now
prog1.Value = 0
prog1.Visible = True
For I = 1 To grid1.Rows - 2
     prog1.Value = Round(I / (grid1.Rows - 2), 2) * 100
    .TextMatrix(I, 3) = RetItemBalance(.TextMatrix(I, 0), xStore.BoundText, xDate.Text, con) & ""
    .TextMatrix(I, 5) = Val(.TextMatrix(I, 4)) - Val(.TextMatrix(I, 3))
    .TextMatrix(I, 7) = Val(.TextMatrix(I, 5)) * Val(.TextMatrix(I, 6))
Next
prog1.Visible = False
If myreplace Then
'MsgBox DateDiff("s", nTime, Now)
    Inform " „ ÷»ÿ «·„” ‰œ »‰Ã«Õ"
End If
End With
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE,CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA " & _
                  " FROM FILE0_10H INNER JOIN FILE0_40 ON FILE0_10H.Store = FILE0_40.CODE "

Generalarray(2) = "Order by [Date]"
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·„Œ“‰"
listarray(0, 1) = "@@Doc_No@@6 or  FILE0_40.DESCA Like '%cFilter%' OR " & _
                  "##[DATE]##"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·„Œ“‰"
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub cmdList_Click()
If Not MYVALID Then Exit Sub
If Not myreplace(, True) Then Exit Sub
openCardTable
myUndo
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
Private Sub cmdPost_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload

On Error GoTo myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 1 WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans

openCardTable
myUndo
MsgBox " „  —ÕÌ· «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdNewInv_Click()
'CardTable.MoveLast
'xDoc_No.Text = RetZero(Val(CardTable!doc_no & ""))
mydefine
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
openCardTable
myUndo
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub cmdunPost_Click()
On Error GoTo myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans
openCardTable
myUndo
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub Command3_Click()
Load PrintGrd
grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = xTotal.Caption
grid1.TextMatrix(grid1.Rows - 1, 0) = "«·«Ã„«·Ì"
PrintGrd.doprint grid1, 1, True
PrintGrd.Show 1
grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = ""
grid1.TextMatrix(grid1.Rows - 1, 0) = ""
End Sub
Private Sub Command1_Click()
doprint "stock1.rpt"
End Sub

Private Sub Command2_Click()
If MsgBox("ÿ»«⁄… »Ì«‰ „⁄ «·«—’œ…", vbYesNo) = vbYes Then
    doprint "stock2.rpt"
Else
    doprint "stock1.rpt"
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
If xDoc_No.Tag = LoadMode Then grid1.SetFocus
Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And (ActiveControl.Name <> xfilter.Name) Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
'clist1 = StrList("select * from file1_50 order by desca")
'cList2 = StrList("select * from file1_10SC order by desca")

Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM File0_10H  ORDER BY DOC_NO ", con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

Set grid1.DataSource = data3
data3.ConnectionString = strCon

data4.ConnectionString = strCon
data4.RecordSource = "ST_LISTH"
Set xList.RowSource = data4
xList.ListField = "Desca"
xList.BoundColumn = "DOC_NO"

'data2.ConnectionString = strCon
'data2.RecordSource = "SELECT Distinct Doc_No, DESCA FROM FILE0_30"

'Set xList.RowSource = data2
'xList.ListField = "Desca"
'xList.BoundColumn = "Doc_No"
'grid1.ColHidden(2) = True
Fixgrd
openCardTable
myUndo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload oSearchDoc
Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim bNew As Boolean
If Col = 0 Then GrdDesc Row, Col
If Not validRow(Row) Then Exit Sub
With grid1
If Row = grid1.Rows - 1 Then
    MyAddItem
    bNew = True
End If
Calctotals

If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then
        xDoc_No.Tag = LoadMode
        xDoc_No.Enabled = False
    End If
    If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then myloadgrd
End If
If bNew Then
    grid1.ShowCell grid1.Rows - 1, 1
End If
End With
End Sub
Private Sub grid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub

Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
ItemsLookupAll Me, osearchitem
End Sub

Private Sub Grid1_EnterCell()
If (grid1.Col = 4 Or grid1.Col = 0) Or grid1.Col >= 8 Then
  grid1.Editable = flexEDKbd
Else
   grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row < 0 Then
    grid1.SetFocus
    grid1.Select 1, 2
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, osearchitem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And cmdSave.Enabled Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from file0_10 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        Calctotals
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
'If foundOther Then Exit Function
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If


If Not bIgMsg Then
    If grid1.Rows < 3 Then
        MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
        Exit Function
    End If
    With grid1
    For I = 1 To .Rows - 2
        If .TextMatrix(I, 0) = "" Then
            .Select I, 0, I, grid1.Cols - 1
            MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
            Exit Function
        Else
            If GetDesca("select item from file1_10 where file1_10.item = " & MyParn(.TextMatrix(I, 0))) = "" Then
                .Select I, 0, I, 2
                MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
                Exit Function
            End If
        End If
    Next
    End With
End If
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xStore.BoundText = CardTable!store & ""
xDesca.Text = CardTable!Desca & ""
Handlecontrols LoadMode
myloadgrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myloadgrd()
cString = "SELECT FILE0_10.ITEM, FILE1_10.DESCA,'' AS DUMMY,FILE0_10.ComputerBal ,File0_10.RealBal,File0_10.differ,File0_10.cost,file0_10.Realbal * file0_10.Cost as Total,ID  " & _
      " FROM FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " order by FILE0_10.ROW"
data3.RecordSource = cString
data3.Refresh
grid1.AddItem ""
Calctotals
Fixgrd
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag("FILE0_10h", "doc_no", con)))
xDate.Text = Format(Date, "YYYY-MM-DD")
xStore.BoundText = ""
xDesca.Text = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
Dim bClosed As Boolean
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
If Not (CardTable.EOF And CardTable.BOF) Then bClosed = CardTable!CLOSED
cmdPost.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
cmdunPost.Enabled = bClosed And bEdit And nMode = LoadMode
cmdfix.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
cmdSave.Enabled = (nMode = DefineMode Or Not bClosed) And bEdit
CmdDelInv.Enabled = (Not bClosed) And bEdit And nMode = LoadMod
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then grid1.RemoveItem grid1.Row
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And Trim(grid1.EditText) <> "" Then
    nFound = FoundOtheritem(Row, Col, Trim(grid1.EditText))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & grid1.TextMatrix(nFound, 1)
        Cancel = True
    End If
    If Not validItem(grid1.EditText, con) Then
        MsgBox "ﬂÊœ «·’‰› €Ì— ”·Ì„"
        Cancel = True
        Exit Sub
    End If
ElseIf Col = 0 Then
    MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
    Cancel = True
End If
End Sub
Private Sub xDate_Change()
'cmdAddList.Enabled = xStore.BoundText <> "" And IsDate(xDate.Text)
End Sub

Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If (CardTable.EOF And CardTable.BOF) Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function MYVALIDfix() As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

If foundOther Then Exit Function
With grid1
For I = 1 To .Rows - 2
    If .TextMatrix(I, 0) = "" Then
        .Select I, 0, I, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Exit Function
    Else
        If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & MyParn(.TextMatrix(I, 0))) = "" Then
            .Select I, 0, I, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
End With
MYVALIDfix = True
End Function
Private Sub GrdDesc(Row, Col)
With grid1
.TextMatrix(Row, 1) = ""
.TextMatrix(Row, 2) = ""
.TextMatrix(Row, 3) = ""
.TextMatrix(Row, 5) = ""
.TextMatrix(Row, 6) = ""
.TextMatrix(Row, 7) = ""
If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Sub
Dim aRet As Variant
aRet = ItemFields(grid1.TextMatrix(Row, 0), con)
If Not IsEmpty(aRet) Then
    .TextMatrix(Row, 1) = retFlag(aRet, "desca") & ""
    .TextMatrix(Row, 2) = retFlag(aRet, "color") & ""
    If IsDate(xDate.Text) And Trim(xStore.BoundText) <> "" Then
        grid1.TextMatrix(Row, 3) = RetItemBalance(grid1.TextMatrix(Row, 0), xStore.BoundText, xDate.Text, con) & ""
    End If
    .TextMatrix(Row, 5) = Val(.TextMatrix(Row, 4)) - Val(.TextMatrix(Row, 3)) & ""
    .TextMatrix(Row, 6) = itemCost(grid1.TextMatrix(Row, 0), Format(xDate.Text, "YYYY-MM-DD"), con) & ""
    .TextMatrix(Row, 7) = Val(.TextMatrix(Row, 4)) * Val(.TextMatrix(Row, 6)) & ""
End If
End With
End Sub
Private Sub xfilter_Change()
FilterGrd xfilter.Text, 1
End Sub
Private Sub xLIST_LostFocus()
If Not xList.MatchedWithList Then xList.BoundText = ""
End Sub

Private Sub xStore_Click(Area As Integer)
'cmdAddList.Enabled = xStore.BoundText <> "" And IsDate(xDate.Text)
End Sub
Private Function RetTotalCost()
For I = 1 To grid1.Rows - 1
    RetTotalCost = RetTotalCost + Val(grid1.TextMatrix(I, 0))
Next
End Function
Private Sub Calctotals()
Dim nTotal As Double
With grid1
For I = 1 To grid1.Rows - 1
    .TextMatrix(I, 5) = Val(.TextMatrix(I, 4)) - Val(.TextMatrix(I, 3))
    .TextMatrix(I, 7) = (Val(grid1.TextMatrix(I, 4)) * Val(grid1.TextMatrix(I, 6)))
    nTotal = nTotal + (Val(grid1.TextMatrix(I, 4)) * Val(grid1.TextMatrix(I, 6)))
Next
End With
xTotal.Caption = Format(nTotal, "#0.00")
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ÊÕœ…|" & "—’Ìœ ﬂÊ„»ÌÊ —|" & "—’Ìœ Ã—œ|" & "«·›—ﬁ|" & " ﬂ·›…|" & "«· ﬁÌ„|"
.ColWidth(0) = 1000
.ColWidth(1) = 5000
.ColWidth(2) = 2500
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000

.ColDataType(0) = flexDTString
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColHidden(2) = True
.ColHidden(.Cols - 1) = True
For I = 0 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub doprint(cReport)
Dim loctable As New ADODB.Recordset
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset

loctable.Open "select file1_10.item,file1_10.[GROUP],file1_50.desca " & _
              " from file1_10 left join file1_50 on file1_10.[GROUP] = file1_50.code", con, adOpenStatic, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!date3 = DateFix(xDate.Text)
    temptable!str21 = turn(xDesca.Text, "", "Ã—œ ") & Trim(xDesca.Text) & turn(xDesca.Text, " ") & "·„Œ“‰ " & xStore.Text
    temptable!str1 = TurnValue(grid1.TextMatrix(I, 0))
    temptable!str2 = TurnValue(grid1.TextMatrix(I, 1))
    temptable!val1 = Val(grid1.TextMatrix(I, 3))
    temptable!val2 = Val(grid1.TextMatrix(I, 4))
    temptable!val11 = I
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\" & cReport
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Function foundOther() As Boolean
For I = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(I, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & I & " Ê " & nRow
        grid1.Select nRow, 0, nRow, grid1.Cols - 1
        If Not grid1.RowIsVisible(nRow) Then grid1.ShowCell nRow, 0
        foundOther = True
        Exit Function
    End If
Next
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = I
            Exit Function
        End If
    End If
Next
End Function
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function
Private Sub myreplaceGrd(nRow, Optional bNewOnly As Boolean = False)
Dim aInsert As Variant
With grid1
    Dim nCost As Double
    For I = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "ComputerBal", Val(grid1.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "RealBal", Val(grid1.TextMatrix(I, 4)))
        aInsert = AddFlag(aInsert, "Differ", Val(grid1.TextMatrix(I, 5)))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "COST", Val(grid1.TextMatrix(I, 6)))
            con.Execute addInsert(aInsert, "FILE0_10")
        ElseIf Not bNewOnly Then
            con.Execute addUpdate(aInsert, "FILE0_10", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Private Function validRow(nRow) As Boolean
If Not MYVALID(True) Then Exit Function
If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Function
validRow = True
End Function
Private Sub FilterGrd(cText, nCol)
Dim bHide As Boolean
For I = 1 To grid1.Rows - 2
     grid1.RowHidden(I) = False
Next
If Trim(cText) = "" Then Exit Sub

Dim aRet As Variant
aRet = Split(Trim(cText))
For I = 1 To grid1.Rows - 2
    For n = 0 To UBound(aRet)
        bHide = InStr(1, Trim(grid1.TextMatrix(I, nCol)), aRet(n)) = 0
        If bHide Then Exit For
    Next
    grid1.RowHidden(I) = bHide
Next
End Sub
Private Sub MyAddItem()
grid1.AddItem ""
End Sub
Private Sub openCardTable()
Set CardTable = New ADODB.Recordset
Dim cString As String
cString = "SELECT * FROM FILE0_10H"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xDoc_No.Text) <> "" Then
        CardTable.Find "DOC_NO = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 5 Then
    If Col < 4 Then
        grid1.Col = 4
    Else
        grid1.Col = Col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, IIf(Trim(grid1.TextMatrix(Row + 1, 0)) <> "", 4, 0)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub


