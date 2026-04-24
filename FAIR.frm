VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form fairfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√ﬂÊ«œ «·«’‰«›"
   ClientHeight    =   9285
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   14160
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   14160
   Begin VB.Frame Frame2 
      Height          =   9195
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
      Width           =   6945
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8925
         Left            =   45
         TabIndex        =   5
         Top             =   180
         Width           =   6855
         _cx             =   12091
         _cy             =   15743
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   675
      Width           =   6990
      Begin VB.CommandButton cmdContainer 
         Caption         =   "..."
         Height          =   330
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1350
         Width           =   345
      End
      Begin VB.TextBox xWeight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1350
         Width           =   2940
      End
      Begin VB.CommandButton cmdTrailer 
         Caption         =   "..."
         Height          =   330
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox xClient 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4050
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo xTrailer 
         Height          =   330
         Left            =   2520
         TabIndex        =   2
         Tag             =   "LL"
         Top             =   990
         Width           =   2940
         _ExtentX        =   5186
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
      Begin VB.Label xClient_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   630
         Width           =   3840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«·Ê“‰"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·„ﬁÿÊ—…"
         DragMode        =   1  'Automatic
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì·"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂÊœ"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   390
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2970
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   14
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
         Picture         =   "FAIR.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "FAIR.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   855
         TabIndex        =   15
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
         Picture         =   "FAIR.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "FAIR.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
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
         Picture         =   "FAIR.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "FAIR.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
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
         Picture         =   "FAIR.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "FAIR.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   6990
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
         Left            =   3465
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAIR.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2325
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAIR.frx":12FD8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAIR.frx":15551
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1185
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAIR.frx":179BD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4590
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAIR.frx":1A257
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5760
         Picture         =   "FAIR.frx":1C803
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   2070
      Top             =   180
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   4770
      Top             =   495
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
   Begin MSAdodcLib.Adodc DATA12 
      Height          =   330
      Left            =   2970
      Top             =   495
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Label xRecord 
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
      Height          =   420
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2520
      Width           =   6990
   End
End
Attribute VB_Name = "Fairfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Public bedit As Boolean, aEditRecord As Variant
Public sCode As String
Dim nRound As Long
Dim con As New ADODB.Connection
Dim cFilter As String, cFilterLookup As String, cList As String
Dim oSearch As New Search3, oSearchClient As New Search3, oSearchPlace As New Search3
Dim formMode As Byte, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdTrailer_Click()
Dim oFlagfrm As New flag_mainfrm, sText As String
sCode = xTrailer.BoundText
oFlagfrm.sTable = "Trailer_codes"
oFlagfrm.sCaption = "«‰Ê«⁄ «·”Ì«—« "
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
DATA2.Refresh
If sText <> "" Then xTrailer.BoundText = sText
If Not xTrailer.MatchedWithList Then xTrailer.BoundText = ""
End Sub

Private Sub cmdContainer_Click()
containerfrm.Show 1
End Sub

Private Sub Form_Activate()
On Error Resume Next
If xCode.Tag = DefineMode Then
    xClient.SetFocus
Else
    grid1.SetFocus
End If
Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyAscii = 0
    End If
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
nRound = 2
bedit = True
openCon con

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Trailer_CODES"
Set xTrailer.RowSource = DATA2
xTrailer.ListField = "Desca"
xTrailer.BoundColumn = "Code"

Set grid1.DataSource = data11
data11.ConnectionString = strCon

cList = StrList("SELECT CODE,DESCA FROM PLACE_CODES ORDER BY DESCA")
Fixgrd
openCardTable
myUndo
End Sub
Private Sub CmdAdd_Click()
mydefine
xCode.Text = ""
grid1.Rows = 1
myAddItem
On Error Resume Next
xClient.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From Fair  Where CODE = " & xCode.Text
    con.Execute "Delete  From Fair_sub  Where CODE = " & xCode.Text
    con.CommitTrans
    If sCode <> "" Then
        Unload Me
        Exit Sub
    End If
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "CODE < " & xCode.Text, , adSearchBackward, adBookmarkLast
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
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If sCode <> "" Then
    Unload Me
    Exit Sub
End If
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
aEditRecord = bedit
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode And aEditRecord)
cmdSave.Enabled = aEditRecord
CmdInform.Enabled = (nMode = LoadMode) And Trim(sCode) = ""
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xClient.Text = ""
xClient_Desca.Caption = ""
xTrailer.BoundText = ""
xWeight.Text = ""
grid1.Rows = 1
myAddItem
Handlecontrols DefineMode
xRecord.Caption = "«÷«›… ”Ã· ÃœÌœ"
End Sub
Sub myload()
xCode.Text = CardTable!Code & ""
xClient.Text = CardTable!CLIENT & ""
xClient_Desca.Caption = CardTable!client_Desca & ""
xTrailer.BoundText = CardTable!trailer & ""
xWeight.Text = Myvalue(CardTable!Weight)
myloadgrd
xRecord.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
Handlecontrols LoadMode
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "CLIENT", addstring(xClient.Text))
aInsert = AddFlag(aInsert, "[Trailer]", addstring(xTrailer.BoundText))
aInsert = AddFlag(aInsert, "[weight]", Val(xWeight.Text))
On Error GoTo myerror
con.BeginTrans
If xCode.Text = "" Then
    xCode.Text = Newflag("FAIR", "CODE")
    aInsert = AddFlag(aInsert, "[CODE]", xCode.Text)
    con.Execute addInsert(aInsert, "FAIR")
Else
    con.Execute addUpdate(aInsert, "FAIR", "CODE = " & xCode.Text)
End If
myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    grid1.TextMatrix(grid1.Row, grid1.Col) = oSearchPlace.grid1.TextMatrix(oSearchPlace.grid1.Row, 0)
    Grid1_AfterEdit grid1.Row, grid1.Col
    Unload oSearchPlace
    CellPos 13, grid1.Row, grid1.Col
ElseIf ActiveControl.Name = CmdInform.Name Then
    xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
    myUndo
ElseIf ActiveControl.Name = xClient.Name Then
    xClient.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    SendKeys "{TAB}"
    oSearchClient.Hide
    Unload oSearchClient
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Unload oSearch
Set oSearch = Nothing
Err.Clear
closeCon con
UnloadAllForms "search3"
Unload itemsfrm
Set itemsfrm = Nothing
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And grid1.Row = grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, 1) = "" And grid1.Col = 0 Then
'    KeyAscii = 0
'    If cmdSave.Enabled Then
'        cmdSave_Click
'        CmdAdd_Click
'    End If
'End If
End Sub

Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Not (Col = 0 Or Col = 1) Then CellPos KeyCode, Row, Col
End Sub

Private Sub xClient_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ClientLookupAll Me, oSearchClient
End Sub
Private Sub xClient_LostFocus()
myLostFocus xClient
xClient_Desca.Caption = ""
If Trim(xClient.Text) = "" Then Exit Sub
xClient.Text = RetZero(xClient.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file3_10 where code = " & MyParn(xClient.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xClient_Desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidInt(xCode.Text) Then Exit Sub
CardTable.Find "CODE = " & xCode.Text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    If xCode.Tag = LoadMode Then mydefine
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Trim(xClient.Text) = "" Then
    If Not bIgMsg Then MsgBox "«·⁄„Ì· €Ì— „”Ã·"
    Exit Function
End If

If Not bIgMsg Then
    If IsEmpty(GetField("select code from file3_10 where code = " & MyParn(xClient.Text))) Then
        MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
    End If
End If
MYVALID = True
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
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
cString = "SELECT FAIR.*,FILE3_10.DESCA AS CLIENT_DESCA FROM FAIR INNER JOIN FILE3_10 ON FAIR.CLIENT = FILE3_10.CODE"
cFilter = ""
If sCode <> "" Then cFilter = cFilter & turn(cFilter, " and ") & " CODE = " & sCode
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY FAIR.[CODE]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myRemove(Row As Long)
grid.RemoveItem Row
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bedit Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "DELETE FROM FAIR_SUB WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 112 And aEditRecord And (grid1.Col = 0 Or grid1.Col = 1) Then
    Places_LookupAll Me, oSearchPlace
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid1
If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then
    myAddItem
ElseIf Row = grid1.Rows - 2 And (Col = 0) Then
    MyEditItem grid1, Row, Col
End If
If myreplace(Row) Then
    Handlecontrols LoadMode
    If grid1.TextMatrix(Row, .Cols - 1) = "" Then
        myloadgrd
    End If
Else
    myloadgrd
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 0)) = Trim(.TextMatrix(Row, 1)) Then Exit Function
If Val(.TextMatrix(Row, 2)) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        grid1.RemoveItem OldRow
    End If
End If
End Sub
Private Sub Grid1_EnterCell()
With grid1
If aEditRecord Then
    grid1.Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub Grid1_GotFocus()
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Grid1_EnterCell
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Or Col = 1 Then
    If Trim(.EditText) = "" Then
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        If IsNumeric(.EditText) Then
            aRet = GetFields("Select code,desca from place_codes where code = " & .EditText)
            If Not IsEmpty(aRet) Then
                .EditText = retFlag(aRet, "desca")
            Else
                .EditText = ""
            End If
        End If
    End If
    If (Col = 0 And Trim(.EditText) = Trim(grid1.Cell(flexcpTextDisplay, Row, 1))) Or (Col = 1 And Trim(.EditText) = Trim(grid1.Cell(flexcpTextDisplay, Row, 0))) Then
        MsgBox "„‰ ‰›” «·„ﬂ«‰ «·Ì ‰›” «·„ﬂ«‰"
        Cancel = True
        Exit Sub
    End If
    
    For I = 1 To grid1.Rows - 1
'        If I <> Row Then
'            If Col = 0 Then
'                If (Trim(.Cell(flexcpTextDisplay, I, 0)) = Trim(.EditText) And Trim(.TextMatrix(I, 1)) = Trim(.TextMatrix(Row, 1))) Or (Trim(.TextMatrix(I, 0)) = Trim(.TextMatrix(Row, 1)) And Trim(.Cell(flexcpTextDisplay, I, 1)) = Trim(.EditText)) Then
'                    MsgBox "«·„”«›… „ﬂ——… ›Ï «·”ÿ— " & I
'                    Cancel = True
'                    Exit Sub
'                End If
'            ElseIf Col = 1 Then
'                If (Trim(.TextMatrix(I, 0)) = Trim(.TextMatrix(Row, 0)) And Trim(.Cell(flexcpTextDisplay, I, 1)) = Trim(.EditText)) Or (Trim(.Cell(flexcpTextDisplay, I, 0)) = Trim(.EditText) And Trim(.TextMatrix(I, 1)) = Trim(.TextMatrix(Row, 0))) Then
'                    MsgBox "«·„”«›… „ﬂ——… ›Ï «·”ÿ— " & I
'                    Cancel = True
'                    Exit Sub
'                End If
'            End If
'        End If
        If I <> Row Then
            If Col = 0 Then
                If (Trim(.Cell(flexcpTextDisplay, I, 0)) = Trim(.EditText) And Trim(.TextMatrix(I, 1)) = Trim(.TextMatrix(Row, 1))) Then
                    MsgBox "«·„”«›… „ﬂ——… ›Ï «·”ÿ— " & I
                    Cancel = True
                    Exit Sub
                End If
            ElseIf Col = 1 Then
                If (Trim(.TextMatrix(I, 0)) = Trim(.TextMatrix(Row, 0)) And Trim(.Cell(flexcpTextDisplay, I, 1)) = Trim(.EditText)) Then
                    MsgBox "«·„”«›… „ﬂ——… ›Ï «·”ÿ— " & I
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    Next
ElseIf Col = 2 Then
    If (Not IsNumeric(grid1.EditText)) Then
        Cancel = True
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„‰|" & "≈·Ì|" & "«·‰Ê·Ê‰|"
.ColWidth(0) = 2300
.ColWidth(1) = 2300
.ColWidth(2) = 1300
.ColWidth(3) = 1000
.ColComboList(0) = cList
.ColComboList(1) = cList
.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 2 Then
    .Col = Col + 1
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, 0, 2)
    .ShowCell Row + 1, 0
End If
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
If grid1.Rows > 2 Then
    .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
End If
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "CODE", xCode.Text)
        aInsert = AddFlag(aInsert, "PLACE1", addvalue(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "PLACE2", addvalue(grid1.TextMatrix(I, 1)))
        aInsert = AddFlag(aInsert, "VALUE", Val(grid1.TextMatrix(I, 2)))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FAIR_SUB")
        Else
            con.Execute addUpdate(aInsert, "FAIR_SUB", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
cString = "SELECT FAIR_SUB.PLACE1,FAIR_SUB.PLACE2,VALUE,FAIR_SUB.ID " & _
          " FROM FAIR_SUB"
cString = cString & turn(cString) & "CODE = " & xCode.Text
data11.RecordSource = cString
data11.Refresh
myAddItem
Fixgrd
End With
End Sub
Private Sub xClient_GotFocus()
myGotFocus xClient
End Sub
Private Sub Xweight_GotFocus()
myGotFocus xWeight
End Sub
Private Sub Xweight_LostFocus()
myLostFocus xWeight
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xTrailer_GotFocus()
myGotFocus xTrailer
End Sub
Private Sub xTrailer_LostFocus()
myLostFocus xTrailer
If Not xTrailer.MatchedWithList Then xTrailer.BoundText = ""
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select fair.code ,FILE3_10.DESCA,Trailer_codes.desca  From Fair inner join file3_10 on fair.client = file3_10.code inner join Trailer_codes on fair.Trailer = Trailer_codes.code "
Generalarray(2) = "Order by fair.code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "(%%FILE3_10.DESCA%%)"

listarray(1, 0) = "‰Ê⁄ «·”Ì«—…"
listarray(1, 1) = "(%%Trailer_CODES.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·⁄„Ì·"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "‰Ê⁄ «·”Ì«—…"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «‰Ê«⁄ «·‰Ê·Ê‰"
oSearch.Show 1
End Sub
