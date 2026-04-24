VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Security 
   Caption         =   "«·’·«ÕÌ« "
   ClientHeight    =   9390
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   11850
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   8550
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   27
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "Security.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   28
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "Security.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   29
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "Security.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   30
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "Security.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":EB26
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   4290
      Begin VB.CommandButton cmdCopy 
         Caption         =   "‰”Œ «·„” Œœ„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   135
         Width           =   2130
      End
      Begin VB.CommandButton cmdFix 
         Caption         =   "÷»ÿ «·ﬁ«∆„…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   135
         Width           =   2040
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   7260
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   5985
         Picture         =   "Security.frx":10C75
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security.frx":13448
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security.frx":159F4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security.frx":1828E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security.frx":1A6FA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Height          =   555
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security.frx":1CC73
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "«·ﬁ«∆„… «·›—⁄Ì…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6945
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2205
      Width           =   5730
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   6555
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   5550
         _cx             =   9790
         _cy             =   11562
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   2
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
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   6075
      TabIndex        =   5
      Top             =   720
      Width           =   5685
      Begin VB.CheckBox xShow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ≈ŸÂ«—"
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
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1080
         Width           =   870
      End
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   90
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   4200
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3510
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   780
      End
      Begin VB.TextBox xPassword 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1530
         MaxLength       =   40
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1035
         Width           =   2760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«”„ «·„” Œœ„ :"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«·—ﬁ„ :"
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
         Left            =   4380
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ·„… «·”— :"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame fmVisa 
      Caption         =   "«·ﬁ«∆„… «·«”«”Ì…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6540
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1980
      Width           =   5910
      Begin VSFlex7Ctl.VSFlexGrid grdMainMenu 
         Height          =   6150
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   5730
         _cx             =   10107
         _cy             =   10848
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
         Rows            =   1
         Cols            =   4
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   -1710
      Top             =   1440
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   661
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
   Begin VB.Frame Frame7 
      Height          =   1230
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   720
      Width           =   5910
      Begin VB.CheckBox xOption5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   810
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.CheckBox xOption3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’·«ÕÌ… »Ê«»…"
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
         Height          =   375
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   495
         Width           =   1545
      End
      Begin VB.CheckBox xOption4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’·«ÕÌ… ‰‘«ÿ —Ì«÷Ì"
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
         Height          =   510
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   450
         Width           =   2130
      End
      Begin VB.CheckBox xOption2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’·«ÕÌ… „«·Ì…"
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
         Height          =   375
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   135
         Width           =   2310
      End
      Begin VB.CheckBox xOption1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’·«ÕÌ… „œÌ—"
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
         Height          =   375
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   135
         Width           =   2220
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   661
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
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public bedit As Boolean
Dim bEditRecord As Boolean
Dim formMode As Byte
Dim CardTable As ADODB.Recordset, oSearch As New Search3
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bedit
cmdAdd.Enabled = (nMode = LoadMode And bEditRecord)
CmdDel.Enabled = (nMode = LoadMode And bEditRecord)
cmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = bedit = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.text = Newflag("users", "code")
xpassword.text = ""
xDesca.text = ""
xOption1.Value = 0
xOption2.Value = 0
xOption3.Value = 0
xOption4.Value = 0
xOption5.Value = 0
With grid1
For I = 1 To grid1.rows - 1
    .TextMatrix(I, 2) = ""
    .TextMatrix(I, 3) = ""
    .TextMatrix(I, 4) = ""
Next
End With
Handlecontrols DefineMode
End Sub
Sub myload()
On Error GoTo myError
xCode.text = CardTable!CODE
xpassword.text = CardTable!PassWord & ""
xDesca.text = CardTable!Desca & ""
xOption1.Value = IIf(CardTable!Option1, 1, 0)
xOption2.Value = IIf(CardTable!Option2, 1, 0)
xOption3.Value = IIf(CardTable!Option3, 1, 0)
xOption4.Value = IIf(CardTable!Option4, 1, 0)
xOption5.Value = IIf(CardTable!Option5, 1, 0)
myloadgrd
Handlecontrols LoadMode
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Function MyReplace() As Boolean
'On Error GoTo myerror
Dim aInsert(8, 1)

aInsert(0, 0) = "Code"
aInsert(0, 1) = addvalue(xCode.text)

aInsert(1, 0) = "[password]"
aInsert(1, 1) = addstring(UCase(xpassword.text))

aInsert(2, 0) = "Desca"
aInsert(2, 1) = addstring(xDesca.text)

aInsert(3, 0) = "option1"
aInsert(3, 1) = IIf(xOption1.Value = 0, "0", "1")

aInsert(4, 0) = "option2"
aInsert(4, 1) = IIf(xOption2.Value = 0, "0", "1")

aInsert(5, 0) = "option3"
aInsert(5, 1) = IIf(xOption3.Value = 0, "0", "1")

aInsert(6, 0) = "option4"
aInsert(6, 1) = IIf(xOption4.Value = 0, "0", "1")

aInsert(7, 0) = "option5"
aInsert(7, 1) = IIf(xOption5.Value = 0, "0", "1")

con.BeginTrans
If xCode.Enabled Then
    xCode.text = Newflag("users", "code")
    aInsert(0, 1) = addvalue(xCode.text)
    con.Execute CreateInsert(aInsert, "users")
Else
    con.Execute CreateUpdate(aInsert, "users", " WHERE code = " & xCode.text)
End If
myreplaceGrd
con.CommitTrans
MyReplace = True
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsNumeric(xCode.text) Then
    If Not bIgMsg Then MsgBox "—ﬁ„ «·„” Œœ„ €Ì— ”·Ì„"
    Exit Function
End If
If xDesca.text = "" Then
    If Not bIgMsg Then MsgBox "«”„ «·„” Œœ„ €Ì— „”Ã·"
    Exit Function
End If

If xpassword.text = "" Then
    If Not bIgMsg Then MsgBox "ﬂ·„… «·”— €Ì— „”Ã·…"
    Exit Function
End If
MYVALID = True
End Function

Private Sub cmdAdd_Click()
mydefine
xCode.SetFocus
End Sub

Private Sub cmdCopy_Click()
Dim cmdNew As New ADODB.Command
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "CODE", xCode.text)
Set cmdNew = myCmdEx("sp_copy_user", con, aPrm)
If IsNull(cmdNew.Parameters("@ERROR_MSG").Value) Then
    If Not IsNull(cmdNew.Parameters("@NEWCODE")) Then
        xCode.text = cmdNew.Parameters("@NEWCODE")
        openCardTable
        myUndo
    Else
        MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ «·‰”Œ"
    End If
Else
    MsgBox cmdNew.Parameters("@ERROR_MSG")
End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdDel_Click()
On Error GoTo myError
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From users  Where code = " & Val(xCode.text)
    con.Execute "delete from MenuSetting where code = " & Val(xCode.text)
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & Val(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myError:
    MsgBox Err.Description
    con.RollbackTrans
    Err.Clear
End Sub
Private Sub CmdFirst_Click()
'On Error GoTo myError
CardTable.MoveFirst
myload
End Sub

Private Sub CmdInform_Click()
    CardLookup
End Sub

Private Sub CmdLast_Click()
'On Error GoTo myError
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
'On Error GoTo myError
CardTable.MoveNext
If CardTable.EOF Then CardTable.MovePrevious Else myload
Exit Sub
End Sub
Private Sub CmdPrevious_Click()
'On Error GoTo myError
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub cmdSave_Click()
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
Private Function FixControl() As Boolean
Dim loctable As New ADODB.Recordset
loctable.Open "select menusetting.Id from menusetting left join menu on menu.control = menusetting.control " & _
              "where (menu.control is null)", con, adOpenStatic, adLockReadOnly, adCmdText
On Error GoTo myError
con.BeginTrans
Do Until loctable.EOF
    con.Execute "delete  from menusetting where ID = " & loctable!ID
    loctable.MoveNext
Loop
con.CommitTrans
loctable.Close
Set loctable = Nothing
FixControl = True
Exit Function
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Private Sub cmdFix_Click()
If FillMainMenu Then
    If FillMenuFile Then
        If FixControl Then Inform " „ ÷»ÿ «·ﬁ«∆„… »‰Ã«Õ"
    End If
    openCardTable
    myUndo
    
    FillMainGrd
    If Me.grdMainMenu.rows > 2 Then grdMainMenu.Select 2, 1
End If
End Sub

Private Sub Command2_Click()
FillMainMenu
End Sub
Private Sub Form_Load()
bedit = True
openCon con
'On Error GoTo myerror
'cmdFix.Visible = bSupermode
Fixgrd

openCardTable
myUndo

FillMainGrd
If Me.grdMainMenu.rows > 2 Then grdMainMenu.Select 2, 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'CardTable.Close
'Set CardTable = Nothing
'    closeCon con
Err.Clear
End Sub

Private Sub grdMainMenu_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> 0 Then
    myloadSub grdMainMenu.TextMatrix(NewRow, 0)
    myloadgrd
End If
End With
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Row = 0 Then
    With grid1
    For I = 1 To .rows - 1
        .TextMatrix(I, Col) = IIf(grid1.Cell(flexcpChecked, 0, Col) = 1, -1, 0)
    Next
    End With
End If

If xCode.Enabled = True Then
    If MYVALID Then
        If MyReplace Then
            xCode.Enabled = False
        End If
    End If
Else
    On Error GoTo myError
    con.BeginTrans
    myreplaceGrd IIf(Row = 0, -1, Row)
    con.CommitTrans
    myloadgrd
End If
Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub Grid1_EnterCell()
If grid1.Col = 2 Then grid1.Editable = flexEDKbdMouse Else grid1.Editable = flexEDNone
End Sub
Private Sub xbox_Validate(Cancel As Boolean)
If Not xbox.MatchedWithList Then xbox.BoundText = ""
End Sub
Private Sub xCode_LostFocus()
If xCode.text = "" Then Exit Sub
CardTable.Find "code = " & xCode.text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub myloadgrd()
Dim bNoAll As Boolean
With grid1
For I = 1 To .rows - 1
    aRet = RetMenu(Val(xCode.text), .TextMatrix(I, 0), con)
    .TextMatrix(I, 2) = IIf(aRet(0) = 0, "0", "-1")
    .TextMatrix(I, 3) = IIf(aRet(1) = 0, "0", "-1")
    .TextMatrix(I, .Cols - 1) = aRet(2) & ""
    If Val(.TextMatrix(I, 2)) = 0 Then bNoAll = True
Next
grid1.Cell(flexcpChecked, 0, 2) = IIf(bNoAll Or .rows = 1, 2, 1)
End With
End Sub
Private Sub myreplaceGrd(Optional nRow As Long = -1)
Dim aGrid(3, 1)

With grid1
 For I = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.rows - 1, nRow)
    If grid1.TextMatrix(I, .Cols - 1) <> "" Then
        con.Execute "Delete FROM MENUSETTING WHERE ID = " & .TextMatrix(I, .Cols - 1)
    End If
    If Val(.TextMatrix(I, 2)) <> 0 Then
        aGrid(0, 0) = "CODE"
        aGrid(0, 1) = Val(xCode.text)
        
        aGrid(1, 0) = "CONTROL"
        aGrid(1, 1) = addstring(.TextMatrix(I, 0))
        
        aGrid(2, 0) = "[VISIBLE]"
        aGrid(2, 1) = IIf(Val(.TextMatrix(I, 2)) = 0, "0", "1")
        
        aGrid(3, 0) = "EDITABLE"
        aGrid(3, 1) = IIf(Val(.TextMatrix(I, 3)) = 0, "0", "1")
        con.Execute CreateInsert(aGrid, "menuSetting")
    End If
Next

End With
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From users "
Generalarray(2) = "Order by desca "
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "»ÕÀ"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 4000


searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ "
oSearch.Show 1
End Sub
Sub myProc()
   CardTable.Find "CODE = " & MyParn(oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
   myload
   Unload oSearch
End Sub
Private Function FillMenuFile() As Boolean
Dim MenuNo As Integer, Order As Integer, MainMenu As String
On Error GoTo myError
con.BeginTrans
con.Execute "Delete  from menu"

For I = 0 To Main.Count - 1
     If TypeOf Main(I) Is Menu Then
        
        If LCase(Mid(Main(I).Name, 1, 2)) = "mn" And MainMenu <> Main(I).Name Then
           MainMenu = Main(I).Name
           Order = 1
           MenuNo = MenuNo + 1
        End If
        If Mid(Main(I).Name, 1, 2) <> "mn" Then
            con.Execute "insert into menu(control,Desca,menuNo,[order],mainmenu)" & _
                           " values(" & _
                           addstring(Main(I).Name) & "," & _
                           addstring(Main(I).Caption) & "," & _
                           MenuNo & "," & _
                           Order & "," & _
                           addstring(MainMenu) & _
                            ")"
            Order = Order + 1
        End If
    End If
Next
con.CommitTrans
FillMenuFile = True
Exit Function
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Private Sub FixMainMenu()
' set column width
With grdMainMenu
.TextMatrix(0, 1) = "«·ﬁ«∆„…"
.TextMatrix(0, 2) = "„—∆Ì…"
.Row = 0
.CellAlignment = flexAlignLeftTop
.ColHidden(3) = True
.RowHidden(1) = True
.IsSubtotal(1) = True

.ColDataType(2) = flexDTBoolean
.ColDataType(3) = flexDTBoolean
.ColWidth(1) = 4000
.ColWidth(2) = 1100
.ColWidth(3) = 1000

.GridLines = flexGridNone
.OutlineCol = 1
.OutlineBar = flexOutlineBarComplete

' behavior
.AllowSelection = False
.ScrollTrack = False
.Editable = True
.BackColorBkg = vbWhite
.SheetBorder = vbWhite
For I = 0 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightTop
Next

.rows = 2
Set GRDTABLE = secdb.OpenRecordset("select * From Menu " & _
                                  " Order by Menu.MenuNo,Menu.Level,Menu.order ")
If GRDTABLE.RecordCount = 0 Then Exit Sub
Do
    cPad = String(IIf(GRDTABLE!Level <= 1, 0, GRDTABLE!Level) * 5, " ")
    .AddItem ""
    .IsSubtotal(.rows - 1) = (GRDTABLE!Order = 0)
     If .IsSubtotal(.rows - 1) Then .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = vbRed
    .RowOutlineLevel(.rows - 1) = GRDTABLE!Level
    .TextMatrix(.rows - 1, 0) = GRDTABLE!Control
    .TextMatrix(.rows - 1, 1) = cPad & GRDTABLE!Desca
    .TextMatrix(.rows - 1, 2) = 0
    .TextMatrix(.rows - 1, 3) = 0
    GRDTABLE.MoveNext
Loop Until GRDTABLE.EOF
.MergeCells = flexMergeSpill
End With
End Sub
Private Function FillMenu() As Boolean
con.BeginTrans
On Error GoTo myError
con.Execute "Delete  from Menu"
For I = 0 To Main.Count - 1
        If TypeOf Main(I) Is Menu Then
            If LCase(Mid(Main(I).Name, 1, 3)) = "mn_" And MainMenu <> Main(I).Name Then
               MainMenu = Main(I).Name
               Order = 1
               MenuNo = MenuNo + 1
            ElseIf Mid(Main(I).Name, 1, 2) <> "mn" Then
                con.Execute "insert into menu(control,Desca,menuNo,[order],mainmenu)" & _
                               " values(" & _
                               addstring(Main(I).Name) & "," & _
                               addstring(Main(I).Caption) & "," & _
                               MenuNo & "," & _
                               Order & "," & _
                               addstring(MainMenu) & _
                                ")"
                Order = Order + 1
            End If
    End If
Next
con.CommitTrans
FillMenu = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
End Function
Private Function FillMainMenu() As Boolean
Dim nRow As Long
con.BeginTrans
On Error GoTo myError
con.Execute "Delete  from MainMenu"
For I = 0 To Main.Count - 1
    If TypeOf Main(I) Is Menu Then
        If LCase(Mid(Main(I).Name, 1, 3)) = "mn_" Then
           nRow = nRow + 1
           nLevel = IIf(LCase(Mid(Main(I).Name, 1, 6)) = "mn_mn_", 2, 1)
           MainMenu = Main(I).Name
           con.Execute "insert into mainmenu(Code,Desca,Level,Row)" & _
                           " values(" & _
                           addstring(Main(I).Name) & "," & _
                           addstring(Main(I).Caption) & "," & _
                           nLevel & "," & _
                           nRow & _
                            ")"
        End If
    End If
Next
con.CommitTrans
FillMainMenu = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
'FillMainMenu = True
End Function
Private Sub FillMainGrd()
' set column width
With grdMainMenu
.TextMatrix(0, 1) = "«·ﬁ«∆„…"
.rows = 2
.Row = 1
.CellAlignment = flexAlignLeftTop
.ColHidden(0) = True
.ColHidden(3) = True
.ColHidden(2) = True
.RowHidden(1) = True
.IsSubtotal(1) = True

.ColDataType(2) = flexDTBoolean
.ColDataType(3) = flexDTBoolean
'.ColHidden(0) = True
'.ColWidth(0) = 4000
.ColWidth(1) = 4000
.ColWidth(2) = 1100
.ColWidth(3) = 1000

.GridLines = flexGridNone
.OutlineCol = 1
.OutlineBar = flexOutlineBarComplete

' behavior
.AllowSelection = False
.ScrollTrack = False
.Editable = True
.BackColorBkg = vbWhite
.SheetBorder = vbWhite
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightTop
Next

.rows = 2

Dim GRDTABLE As New ADODB.Recordset, cString As String
cString = "Select * FROM MAINMENU ORDER BY ROW"
GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until GRDTABLE.EOF
    cPad = String(IIf(GRDTABLE!Level <= 1, 0, GRDTABLE!Level) * 5, " ")
    .AddItem ""
    .IsSubtotal(.rows - 1) = (GRDTABLE!Level = 1)
     If .IsSubtotal(.rows - 1) Then .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = vbRed
    .RowOutlineLevel(.rows - 1) = GRDTABLE!Level
    .TextMatrix(.rows - 1, 0) = GRDTABLE!CODE
    .TextMatrix(.rows - 1, 1) = cPad & GRDTABLE!Desca
    .TextMatrix(.rows - 1, 2) = 0
    .TextMatrix(.rows - 1, 3) = 0
    GRDTABLE.MoveNext
Loop
.MergeCells = flexMergeSpill
End With
End Sub
Private Sub myloadSub(pControl As String)
grid1.rows = 1
Dim GRDTABLE As New ADODB.Recordset
cString = "Select * from menu "
cString = cString & turn(cString) & " mainMenu = " & MyParn(pControl)
cString = cString & " Order by [Order]"
GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly
With grid1
    Do Until GRDTABLE.EOF
        .AddItem ""
        .TextMatrix(.rows - 1, 0) = GRDTABLE!Control & ""
        .TextMatrix(.rows - 1, 1) = GRDTABLE!Desca & ""
        .TextMatrix(.rows - 1, 2) = 0
        .TextMatrix(.rows - 1, 3) = 0
        GRDTABLE.MoveNext
    Loop
End With
GRDTABLE.Close
Set GRDTABLE = Nothing
End Sub
Private Sub Fixgrd()
    grid1.Cols = 5
    grid1.ColWidth(1) = 4000
    grid1.ColWidth(2) = 1000
    grid1.ColWidth(3) = 1000
    grid1.ColDataType(2) = flexDTBoolean
    grid1.ColDataType(3) = flexDTBoolean
    grid1.TextMatrix(0, 1) = "«·»‰œ"
    grid1.TextMatrix(0, 2) = "≈ŸÂ«—"
    grid1.TextMatrix(0, 3) = " ⁄œÌ·"
    grid1.ColHidden(0) = True
    grid1.ColHidden(3) = True
    grid1.ColHidden(grid1.Cols - 1) = True
    For I = 0 To 3
        grid1.ColAlignment(I) = flexAlignRightCenter
    Next
    grid1.Cell(flexcpChecked, 0, 2, 0, grid1.Cols - 1) = 2
    
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xCode.text) <> "" Then
        CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT * FROM USERS"
cString = cString & " order by code"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub

Private Sub xShow_Click()
xpassword.PasswordChar = IIf(xShow.Value = 1, "", "*")
End Sub

