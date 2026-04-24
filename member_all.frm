VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form member_allfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «⁄÷«¡ «·‰«œÌ"
   ClientHeight    =   9930
   ClientLeft      =   615
   ClientTop       =   1320
   ClientWidth     =   15615
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
   ForeColor       =   &H80000017&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   15615
   WindowState     =   2  'Maximized
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   2265
      Left            =   135
      TabIndex        =   26
      Top             =   6525
      Width           =   15360
      _cx             =   27093
      _cy             =   3995
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483640
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   8820
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   47
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
         Picture         =   "member_all.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_all.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   48
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
         Picture         =   "member_all.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_all.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   49
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
         Picture         =   "member_all.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_all.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   50
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
         Picture         =   "member_all.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_all.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1320
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   5175
      Width           =   2625
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "«Œ— «Ì’«·"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   225
         Width           =   825
      End
      Begin VB.Label xdate_paid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label15 
         Caption         =   " «Œ— ”œ«œ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   945
         Width           =   780
      End
      Begin VB.Label xDate_print 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label20 
         Caption         =   "«Œ— ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   585
         Width           =   825
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   1545
      Left            =   8325
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   9855
      Visible         =   0   'False
      Width           =   5550
      Begin VB.CommandButton Command2 
         Caption         =   "«÷«›… «·«⁄÷«¡"
         Height          =   600
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   765
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«÷«›… «·’Ê—"
         Height          =   600
         Left            =   -405
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1485
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "«÷«›… «· Ê«»⁄"
         Height          =   600
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   270
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command4 
         Caption         =   "«÷«›… ’Ê— «· Ê«»⁄"
         Height          =   600
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   495
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command5"
         Height          =   420
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   405
         Visible         =   0   'False
         Width           =   3075
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   2745
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   5490
      Width           =   3615
      Begin VB.TextBox xDate_Begin 
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
         Height          =   330
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Tag             =   "D"
         Top             =   180
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   90
         TabIndex        =   25
         Top             =   540
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.Label Label25 
         Caption         =   " «—ÌŒ «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label28 
         Caption         =   "‰Ê⁄ «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   585
         Width           =   990
      End
   End
   Begin VB.Frame Frame9 
      Height          =   2040
      Left            =   4545
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   2745
      Width           =   10905
      Begin VB.TextBox xNotes 
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1620
         Width           =   7485
      End
      Begin VB.TextBox xMail 
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
         Height          =   330
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1260
         Width           =   7485
      End
      Begin VB.TextBox xAddress 
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   7485
      End
      Begin VB.TextBox xPhone 
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   540
         Width           =   7485
      End
      Begin VB.TextBox xMobil 
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   900
         Width           =   7485
      End
      Begin VB.Label Label22 
         Caption         =   "„·«ÕŸ« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   1710
         Width           =   1170
      End
      Begin VB.Label Label18 
         Caption         =   "»—Ìœ «·Ìﬂ —Ê‰Ï"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1305
         Width           =   1170
      End
      Begin VB.Label Label3 
         Caption         =   "⁄‰Ê«‰ «·⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   225
         Width           =   1170
      End
      Begin VB.Label Label9 
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label13 
         Caption         =   "«· ·Ì›Ê‰ «·«—÷Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   585
         Width           =   1440
      End
   End
   Begin VB.Frame Frame7 
      Height          =   690
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   0
      Width           =   8700
      Begin VB.CommandButton cmdInform_rel 
         Caption         =   "«” ⁄·«„  «»⁄"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6210
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1230
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
         Height          =   510
         Left            =   3735
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_all.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2505
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_all.frx":12FD8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_all.frx":15551
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_all.frx":179BD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4965
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_all.frx":1A257
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton cmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   7425
         Picture         =   "member_all.frx":1C803
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   4545
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   675
      Width           =   10905
      Begin VB.TextBox xId_no 
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
         Height          =   330
         Left            =   6525
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1620
         Width           =   2805
      End
      Begin VB.TextBox xDate_birth 
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
         Height          =   330
         Left            =   6525
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   900
         Width           =   2805
      End
      Begin VB.TextBox xTitle 
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
         Height          =   330
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   2715
      End
      Begin VB.TextBox xDesca 
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
         Height          =   330
         Left            =   5310
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   4020
      End
      Begin VB.TextBox xCode 
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
         Height          =   330
         Left            =   8010
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   585
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xReligion 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   945
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xSocial 
         Height          =   330
         Left            =   6525
         TabIndex        =   5
         Top             =   1260
         Width           =   2805
         _ExtentX        =   4948
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
      Begin VB.Label Label21 
         Caption         =   "«·—ﬁ„ «·ﬁÊ„Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label19 
         Caption         =   "«·Õ«·… «·«Ã „«⁄Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1305
         Width           =   1440
      End
      Begin VB.Label Label6 
         Caption         =   "«·œÌ«‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label Label12 
         Caption         =   "«·‰Ê⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label14 
         Caption         =   " «—ÌŒ «·„Ì·«œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   945
         Width           =   1125
      End
      Begin VB.Label Label11 
         Caption         =   "«··ﬁ»"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label7 
         Caption         =   "ﬂÊœ «·⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "≈”„ «·⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   585
         Width           =   1005
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   810
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   225
      Top             =   405
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   741
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
      Height          =   375
      Left            =   2520
      Top             =   7020
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   375
      Left            =   2115
      Top             =   -135
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Caption         =   "DATA7"
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
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   81
      Top             =   9585
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin VB.Frame Frame3 
      Height          =   3795
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   990
      Width           =   4380
      Begin Threed.SSCommand cmdScan 
         Height          =   825
         Left            =   90
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   1455
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_all.frx":1EFD6
         Caption         =   "„”Õ ÷Ê∆Ì"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_all.frx":21770
      End
      Begin VB.Image xAppendPhoto 
         Appearance      =   0  'Flat
         Height          =   2655
         Left            =   90
         Stretch         =   -1  'True
         Top             =   180
         Width           =   2115
      End
      Begin VB.Image xMemberPhoto 
         Appearance      =   0  'Flat
         Height          =   2640
         Left            =   2250
         Stretch         =   -1  'True
         Top             =   180
         Width           =   2025
      End
   End
   Begin ImgXCtrl6.ImgXCtrl imgx1 
      DragIcon        =   "member_all.frx":244CB
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   -45
      TabIndex        =   40
      Tag             =   "-1"
      Top             =   1665
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3678
      BorderStyle     =   1
      AutoZoom        =   -1  'True
      LicenseUserName =   "mrmind71"
      LicenseRegCode  =   "íß“Ωª∫≠Ω≥´±“™ºØ´¥æÆØUBOR-FEOEONZI-EPCP6gI"
   End
   Begin VB.Frame Frame6 
      Height          =   1725
      Left            =   6390
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   4770
      Width           =   9105
      Begin VB.TextBox xDate_Degree 
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
         Height          =   345
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Tag             =   "D"
         Top             =   585
         Width           =   1455
      End
      Begin VB.CommandButton cmdStatus 
         Caption         =   "..."
         Height          =   330
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox xDate_Job_begin 
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
         Height          =   345
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Tag             =   "D"
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox xDate_job 
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
         Height          =   345
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Tag             =   "D"
         Top             =   945
         Width           =   1455
      End
      Begin VB.CommandButton cmdSection 
         Caption         =   "..."
         Height          =   330
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   900
         Width           =   375
      End
      Begin VB.CommandButton cmdDegree 
         Caption         =   "..."
         Height          =   330
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
      End
      Begin VB.CommandButton cmdJob 
         Caption         =   "..."
         Height          =   330
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   375
      End
      Begin MSDataListLib.DataCombo xJob 
         Height          =   330
         Left            =   4185
         TabIndex        =   13
         Top             =   180
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo xDegree 
         Height          =   330
         Left            =   4185
         TabIndex        =   15
         Top             =   540
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo xSection 
         Height          =   330
         Left            =   4185
         TabIndex        =   17
         Top             =   900
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo xStatus 
         Height          =   330
         Left            =   4185
         TabIndex        =   19
         Top             =   1260
         Width           =   3345
         _ExtentX        =   5900
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
      Begin VB.Label Label1 
         Caption         =   " «—ÌŒ «·„ƒÂ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   630
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   " «—ÌŒ ‘€· «·ÊŸÌ›…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label Label10 
         Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "«·Õ«·…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label24 
         Caption         =   "«·„ƒÂ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   585
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "«·«œ«—… «·⁄«„…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "«·ÊŸÌ›… «·Õ«·Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   225
         Width           =   1170
      End
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
      Height          =   510
      Left            =   3285
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   8910
      Width           =   4020
   End
End
Attribute VB_Name = "member_allfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim fs As New FileSystemObject
Public bEdit As Boolean, bEditRecord As Boolean
Dim con As New ADODB.Connection
Dim fs As New FileSystemObject
Dim WithEvents twain As ImgXTwain, nphoto As Long
Attribute twain.VB_VarHelpID = -1
Dim cRelStr As String, cGenderStr As String
Dim formMode As Byte
Dim oSearch As New Search3, oSearchRel As New Search3
Dim CardTable As ADODB.Recordset, bAct As Boolean
Public sCode As String
Dim cFilter As String, cFilterLookup As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdAdd.Enabled = (nMode = LoadMode And bEditRecord)
CmdDel.Enabled = (nMode = LoadMode And bEditRecord)
cmdSave.Enabled = bEditRecord
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = bEdit = Not (nMode = LoadMode)
cmdScan.Enabled = nMode = LoadMode And bEditRecord
'cmdScan2.Enabled = nMode = LoadMode And bEditRecord
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = Newflag("file1_10", "code")
xTitle.Text = ""
xDescA.Text = ""
xDate_birth.Text = ""
xDate_Job_begin.Text = ""
xGender.BoundText = ""
xSocial.BoundText = ""
xReligion.BoundText = ""
xId_no.Text = ""
xAddress.Text = ""
xPhone.Text = ""
xMobil.Text = ""
xMail.Text = ""
xJob.BoundText = ""
xDegree.BoundText = ""
xSection.BoundText = ""
xStatus.BoundText = ""
xDate_Job_begin.Text = ""
xDate_Degree.Text = ""
xDate_job.Text = ""
xType.BoundText = ""
'xDate_Last.Text = ""
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")

xDate_print.Caption = ""
xdate_paid.Caption = ""
xdoc_no.Caption = ""
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
StatusBar1.Panels(3).Text = ""
StatusBar1.Panels(4).Text = ""
grid1.rows = 1
myAddItem
Handlecontrols DefineMode
xRecordNo.Caption = "«÷«›… ”Ã· ÃœÌœ " & "(" & CardTable.RecordCount & ")"
On Error Resume Next
CellPos 13, grid1.rows - 2, grid1.Cols - 1
grid1.SetFocus
Err.Clear
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.row, 0)
    Unload oSearch
ElseIf ActiveControl.Name = Me.cmdInform_rel.Name Then
    xCode.Text = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.row, 0)
    Unload oSearchRel
End If
myUndo
End Sub
Private Sub myload()
xCode.Text = CardTable!CODE & ""
xTitle.Text = CardTable!Title & ""
xDescA.Text = CardTable!Desca & ""
xDate_birth.Text = myFormat_p(CardTable!DATE_BIRTH)
xGender.BoundText = CardTable!GENDER & ""
xSocial.BoundText = CardTable!Social & ""
xReligion.BoundText = CardTable!Religion & ""
xId_no.Text = CardTable!id_no & ""
xAddress.Text = CardTable!Address & ""
xPhone.Text = CardTable!Phone & ""
xMobil.Text = CardTable!MOBIL & ""
xMail.Text = CardTable!Mail & ""
xJob.BoundText = CardTable!job & ""
xDegree.BoundText = CardTable!Degree & ""
xSection.BoundText = CardTable!Section & ""
 xStatus.BoundText = CardTable!Status & ""
xDate_Job_begin.Text = myFormat_p(CardTable!Date_Job_begin)
xDate_Degree.Text = myFormat_p(CardTable!Date_Degree)
xDate_Begin.Text = myFormat_p(CardTable!Date_begin)
xDate_job.Text = myFormat_p(CardTable!Date_job)
xType.BoundText = CardTable!Type & ""
Handlecontrols LoadMode
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")

StatusBar1.Panels(1).Text = CardTable!UserName & ""
StatusBar1.Panels(2).Text = myFormat_p(CardTable!Time, True)
StatusBar1.Panels(3).Text = CardTable!UserName2 & ""
StatusBar1.Panels(4).Text = myFormat_p(CardTable!Time2, True)
xRecordNo.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
If validPhoto(RetPhoto(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto(xCode.Text))
myloadgrd
grid1.Select 1, 0
aret = LastDoc(xCode.Text, con)
xdoc_no.Caption = retFlag(aret, "FORM_NO") & ""
xdate_paid.Caption = myFormat_p(retFlag(aret, "date"))
xDate_print.Caption = myFormat_p(CardTable!Date_print)
'If bNoGrid Then Exit Sub
On Error Resume Next
CellPos 13, grid1.rows - 2, grid1.Cols - 1
grid1.SetFocus
Err.Clear
End Sub
Private Function MyReplace(Optional row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "Title", addstring(xTitle.Text))
aInsert = AddFlag(aInsert, "Desca", addstring(xDescA.Text))
aInsert = AddFlag(aInsert, "Date_birth", addstring(xDate_birth.Text))
aInsert = AddFlag(aInsert, "Date_Begin", addDate(xDate_Begin.Text))
aInsert = AddFlag(aInsert, "Gender", addvalue(xGender.BoundText))
aInsert = AddFlag(aInsert, "Social", addvalue(xSocial.BoundText))
aInsert = AddFlag(aInsert, "Religion", addvalue(xReligion.BoundText))
aInsert = AddFlag(aInsert, "Id_no", addstring(xId_no.Text))
aInsert = AddFlag(aInsert, "Address", addstring(xAddress.Text))
aInsert = AddFlag(aInsert, "Phone", addstring(xPhone.Text))
aInsert = AddFlag(aInsert, "Mobil", addstring(xMobil.Text))
aInsert = AddFlag(aInsert, "Mail", addstring(xMail.Text))
aInsert = AddFlag(aInsert, "Job", addvalue(xJob.BoundText))
aInsert = AddFlag(aInsert, "Degree", addvalue(xDegree.BoundText))
aInsert = AddFlag(aInsert, "Section", addvalue(xSection.BoundText))
aInsert = AddFlag(aInsert, "Status", addvalue(xStatus.BoundText))
aInsert = AddFlag(aInsert, "Date_Job_begin", addDate(xDate_Job_begin.Text))
aInsert = AddFlag(aInsert, "Date_Degree", addDate(xDate_Degree.Text))
aInsert = AddFlag(aInsert, "Date_job", addstring(xDate_job.Text))
aInsert = AddFlag(aInsert, "Type", addvalue(xType.BoundText))
con.BeginTrans
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    aInsert = AddFlag(aInsert, "Code", addstring(xCode.Text))
    con.Execute addInsert(aInsert, "FILE1_10")
Else
    con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.CODE = " & xCode.Text)
End If
myreplaceGrd row
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CmdAdd_Click()
mydefine
xCode.SetFocus
End Sub

Private Sub cmdDegree_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xDegree.BoundText
oFlagfrm.sTable = "Degree_CODES"
oFlagfrm.sCaption = "«·ÊŸÌ›…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
Set Data3.Recordset = myRecordSet("select * from Degree_Codes", con)
xDegree.BoundText = sBound
If Not xDegree.MatchedWithList Then xDegree.BoundText = ""
End Sub

Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If grid1.rows > 2 Then
        MsgBox "«·⁄÷Ê ·Â  Ê«»⁄ ÌÃ» Õ–›Â„ «Ê·«"
        Exit Sub
    End If
    con.BeginTrans
    con.Execute "Delete  From FILE1_10 Where code = " & xCode.Text
    
    Dim fs As New FileSystemObject
    If fs.FileExists(RetPhoto(xCode.Text)) Then
        fs.DeleteFile RetPhoto(xCode.Text)
    End If
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & xCode.Text, , adSearchBackward, adBookmarkLast
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
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
If Trim(xCode.Text) = "" Then Exit Sub
If xCode.Tag = DefineMode Then Exit Sub
Set fs = CreateObject("Scripting.FileSystemObject")
Dim cFile As String, cNewFile As String
On Error GoTo myerror
Common1.FileName = ""
Common1.InitDir = App.Path & "\PICT"
Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
Common1.ShowOpen
If Common1.FileTitle <> "" Then
    cFile = Common1.FileName
    If cFile <> "" Then
        fs.CopyFile cFile, RetPhoto(xCode.Text)
    End If
    myload
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
MemberLookupAll Me, oSearch, cFilter
End Sub
Private Sub cmdInform_rel_Click()
relLookupAll Me, oSearchRel
End Sub

Private Sub cmdJob_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xJob.BoundText
oFlagfrm.sTable = "JOB_CODES"
oFlagfrm.sCaption = "«·ÊŸÌ›…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
Set Data1.Recordset = myRecordSet("select * from JOB_Codes", con)
xJob.BoundText = sBound
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub

Private Sub cmdSection_Click()
Dim oFlagfrm As New flag_mainfrm, sBoundText As String
sBoundText = xSection.BoundText
oFlagfrm.sTable = "SECTION_CODES"
oFlagfrm.sCaption = "«·‘⁄»…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
Set Data1.Recordset = myRecordSet("select * from section_Codes", con)
xSection.BoundText = sBoundText
If Not xSection.MatchedWithList Then xSection.BoundText = ""
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

Private Sub cmdQual_Click()
Dim myPublic(5)
nCode = xQUAL_CODE.BoundText
myPublic(0) = "Qual_codes"
myPublic(1) = "Code"
myPublic(2) = "Desca"
myPublic(3) = "ﬂÊœ «·„ƒÂ·"
myPublic(4) = "«·„ƒÂ·"
myPublic(5) = "«ﬂÊ«œ «·„ƒÂ·« "
FlagFrm.bEdit = True
FlagFrm.myPublic = myPublic
FlagFrm.Show 1
Data3.Refresh
xQUAL_CODE.BoundText = nCode
If Not xQUAL_CODE.MatchedWithList Then xQUAL_CODE.BoundText = ""
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub cmdScan_Click()
Scan.sCode = xCode.Text
Scan.Show 1
If validPhoto(RetPhoto(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto(xCode.Text))
If grid1.TextMatrix(grid1.row, 0) <> "" And grid1.row <> 0 Then
    If validPhoto(RetAppendPhoto(xCode.Text, grid1.row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.row, 0)))
End If
myload
End Sub

Private Sub cmdStatus_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xStatus.BoundText
oFlagfrm.sTable = "Status_CODES"
oFlagfrm.sCaption = "«·Õ«·…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
Set Data4.Recordset = myRecordSet("select * from Status_Codes", con)
xStatus.BoundText = sBound
If Not xStatus.MatchedWithList Then xStatus.BoundText = ""
End Sub

Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub cmdScan2_Click()
nphoto = 0
ScanImage
On Error Resume Next
If validPhoto(RetPhoto(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto(xCode.Text))
If grid1.TextMatrix(grid1.row, 0) <> "" And grid1.row <> 0 Then
    If validPhoto(RetAppendPhoto(xCode.Text, grid1.row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.row, 0)))
End If
Err.Clear
End Sub
Private Sub Command1_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
'Set f = fs.GetFolder(App.Path & "\photo\")
'Set fc = f.Files
'nCount = fc.Count
Dim cString As String, I As Long, cFile As String, nRecordCount As Long, cCaption As String
Dim loctable As New ADODB.Recordset
'loctable.Open "select * from file1_10 where NEWDATA = true", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.Open "select * from file1_10", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.MoveLast
nRecordCount = loctable.RecordCount
loctable.MoveFirst
cCaption = Me.Caption
Do Until loctable.EOF
    I = I + 1
    Me.Caption = cCaption & I & " from " & nRecordCount
    If Not IsNull(loctable!PHOTO_CODE) Then
        cFile = App.Path & "\person\" & loctable!PHOTO_CODE
        If fs.FileExists(cFile) Then
            fs.CopyFile cFile, RetPhoto(loctable!CODE)
        End If
    End If
    loctable.MoveNext
Loop
MsgBox "Done"
End Sub

Private Sub Command2_Click()
AddMember
End Sub
Private Sub AddMember()
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM FILE1_10", conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
sCaption = Me.Caption
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordCount
    aInsert = AddFlag(Empty, "CODE", loctable!CODE)
    aInsert = AddFlag(aInsert, "MEMBERID", addstring(loctable!MEMBERID))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca & ""))
    aInsert = AddFlag(aInsert, "SECTION", addvalue(loctable!Section & ""))
    aInsert = AddFlag(aInsert, "DATE_BIRTH", addDate(Format(loctable!DATE_BIRTH, "YYYY-MM-DD")))
    aInsert = AddFlag(aInsert, "union_reg", addstring(loctable!Union_reg & ""))
    aInsert = AddFlag(aInsert, "NOTES", addstring(loctable!notes & ""))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address & ""))
    aInsert = AddFlag(aInsert, "JOB_CODE", addstring(loctable!JOB_CODE & ""))
    aInsert = AddFlag(aInsert, "PHONE", addstring(loctable!Phone & ""))
    aInsert = AddFlag(aInsert, "MOBIL", addstring(loctable!MOBIL & ""))
    aInsert = AddFlag(aInsert, "PHOTO_CODE", addstring(loctable!PHOTO_CODE & ""))
    con.Execute addInsert(aInsert, "FILE1_10")
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
conmdb.Close
Set conmdb = Nothing
Exit Sub
myerror:
MsgBox Err.Description
End Sub
Private Sub addRelation()
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM FILE1_11", conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
sCaption = Me.Caption
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordCount
    Dim aSep As Variant
    aSep = Split(loctable!CODE_ALL, "-")
    aInsert = AddFlag(Empty, "CODE", aSep(1))
    aInsert = AddFlag(aInsert, "MEMBER_SPLIT", aSep(0))
    aInsert = AddFlag(aInsert, "CODE_ALL", loctable!CODE_ALL)
    aInsert = AddFlag(aInsert, "MEMBER", addvalue(loctable!Member))
    aInsert = AddFlag(aInsert, "MEMBERID", addstring(loctable!MEMBERID))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca & ""))
    aInsert = AddFlag(aInsert, "DATE_BIRTH", addDate(Format(loctable!DATE_BIRTH, "YYYY-MM-DD")))
    aInsert = AddFlag(aInsert, "RELATION", addvalue(loctable!Relation & ""))
    aInsert = AddFlag(aInsert, "SECTION", addvalue(loctable!Section & ""))
    aInsert = AddFlag(aInsert, "GENDER", addvalue(loctable!GENDER & ""))
    aInsert = AddFlag(aInsert, "union_reg", addstring(loctable!Union_reg & ""))
    aInsert = AddFlag(aInsert, "NOTES", addstring(loctable!notes & ""))
    aInsert = AddFlag(aInsert, "JOB_CODE", addstring(loctable!JOB_CODE & ""))
    aInsert = AddFlag(aInsert, "PHOTO_CODE", addstring(loctable!PHOTO_CODE & ""))
    con.Execute addInsert(aInsert, "FILE1_11")
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
conmdb.Close
Set conmdb = Nothing
MsgBox "Done"
Exit Sub
myerror:
MsgBox Err.Description
End Sub

Private Sub Command3_Click()
addRelation
End Sub

Private Sub Command4_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
'Set f = fs.GetFolder(App.Path & "\photo\")
'Set fc = f.Files
'nCount = fc.Count
Dim cString As String, I As Long, cFile As String, nRecordCount As Long, cCaption As String
Dim loctable As New ADODB.Recordset
loctable.Open "select * from file1_11", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.MoveLast
nRecordCount = loctable.RecordCount
loctable.MoveFirst
cCaption = Me.Caption
Do Until loctable.EOF
    I = I + 1
    Me.Caption = cCaption & I & " from " & nRecordCount
    If Not IsNull(loctable!PHOTO_CODE) Then
        cFile = App.Path & "\person\" & loctable!PHOTO_CODE
        If fs.FileExists(cFile) Then
            fs.CopyFile cFile, RetAppendPhoto(loctable!Member, loctable!CODE)
        End If
    End If
    loctable.MoveNext
Loop
MsgBox "Done"
End Sub
Private Sub Command5_Click()
Dim loctable As ADODB.Recordset, cString As String, cPhoto As String
'Me.MousePointer = 11
Dim fs As New FileSystemObject
cString = "select * from FILE1_10 ORDER BY CODE"
Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    I = I + 1
    Me.Caption = I & " From "
    cPhoto = App.Path & "\photo_all\" & loctable!Register & ".jpg"
    If validPhoto(cPhoto) Then
        fs.MoveFile cPhoto, RetPhoto(loctable!CODE)
    End If
    loctable.MoveNext
Loop

cString = "select * from FILE1_11 ORDER BY MEMBER,CODE"
Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    cPhoto = App.Path & "\photo_all\" & loctable!photo & ".jpg"
    If validPhoto(cPhoto) Then
        fs.MoveFile cPhoto, RetAppendPhoto(loctable!Member, loctable!CODE)
        I = I + 1
        Me.Caption = I & " From "
    End If
    loctable.MoveNext
Loop
MsgBox " „ ”Õ» «·’Ê— »‰Ã«Õ"
End Sub

Private Sub Command6_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
Set f = fs.GetFolder(App.Path & "\photo_fix")
Set fc = f.Files
nCount = fc.Count
Dim cString As String, I As Long
bCrypt = True
For Each f1 In fc
    I = I + 1
    Me.Caption = I
    If InStr(1, LCase(App.Path & "\photo_fix\" & f1.Name), "jpg") <> 0 Then
        cFile = RetPhoto(Replace(LCase(f1.Name), ".jpg", ""))
        If cFile <> "" Then
            fs.CopyFile App.Path & "\photo_fix\" & f1.Name, cFile
        Else
           con.Execute "INSERT INTO TEST(CODE) " & _
                        "VALUES(" & _
                        addstring(f1.Name) & _
                        ")"
        End If
    End If
Next
MsgBox "done..."
End Sub

Private Sub Command7_Click()
Dim loctable As New ADODB.Recordset, sCaption As String
On Error GoTo myerror

loctable.Open "SELECT * FROM FILE1_10", con, adOpenStatic, adLockReadOnly, adCmdText

Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
sCaption = Me.Caption
Dim aInsert As Variant, bNoCard As Boolean
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordCount
    bNoCard = Not validPhoto(RetPhoto(loctable!CODE))
    aInsert = AddFlag(Empty, "nocard", IIf(bNoCard, "1", "0"))
    con.Execute addUpdate(aInsert, "FILE1_10", "CODE = " & loctable!CODE)
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
Exit Sub
myerror:
MsgBox Err.Description
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xCode.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xCode.SetFocus
    End If
    Err.Clear
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub Form_Load()
'makeMyLoad Me
'makeMyReplace Me
'makeMyDefine Me
'LostFocus Me
'makeMyValidate MeLoadText Me
'LoadText Me
'nMemValue = Val(GetDesca("Select value from rel_codes where code = -1"))
openCon con

cRelStr = StrList2("Select Code,Desca From relation_codes order by desca")
cGenderStr = StrList2("Select Code,Desca From gender_codes order by Code")

Set Data1.Recordset = myRecordSet("select * from section_Codes", con)
Set xSection.RowSource = Data1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

Set Data2.Recordset = myRecordSet("select * from Job_Codes", con)
Set xJob.RowSource = Data2
xJob.ListField = "Desca"
xJob.BoundColumn = "Code"

Set Data3.Recordset = myRecordSet("select * from Degree_Codes", con)
Set xDegree.RowSource = Data3
xDegree.ListField = "Desca"
xDegree.BoundColumn = "Code"

Set Data4.Recordset = myRecordSet("select * from Status_Codes", con)
Set xStatus.RowSource = Data4
xStatus.ListField = "Desca"
xStatus.BoundColumn = "Code"

Set DATA5.Recordset = myRecordSet("select * from Gender_Codes", con)
Set xGender.RowSource = DATA5
xGender.ListField = "Desca"
xGender.BoundColumn = "Code"

Set DATA6.Recordset = myRecordSet("select * from type_Codes", con)
Set xType.RowSource = DATA6
xType.ListField = "Desca"
xType.BoundColumn = "Code"

Set DATA7.Recordset = myRecordSet("select * from religion_Codes", con)
Set xReligion.RowSource = DATA7
xReligion.ListField = "Desca"
xReligion.BoundColumn = "Code"

Set grid1.DataSource = DATA11
'data10.ConnectionString = con.ConnectionString

bEdit = Not retFlag(aSec, "INFORM")
Fixgrd
openCardTable
myUndo
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (grid1.Col <> 1 And grid1.Col <> 2) Then KeyAscii = 0
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidInt(xCode.Text) Then Exit Sub
CardTable.Find "code = " & xCode.Text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xCode.Tag = LoadMode Then
    mydefine
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not ValidInt(xCode.Text) Then
    If Not igMsg Then MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

If xDescA.Text = "" Then
    If Not igMsg Then MsgBox "√”„ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

For I = 1 To grid1.rows - 2
    If Not ValidInt(grid1.TextMatrix(I, 0)) Then
        MsgBox "ﬂÊœ «· «»⁄ €Ì— „”Ã·"
        Exit Function
    End If

   
    If Not ValidInt(grid1.TextMatrix(I, 1)) Then
        MsgBox "‰Ê⁄ «· »⁄Ì… €Ì— „”Ã·…"
        Exit Function
    End If
Next
MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveText Me, , Array(xCode1.Name, xCode2.Name)
CardTable.Close
CaseTable.Close
Set CardTable = Nothing
Set CaseTable = Nothing
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo myError
'If Not bSupermode Then Exit Sub
If KeyCode = 46 And grid1.row <> grid1.rows - 1 And bEditRecord Then
    If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        If grid1.TextMatrix(grid1.row, grid1.Cols - 1) <> "" Then
            Dim fs As New FileSystemObject
            If fs.FileExists(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.row, 0))) Then
                fs.DeleteFile RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.row, 0))
            End If
            con.BeginTrans
            con.Execute "Delete  from file1_11 where id = " & grid1.TextMatrix(grid1.row, grid1.Cols - 1)
            con.CommitTrans
            xAppendPhoto.Picture = LoadPicture("")
        End If
        grid1.RemoveItem grid1.row
        Grid1_EnterCell
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub PrintMembers()
Dim cString As String, temptable As New ADODB.Recordset, loctable As New ADODB.Recordset

contemp.Execute "delete  from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

cString = "SELECT FILE1_10.*, FILE1_11.MEMBER, FILE1_11.DESCA AS DESCA_REL, FILE1_11.DATE_BIRTH AS DATE_BIRTH_REL, FILE1_11.PRINT_DATE AS PRINT_DATE_REL, REL_CODES.DESCA AS REL_CODE_DESCA" & _
          " FROM (FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.CODE = FILE1_11.MEMBER) LEFT JOIN REL_CODES ON FILE1_11.RELATION = REL_CODES.CODE"

If IsNumeric(xCode1.Text) Then
    cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.Text), " >= ", " = ") & xCode1.Text
End If

If IsNumeric(xCode2.Text) Then
    cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.Text
End If
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Do Until loctable.EOF
    temptable.AddNew
    temptable!val1 = loctable!CODE
    temptable!str1 = ArbString(loctable!CODE)
    temptable!str2 = loctable!vip
    temptable!Str3 = loctable!Name
    temptable!str4 = loctable!Title
    temptable!str5 = loctable!Address
    If Not IsNull(loctable!Degree) Then
        temptable!STR6 = GetField("select desca from degree_Codes where code = " & UnCodeSerial(CardTable!Degree, 71))
    End If
    temptable!str7 = loctable!Address
    temptable!str8 = loctable!phone1
    temptable!str9 = loctable!MOBIL
    temptable!Str10 = loctable!Union
    
    temptable!str11 = TurnValue(ArbString(Format(loctable!DATE_BIRTH, "yyyy/mm/dd")))
    temptable!str12 = TurnValue(ArbString(loctable!receipt & ""))
    temptable!str13 = TurnValue(ArbString(Format(loctable!Print_date, "yyyy/mm/dd")))
    
    temptable!val2 = loctable!Member
    temptable!str16 = loctable!Desca_rel
    temptable!str17 = loctable!REL_CODE_DESCA
    temptable!str18 = TurnValue(ArbString(Format(loctable!Print_date_rel, "yyyy/mm/dd")))
    temptable!str19 = TurnValue(ArbString(Format(loctable!DATE_BIRTH_rel, "yyyy/mm/dd")))
    temptable!Val3 = retPaid(loctable!CODE)
    temptable.Update
    loctable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    temptable.Requery
    con.BeginTrans
    con.CommitTrans
    REPORT1.ReportFileName = MainPath & "\rpt\Member_data.rpt"
    REPORT1.DataFiles(0) = cTempPath
    REPORT1.Action = 1
End If
Set temptable = Nothing
Set loctable = Nothing
End Sub
Private Sub Calctotals()
Dim nValue
For I = 1 To grid1.rows - 2
    nValue = Val(grid1.TextMatrix(I, 7)) + nValue
Next
If xDied.Value = 0 Then nValue = nValue + nMemValue
xTotal.Caption = Format(nValue, "fixed")
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT FILE1_10.* FROM FILE1_10"
If IsNumeric(sCode) Then cString = cString & turn(cString) & " CODE = " & sCode
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY  FILE1_10.code"
CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
'On Error GoTo myError
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If IsNumeric(xCode.Text) Then
        CardTable.Find "code = " & xCode.Text, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub ScanImage()
On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hwnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
If twain.Sources.Count > 1 Then twain.SelectSource
twain.Acquire False, Me.hwnd
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub Twain_ImageAcquired(Image As ImgX_Image)
If Not IsNumeric(xCode.Text) Then Exit Sub
If nphoto = 0 And xCode.Text Then
    ReplaceFromImage Image, RetPhoto(xCode.Text)
Else
    If nphoto <= grid1.rows - 1 Then
        If IsNumeric(grid1.TextMatrix(nphoto, 0)) Then
            ReplaceFromImage Image, RetPhoto(xCode.Text & "-" & grid1.TextMatrix(nphoto, 0))
        End If
    End If
End If
nphoto = nphoto + 1
End Sub

Private Sub Twain_TwainError(ByVal erNum As Long, ByVal erSource As String, ByVal Description As String)
MsgBox "Error Number:  " & erNum & vbCrLf & Description, vbInformation, erSource
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub
Private Sub ReplaceFromImage(Image As ImgX_Image, cPhoto)
On Error GoTo myerror
imgx1.Images.Replace Image, , False
imgx1.Refresh
imgx1.Export.ToFile cPhoto, ixfsJPG
Exit Sub
myerror:
imgx1.Images.Clear
Err.Clear
End Sub
Private Function retPaid(pMember As String) As Double
Dim aret As Variant, cString As String
aret = GetField("Select code from file1_10 where (not died) and  code = " & pMember)
If Not IsEmpty(aret) Then retPaid = nMemValue

cString = "SELECT SUM(REL_CODES.[VALUE])" & _
          " FROM FILE1_11 INNER JOIN REL_CODES ON FILE1_11.RELATION = REL_CODES.CODE WHERE FILE1_11.MEMBER= " & pMember
aret = GetField(cString)
If Not IsEmpty(cString) Then
    retPaid = retPaid + Val(aret & "")
End If
End Function

Private Sub grid1_KeyUpEdit(ByVal row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And (Col <> 1 And Col <> 2) Then CellPos KeyCode, row, Col
End Sub
Private Sub Grid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid1
'If Row = .Rows - 1 Then
'    xAppendPhoto.Picture = LoadPicture("")
'    If validPhoto(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0))) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0)))
'End If
If Not MYVALID Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myloadgrd
    If row < grid1.rows - 1 Then
        grid1.Select row, Col
    Else
        CellPos 13, grid1.rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If
If Not validRow(row) Then Exit Sub
If row = .rows - 1 Then
    myAddItem
End If
'Calctotals
If MyReplace(row) Then
    'Handlecontrols LoadMode
    xCode.Tag = LoadMode
    xCode.Enabled = False
    If grid1.TextMatrix(row, .Cols - 1) = "" Then
        myloadgrd
        .ShowCell grid1.rows - 1, 0
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
Private Function validRow(row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
'If Not MYVALID(True) Then Exit Function
If Trim(.TextMatrix(row, 0)) = "" Then Exit Function
If Trim(.TextMatrix(row, 1)) = "" Then Exit Function
'If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function
If Trim(.TextMatrix(row, 4)) = "" Then Exit Function

'If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function

End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
    End If
End If
On Error GoTo myerror
If OldRow <> NewRow Then
    xAppendPhoto.Picture = LoadPicture("")
    If validPhoto(RetAppendPhoto(xCode.Text, grid1.TextMatrix(NewRow, 0))) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(NewRow, 0)))
End If
Exit Sub
myerror:
xAppendPhoto.Picture = LoadPicture("")
End Sub
Private Sub Grid1_EnterCell()
With grid1
If (.Col = 0 And Trim(grid1.TextMatrix(grid1.row, grid1.Cols - 1)) = "") Or grid1.Col = 7 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub Grid1_GotFocus()
'CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Grid1_EnterCell
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If Not validRow(grid1.row) And grid1.row <> grid1.rows - 1 And grid1.TextMatrix(grid1.row, grid1.Cols - 1) = "" Then
    myRemove grid1.row
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If Trim(.EditText) = "" Then
        If grid1.row = grid1.rows - 1 Then Exit Sub
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        nFound = FoundOtheritem(grid1, row, 0, Trim(grid1.EditText))
        If nFound <> -1 Then
            MsgBox "«·ﬂÊœ „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
            Cancel = True
            Exit Sub
        End If
    End If
ElseIf Col = 2 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
    End If
ElseIf Col = 5 Then
    If (Not IsDate(grid1.EditText)) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    Else
        grid1.EditText = Format(grid1.EditText, "yyyy/mm/dd")
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "«·—ﬁ„|" & "«·ﬁ—«»…|" & "«·‰Ê⁄|" & "«·’›…|" & "«·«”„|" & " «—ÌŒ «·„Ì·«œ|" & "„·«ÕŸ« |" & " «—ÌŒ «·ÿ»«⁄…|"
.ColWidth(0) = 800
.ColWidth(1) = 1400
.ColWidth(2) = 1400
.ColWidth(3) = 1200
.ColWidth(4) = 3000
.ColWidth(5) = 1350
.ColWidth(6) = 3000
.ColWidth(7) = 1300
.ColWidth(8) = 950
'.ColHidden(2) = True
'.ColHidden(.Cols - 4) = True
'.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ColComboList(1) = cRelStr
.ColComboList(2) = cGenderStr
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 3 Then
    If .Col = 0 Or .Col = 1 Then
        .Col = NextEmpty(grid1, row, Col + 1, 3)
    Else
        .Col = Col + 1
    End If
ElseIf row < .rows - 1 Then
    .Select row + 1, NextEmpty(grid1, row + 1, 0, 3)
    .ShowCell row + 1, 0
End If
End With
End Sub
Private Sub myAddItem()
With grid1
'Dim nMax As Long
'For i = 1 To grid1.Rows - 1
'    nMax = IIf(Val(.TextMatrix(i, 0)) > nMax, Val(.TextMatrix(i, 0)), nMax)
'Next
.AddItem ""
If grid1.rows > 2 Then
    .TextMatrix(.rows - 1, 0) = Val(grid1.TextMatrix(.rows - 2, 0)) + 1
Else
    .TextMatrix(.rows - 1, 0) = "1"
End If
'.TextMatrix(.Rows - 1, 0) = nMax + 1
End With
End Sub
Private Function myreplaceGrd(row) As Boolean
Dim aInsert As Variant
With grid1
    For I = IIf(row = -1, 1, row) To IIf(row = -1, grid1.rows - 2, row)
        aInsert = AddFlag(Empty, "MEMBER", addvalue(xCode.Text))
        aInsert = AddFlag(aInsert, "CODE", addvalue(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "RELATION", addvalue(grid1.TextMatrix(I, 1)))
        aInsert = AddFlag(aInsert, "GENDER", addvalue(grid1.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "TITLE", addstring(grid1.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(I, 4)))
        aInsert = AddFlag(aInsert, "DATE_BIRTH", addDate(grid1.TextMatrix(I, 5)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(grid1.TextMatrix(I, 6)))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_11")
        Else
            con.Execute addUpdate(aInsert, "FILE1_11", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
Dim cString As String
cString = "SELECT FILE1_11.CODE,FILE1_11.RELATION,FILE1_11.GENDER,FILE1_11.TITLE,FILE1_11.DESCA,CONVERT(VARCHAR(10),FILE1_11.DATE_BIRTH,111),FILE1_11.NOTES,CONVERT(VARCHAR(10),FILE1_11.DATE_PRINT,111),FILE1_11.ID " & _
          " FROM FILE1_11"
cString = cString & turn(cString) & "FILE1_11.MEMBER = " & xCode.Text
cString = cString & " ORDER BY CODE"
Set DATA11.Recordset = myRecordSet(cString, con)
myAddItem
Fixgrd
End With
End Sub
Private Sub myRemove(row As Long)
grid1.RemoveItem row
'Calctotals
End Sub
Private Function FoundOtheritem(grid1 As Variant, nrow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.rows - 2
    If I <> nrow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function
Private Sub xUnion_reg_Validate(Cancel As Boolean)
'If Trim(xUnion_reg.Text) <> "" Then
'    Dim sCode_union As Variant
'    sCode_union = GetField("Select code from  file1_10 where  union_reg = " & MyParn(xUnion_reg.Text) & " and code <> " & xCode.Text)
'    If Not IsEmpty(sCode_union) Then
'         MsgBox "—ﬁ„ «·ﬁÌœ „ﬁÌœ „‰ ﬁ»· ··⁄÷Ê —ﬁ„ " & sCode_union, , systemName
'         Cancel = True
'    End If
'End If
End Sub
Private Sub xDate_Last_GotFocus()
myGotFocus xDate_Last
End Sub
Private Sub xDate_Last_LostFocus()
myLostFocus xDate_Last
myValidDate xDate_Last
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xMail_GotFocus()
myGotFocus xMail
End Sub
Private Sub xMail_LostFocus()
myLostFocus xMail
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub
Private Sub xPhone_LostFocus()
myLostFocus xPhone
End Sub
Private Sub xMobil_GotFocus()
myGotFocus xMobil
End Sub
Private Sub xMobil_LostFocus()
myLostFocus xMobil
End Sub
Private Sub xId_no_GotFocus()
myGotFocus xId_no
End Sub
Private Sub xId_no_LostFocus()
myLostFocus xId_no
End Sub
Private Sub xDate_birth_GotFocus()
myGotFocus xDate_birth
End Sub
Private Sub xDate_birth_LostFocus()
myLostFocus xDate_birth
myValidDate xDate_birth
End Sub
Private Sub xTitle_GotFocus()
myGotFocus xTitle
End Sub
Private Sub xTitle_LostFocus()
myLostFocus xTitle
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
Private Sub xSocial_GotFocus()
myGotFocus xSocial
End Sub
Private Sub xSocial_LostFocus()
myLostFocus xSocial
If Not xSocial.MatchedWithList Then xSocial.BoundText = ""
End Sub
Private Sub xDate_Degree_GotFocus()
myGotFocus xDate_Degree
End Sub
Private Sub xDate_Degree_LostFocus()
myLostFocus xDate_Degree
myValidDate xDate_Degree
End Sub
Private Sub xDate_Job_begin_GotFocus()
myGotFocus xDate_Job_begin
End Sub
Private Sub xDate_Job_begin_LostFocus()
myLostFocus xDate_Job_begin
myValidDate xDate_Job_begin
End Sub
Private Sub xDate_job_GotFocus()
myGotFocus xDate_job
End Sub
Private Sub xDate_job_LostFocus()
myLostFocus xDate_job
myValidDate xDate_job
End Sub
Private Sub xDate_Begin_GotFocus()
myGotFocus xDate_Begin
End Sub
Private Sub xDate_Begin_LostFocus()
myLostFocus xDate_Begin
myValidDate xDate_Begin
End Sub
Private Sub xjob_GotFocus()
myGotFocus xJob
End Sub
Private Sub xjob_LostFocus()
myLostFocus xJob
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub
Private Sub xDegree_GotFocus()
myGotFocus xDegree
End Sub
Private Sub xDegree_LostFocus()
myLostFocus xDegree
If Not xDegree.MatchedWithList Then xDegree.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xStatus_GotFocus()
myGotFocus xStatus
End Sub
Private Sub xStatus_LostFocus()
myLostFocus xStatus
If Not xStatus.MatchedWithList Then xStatus.BoundText = ""
End Sub
