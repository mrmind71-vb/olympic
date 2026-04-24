VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form paid_totalfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ì’«·«  ”œ«œ"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   18330
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   18330
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin Olymbic.CSubclass CSubclass1 
      Left            =   360
      Top             =   495
      _ExtentX        =   1376
      _ExtentY        =   1376
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   0
      Width           =   3930
      Begin Threed.SSCommand cmdPay 
         Height          =   510
         Left            =   1980
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   135
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«÷«›… «·”œ«œ"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdUnPay 
         Height          =   510
         Left            =   45
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   135
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ–› «·”œ«œ"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   0
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   765
      Width           =   2175
      Begin Threed.SSCommand cmdAddItems 
         Height          =   1050
         Left            =   45
         TabIndex        =   31
         Top             =   135
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1852
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«÷«›… «·„ÿ«·»… «·Ã„«⁄Ì…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   180
      TabIndex        =   26
      Top             =   2025
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "√Œÿ«¡ «·”œ«œ"
      TabPicture(0)   =   "PAID_TOTAL.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "”œ«œ «·«⁄÷«¡"
      TabPicture(1)   =   "PAID_TOTAL.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   6045
         Left            =   -74955
         TabIndex        =   28
         Top             =   360
         Width           =   17880
         _cx             =   31538
         _cy             =   10663
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
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   6045
         Left            =   90
         TabIndex        =   29
         Top             =   360
         Width           =   17835
         _cx             =   31459
         _cy             =   10663
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
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8460
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":0038
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID_TOTAL.frx":21DF
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":4226
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID_TOTAL.frx":6311
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":830B
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID_TOTAL.frx":A41C
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":C416
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID_TOTAL.frx":E63A
      End
   End
   Begin VB.CheckBox xCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ".«·”‰… «·Õ«·Ì… ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   16110
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8055
      Width           =   2040
   End
   Begin VB.CheckBox xAdded 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   225
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Caption         =   "„” ‰œ „€·ﬁ"
      Enabled         =   0   'False
      Height          =   450
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   -315
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   12015
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   6180
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4995
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":1070B
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID_TOTAL.frx":12AD6
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   3735
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":14B7F
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID_TOTAL.frx":16B87
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2475
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":18B3E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID_TOTAL.frx":1B2DA
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":1D76E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1260
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
         _Version        =   196610
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":1FA91
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID_TOTAL.frx":21E07
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   12795
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9765
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9765
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "1"
         Top             =   180
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xCompany 
         Height          =   330
         Left            =   8640
         TabIndex        =   13
         Top             =   900
         Width           =   2760
         _ExtentX        =   4868
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
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‘—ﬂ…"
         Height          =   285
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label xYear_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê”„"
         Height          =   240
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„” ‰œ"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   510
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   4
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   765
      Width           =   1725
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "PAID_TOTAL.frx":23F8A
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID_TOTAL.frx":268AF
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   675
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
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
         Picture         =   "PAID_TOTAL.frx":29103
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID_TOTAL.frx":2B263
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   990
      Top             =   180
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      CopiesToPrinter =   2
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   1755
      Top             =   630
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   36
      Top             =   9240
      Visible         =   0   'False
      Width           =   18330
      _ExtentX        =   32332
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label xYear_code 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -405
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   -45
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label xBranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -135
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   -45
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "paid_totalfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte
Dim CardTable As ADODB.Recordset, loctable As ADODB.Recordset
Dim oSearchDoc As New Search
Dim bEditRecord As Boolean, bAct As Boolean, aPen As Variant
Dim formMode
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE6_20 FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO where file6_20h.Doc_No_ADD = " & addvalue(xdoc_no.text)
    con.Execute "Delete From FILE6_20H where file6_20h.Doc_No_ADD = " & addvalue(xdoc_no.text)
    con.Execute "Delete From FILE6_91  where file6_91.Doc_No = " & addvalue(xdoc_no.text)
    con.Execute "Delete From FILE6_90H where Doc_No = " & addvalue(xdoc_no.text)
    con.CommitTrans
    openCardTable
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & addvalue(xdoc_no.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Function retAll(aMember As Variant) As Integer
retAll = IIf(retFlag(aMember, "Died", False), 0, 1)
Dim cString As String
cString = "SELECT SUM(1) FROM FILE1_11"
cString = cString & " WHERE FILE1_11.MEMBER = " & xCode.text
retAll = retAll + Val(GetField(cString, con) & "")
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
cString = "SELECT TOP 2000 FILE6_90H.DOC_NO,FILE6_90H.[YEAR_CODE],CONVERT(VARCHAR(10),FILE6_90H.DATE,111),COMPANY_CODES.DESCA" & _
          "  FROM  FILE6_90H INNER JOIN COMPANY_CODES ON FILE6_90H.COMPANY = COMPANY_CODES.CODE"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter

Generalarray(1) = cString
Generalarray(2) = " ORDER BY FILE6_90H.DATE"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·«” „«—…-«·‘—ﬂ…"
listarray(0, 1) = "(**FILE6_60H.DOC_NO** OR %%COMPANY_CODES.DESCA%%)"

listarray(1, 0) = " «—ÌŒ «·„” ‰œ"
listarray(1, 1) = "(##FILE6_20.Date##)"

listarray(2, 0) = "«·”‰…"
listarray(2, 1) = "(**FILE6_90H.YEAR_CODE**)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·”‰…"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(2, 1) = 1400

GrdArray(3, 0) = "«·‘—ﬂ…"
GrdArray(3, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„  „ÿ«·»«  «·‘—ﬂ« "
oSearchDoc.Show 1
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
Private Sub cmdPay_Click()
If AddPay Then
    Inform " „ «÷«›… «·”œ«œ"
    myLoadGrd
    'FixSerial
End If
End Sub
Private Function AddPay(Optional bAdd As Boolean = True) As Boolean
Dim loctable As New ADODB.Recordset, cString As String, aInsert As Variant, nForm_no As Long
cString = "select FILE6_20H.* FROM FILE6_20H"
cString = cString & " WHERE FILE6_20H.DOC_NO_ADD = " & addvalue(xdoc_no.text)
cString = cString & " ORDER BY FILE6_20H.DATE,FILE6_20H.DOC_NO"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
nForm_no = Newflag("FILE6_20H", "FORM_NO", con) = 0
nForm_no = IIf(nForm_no = 0, 1, nForm_no)
Do Until loctable.EOF
    aInsert = AddFlag(Empty, "FORM_NO", IIf(bAdd, nForm_no, "NULL"))
    cInsert = cInsert & addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & loctable!DOC_NO) & ";"
    nForm_no = nForm_no + 1
    loctable.MoveNext
Loop
con.BeginTrans
On Error GoTo myerror
If cInsert <> "" Then con.Execute cInsert
con.CommitTrans
AddPay = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub cmdAdd_Click()
mydefine
'On Error Resume Next
'xCode.SetFocus
Err.Clear
End Sub

Private Sub cmdPrint_Click()
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 3 + 1)
grid1.ColHidden(3 + 1) = True
grid1.ColHidden(4 + 1) = True
grid1.ColHidden(6 + 1) = True

Set PrintGrdNew.myForm = Me
PrintGrdNew.doprint grid1, grdRate(grid1, 15000), 0, "„ÿ«·»«  «⁄÷«¡ ‘—ﬂ… " & xCompany.text & " ·⁄«„ " & xYear_Desca.Caption, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, True, 12, , aRow
grid1.ColHidden(3 + 1) = False
grid1.ColHidden(4 + 1) = False
grid1.ColHidden(6 + 1) = False
PrintGrdNew.Show 1
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant, i As Integer
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[COMPANY]", addvalue(xCompany.BoundText))
aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(Ret_Year(xDate.text, "code", con, Year(xDate.text))))
con.BeginTrans
On Error GoTo myerror
If xdoc_no.Tag = DefineMode Then
    xdoc_no.text = Newflag("FILE6_90H", "DOC_NO")
    aInsert = AddFlag(aInsert, "DOC_NO", addvalue(xdoc_no.text))
    con.Execute addInsert(aInsert, "FILE6_90H")
Else
    con.Execute addUpdate(aInsert, "FILE6_90H", "doc_no = " & addvalue(xdoc_no.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub

Private Sub Command2_Click()
End Sub

Private Sub cmdUnPay_Click()
If AddPay(False) Then
    Inform " „ «·€«¡ «·”œ«œ"
    myLoadGrd
End If
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xdoc_no.Tag = LoadMode Then grid1.SetFocus Else xCompany.SetFocus
    Err.Clear
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
Dim i As Long

CSubclass1.SubClassMe SSTab1.hwnd, 0, , vbWhite      '//--- Begin SubClassing
openCon con

bedit = True

Set data1.Recordset = myRecordSet("select * from company_Codes", con)
Set xCompany.RowSource = data1
xCompany.ListField = "Desca"
xCompany.BoundColumn = "Code"

Set grid1.DataSource = DATA10
Set grid2.DataSource = DATA11


openCardTable
myUndo

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


MYVALID = True
End Function
Private Sub myload()
Dim i As Integer
xdoc_no.text = CardTable!DOC_NO
xDate.text = myFormat_p(CardTable!Date)
xCompany.BoundText = CardTable!Company & ""

'xSeason.Caption = CardTable!SEASON
xdate_LostFocus
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
Handlecontrols LoadMode

myLoadGrd
myloadgrd2
End Sub
Private Sub myLoadGrd()
Dim cString As String
cString = "SELECT FILE6_20H.[CODE],FILE1_10.DESCA,FILE1_10.REGISTER,FILE6_20H.DOC_NO,FILE6_20H.DATE,SUM(TOTAL),FILE6_20H.FORM_NO " & _
          " From [FILE6_20H] INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO"
cString = cString & turn(cString) & "FILE6_20H.DOC_NO_ADD = " & addvalue(xdoc_no.text)
cString = cString & " GROUP BY FILE6_20H.[CODE],FILE1_10.DESCA,FILE6_20H.DOC_NO,FILE6_20H.DATE,FILE1_10.REGISTER,FILE6_20H.FORM_NO"
cString = cString & " ORDER BY FILE6_20H.[CODE]"


Set DATA10.Recordset = myRecordSet(cString, con)
Fixgrd
End Sub
Private Sub myloadgrd2()
Dim cString As String
cString = "SELECT FILE6_91.[CODE],FILE1_10.DESCA,FILE6_91.DESCA " & _
          " From FILE6_91 INNER JOIN FILE1_10 ON FILE6_91.CODE = FILE1_10.CODE"
cString = cString & turn(cString) & "FILE6_91.DOC_NO = " & addvalue(xdoc_no.text)
Set DATA11.Recordset = myRecordSet(cString, con)
fixgrd2
End Sub
Private Sub mydefine()
Dim i As Integer, aRet As Variant
xdoc_no.text = Newflag("FILE6_90H", "DOC_NO")
xDate.text = myFormat_p(Date)
xCompany.BoundText = ""
xYear_code.Caption = ""
Fixgrd
grid1.rows = 1
fixgrd2
grid2.rows = 1
Handlecontrols DefineMode
'Calctotals
On Error Resume Next
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bedit And xClosed.Value = 0
xCompany.Enabled = bEditRecord
cmdAddItems.Enabled = bEditRecord And nMode = LoadMode
cmdAdd.Enabled = nMode = LoadMode
cmdSave.Enabled = bEditRecord
CmdDel.Enabled = nMode = LoadMode And bEditRecord
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And DOC_NO = ""
cmdInform.Enabled = nMode = LoadMode
cmdAddItems.Enabled = nMode = LoadMode And bEditRecord
cmdPay.Enabled = nMode = LoadModeAnd And bEditRecord
cmdUnPay.Enabled = nMode = LoadMode And bEditRecord
xdoc_no.Enabled = (nMode = DefineMode)
xdoc_no.Tag = nMode
End Sub

Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
With grid1
If Col < .Cols - 2 Then
    .Col = Col + 1
ElseIf Row < .rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, grid1.Cols - 1, grid1.Cols - 1)
    .ShowCell .Row, 0
Else
    .Select Row, Col
End If
End With
End Sub
Private Function validRow(Row) As Boolean
If (Not ValidNum(grid1.TextMatrix(Row, grid1.Cols - 1))) And grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then Exit Function
validRow = True
End Function
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then
    myLoadGrd
    Exit Sub
End If
With grid1
    If Not myreplaceGrd(Row) Then
        myLoadGrd
    End If
End With
End Sub
Private Function myreplaceGrd(Row As Long) As Boolean
Dim aInsert As Variant, cWhere As String, i As Long
With grid1
    con.BeginTrans
    On Error GoTo myerror
    For i = IIf(Row = -1, 2, Row) To IIf(Row = -1, grid1.rows - 1, Row)
        aInsert = AddFlag(Empty, "FORM_NO", addvalue(.TextMatrix(i, .Cols - 1)))
        con.Execute addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & grid1.TextMatrix(i, 0))
    Next
    con.CommitTrans
    myreplaceGrd = True
End With
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub grid1_EnterCell()
If grid1.Col = grid1.Cols - 1 And grid1.Row > 1 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub

Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
If Trim(xdoc_no.text) = "" Then
    If xdoc_no.Tag = LoadMode Then mydefine
    Exit Sub
End If
CardTable.Find "Doc_no = " & xdoc_no.text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xdoc_no.Tag = LoadMode Then
    mydefine
End If
End Sub
Private Sub xForm_no_LostFocus()
Dim sDoc As String
If Trim(xForm_No.text) = "" Then Exit Sub
'xDoc_No.Text = GetField("select top 1 doc_no from FILE6_90H where form_no = " & xForm_no.Text & " and season = " & xSeason.Caption)
'If xDoc_No.Text = "" Then xDoc_No.Text = GetField("select top 1 doc_no from FILE6_90H where form_no = " & xForm_no.Text)
'xDoc_No_LostFocus
End Sub
Private Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From file2_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchItems.Caption = "≈” ⁄·«„ «·»‰Ê\"
oSearchItems.Show 1
End Sub
Private Function CalcTotals(Optional bOverRide As Boolean = False)
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
'If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
Dim i As Long
With grid1
.FormatString = "„|" & "ﬂÊœ  «·⁄÷Ê|" & "≈”„ «·⁄÷Ê|" & "—ﬁ„ «·ﬁÌœ|" & "—ﬁ„ «·„” ‰œ|" & " «—ÌŒ «·„” ‰œ|" & "≈Ã„«·Ï «·„»·€|" & "—ﬁ„ «·«” „«—…"
.ColWidth(0) = 700
.ColWidth(0 + 1) = 1500
.ColWidth(1 + 1) = 6000
.ColWidth(2 + 1) = 1500
.ColWidth(2 + 1 + 1) = 1500
.ColWidth(3 + 1 + 1) = 1800
.ColWidth(4 + 1 + 1) = 1800
.ColWidth(5 + 1 + 1) = 1500
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
If .rows > 1 Then
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 5 + 1, "#0.00", &HC0FFC0, vbBlack, True, "  "
    For i = 0 To 3
        .TextMatrix(1, i) = "«·≈Ã„«·Ï"
    Next
    .MergeCells = flexMergeFree
    .MergeRow(1) = True
End If
FixSerial
End With
End Sub
Private Sub fixgrd2()
With grid2
.FormatString = "ﬂÊœ  «·⁄÷Ê|" & "≈”„ «·⁄÷Ê|" & "«·”»»"
.ColWidth(0) = 1500
.ColWidth(1) = 6000
.ColWidth(2) = 5000
For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT FILE6_90H.* FROM FILE6_90H"
cFilter = ""
If xCurrent.Value = 1 Then
    Dim aRet As Variant
    aRet = Year_Load(sSeason, , con)
    cFilter = "FILE6_90H.DATE >= " & DateSq(retFlag(aRet, "DATE1"))
    cFilter = cFilter & " AND FILE6_90H.DATE >= " & DateSq(retFlag(aRet, "DATE2"))
End If
If sDoc_no <> "" Then cFilter = " DOC_NO = " & MyParn(sDoc_no)
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " Order by FILE6_90H.DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
'On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xdoc_no.text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
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
Private Sub MyAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
If IsDate(xDate.text) Then
    xYear_Desca.Caption = Ret_Year(xDate.text, "Desca_r", con, Year(xDate.text))
End If
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub myRemove(Index As Integer, Row As Long)
CalcTotals
End Sub
Private Function doprint()
'If Not MYVALID Then Exit Function
'Dim loctable As New ADODB.Recordset, cString As String
'Dim temptable As New ADODB.Recordset
'cString = "SELECT FILE6_20.DOC_NO,FILE6_90H.DATE,CASE WHEN USERS.DESCA IS NULL THEN FILE6_90H.USERNAME ELSE USERS.DESCA END AS USER_NAME,FILE6_90H.DATE2,FILE6_90H.CODE AS CODE_MEMBER,FILE1_10.DESCA AS DESCA_MEMBER, FILE6_20.CODE,FILE1_20.DESCA AS ITEM_DESCA,FILE6_20.DESCA," & _
'          "FILE6_20.VALUE,FORM_NO,FILE6_20.[NOTES]" & _
'          " FROM FILE6_20 INNER JOIN FILE6_90H ON FILE6_20.DOC_NO = FILE6_90H.DOC_NO " & _
'          " INNER JOIN FILE1_10 ON FILE6_90H.CODE = FILE1_10.CODE" & _
'          " INNER JOIN FILE1_20 ON FILE6_20.CODE = FILE1_20.CODE" & _
'          " LEFT JOIN USERS ON FILE6_90H.USERCODE = USERS.CODE"
'cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & xDoc_No.Text
'
'Dim aTotal As Variant
'aTotal = GetFields("Select sum(file6_20.total) as total from file6_20 where doc_no = " & xDoc_No.Text)
'loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
'contemp.Execute "DELETE * FROM TEMP"
'temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
'
'Dim I As Long
'With loctable
'Do Until loctable.EOF
'    temptable.AddNew
'    I = I + 1
'    temptable!str1 = ArbString(Val(loctable!FORM_NO & ""))
'    temptable!str2 = ArbString(Format(loctable!Date, "YYYY-MM-DD"))
'    temptable!Str3 = TurnValue(ArbString(loctable!CODE_MEMBER))
'    temptable!str4 = TurnValue(ArbString(loctable!Desca_Member))
'    temptable!str5 = TurnValue(ArbString(Format(loctable!date2, "YYYY-MM-DD")))
'    temptable!STR6 = Format(Now, "HH:NN")
'    temptable!str11 = TurnValue(loctable!Item_Desca)
'    temptable!str12 = TurnValue(loctable!Desca)
'    temptable!str13 = TurnValue(loctable!notes)
'    temptable!str13 = TurnValue(loctable!notes)
'    temptable!str14 = TurnValue(loctable!user_name)
'    temptable!str21 = "≈Ì’«· ”œ«œ Ê«” ·«„"
'    temptable!val1 = Val(loctable!Value & "")
'    temptable!Str10 = MyOnly(Val(retFlag(aTotal, "total") & ""))
'    temptable!Val10 = I
'    temptable.Update
'    loctable.MoveNext
'Loop
'End With
'contemp.BeginTrans
'contemp.CommitTrans
'
'REPORT1.Reset
'REPORT1.WindowState = crptMaximized
'REPORT1.ReportFileName = App.Path & "\Reports\paid.rpt"
'REPORT1.DataFiles(0) = tempFile
'REPORT1.ProgressDialog = False
'REPORT1.CopiesToPrinter = 1
''REPORT1.Destination = crptToPrinter
'REPORT1.Action = 1
'temptable.Close
'Set temptable = Nothing
End Function
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    xdoc_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    oSearchDoc.Hide
    myUndo
End If
End Sub
Private Sub FixSerial()
Dim i As Long
For i = 2 To grid1.rows - 1
    grid1.TextMatrix(i, 0) = i - 1
Next
End Sub
